#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import signal
import socket
import re
import sys
import time
import urllib.error
import urllib.parse
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import server  # noqa: E402
from tools.update_mops_seed import (  # noqa: E402
    fetch_official_pdf_notice_info,
    has_pickup_details,
    load_official_pdf_sources,
    load_seed,
    save_seed,
)

TWSE_COMPANY_API = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
TPEX_COMPANY_API = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"
OFFICIAL_SOURCES_PATH = ROOT / "data" / "official_notice_sources.json"
OFFICIAL_SCAN_CACHE_PATH = ROOT / "data" / "official_site_scan_cache.json"
COMPANY_SITES_CACHE_PATH = ROOT / "data" / "company_sites_cache.json"
PAGE_TIMEOUT_SECONDS = 10
COMPANY_SCAN_TIMEOUT_SECONDS = 25
COMPANY_SITES_CACHE_TTL_SECONDS = 24 * 3600
FETCH_ENGINE_AUTO = "auto"
FETCH_ENGINE_URLLIB = "urllib"
FETCH_ENGINE_SCRAPLING = "scrapling"

PAGE_KEYWORDS = (
    "投資人",
    "股東",
    "股東會",
    "公司治理",
    "下載",
    "檔案",
    "公告",
    "專區",
    "relations",
    "download",
    "files",
    "reports",
    "財務",
    "ir",
    "investor",
    "shareholder",
    "governance",
)
PDF_KEYWORDS = (
    "開會通知",
    "通知書",
    "股東會",
    "紀念品",
    "meeting",
    "notice",
    "shareholder",
)
CURRENT_ROC_YEAR = str(server.CURRENT_YEAR - 1911)
CURRENT_AD_YEAR = str(server.CURRENT_YEAR)
NEGATIVE_PDF_KEYWORDS = (
    "年報",
    "annual",
    "議事手冊",
    "議案參考資料",
    "財報",
    "financial",
)
PRIORITY_PAGE_KEYWORDS = (
    "投資人專區",
    "股東專區",
    "股東會",
    "公司治理",
    "檔案下載",
    "下載專區",
    "投資人關係",
    "investor relations",
    "shareholder meeting",
    "corporate governance",
    "download",
    "ir",
)
PDF_URL_RE = re.compile(r"""(?P<url>https?://[^\s"'<>]+?\.pdf(?:\?[^\s"'<>]*)?)""", re.I)

try:
    from scrapling.fetchers import Fetcher as ScraplingFetcher
except Exception:  # pragma: no cover - optional dependency
    ScraplingFetcher = None


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.links: list[dict[str, str]] = []
        self._current_href = ""
        self._current_text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        attrs_dict = {key.lower(): value or "" for key, value in attrs}
        self._current_href = attrs_dict.get("href", "")
        self._current_text = []

    def handle_data(self, data: str) -> None:
        if self._current_href:
            self._current_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "a" and self._current_href:
            self.links.append(
                {
                    "href": self._current_href,
                    "text": server.normalize_text(" ".join(self._current_text)),
                }
            )
            self._current_href = ""
            self._current_text = []


class CompanyScanTimeoutError(TimeoutError):
    pass


def _company_scan_timeout_handler(signum: int, frame: object) -> None:
    raise CompanyScanTimeoutError("company site scan timed out")


def fetch_json(url: str) -> list[dict[str, Any]]:
    return json.loads(server.fetch_bytes(url).decode("utf-8", "ignore"))


def classify_error(error: Exception | str) -> str:
    message = str(error or "").strip()
    lowered = message.lower()
    if isinstance(error, socket.gaierror) or "nodename nor servname provided" in lowered or "name or service not known" in lowered:
        return "dns"
    if isinstance(error, TimeoutError) or "timed out" in lowered or "timeout" in lowered:
        return "timeout"
    if isinstance(error, urllib.error.HTTPError):
        return f"http_{error.code}"
    if "http error 403" in lowered or "forbidden" in lowered:
        return "http_403"
    if "http error 404" in lowered or "not found" in lowered:
        return "http_404"
    if "ssl" in lowered or "certificate" in lowered or "tls" in lowered:
        return "ssl"
    if "connection refused" in lowered:
        return "connection_refused"
    if "connection reset" in lowered:
        return "connection_reset"
    if "scrapling is not installed" in lowered:
        return "scrapling_missing"
    return "fetch_error"


def normalize_company_url(value: str) -> str:
    value = (value or "").strip().strip("　")
    value = value.replace("：", ":").replace("／", "/")
    if not value:
        return ""
    if not re.match(r"https?://", value, re.I):
        value = f"https://{value}"
    try:
        parsed = urllib.parse.urlparse(value)
    except ValueError:
        return ""
    if not parsed.netloc:
        return ""
    return urllib.parse.urlunparse((parsed.scheme, parsed.netloc, parsed.path or "/", "", "", ""))


def same_site(base_url: str, target_url: str) -> bool:
    base_host = urllib.parse.urlparse(base_url).netloc.lower().removeprefix("www.")
    target_host = urllib.parse.urlparse(target_url).netloc.lower().removeprefix("www.")
    return bool(base_host and target_host and (target_host == base_host or target_host.endswith(f".{base_host}")))


def absolute_url(base_url: str, href: str) -> str:
    href = (href or "").strip()
    if not href or href.startswith(("mailto:", "tel:", "javascript:")):
        return ""
    return quote_url(urllib.parse.urljoin(base_url, href))


def quote_url(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    return urllib.parse.urlunsplit(
        (
            parsed.scheme,
            parsed.netloc.encode("idna").decode("ascii") if parsed.netloc else "",
            urllib.parse.quote(urllib.parse.unquote(parsed.path), safe="/%:@"),
            urllib.parse.quote_plus(urllib.parse.unquote_plus(parsed.query), safe="=&%:@/"),
            urllib.parse.quote(urllib.parse.unquote(parsed.fragment), safe="%:@/"),
        )
    )


def compact_for_match(value: str) -> str:
    return urllib.parse.unquote(value or "").lower()


def looks_like_page_candidate(link: dict[str, str], url: str, base_url: str) -> bool:
    if not same_site(base_url, url):
        return False
    if re.search(r"\.(?:pdf|doc|docx|xls|xlsx|zip|jpg|png)(?:$|\?)", url, re.I):
        return False
    haystack = compact_for_match(f"{link.get('text', '')} {url}")
    return any(keyword.lower() in haystack for keyword in PAGE_KEYWORDS)


def score_page_candidate(link: dict[str, str], url: str, base_url: str) -> int:
    if not same_site(base_url, url):
        return 0
    if re.search(r"\.(?:pdf|doc|docx|xls|xlsx|zip|jpg|png)(?:$|\?)", url, re.I):
        return 0
    haystack = compact_for_match(f"{link.get('text', '')} {url}")
    score = 0
    for keyword in PAGE_KEYWORDS:
        if keyword.lower() in haystack:
            score += 2
    for keyword in PRIORITY_PAGE_KEYWORDS:
        if keyword.lower() in haystack:
            score += 4
    if re.search(r"investor|shareholder|governance|download|ir|pdf|meeting|notice|股東|治理|下載|公告", haystack, re.I):
        score += 2
    return score


def score_pdf_candidate(link: dict[str, str], url: str, code: str) -> int:
    haystack = compact_for_match(f"{link.get('text', '')} {url}")
    if CURRENT_ROC_YEAR not in haystack and CURRENT_AD_YEAR not in haystack:
        return 0
    if not re.search(r"開會通知|通知書|股東會|meeting|notice|shareholder", haystack, re.I):
        return 0
    score = 0
    for keyword in PDF_KEYWORDS:
        if keyword.lower() in haystack:
            score += 3
    if code in haystack:
        score += 2
    if re.search(r"(?:^|[^0-9])115(?:[^0-9]|$)|2026", haystack):
        score += 4
    if re.search(r"notice|stockholder|shareholder|meeting|股東會|開會通知", haystack):
        score += 5
    for keyword in NEGATIVE_PDF_KEYWORDS:
        if keyword.lower() in haystack:
            score -= 5
    return score


def extract_pdf_candidates_from_html(html_text: str, page_url: str, code: str, company_name: str) -> list[dict[str, str]]:
    candidates: dict[str, dict[str, str | int]] = {}
    for match in PDF_URL_RE.finditer(html_text):
        raw_url = match.group("url")
        url = absolute_url(page_url, raw_url) or quote_url(raw_url)
        if not url:
            continue
        score = score_pdf_candidate({"text": ""}, url, code)
        if score < 5:
            continue
        candidates[url] = {
            "url": url,
            "label": f"{company_name or code}官網開會通知書",
            "sourceType": "company_pdf",
            "score": score,
        }
    ranked = sorted(candidates.values(), key=lambda item: int(item["score"]), reverse=True)
    return [{key: str(value) for key, value in item.items() if key != "score"} for item in ranked]


def fetch_page(url: str, fetch_engine: str) -> str:
    if fetch_engine in (FETCH_ENGINE_AUTO, FETCH_ENGINE_SCRAPLING) and ScraplingFetcher is not None:
        try:
            return fetch_page_with_scrapling(url)
        except Exception:
            if fetch_engine == FETCH_ENGINE_SCRAPLING:
                raise
    return server.fetch_text(url, timeout=PAGE_TIMEOUT_SECONDS)


def parse_links(html_text: str, page_url: str) -> list[dict[str, str]]:
    parser = LinkParser()
    parser.feed(html_text)
    links: list[dict[str, str]] = []
    for link in parser.links:
        url = absolute_url(page_url, link["href"])
        if not url:
            continue
        links.append({**link, "url": url})
    return links


def load_company_sites_cache(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def save_company_sites_cache(path: Path, companies: dict[str, dict[str, str]]) -> None:
    path.parent.mkdir(exist_ok=True)
    payload = {
        "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "generatedAtEpoch": time.time(),
        "companies": {code: companies[code] for code in sorted(companies)},
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def load_company_sites(use_cache: bool = True, cache_ttl_seconds: int = COMPANY_SITES_CACHE_TTL_SECONDS) -> dict[str, dict[str, str]]:
    stale_companies: dict[str, dict[str, str]] = {}
    if use_cache:
        cache_payload = load_company_sites_cache(COMPANY_SITES_CACHE_PATH)
        cached_companies = cache_payload.get("companies")
        generated_at = float(cache_payload.get("generatedAtEpoch") or 0)
        if isinstance(cached_companies, dict):
            stale_companies = {str(code): item for code, item in cached_companies.items() if isinstance(item, dict)}
        if stale_companies and generated_at and time.time() - generated_at < cache_ttl_seconds:
            return stale_companies

    companies: dict[str, dict[str, str]] = {}
    try:
        for row in fetch_json(TWSE_COMPANY_API):
            code = str(row.get("公司代號", "")).strip()
            url = normalize_company_url(str(row.get("網址", "")))
            if code and url:
                companies[code] = {"code": code, "name": str(row.get("公司簡稱") or row.get("公司名稱") or ""), "url": url, "market": "TWSE"}
        for row in fetch_json(TPEX_COMPANY_API):
            code = str(row.get("SecuritiesCompanyCode", "")).strip()
            url = normalize_company_url(str(row.get("WebAddress", "")))
            if code and url:
                companies[code] = {
                    "code": code,
                    "name": str(row.get("CompanyAbbreviation") or row.get("CompanyName") or ""),
                    "url": url,
                    "market": "TPEx",
                }
    except Exception:
        if stale_companies:
            return stale_companies
        raise
    save_company_sites_cache(COMPANY_SITES_CACHE_PATH, companies)
    return companies


def read_codes_file(path: Path) -> list[str]:
    if not path.exists():
        return []
    return server.clean_codes(path.read_text(encoding="utf-8"))


def unique_preserve(values: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


def rotate_codes(codes: list[str], offset: int) -> list[str]:
    if not codes:
        return []
    real_offset = offset % len(codes)
    return [*codes[real_offset:], *codes[:real_offset]]


def fetch_page_with_scrapling(url: str) -> str:
    if ScraplingFetcher is None:
        raise RuntimeError("Scrapling is not installed")

    attempts = (
        {"timeout": PAGE_TIMEOUT_SECONDS, "follow_redirects": True, "stealthy_headers": True},
        {"timeout": PAGE_TIMEOUT_SECONDS, "follow_redirects": True},
    )
    last_error: Exception | None = None
    for kwargs in attempts:
        try:
            page = ScraplingFetcher.get(url, **kwargs)
            text = getattr(page, "text", "")
            if callable(text):
                text = text()
            if isinstance(text, str) and text.strip():
                return text
            html = getattr(page, "html_content", "") or getattr(page, "html", "")
            if callable(html):
                html = html()
            if isinstance(html, str) and html.strip():
                return html
            rendered = str(page)
            if rendered.strip():
                return rendered
        except TypeError as error:
            last_error = error
            continue
        except Exception as error:
            last_error = error
            continue
    raise RuntimeError(str(last_error or "Scrapling could not fetch page"))


def discover_company_pdfs(code: str, company: dict[str, str], max_pages: int, fetch_engine: str) -> list[dict[str, str]]:
    base_url = company["url"]
    queue: list[tuple[int, str]] = [(100, base_url)]
    visited: set[str] = set()
    candidates: dict[str, dict[str, str | int]] = {}

    while queue and len(visited) < max_pages:
        queue.sort(key=lambda item: item[0], reverse=True)
        _, page_url = queue.pop(0)
        if page_url in visited:
            continue
        visited.add(page_url)
        try:
            html_text = fetch_page(page_url, fetch_engine)
            links = parse_links(html_text, page_url)
        except Exception:
            continue

        for candidate in extract_pdf_candidates_from_html(html_text, page_url, code, company.get("name", "")):
            existing = candidates.get(candidate["url"])
            if existing:
                continue
            candidates[candidate["url"]] = {**candidate, "score": int(score_pdf_candidate({"text": ""}, candidate["url"], code))}

        for link in links:
            url = link["url"]
            if re.search(r"\.pdf(?:$|\?)", urllib.parse.urlparse(url).path, re.I) or ".pdf" in url.lower():
                score = score_pdf_candidate(link, url, code)
                if score >= 5:
                    candidates[url] = {
                        "url": url,
                        "label": f"{company.get('name') or code}官網開會通知書",
                        "sourceType": "company_pdf",
                        "score": score,
                    }
                continue

            page_score = score_page_candidate(link, url, base_url)
            if len(queue) + len(visited) < max_pages and page_score > 0:
                queue.append((page_score, url))

    ranked = sorted(candidates.values(), key=lambda item: int(item["score"]), reverse=True)
    return [{key: str(value) for key, value in item.items() if key != "score"} for item in ranked]


def load_sources_payload(path: Path) -> dict[str, list[dict[str, str]]]:
    if not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        return {}
    normalized: dict[str, list[dict[str, str]]] = {}
    for raw_code, raw_items in payload.items():
        codes = server.clean_codes(str(raw_code))
        if not codes:
            continue
        items = raw_items if isinstance(raw_items, list) else [raw_items]
        normalized[codes[0]] = [item for item in items if isinstance(item, dict) and item.get("url")]
    return normalized


def save_sources_payload(path: Path, payload: dict[str, list[dict[str, str]]]) -> None:
    path.parent.mkdir(exist_ok=True)
    ordered = {code: payload[code] for code in sorted(payload)}
    path.write_text(json.dumps(ordered, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def load_scan_cache(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def save_scan_cache(path: Path, payload: dict[str, dict[str, Any]]) -> None:
    path.parent.mkdir(exist_ok=True)
    ordered = {code: payload[code] for code in sorted(payload)}
    path.write_text(json.dumps(ordered, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def merge_source(payload: dict[str, list[dict[str, str]]], code: str, source: dict[str, str]) -> bool:
    existing = payload.setdefault(code, [])
    source_key = canonical_url_key(source["url"])
    if any(canonical_url_key(item.get("url", "")) == source_key for item in existing):
        return False
    existing.append(source)
    return True


def canonical_url_key(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    return urllib.parse.urlunsplit(
        (
            "",
            parsed.netloc.lower().removeprefix("www."),
            urllib.parse.unquote(parsed.path),
            urllib.parse.unquote_plus(parsed.query),
            "",
        )
    )


def find_existing_source(payload: dict[str, list[dict[str, str]]], code: str, url: str) -> dict[str, str] | None:
    source_key = canonical_url_key(url)
    for item in payload.get(code, []):
        if canonical_url_key(item.get("url", "")) == source_key:
            return item
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Discover official company notice PDFs before falling back to MOPS.")
    parser.add_argument("codes", nargs="*", help="Stock codes or mixed text.")
    parser.add_argument("--watchlist-file", default="", help="Text file with stock codes.")
    parser.add_argument("--sources-file", default=str(OFFICIAL_SOURCES_PATH), help="Official PDF source JSON to update.")
    parser.add_argument("--scan-cache-file", default=str(OFFICIAL_SCAN_CACHE_PATH), help="JSON file recording official website scan attempts.")
    parser.add_argument("--limit", type=int, default=0, help="Maximum company websites to scan. 0 means all selected codes.")
    parser.add_argument("--max-pages-per-company", type=int, default=10, help="Maximum pages to crawl per company site.")
    parser.add_argument("--company-timeout-seconds", type=int, default=COMPANY_SCAN_TIMEOUT_SECONDS, help="Maximum seconds to spend on a single company website before skipping it.")
    parser.add_argument("--sleep", type=float, default=0.7, help="Seconds to sleep between company sites.")
    parser.add_argument("--skip-existing-source", action="store_true", help="Skip codes already listed in official sources.")
    parser.add_argument("--skip-cached-pickup", action="store_true", help="Skip codes already cached with pickup details.")
    parser.add_argument("--skip-recent-attempt-hours", type=float, default=0, help="Skip company sites already scanned within this many hours.")
    parser.add_argument("--update-seed", action="store_true", help="Parse useful PDFs and update the deployable seed cache.")
    parser.add_argument("--refresh-company-sites", action="store_true", help="Refresh company site list from TWSE/TPEx instead of using local cache.")
    parser.add_argument(
        "--fetch-engine",
        choices=(FETCH_ENGINE_AUTO, FETCH_ENGINE_URLLIB, FETCH_ENGINE_SCRAPLING),
        default=FETCH_ENGINE_AUTO,
        help="Fetcher to use for company pages. auto prefers Scrapling when installed, else urllib.",
    )
    parser.add_argument("--rotate-offset", type=int, default=0, help="Rotate selected company codes before applying limit.")
    args = parser.parse_args()

    codes = server.clean_codes(" ".join(args.codes))
    if args.watchlist_file:
        codes.extend(read_codes_file(Path(args.watchlist_file)))
    codes = unique_preserve(codes)
    if not codes:
        print("No stock codes found.")
        return 0

    company_sites = load_company_sites(use_cache=not args.refresh_company_sites)
    sources_path = Path(args.sources_file)
    scan_cache_path = Path(args.scan_cache_file)
    sources_payload = load_sources_payload(sources_path)
    scan_cache = load_scan_cache(scan_cache_path)
    loaded_sources = load_official_pdf_sources(sources_path)
    seed = load_seed()
    now_epoch = time.time()

    selected: list[str] = []
    for code in codes:
        if code not in company_sites:
            continue
        if args.skip_existing_source and code in loaded_sources:
            continue
        if args.skip_cached_pickup and has_pickup_details(seed.get(code)):
            continue
        last_attempt = float((scan_cache.get(code) or {}).get("attemptedAtEpoch") or 0)
        if args.skip_recent_attempt_hours > 0 and last_attempt and now_epoch - last_attempt < args.skip_recent_attempt_hours * 3600:
            continue
        selected.append(code)

    if args.rotate_offset:
        selected = rotate_codes(selected, args.rotate_offset)
    if args.limit > 0:
        selected = selected[: args.limit]

    selected_engine = args.fetch_engine
    if selected_engine == FETCH_ENGINE_AUTO:
        selected_engine = FETCH_ENGINE_SCRAPLING if ScraplingFetcher is not None else FETCH_ENGINE_URLLIB
    print(f"Using fetch engine: {selected_engine}", flush=True)

    added = 0
    parsed = 0
    for index, code in enumerate(selected, 1):
        company = company_sites[code]
        print(f"[{index}/{len(selected)}] {code} {company.get('name', '')} {company['url']}", flush=True)
        previous_handler = signal.getsignal(signal.SIGALRM)
        signal.signal(signal.SIGALRM, _company_scan_timeout_handler)
        signal.alarm(max(1, int(args.company_timeout_seconds)))
        try:
            candidates = discover_company_pdfs(code, company, args.max_pages_per_company, selected_engine)
        except urllib.error.URLError as error:
            print(f"  skipped website: {error}", flush=True)
            scan_cache[code] = {
                "code": code,
                "companyName": company.get("name", ""),
                "companyUrl": company["url"],
                "attemptedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "attemptedAtEpoch": time.time(),
                "candidateCount": 0,
                "foundUseful": False,
                "error": str(error),
                "errorType": classify_error(error),
                "fetchEngine": selected_engine,
            }
            continue
        except CompanyScanTimeoutError as error:
            print(f"  skipped website: {error}", flush=True)
            scan_cache[code] = {
                "code": code,
                "companyName": company.get("name", ""),
                "companyUrl": company["url"],
                "attemptedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "attemptedAtEpoch": time.time(),
                "candidateCount": 0,
                "foundUseful": False,
                "error": str(error),
                "errorType": "company_timeout",
                "fetchEngine": selected_engine,
            }
            continue
        except RuntimeError as error:
            print(f"  skipped website: {error}", flush=True)
            scan_cache[code] = {
                "code": code,
                "companyName": company.get("name", ""),
                "companyUrl": company["url"],
                "attemptedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "attemptedAtEpoch": time.time(),
                "candidateCount": 0,
                "foundUseful": False,
                "error": str(error),
                "errorType": classify_error(error),
                "fetchEngine": selected_engine,
            }
            continue
        finally:
            signal.alarm(0)
            signal.signal(signal.SIGALRM, previous_handler)

        if not candidates:
            print("  no likely official notice PDF found", flush=True)
        found_useful = False
        for source in candidates[:3]:
            source_for_parse = find_existing_source(sources_payload, code, source["url"]) or source
            info, error = fetch_official_pdf_notice_info(code, source_for_parse)
            if not info:
                print(f"  candidate skipped: {source['url']} ({error})", flush=True)
                continue
            if merge_source(sources_payload, code, source):
                added += 1
            if args.update_seed:
                seed[code] = info
                parsed += 1
            found_useful = True
            start = info.get("evotePickupStartDate") or "-"
            end = info.get("evotePickupEndDate") or "-"
            print(f"  official notice cached: pickup={start}~{end} {source['url']}", flush=True)
            break

        scan_cache[code] = {
            "code": code,
            "companyName": company.get("name", ""),
            "companyUrl": company["url"],
            "attemptedAt": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
            "attemptedAtEpoch": time.time(),
            "candidateCount": len(candidates),
            "foundUseful": found_useful,
            "error": "" if found_useful or candidates else "no likely official notice PDF found",
            "errorType": "" if found_useful else ("no_candidate" if not candidates else "no_useful_candidate"),
            "fetchEngine": selected_engine,
        }

        if args.sleep > 0 and index < len(selected):
            time.sleep(args.sleep)

    save_sources_payload(sources_path, sources_payload)
    save_scan_cache(scan_cache_path, scan_cache)
    if args.update_seed:
        save_seed(seed)
    print(f"Official sources added: {added}", flush=True)
    if args.update_seed:
        print(f"Seed records parsed from official PDFs: {parsed}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

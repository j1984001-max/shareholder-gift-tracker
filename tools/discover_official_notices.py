#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
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
PAGE_TIMEOUT_SECONDS = 10

PAGE_KEYWORDS = (
    "投資人",
    "股東",
    "股東會",
    "公司治理",
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


def fetch_json(url: str) -> list[dict[str, Any]]:
    return json.loads(server.fetch_bytes(url).decode("utf-8", "ignore"))


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


def fetch_page(url: str) -> str:
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


def load_company_sites() -> dict[str, dict[str, str]]:
    companies: dict[str, dict[str, str]] = {}
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


def discover_company_pdfs(code: str, company: dict[str, str], max_pages: int) -> list[dict[str, str]]:
    base_url = company["url"]
    queue = [base_url]
    visited: set[str] = set()
    candidates: dict[str, dict[str, str | int]] = {}

    while queue and len(visited) < max_pages:
        page_url = queue.pop(0)
        if page_url in visited:
            continue
        visited.add(page_url)
        try:
            links = parse_links(fetch_page(page_url), page_url)
        except Exception:
            continue

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

            if len(queue) + len(visited) < max_pages and looks_like_page_candidate(link, url, base_url):
                queue.append(url)

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
    parser.add_argument("--sleep", type=float, default=0.7, help="Seconds to sleep between company sites.")
    parser.add_argument("--skip-existing-source", action="store_true", help="Skip codes already listed in official sources.")
    parser.add_argument("--skip-cached-pickup", action="store_true", help="Skip codes already cached with pickup details.")
    parser.add_argument("--skip-recent-attempt-hours", type=float, default=0, help="Skip company sites already scanned within this many hours.")
    parser.add_argument("--update-seed", action="store_true", help="Parse useful PDFs and update the deployable seed cache.")
    args = parser.parse_args()

    codes = server.clean_codes(" ".join(args.codes))
    if args.watchlist_file:
        codes.extend(read_codes_file(Path(args.watchlist_file)))
    codes = unique_preserve(codes)
    if not codes:
        print("No stock codes found.")
        return 0

    company_sites = load_company_sites()
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

    if args.limit > 0:
        selected = selected[: args.limit]

    added = 0
    parsed = 0
    for index, code in enumerate(selected, 1):
        company = company_sites[code]
        print(f"[{index}/{len(selected)}] {code} {company.get('name', '')} {company['url']}", flush=True)
        try:
            candidates = discover_company_pdfs(code, company, args.max_pages_per_company)
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
            }
            continue

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
            "error": "",
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

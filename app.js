const WATCHLIST_KEY = "shareholder-gift-tracker-watchlist-v1";
const SNAPSHOT_KEY = "shareholder-gift-tracker-snapshot-v1";
const AUTO_REFRESH_MS = 15 * 60 * 1000;
const CURRENT_ROC_YEAR = new Date().getFullYear() - 1911;
const COMPARE_YEARS = [CURRENT_ROC_YEAR, CURRENT_ROC_YEAR - 1, CURRENT_ROC_YEAR - 2];

const sampleCodes = ["2317", "2409", "3037", "9938", "2006", "1416"];

const codesInput = document.getElementById("codesInput");
const lookupBtn = document.getElementById("lookupBtn");
const saveWatchlistBtn = document.getElementById("saveWatchlistBtn");
const currentModeBtn = document.getElementById("currentModeBtn");
const compareModeBtn = document.getElementById("compareModeBtn");
const loadSampleBtn = document.getElementById("loadSampleBtn");
const clearBtn = document.getElementById("clearBtn");
const refreshBtn = document.getElementById("refreshBtn");
const exportBtn = document.getElementById("exportBtn");
const updatedAtText = document.getElementById("updatedAtText");
const noticeProgressText = document.getElementById("noticeProgressText");
const statusText = document.getElementById("statusText");
const resultsBody = document.getElementById("resultsBody");
const summaryStrip = document.getElementById("summaryStrip");
const summaryCardTemplate = document.getElementById("summaryCardTemplate");
const codeCounter = document.getElementById("codeCounter");
const resultSearchInput = document.getElementById("resultSearchInput");
const filterTabs = [...document.querySelectorAll(".filter-tab")];

let activeCodes = [];
let lastResponse = null;
let activeFilter = "all";
let viewMode = "current";
exportBtn.disabled = true;

function extractCodes(raw) {
  const normalized = raw
    .replace(/[（［【「『]/g, "(")
    .replace(/[）］】」』]/g, ")")
    .replace(/[\u3000,，、；;]/g, " ");

  const bracketMatches = [...normalized.matchAll(/\((\d{3,6})\)/g)].map((match) => match[1]);
  const plainMatches = [...normalized.matchAll(/\d{3,6}/g)].map((match) => match[0]);
  return [...new Set([...bracketMatches, ...plainMatches])];
}

function compareYearsParam() {
  return COMPARE_YEARS.join(",");
}

function setViewMode(mode) {
  viewMode = mode === "compare" ? "compare" : "current";
  currentModeBtn?.classList.toggle("active", viewMode === "current");
  compareModeBtn?.classList.toggle("active", viewMode === "compare");
  lookupBtn.textContent = viewMode === "compare" ? "查詢三年度比較" : "查詢最新資料";
  exportBtn.textContent = viewMode === "compare" ? "下載今年 Excel" : "下載 Excel";
  if (viewMode === "compare") {
    statusText.textContent = "三年度比較會讀取本機已建好的歷史快照，不會在線上臨時爬 MOPS。";
  }
}

function saveWatchlist(codes) {
  localStorage.setItem(WATCHLIST_KEY, codes.join("\n"));
}

function loadWatchlist() {
  return localStorage.getItem(WATCHLIST_KEY) || "";
}

function loadSnapshots() {
  try {
    return JSON.parse(localStorage.getItem(SNAPSHOT_KEY) || "{}");
  } catch {
    return {};
  }
}

function saveSnapshots(snapshot) {
  localStorage.setItem(SNAPSHOT_KEY, JSON.stringify(snapshot));
}

async function refreshNoticeProgress() {
  if (!noticeProgressText) return;
  try {
    const response = await fetch("/api/notice-progress");
    const progress = await response.json();
    if (!progress.ok) throw new Error(progress.error || "progress unavailable");
    noticeProgressText.textContent = `通知書 ${progress.noticeCached}/${progress.watchlistTotal}，電投日期 ${progress.pickupDateCached}，MOPS已嘗試/快取 ${progress.mopsAttemptedOrCached || 0}`;
    noticeProgressText.title = `MOPS紀錄嘗試：${progress.mopsAttemptLogged || 0}；MOPS尚未嘗試：${progress.mopsNeverAttempted || 0}；MOPS限流：${progress.mopsRateLimited || 0}；官網掃描：${progress.officialSiteScanned || 0}/${progress.watchlistTotal}；官網/官方 PDF：${progress.officialPdfCached}；官網找到：${progress.officialSiteFound || 0}；尚缺通知書：${progress.missingNotice}；尚缺電投日期：${progress.missingPickupDate}`;
  } catch {
    noticeProgressText.textContent = "暫時無法讀取";
  }
}

function formatDate(value) {
  if (!value) return "未提供";
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("zh-TW", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);
}

function formatRange(start, end) {
  if (!start && !end) return "未提供";
  if (start && end) return `${formatDate(start)} - ${formatDate(end)}`;
  return formatDate(start || end);
}

function rocDateToIso(year, month, day) {
  const fullYear = Number(year) + 1911;
  return `${fullYear}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function inferVotePeriodFromText(text) {
  const compact = (text || "").replace(/\s+/g, "");
  if (!compact || !compact.includes("電子投票")) return null;

  const patterns = [
    /(\d{2,3})年(\d{1,2})月(\d{1,2})日(?:起)?(?:至|到|迄)(?:(\d{2,3})年)?(\d{1,2})月(\d{1,2})日(?:止|前)?[^。；]{0,50}(?:完成)?電子投票/,
    /(?:電子投票|電子方式行使表決權)[^。；]{0,90}?(\d{2,3})年(\d{1,2})月(\d{1,2})日(?:起)?(?:至|到|迄)(?:(\d{2,3})年)?(\d{1,2})月(\d{1,2})日/,
  ];

  for (const pattern of patterns) {
    const match = compact.match(pattern);
    if (match) {
      const [, startYear, startMonth, startDay, endYear, endMonth, endDay] = match;
      return {
        start: rocDateToIso(startYear, startMonth, startDay),
        end: rocDateToIso(endYear || startYear, endMonth, endDay),
      };
    }
  }

  return null;
}

function getVotePeriod(item) {
  if (item.evoteStartDate || item.evoteEndDate) {
    return {
      text: formatRange(item.evoteStartDate, item.evoteEndDate),
      source: "",
      hasPeriod: true,
    };
  }

  const inferred = inferVotePeriodFromText(
    [
      item.evotePickupRule,
      item.noticeSummary,
      item.noticeGiftSummary,
      item.meetingDistributionRule,
    ]
      .filter(Boolean)
      .join(" "),
  );
  if (inferred) {
    return {
      text: formatRange(inferred.start, inferred.end),
      source: "（由通知書條件推估）",
      hasPeriod: true,
    };
  }

  return {
    text: hasPickupPeriod(item) ? "未補到投票期間（已補到領取期）" : "未提供",
    source: "",
    hasPeriod: false,
  };
}

function formatPickupPeriod(item) {
  const range = formatRange(item.evotePickupStartDate, item.evotePickupEndDate);
  if (range !== "未提供") return range;
  return item.noticeEvotePickupPeriodText || "未提供";
}

function hasPickupPeriod(item) {
  return Boolean(item.evotePickupStartDate || item.evotePickupEndDate || item.noticeEvotePickupPeriodText);
}

function hasNotice(item) {
  return Boolean(item.noticeFilename || item.noticeSourceLabel);
}

function hasVotePeriod(item) {
  return getVotePeriod(item).hasPeriod;
}

function updateCodeCounter() {
  const count = extractCodes(codesInput.value).length;
  codeCounter.textContent = `${count} 檔`;
}

function toSignature(item) {
  return [
    item.status,
    item.souvenirName,
    item.lastBuyDate,
    item.meetingDate,
    item.evoteStartDate,
    item.evoteEndDate,
    item.evotePickupStartDate,
    item.evotePickupEndDate,
    item.evotePickupRule,
  ].join("|");
}

function computeDiffs(results) {
  const previous = loadSnapshots();
  const next = {};

  const enriched = results.map((item) => {
    const signature = toSignature(item);
    const changed = previous[item.code] && previous[item.code] !== signature;
    next[item.code] = signature;
    return { ...item, changed };
  });

  saveSnapshots(next);
  return enriched;
}


function previewDiffs(results) {
  const previous = loadSnapshots();
  return results.map((item) => {
    const signature = toSignature(item);
    const changed = previous[item.code] && previous[item.code] !== signature;
    return { ...item, changed };
  });
}

function filterResults(results) {
  const keyword = (resultSearchInput?.value || "").trim().toLowerCase();
  return results.filter((item) => {
    if (activeFilter === "pickup" && !hasPickupPeriod(item)) return false;
    if (activeFilter === "missingPickup" && hasPickupPeriod(item)) return false;
    if (activeFilter === "notice" && !hasNotice(item)) return false;
    if (activeFilter === "changed" && !item.changed) return false;
    if (!keyword) return true;

    return [
      item.code,
      item.companyName,
      item.souvenirName,
      item.transferAgentName,
      item.transferAgentShort,
      item.evotePickupLocation,
      item.evotePickupDocuments,
      item.noticeSummary,
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase()
      .includes(keyword);
  });
}

function latestYearRecord(row) {
  return row.years?.[0] || {};
}

function rowHasPickupPeriod(row) {
  return row.years?.some((item) => hasPickupPeriod(item));
}

function rowHasNotice(row) {
  return row.years?.some((item) => hasNotice(item));
}

function filterCompareResults(rows) {
  const keyword = (resultSearchInput?.value || "").trim().toLowerCase();
  return rows.filter((row) => {
    const latest = latestYearRecord(row);
    if (activeFilter === "pickup" && !rowHasPickupPeriod(row)) return false;
    if (activeFilter === "missingPickup" && hasPickupPeriod(latest)) return false;
    if (activeFilter === "notice" && !rowHasNotice(row)) return false;
    if (activeFilter === "changed") return true;
    if (!keyword) return true;

    return [
      row.code,
      row.companyName,
      ...((row.years || []).flatMap((item) => [
        item.companyName,
        item.souvenirName,
        item.transferAgentName,
        item.transferAgentShort,
        item.evotePickupLocation,
        item.evotePickupDocuments,
        item.noticeSummary,
      ])),
    ]
      .filter(Boolean)
      .join(" ")
      .toLowerCase()
      .includes(keyword);
  });
}

function renderResultSet(results) {
  renderSummary(results);
  renderRows(filterResults(results));
}

function renderCompareResultSet(rows) {
  renderCompareSummary(rows);
  renderCompareRows(filterCompareResults(rows));
}

function renderSummary(results) {
  summaryStrip.innerHTML = "";
  const published = results.filter((item) => item.status === "published").length;
  const unpublished = results.filter((item) => item.status === "unpublished").length;
  const changed = results.filter((item) => item.changed).length;
  const withVote = results.filter((item) => hasVotePeriod(item)).length;
  const withPickupDate = results.filter(
    (item) => hasPickupPeriod(item),
  ).length;
  const withNotice = results.filter((item) => hasNotice(item)).length;
  const missingPickupDate = Math.max(0, results.length - withPickupDate);

  const cards = [
    ["追蹤代號", results.length, "這次查詢的股票代號數量"],
    ["已公告", published, "已抓到紀念品或會議資料"],
    ["有電投日期", withPickupDate, "已抓到電子投票紀念品領取日期"],
    ["缺電投日期", missingPickupDate, "優先追蹤這些標的"],
    ["已抓通知書", withNotice, "已從 MOPS 或官方 PDF 補到通知書"],
    ["有新變化", changed, "和上次查詢相比有欄位變動"],
    ["有投票期間", withVote, "已補到電子投票起訖或通知書條件"],
  ];

  cards.forEach(([label, value, note]) => {
    const node = summaryCardTemplate.content.firstElementChild.cloneNode(true);
    node.querySelector(".summary-label").textContent = label;
    node.querySelector(".summary-value").textContent = value;
    node.querySelector(".summary-note").textContent = note;
    summaryStrip.appendChild(node);
  });
}

function renderCompareSummary(rows) {
  summaryStrip.innerHTML = "";
  const latestRecords = rows.map((row) => latestYearRecord(row));
  const withCurrentPickup = latestRecords.filter((item) => hasPickupPeriod(item)).length;
  const withAnyNotice = rows.filter((row) => rowHasNotice(row)).length;
  const withThreeYearGift = rows.filter((row) =>
    (row.years || []).filter((item) => item.souvenirName).length >= 3,
  ).length;
  const missingCurrentPickup = Math.max(0, rows.length - withCurrentPickup);

  const cards = [
    ["追蹤代號", rows.length, "這次比較的股票數"],
    ["比較年度", COMPARE_YEARS.join(" / "), "今年與前兩年度"],
    ["今年有電投日期", withCurrentPickup, "目前年度已補到領取日期"],
    ["今年缺電投日期", missingCurrentPickup, "仍要優先補通知書"],
    ["三年都有紀念品", withThreeYearGift, "適合做年度變化比較"],
    ["任一年有通知書", withAnyNotice, "MOPS 或官方 PDF 已補到"],
  ];

  cards.forEach(([label, value, note]) => {
    const node = summaryCardTemplate.content.firstElementChild.cloneNode(true);
    node.querySelector(".summary-label").textContent = label;
    node.querySelector(".summary-value").textContent = value;
    node.querySelector(".summary-note").textContent = note;
    summaryStrip.appendChild(node);
  });
}

function buildStatusBadge(item) {
  const badge = document.createElement("div");
  badge.className = `status-badge ${item.status}`;
  badge.textContent =
    item.status === "published"
      ? "已公告"
      : item.status === "partial"
        ? "部分資料"
        : "未公告";

  if (item.changed) {
    const flag = document.createElement("span");
    flag.className = "new-flag";
    flag.textContent = "NEW";
    badge.appendChild(flag);
  }

  return badge;
}

function renderSources(item) {
  if (!item.sources?.length) return '<span class="source-empty">未標示</span>';
  return item.sources
    .map((source) => `<a class="source-pill" href="${source.url}" target="_blank" rel="noreferrer">${source.label}</a>`)
    .join("");
}

function renderRows(results) {
  if (!results.length) {
    resultsBody.innerHTML = '<tr><td colspan="5" class="empty-cell">查無結果</td></tr>';
    return;
  }

  resultsBody.innerHTML = results
    .map((item) => {
      const souvenirText = item.souvenirName || "尚未公布";
      const pickupPeriod = formatPickupPeriod(item);
      const pickupLocation = item.evotePickupLocation || item.evotePickupPlace || "未補到地點";
      const pickupDocuments = item.evotePickupDocuments || "未補到攜帶資料";
      const pickupRule = item.evotePickupRule || item.noticeSummary || item.meetingDistributionRule || "未補到更細資訊";
      const votePeriod = getVotePeriod(item);
      const dateBlock = `
        <div class="cell-stack">
          <span><strong>最後買進：</strong>${formatDate(item.lastBuyDate)}</span>
          <span><strong>股東會：</strong>${formatDate(item.meetingDate)}</span>
          <span><strong>地點：</strong>${item.meetingCity || "未提供"}</span>
        </div>
      `;
      const evoteBlock = `
        <div class="cell-stack">
          <span><strong>電子投票期間：</strong>${votePeriod.text}${votePeriod.source}</span>
          <div class="pickup-focus ${hasPickupPeriod(item) ? "ready" : "missing"}">
            <span><strong>電投領取期</strong>${pickupPeriod}</span>
            <span><strong>領取地點</strong>${pickupLocation}</span>
            <span><strong>攜帶資料</strong>${pickupDocuments}</span>
          </div>
          <span><strong>領取資訊：</strong>${pickupRule}</span>
          <span><strong>領取來源：</strong>${item.evotePickupSource || "未標示"}</span>
          <span><strong>通知書來源：</strong>${item.noticeSourceLabel || "未抓到通知書"}</span>
          <span><strong>通知書摘要：</strong>${item.noticeSummary || item.noticeGiftSummary || "未補到摘要"}</span>
          <span><strong>通知書快取：</strong>${
            item.noticeCacheStatus === "hit"
              ? "已快取"
              : item.noticeCacheStatus === "miss"
                ? "本次新抓"
                : "未使用"
          }</span>
          <div class="merged-meta">
            <span><strong>股代：</strong>${item.transferAgentName || item.transferAgentShort || "未提供"}</span>
            <span><strong>電話：</strong>${item.transferAgentPhone || "未提供"}</span>
            <span><strong>零股寄單：</strong>${item.oddLotMail || "未提供"}</span>
            <span><strong>備註：</strong>${item.notes || item.proxyPeriodText || "—"}</span>
          </div>
          <div class="source-links">
            <strong>來源</strong>
            <div>${renderSources(item)}</div>
          </div>
        </div>
      `;

      return `
        <tr class="${item.changed ? "row-changed" : ""} ${hasPickupPeriod(item) ? "" : "row-missing-pickup"}">
          <td>
            <div class="stock-block">
              <strong>${item.code}</strong>
              <span>${item.companyName || "尚未比對到公司名稱"}</span>
            </div>
          </td>
          <td></td>
          <td>${souvenirText}</td>
          <td>${dateBlock}</td>
          <td>${evoteBlock}</td>
        </tr>
      `;
    })
    .join("");

  [...resultsBody.querySelectorAll("tr")].forEach((row, index) => {
    const cell = row.children[1];
    if (cell) {
      cell.appendChild(buildStatusBadge(results[index]));
    }
  });
}

function yearCard(item, section = "gift") {
  const rocYear = item.rocYear || (item.year ? Number(item.year) - 1911 : "");
  const isCurrent = Number(rocYear) === CURRENT_ROC_YEAR;
  const hasData = Boolean(item.souvenirName || item.meetingDate || item.noticeFilename || item.noticeSourceLabel);
  const statusText =
    item.status === "published"
      ? "已公告"
      : item.status === "partial"
        ? "部分資料"
        : "未補齊";
  const className = `year-card ${isCurrent ? "current" : ""} ${hasData ? "" : "missing"}`;

  if (section === "dates") {
    return `
      <div class="${className}">
        <div class="year-card-head"><strong>${rocYear} 年</strong><span>${statusText}</span></div>
        <p>最後買進：${formatDate(item.lastBuyDate)}</p>
        <p>股東會：${formatDate(item.meetingDate)}</p>
        <p>地點：${item.meetingCity || "未提供"}</p>
      </div>
    `;
  }

  if (section === "pickup") {
    return `
      <div class="${className}">
        <div class="year-card-head"><strong>${rocYear} 年</strong><span>${hasNotice(item) ? "有通知書" : "無通知書"}</span></div>
        <p>電投領取期：${formatPickupPeriod(item)}</p>
        <p>地點：${item.evotePickupLocation || item.evotePickupPlace || "未補到地點"}</p>
        <p>攜帶：${item.evotePickupDocuments || "未補到攜帶資料"}</p>
        <p>摘要：${item.noticeSummary || item.evotePickupRule || "未補到摘要"}</p>
      </div>
    `;
  }

  return `
    <div class="${className}">
      <div class="year-card-head"><strong>${rocYear} 年</strong><span>${statusText}</span></div>
      <p>紀念品：${item.souvenirName || "尚未補齊"}</p>
      <p>股代：${item.transferAgentName || item.transferAgentShort || "未提供"}</p>
    </div>
  `;
}

function renderCompareRows(rows) {
  if (!rows.length) {
    resultsBody.innerHTML = '<tr><td colspan="5" class="empty-cell">查無比較結果</td></tr>';
    return;
  }

  resultsBody.innerHTML = rows
    .map((row) => {
      const latest = latestYearRecord(row);
      const history = row.years || [];
      return `
        <tr class="${hasPickupPeriod(latest) ? "" : "row-missing-pickup"}">
          <td>
            <div class="stock-block">
              <strong>${row.code}</strong>
              <span>${row.companyName || latest.companyName || "尚未比對到公司名稱"}</span>
            </div>
          </td>
          <td></td>
          <td><div class="year-comparison">${history.map((item) => yearCard(item, "gift")).join("")}</div></td>
          <td><div class="year-comparison">${history.map((item) => yearCard(item, "dates")).join("")}</div></td>
          <td><div class="year-comparison">${history.map((item) => yearCard(item, "pickup")).join("")}</div></td>
        </tr>
      `;
    })
    .join("");

  [...resultsBody.querySelectorAll("tr")].forEach((row, index) => {
    const cell = row.children[1];
    const latest = latestYearRecord(rows[index]);
    if (cell) {
      cell.appendChild(buildStatusBadge(latest));
    }
  });
}

async function lookupCompare(codes) {
  if (!codes.length) {
    statusText.textContent = "請先輸入至少一筆可辨識的股票代號";
    return;
  }

  activeCodes = codes;
  lookupBtn.disabled = true;
  refreshBtn.disabled = true;
  exportBtn.disabled = true;
  statusText.textContent = `正在讀取 ${COMPARE_YEARS.join(" / ")} 年比較快照...`;
  resultsBody.innerHTML = '<tr><td colspan="5" class="empty-cell">正在整理三年度比較...</td></tr>';

  try {
    const response = await fetch(
      `/api/compare?codes=${encodeURIComponent(codes.join(" "))}&years=${encodeURIComponent(compareYearsParam())}`,
    );
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || "三年度比較讀取失敗");
    }
    lastResponse = { mode: "compare", requestedCodes: codes, results: payload.results || [] };
    updatedAtText.textContent = new Date().toLocaleString("zh-TW");
    statusText.textContent = `已完成 ${lastResponse.results.length} 檔三年度比較。`;
    renderCompareResultSet(lastResponse.results);
    exportBtn.disabled = false;
  } catch (error) {
    exportBtn.disabled = true;
    statusText.textContent = error.message || "三年度比較讀取失敗";
    resultsBody.innerHTML = `<tr><td colspan="5" class="empty-cell">${statusText.textContent}</td></tr>`;
  } finally {
    lookupBtn.disabled = false;
    refreshBtn.disabled = false;
  }
}

async function lookup(codes) {
  if (viewMode === "compare") {
    await lookupCompare(codes);
    return;
  }

  if (!codes.length) {
    statusText.textContent = "請先輸入至少一筆可辨識的股票代號";
    return;
  }

  activeCodes = codes;
  lookupBtn.disabled = true;
  refreshBtn.disabled = true;
  exportBtn.disabled = true;
  statusText.textContent = "正在抓取最新資料...";
  resultsBody.innerHTML = '<tr><td colspan="5" class="empty-cell">準備開始逐筆查詢...</td></tr>';

  try {
    const collected = [];

    for (let index = 0; index < codes.length; index += 1) {
      const code = codes[index];
      statusText.textContent = `正在查詢 ${index + 1} / ${codes.length}：${code}`;

      try {
        const response = await fetch(`/api/lookup?codes=${encodeURIComponent(code)}`);
        const payload = await response.json();
        if (!response.ok || !payload.ok) {
          throw new Error(payload.error || "查詢失敗");
        }
        if (payload.results?.[0]) {
          collected.push(payload.results[0]);
        }
      } catch (error) {
        collected.push({
          code,
          companyName: "",
          status: "partial",
          isPublished: false,
          souvenirName: "",
          meetingDate: "",
          lastBuyDate: "",
          meetingCity: "",
          priceText: "",
          transferAgentName: "",
          transferAgentPhone: "",
          transferAgentShort: "",
          oddLotMail: "",
          reelection: "",
          needVote: null,
          fractionalOk: null,
          evoteStartDate: "",
          evoteEndDate: "",
          evotePickupStartDate: "",
          evotePickupEndDate: "",
          evotePickupPlace: "",
          evotePickupLocation: "",
          evotePickupDocuments: "",
          evotePickupRule: "",
          evotePickupSource: "",
          noticeSummary: "",
          noticeGiftSummary: "",
          noticeEvotePickupPeriodText: "",
          noticeCacheStatus: "",
          noticeSourceLabel: "",
          noticeSourceType: "",
          notes: `抓取失敗：${error.message || "未知錯誤"}`,
          sources: [],
        });
      }

      const preview = previewDiffs(collected);
      lastResponse = { mode: "current", requestedCodes: codes, results: preview };
      renderResultSet(preview);
    }

    const enriched = computeDiffs(collected);
    lastResponse = { mode: "current", requestedCodes: codes, results: enriched };
    exportBtn.disabled = !enriched.length;
    updatedAtText.textContent = new Date().toLocaleString("zh-TW");
    statusText.textContent = `已完成 ${enriched.length} 檔查詢，資料已逐筆載入。`;
    renderResultSet(enriched);
  } catch (error) {
    exportBtn.disabled = true;
    statusText.textContent = error.message || "查詢失敗";
    resultsBody.innerHTML = `<tr><td colspan="5" class="empty-cell">${statusText.textContent}</td></tr>`;
  } finally {
    lookupBtn.disabled = false;
    refreshBtn.disabled = false;
  }
}

lookupBtn.addEventListener("click", () => {
  const codes = extractCodes(codesInput.value);
  lookup(codes);
});

currentModeBtn?.addEventListener("click", () => {
  setViewMode("current");
  if (lastResponse?.results) {
    lookup(activeCodes.length ? activeCodes : extractCodes(codesInput.value));
  }
});

compareModeBtn?.addEventListener("click", () => {
  setViewMode("compare");
  const codes = activeCodes.length ? activeCodes : extractCodes(codesInput.value);
  if (codes.length) {
    lookup(codes);
  }
});

codesInput.addEventListener("input", updateCodeCounter);

filterTabs.forEach((button) => {
  button.addEventListener("click", () => {
    activeFilter = button.dataset.filter || "all";
    filterTabs.forEach((tab) => tab.classList.toggle("active", tab === button));
    if (lastResponse?.results) {
      if (lastResponse.mode === "compare") {
        renderCompareResultSet(lastResponse.results);
      } else {
        renderResultSet(lastResponse.results);
      }
    }
  });
});

resultSearchInput?.addEventListener("input", () => {
  if (lastResponse?.results) {
    if (lastResponse.mode === "compare") {
      renderCompareResultSet(lastResponse.results);
    } else {
      renderResultSet(lastResponse.results);
    }
  }
});

refreshBtn.addEventListener("click", () => {
  const codes = activeCodes.length ? activeCodes : extractCodes(codesInput.value);
  lookup(codes);
});

exportBtn.addEventListener("click", async () => {
  const codes = activeCodes.length ? activeCodes : extractCodes(codesInput.value);
  if (!codes.length) {
    statusText.textContent = "請先查詢資料後再下載 Excel";
    return;
  }
  exportBtn.disabled = true;
  statusText.textContent = viewMode === "compare" ? "正在產生今年資料 Excel..." : "正在產生 Excel...";
  try {
    const response = await fetch("/api/export.xlsx", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
      },
      body: new URLSearchParams({ codes: codes.join(" ") }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      let message = "Excel 匯出失敗";
      try {
        message = JSON.parse(errorText).error || message;
      } catch {
        message = errorText || message;
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    if (!blob.size) throw new Error("Excel 匯出失敗：下載內容是空的");

    const disposition = response.headers.get("Content-Disposition") || "";
    const filenameMatch = disposition.match(/filename\*?=(?:UTF-8''|")?([^";]+)/i);
    const filename = decodeURIComponent(filenameMatch?.[1] || `shareholder-gifts-${Date.now()}.xlsx`);
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = filename.replace(/[\\/:*?"<>|]/g, "-");
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
    statusText.textContent = `Excel 已下載：${link.download}`;
  } catch (error) {
    statusText.textContent = error.message || "Excel 匯出失敗";
  } finally {
    exportBtn.disabled = !lastResponse?.results?.length;
  }
});

saveWatchlistBtn.addEventListener("click", () => {
  const codes = extractCodes(codesInput.value);
  saveWatchlist(codes);
  statusText.textContent = `已儲存 ${codes.length} 筆 watchlist。`;
});

loadSampleBtn.addEventListener("click", () => {
  codesInput.value = sampleCodes.join("\n");
  updateCodeCounter();
});

clearBtn.addEventListener("click", () => {
  codesInput.value = "";
  activeCodes = [];
  exportBtn.disabled = true;
  updateCodeCounter();
  statusText.textContent = "已清空輸入";
});

setInterval(() => {
  if (activeCodes.length) {
    lookup(activeCodes);
  }
}, AUTO_REFRESH_MS);

function bootstrap() {
  setViewMode("current");
  refreshNoticeProgress();
  const saved = loadWatchlist();
  if (saved) {
    codesInput.value = saved;
    updateCodeCounter();
    const codes = extractCodes(saved);
    if (codes.length) {
      lookup(codes);
    }
  }
  updateCodeCounter();
}

bootstrap();

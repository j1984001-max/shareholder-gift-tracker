const WATCHLIST_KEY = "shareholder-gift-tracker-watchlist-v1";
const SNAPSHOT_KEY = "shareholder-gift-tracker-snapshot-v1";
const AUTO_REFRESH_MS = 15 * 60 * 1000;

const sampleCodes = ["2317", "2409", "3037", "9938", "2006", "1416"];

const codesInput = document.getElementById("codesInput");
const lookupBtn = document.getElementById("lookupBtn");
const saveWatchlistBtn = document.getElementById("saveWatchlistBtn");
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
    noticeProgressText.textContent = `通知書 ${progress.noticeCached}/${progress.watchlistTotal}，電投日期 ${progress.pickupDateCached}/${progress.watchlistTotal}，官網掃描 ${progress.officialSiteScanned || 0}/${progress.watchlistTotal}`;
    noticeProgressText.title = `官網/官方 PDF：${progress.officialPdfCached}；官網找到：${progress.officialSiteFound || 0}；尚缺通知書：${progress.missingNotice}；尚缺電投日期：${progress.missingPickupDate}`;
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

function renderResultSet(results) {
  renderSummary(results);
  renderRows(filterResults(results));
}

function renderSummary(results) {
  summaryStrip.innerHTML = "";
  const published = results.filter((item) => item.status === "published").length;
  const unpublished = results.filter((item) => item.status === "unpublished").length;
  const changed = results.filter((item) => item.changed).length;
  const withVote = results.filter((item) => item.evoteStartDate || item.evoteEndDate).length;
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
    ["含電子投票", withVote, "已補到電子投票起訖資訊"],
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
  if (!item.sources?.length) return "—";
  return item.sources
    .map((source) => `<a href="${source.url}" target="_blank" rel="noreferrer">${source.label}</a>`)
    .join("<br>");
}

function renderRows(results) {
  if (!results.length) {
    resultsBody.innerHTML = '<tr><td colspan="7" class="empty-cell">查無結果</td></tr>';
    return;
  }

  resultsBody.innerHTML = results
    .map((item) => {
      const souvenirText = item.souvenirName || "尚未公布";
      const pickupPeriod = formatPickupPeriod(item);
      const pickupLocation = item.evotePickupLocation || item.evotePickupPlace || "未補到地點";
      const pickupDocuments = item.evotePickupDocuments || "未補到攜帶資料";
      const pickupRule = item.evotePickupRule || item.noticeSummary || item.meetingDistributionRule || "未補到更細資訊";
      const dateBlock = `
        <div class="cell-stack">
          <span><strong>最後買進：</strong>${formatDate(item.lastBuyDate)}</span>
          <span><strong>股東會：</strong>${formatDate(item.meetingDate)}</span>
          <span><strong>地點：</strong>${item.meetingCity || "未提供"}</span>
        </div>
      `;
      const evoteBlock = `
        <div class="cell-stack">
          <span><strong>電子投票：</strong>${formatRange(item.evoteStartDate, item.evoteEndDate)}</span>
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
        </div>
      `;
      const agentBlock = `
        <div class="cell-stack">
          <span><strong>股代：</strong>${item.transferAgentName || item.transferAgentShort || "未提供"}</span>
          <span><strong>電話：</strong>${item.transferAgentPhone || "未提供"}</span>
          <span><strong>零股寄單：</strong>${item.oddLotMail || "未提供"}</span>
          <span><strong>備註：</strong>${item.notes || item.proxyPeriodText || "—"}</span>
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
          <td>${agentBlock}</td>
          <td class="sources-cell">${renderSources(item)}</td>
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

async function lookup(codes) {
  if (!codes.length) {
    statusText.textContent = "請先輸入至少一筆可辨識的股票代號";
    return;
  }

  activeCodes = codes;
  lookupBtn.disabled = true;
  refreshBtn.disabled = true;
  exportBtn.disabled = true;
  statusText.textContent = "正在抓取最新資料...";
  resultsBody.innerHTML = '<tr><td colspan="7" class="empty-cell">準備開始逐筆查詢...</td></tr>';

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
      lastResponse = { requestedCodes: codes, results: preview };
      renderResultSet(preview);
    }

    const enriched = computeDiffs(collected);
    lastResponse = { requestedCodes: codes, results: enriched };
    exportBtn.disabled = !enriched.length;
    updatedAtText.textContent = new Date().toLocaleString("zh-TW");
    statusText.textContent = `已完成 ${enriched.length} 檔查詢，資料已逐筆載入。`;
    renderResultSet(enriched);
  } catch (error) {
    exportBtn.disabled = true;
    statusText.textContent = error.message || "查詢失敗";
    resultsBody.innerHTML = `<tr><td colspan="7" class="empty-cell">${statusText.textContent}</td></tr>`;
  } finally {
    lookupBtn.disabled = false;
    refreshBtn.disabled = false;
  }
}

lookupBtn.addEventListener("click", () => {
  const codes = extractCodes(codesInput.value);
  lookup(codes);
});

codesInput.addEventListener("input", updateCodeCounter);

filterTabs.forEach((button) => {
  button.addEventListener("click", () => {
    activeFilter = button.dataset.filter || "all";
    filterTabs.forEach((tab) => tab.classList.toggle("active", tab === button));
    if (lastResponse?.results) {
      renderResultSet(lastResponse.results);
    }
  });
});

resultSearchInput?.addEventListener("input", () => {
  if (lastResponse?.results) {
    renderResultSet(lastResponse.results);
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
  statusText.textContent = "正在產生 Excel...";
  try {
    const form = document.createElement("form");
    form.method = "POST";
    form.action = "/api/export.xlsx";
    form.target = "excelDownloadFrame";
    form.style.display = "none";

    const input = document.createElement("input");
    input.type = "hidden";
    input.name = "codes";
    input.value = codes.join(" ");
    form.appendChild(input);

    let frame = document.querySelector('iframe[name="excelDownloadFrame"]');
    if (!frame) {
      frame = document.createElement("iframe");
      frame.name = "excelDownloadFrame";
      frame.style.display = "none";
      document.body.appendChild(frame);
    }

    document.body.appendChild(form);
    form.submit();
    form.remove();
    statusText.textContent = "Excel 已開始下載";
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

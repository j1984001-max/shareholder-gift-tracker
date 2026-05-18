# 股東會紀念品追蹤器

批次輸入股票代號後，整理今年度股東會紀念品、最後買進日、電子投票期間與可補到的領取資訊。

## 啟動方式

```bash
cd "/Users/wujohnson/Documents/New project/shareholder-gift-tracker"
python3 server.py
```

預設會開在：

```text
http://127.0.0.1:8765
```

如果只是想在這台電腦或同區網裝置開：

- 本機：`http://127.0.0.1:8765`
- 同網路其他裝置：`http://你的電腦區網IP:8765`

## 目前整合的來源

- 撿股讚：紀念品、開會日期、最後買進日、股代、零股寄單
- 股東禮簿：電子投票起訖、零股可否領取、股代名稱
- 宏遠股代：部分個股的電投領取期間與發放條件
- 公開資訊觀測站／官方 PDF 快取：開會通知書內的電投領取日期、地點、攜帶資料

## 這版可以做什麼

- 批次貼上多筆股票代號
- 儲存 watchlist 到瀏覽器 localStorage
- 網站優先讀取本機預先建好的查詢快照，不把 Render 當成臨時爬蟲
- 和上次查詢結果比對，欄位有異動時標示 `NEW`
- 對於尚未公告的股票，保留在 watchlist 方便持續追蹤

## 本機優先架構

通知書下載與解析建議都在你的本機跑完，再把結果推到 GitHub 讓 Render 自動部署。部署後網站查詢會優先讀 [data/lookup_snapshot.json](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/data/lookup_snapshot.json) 與 [data/mops_notice_seed_cache.json](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/data/mops_notice_seed_cache.json)，避免 Render 線上臨時打外部網站。

推薦的本機更新流程：

```bash
cd "/Users/wujohnson/Documents/New project/shareholder-gift-tracker"
PYTHONPATH=.vendor python3 tools/local_refresh_pipeline.py --official-limit 120 --mops-limit 5
```

這支腳本會依序做三件事：

- 掃公司官網 / 官方 PDF
- 小批次補公開資訊觀測站或官方通知書快取
- 重建部署用的 [data/lookup_snapshot.json](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/data/lookup_snapshot.json)

如果你想讓本機每 5 分鐘自動補一次，而且已經有完整資料的股票不要重複查 MOPS，可以直接用這支單次 worker：

```bash
cd "/Users/wujohnson/Documents/New project/shareholder-gift-tracker"
PYTHONPATH=.vendor python3 tools/run_local_mops_cycle.py
```

這支 worker 會：

- 只跑 MOPS / 官方通知書補資料
- 每次最多抓 `5` 檔
- 用 `--skip-existing --retry-empty`，已有完整領取資料的不會重抓
- 補完後自動重建 `lookup_snapshot.json`
- 內建 lock，避免上一輪還沒跑完又重複啟動

如果要讓 Mac 自動每 5 分鐘跑一次：

```bash
cp "/Users/wujohnson/Documents/New project/shareholder-gift-tracker/ops/shareholder-gift-tracker.local-mops.plist" ~/Library/LaunchAgents/
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/shareholder-gift-tracker.local-mops.plist
launchctl kickstart -k gui/$(id -u)/com.wujohnson.shareholder-gift-tracker.local-mops
```

查看 log：

```bash
tail -f "/Users/wujohnson/Documents/New project/shareholder-gift-tracker/.cache/local-mops-cycle.log"
```

如果專案放在 `Documents` 下面，macOS 可能會擋掉背景 `launchd` 行程直接讀這個資料夾。遇到這種情況，最穩的做法是把 repo 移到像 `~/workspace/shareholder-gift-tracker` 這類非受保護路徑，再啟用 `launchd`。不搬資料夾的話，就改成在已授權的 Terminal 裡手動或常駐執行這支 worker。

跑完後再推上去：

```bash
git add data/mops_notice_seed_cache.json data/official_notice_sources.json data/official_site_scan_cache.json data/lookup_snapshot.json
git commit -m "Refresh shareholder lookup snapshot"
git push origin main
```

## 後台補資料

GitHub Actions 仍會定期執行 [tools/update_mops_seed.py](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/tools/update_mops_seed.py) 與 [tools/build_lookup_snapshot.py](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/tools/build_lookup_snapshot.py)，把通知書解析結果與部署用快照一起更新。網站查詢時會優先讀這些已建好的資料檔。

如果公開資訊觀測站查詢過量，可以把公司官網或股代公告的官方 PDF 放到 [data/official_notice_sources.json](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/data/official_notice_sources.json)，排程會優先解析這些 PDF，例如：

```json
{
  "1101": [
    {
      "label": "公司官網開會通知書",
      "url": "https://example.com/1101_notice.pdf",
      "sourceType": "company_pdf"
    }
  ]
}
```

## 限制

- 電投領取細節目前只有在整合來源有公開欄位時才能補出來，並不是每家股代網站都會直接提供。
- 資料來源都是公開網頁，若對方網站改版，可能需要微調解析邏輯。
- 實際領取規則仍應以公司通知書、公開資訊觀測站與股代公告為準。

## Render 上線

這個資料夾已包含：

- [Dockerfile](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/Dockerfile)
- [render.yaml](/Users/wujohnson/Documents/New project/shareholder-gift-tracker/render.yaml)

如果要公開上線到 Render：

1. 把 `shareholder-gift-tracker/` 推到 GitHub repo。
2. 登入 Render。
3. 點 `New +` -> `Blueprint`。
4. 選你的 GitHub repo。
5. Render 會自動讀取 `render.yaml` 建立網站。
6. 部署完成後，直接用 Render 提供的網址開啟。

# running

富士山マラソン 2026（12/14）訓練課表與紀錄。

- 📱 手機看這裡：**https://iamwuwei.github.io/running/**
- `富士山馬拉松_2026_課表.xlsx`：資料來源
  - 工作表「課表」：每日計畫與實際紀錄
  - 工作表「週總覽」：每週計畫 / 實際 km、達成率、重點
- `build_site.py`：從 xlsx 產生 `docs/index.html`（每日報告與分析、下次練習、週進度、歷史紀錄）
- `.github/workflows/pages.yml`：push 到 main 時自動重新產生並部署到 GitHub Pages

## 更新流程

1. 在「課表」填入當天的實際 km、時間、配速、心拍等
2. `python3 build_site.py`
3. commit xlsx 與 `docs/index.html` 一起 push 到 main

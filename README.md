# 巧固球賽程自動生成器

## 本機運行
1. 安裝 Python 3.10 或以上版本。
2. 在專案資料夾開啟終端機。
3. 執行：`pip install -r requirements.txt`
4. 執行：`streamlit run app.py`

## GitHub + Streamlit Community Cloud 部署
1. 在 GitHub 建立一個新的 Public repository，例如 `tchoukball-scheduler`。
2. 選擇 Add file → Upload files。
3. 上載 `app.py` 及 `requirements.txt`（README.md 可選）。
4. 按 Commit changes。
5. 前往 https://share.streamlit.io/ 並以 GitHub 登入。
6. 選擇 Create app，揀選剛才的 repository、branch（通常是 main）及主檔案 `app.py`。
7. 按 Deploy。部署完成後會得到一條可分享的網址。

## 目前版本的規則
- 15 分鐘 × 3 節：每場佔用 60 分鐘。
- 12 分鐘 × 3 節：每場佔用 45 分鐘。
- 5 隊或以下：單循環，可選純單循環、四強及冠軍／季軍賽、或直接冠軍／季軍賽。
- 6–7 隊：兩組，每組頭兩名晉級四強。
- 8 隊或以上：兩組，每組頭四名晉級八強；之後四強、冠軍賽及季軍賽。
- 可生成單場地、雙場地或兩個版本。
- 每場之間不額外加入休息時間。
- 生成 Excel / CSV。

## 重要
此版本是第一個可運行版本。正式比賽前，請核對分組、晉級名額及時間表。淘汰賽對賽目前使用排名佔位符，尚未提供賽果輸入後自動更新下一輪隊名的功能。

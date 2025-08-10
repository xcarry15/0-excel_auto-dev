## Excel/CSV 批量合并工具（Web 版）

基于 Streamlit 的可视化网页工具，支持一次上传多个 CSV/Excel 文件，设置“跳过行数”（标题区行数），在线预览并下载合并结果。

### 运行环境
- Python 3.10+
- Windows/macOS/Linux

### 安装
```bash
pip install -r requirements.txt
```

### 启动
```bash
streamlit run web_app.py
```

启动后浏览器会自动打开本地页面。若未自动打开，可访问 `http://localhost:8501`。

### 使用说明
- 在侧边栏设置“跳过行数”（即每个文件顶部非数据的标题区行数）。
- 点击“选择要合并的文件”，可多选上传 `.csv` 与 `.xlsx` 文件（可混合）。
- 点击“开始合并”后，将分别对 CSV 与 Excel 两类文件分组合并：
  - 合并时以首个文件的数据区首行确定“列数”（避免把数据当列名），其“标题区最后一行”仅用于展示预览。
  - 若不同文件列数不一致，多余列将被截断，缺失列会补空。
- 完成后可在线预览前 100 行，并下载生成文件。

### 常见问题
- 无法读取 CSV：尝试了多种常见编码（utf-8/gbk/gb2312/utf-8-sig/latin1）。若依旧失败，请手动转换编码后重试。
- Excel 标题加粗：当“跳过行数”>0 时，会将输出 Excel 的标题区最后一行加粗；CSV 不含样式。
- 大文件：如文件很多或体积较大，等待时间会增加。建议按类型分批上传。

### 推送到 GitHub 并在 Streamlit Community Cloud 部署
1. 初始化 Git 仓库并提交代码：
   ```bash
   git init
   git add .
   git commit -m "feat: web app 初版"
   ```
2. 在 GitHub 新建仓库（建议公开仓库以便 Streamlit 部署），将本地仓库推送：
   ```bash
   git remote add origin https://github.com/<your-org-or-user>/<repo-name>.git
   git branch -M main
   git push -u origin main
   ```
3. 打开 Streamlit Community Cloud（需要 GitHub 账户授权），新建应用：
   - 选择仓库与分支：`<your-org-or-user>/<repo-name> @ main`
   - 指定入口脚本：`web_app.py`
   - Python 版本：与本地一致（如 3.10+）
4. 依赖管理：平台会自动识别 `requirements.txt` 并安装。
5. 可选：如需私密配置（例如 API Key），请在 Streamlit 平台的 `Secrets` 中设置，不要把敏感信息提交到仓库。


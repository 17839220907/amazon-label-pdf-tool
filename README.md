# 亚马逊标签 PDF 飞书上传工具

这是 Streamlit 在线版。

功能：

- 上传一个或多个标签 PDF
- 自动拆分多页 PDF
- 从飞书在线索引表读取 FNSKU / SKU / 款式等信息
- 遮住旧文字并写入新标签文字
- 上传到指定飞书文件夹
- 同一个 FNSKU 只保留一个文件
- 生成处理日志

## 本地运行

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Cloud 部署

主文件填写：

```text
streamlit_app.py
```

需要在 Streamlit Cloud 的 App secrets 里填写：

```toml
[feishu]
app_id = "你的飞书应用 app_id"
app_secret = "你的飞书应用 app_secret"
spreadsheet_token = "飞书在线索引表 token"
output_folder_token = "飞书目标文件夹 token"
index_range = "A1:G50000"
```

## 使用

1. 上传一个或多个标签 PDF。
2. 点击开始处理并上传飞书。
3. 处理完成后查看成功和异常数量。
4. 下载处理日志 Excel。

异常标签不会上传成功，只会写入处理日志。

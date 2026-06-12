# 亚马逊标签 PDF 自动整理工具

这是 Streamlit 在线版。

## 本地运行

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Cloud 部署

1. 把本文件夹上传到 GitHub 仓库。
2. 打开 Streamlit Cloud，创建新应用。
3. 选择仓库，主文件填写 `streamlit_app.py`。
4. 点击 Deploy。

## 使用

1. 上传 `标签索引表.xlsx`。
2. 上传一个或多个标签 PDF。
3. 点击开始处理。
4. 下载处理结果 ZIP。

异常标签不会放入成功 PDF 文件夹，只会写入处理日志。

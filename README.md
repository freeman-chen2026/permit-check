# 国内批复核对工具

一个基于 Streamlit 的国内批复自动核对工具。

## 功能

- 自动从批复汇总表 docx 中识别类似 `B8160 GLF5 VRDA1500 2245ZBAD ON 01OCT2026 U/H 中国籍` 的批复行
- 按注册号自动判断时区（B 注册按北京时间；其他按世界时 UTC +8）
- 与航段数据导出 xlsx 进行起飞/落地机场、时间、日期、用途比对
- 与文本版航班信息对比机组，判断“中国籍”备注
- 差异处标红，并支持下载标红后的 docx

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 Streamlit Cloud

1. 把本项目 push 到 GitHub
2. 打开 https://share.streamlit.io
3. New app → 选择仓库 → Main file path 填 `app.py` → Deploy

## 使用

页面上传四个文件：

1. 国内批复信息汇总表 (.docx)
2. 航段数据导出 (.xlsx)
3. 文本版航班信息 (.txt)
4. 飞行员名单 (.txt)

点击 **开始核对**，查看差异并下载标红 docx。

## 关键规则

- `U/H` → 载客；`N/M` → 调机
- B 注册号时间按北京时间处理
- 非 B 注册号时间按世界时（UTC），加 8 小时转为北京时间
- 备注含“中国籍”即期望机组全部为中国籍飞行员

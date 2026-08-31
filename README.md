# BP 项目初审工具

自动批量审核 BP（商业计划书）PDF，调用 DeepSeek 大模型逐份审阅并生成 **Excel 汇总报告** 的桌面工具。

## 功能

- 批量读取 `BP/` 文件夹中的 PDF 商业计划书
- 调用 DeepSeek API 逐份审核
- 审核完成后在 `output/` 文件夹生成 Excel 汇总报告

## 使用说明

1. 把商业计划书 PDF 全部放入 **BP/** 文件夹
2. 编辑 **.env** 文件，把 `DEEPSEEK_API_KEY` 改成你自己的 DeepSeek API Key（`.env` 已被 git 忽略，不会上传到仓库；可参照 `.env.example` 模板）
3. 双击 **FX.exe** 运行程序
4. 运行完成后自动新建 `output/` 文件夹存放 Excel 汇总报告

> ⚠️ 运行过程中不要打开 `output/` 里的 Excel，否则保存会报错。
> 如果生成报告为空白，一般是 PDF 是扫描版导致，可加装文字识别（OCR）库。

## 依赖

```bash
pip install PyPDF2 python-dotenv pandas openai pycryptodome openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 文件说明

| 文件 | 说明 |
| --- | --- |
| `FX.exe` | 主程序（约 139MB，因超过 GitHub 单文件 100MB 限制，通过 Releases 附件分发） |
| `.env.example` | API Key 配置模板（复制为 `.env` 后填入真实 Key） |
| `前置库.txt` | Python 依赖清单 |
| `说明.txt` | 原版使用说明 |
| `BP/` | 放置待审 BP PDF 的文件夹 |

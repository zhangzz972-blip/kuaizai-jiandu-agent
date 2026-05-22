# 快哉荐读 agent v1.0

徐州图书馆「国潮汉风、快哉荐读」栏目智能荐读助手。

## 快速开始

```powershell
$env:DEEPSEEK_API_KEY = "sk-xxx"
pip install flask openai pandas openpyxl pillow python-docx playwright
playwright install chromium
python app.py
```

浏览器打开 `http://127.0.0.1:5000`

## 功能流程

```
输入主题 -> AI 语义展开关键词 -> 书库检索(3万条) -> AI 精选书目
    -> 撰写荐读稿 -> 润色审校 -> 导出 Word/HTML/MD/海报
```

## 荐读稿结构

- **卷首语** — 2-3段引言，连接主题与阅读
- **推荐书籍** — 书名、作者、索书号、馆藏位置（徐州图书馆参考咨询室）、荐读理由（150-300字）
- **结语** — 「书香徐来、开卷快哉」+ 徐州图书馆

## 项目文件

| 文件 | 说明 |
|------|------|
| `app.py` | Flask Web 应用（主入口） |
| `jiandu_deepseek.py` | 核心引擎（可导入模块 + CLI） |
| `templates/index.html` | Web 前端界面 |
| `参考咨询书库馆藏清单.xlsx` | 馆藏书库（约30,000条） |
| `馆标黑版.png` | 徐州图书馆馆标 |

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `DEEPSEEK_API_KEY` | **必填** | Deepseek API 密钥 |
| `DEEPSEEK_BASE_URL` | `https://api.deepseek.com/v1` | API 地址 |
| `DEEPSEEK_MODEL` | `deepseek-v4-pro` | 模型名称 |

## Web 界面功能

| 功能 | 说明 |
|------|------|
| 风格切换 | 典雅中国风 / 现代简约 / 清雅书香 / 水墨意境 |
| 预览/编辑双模式 | 预览查看渲染效果，编辑直接改文字 |
| AI 智能精选 | 从候选书中自动挑选最贴题的书 |
| AI 按意见重写 | 输入修改意见，AI 重写整篇 |
| 多格式导出 | Word / HTML / Markdown 一键下载 |

## CLI 用法

```powershell
python jiandu_deepseek.py --theme "端午节" --max 8
```

## 依赖

Python 3.10+, flask, openai, pandas, openpyxl, pillow, python-docx, playwright

## License

徐州图书馆内部工具

# WIPO Global Brand Database Spider

从 [WIPO Global Brand Database](https://branddb.wipo.int/) 批量下载品牌商标 PDF。

## 功能

- 从本地 Excel 文件读取品牌名称（支持中文路径）
- 调用 WIPO 搜索 API 获取全部匹配商标记录（自动处理分页）
- 批量下载每条商标记录的 PDF 文件
- 自动处理速率限制（HTTP 429）与服务器错误，支持指数退避重试
- 保存进度文件，支持断点续传（中断后重新运行可跳过已完成品牌）
- PDF 按品牌名分目录保存，文件名包含办公室代码和注册号

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 直接运行（使用默认配置）

修改 `wipo_brand_spider.py` 顶部的配置常量：

```python
EXCEL_PATH  = os.path.join(os.path.expanduser("~"), "Desktop", "Fashion_jet", "可挖掘新品牌.xlsx")
SHEET_NAME  = "可挖掘新品牌"
COLUMN_NAME = "brand"
OUTPUT_DIR  = "downloaded_pdfs"   # PDF 保存目录（相对路径或绝对路径）
```

然后运行：

```bash
python wipo_brand_spider.py
```

### 命令行参数

也可通过命令行参数覆盖默认配置：

```bash
python wipo_brand_spider.py \
  --excel  "/path/to/your/brands.xlsx" \
  --sheet  "可挖掘新品牌" \
  --column "brand" \
  --output "downloaded_pdfs" \
  --delay  2.0
```

| 参数         | 说明                        | 默认值                 |
|--------------|-----------------------------|------------------------|
| `--excel`    | Excel 文件路径              | 见代码配置常量         |
| `--sheet`    | Sheet 名称                  | `可挖掘新品牌`         |
| `--column`   | 品牌列名                    | `brand`                |
| `--output`   | PDF 保存根目录              | `downloaded_pdfs`      |
| `--progress` | 进度文件路径                | `progress.json`        |
| `--delay`    | 每次请求间隔（秒）          | `2.0`                  |

## 输出结构

```
downloaded_pdfs/
├── NIKE/
│   ├── NIKE_US_1234567_1234567.pdf
│   └── NIKE_EM_9876543_9876543.pdf
├── ADIDAS/
│   └── ADIDAS_DE_111111_111111.pdf
└── ...
progress.json       # 进度文件
wipo_spider.log     # 运行日志
```

## 速率限制与注意事项

- WIPO 对频繁请求有速率限制（HTTP 429），脚本会自动等待后重试。
- 默认每次请求间隔 **2 秒**，如遇频繁 429 可增大 `--delay` 参数。
- 脚本会在 `progress.json` 中记录已完成品牌，中断后重跑可自动续传。
- 运行日志保存在 `wipo_spider.log`，可用于排查失败品牌。
- 对于大批量品牌列表，建议在稳定网络环境下分批运行或挂机运行。

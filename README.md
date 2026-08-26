# ERP 实发数量与退货率同步

`erp_excel_sync.py` 从快麦 ERP 的“销售主题报表/按款”接口抓取数据，将：

- `actualSysConsignCount` 写入工作表 Y 列（8月实发）
- `customA1ED4F3EEFEF30DBB8E9A9A4823B79A3` 写入 AA 列（退货率）

脚本直接修改 XLSX 内部 XML，只重新压缩目标工作表，适合当前约 283MB、包含大量图片的文件。无需打开 WPS。

## 一键运行

1. 打开 `config.json`，修改最上面的日期：

```json
"start_date": "2026-08-01",
"end_date": "2026-08-24"
```

需要时也可修改 `input`、`output`、`skus_file` 等运行参数。配置中的相对路径以 `config.json` 所在目录为基准。

2. 双击 `run_sync.bat`。

脚本会读取 `config.json` 并开始同步，窗口会保留执行结果。没有设置 `ERP_COOKIE` 时，会提示粘贴浏览器登录后的 Cookie。

当前配置的重要参数：

- `input`：源 Excel 文件。
- `output`：生成的 Excel 文件，源文件不会被覆盖。
- `start_date` / `end_date`：订单创建时间范围，包含当天全天。
- `skus_file`：需要同步的货号清单；设为 `null` 时处理表内全部货号。
- `dry_run`：设为 `true` 时只抓取和生成审核报告，不生成 Excel。

## 命令行运行

使用 Python 3.10+，仅依赖标准库。先设置浏览器登录后的 Cookie：

```powershell
$env:ERP_COOKIE = '从浏览器请求中复制的 Cookie'
```

然后运行：

```powershell
python .\erp_excel_sync.py `
  --input "$env:USERPROFILE\Desktop\26年8月分级总表.xlsx" `
  --output ".\outputs\26年8月分级总表_API同步.xlsx" `
  --start 2026-08-01 `
  --end 2026-08-24 `
  --skus-file .\target_skus.txt
```

脚本默认读取同目录的 `config.json`。命令行参数优先级更高，可临时覆盖配置而不修改文件：

```powershell
python .\erp_excel_sync.py --start 2026-08-01 --end 2026-08-25
```

也可以指定另一份配置：

```powershell
python .\erp_excel_sync.py --config .\config.json
```

未设置 `ERP_COOKIE` 时，脚本会安全提示输入。关闭网页不会中断脚本；Cookie 过期后需要重新登录并复制。

日期筛选使用订单“创建时间”。即使开始/结束日期不变，订单后续实发、退款和售后状态变化仍会让历史区间的接口结果发生变化，因此每次运行都会保存快照并生成字段差异清单。

## 匹配与审核

- 精确货号、`sku_aliases.json` 中已确认的别名：自动写入。
- 只有轻微差异的模糊匹配：默认不写入，列入 `reports/manual_review.csv`。
- 同一货号有多条记录时，使用商品名称消歧；仍不确定则人工审核。
- `reports/sync_changes.csv`：Excel 原值、新值和发生变化的列。
- `reports/api_changes.csv`：本次 API 与上次快照发生变化的字段。
- `reports/summary.json`：执行汇总。

审核确认后，将映射加入 `sku_aliases.json`，下次即可自动写入。

只检查、不生成 Excel：

```powershell
python .\erp_excel_sync.py --input "原表.xlsx" --start 2026-08-01 --end 2026-08-24 --dry-run
```

使用已保存的接口响应进行离线测试：

```powershell
python .\erp_excel_sync.py --input "原表.xlsx" --output "结果.xlsx" `
  --start 2026-08-01 --end 2026-08-24 --api-json .\erp_dimensions.json
```

## 测试

```powershell
python -m unittest tests.test_erp_excel_sync -v
```

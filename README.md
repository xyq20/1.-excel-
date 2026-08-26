# ERP 月度实发数量与退货率同步

`erp_excel_sync.py` 根据运行日期自动选择每月1日或15日节点，分别请求ERP的实发数量和退货率，全量更新“分级总表”。脚本直接补丁 XLSX 内部 XML，不需要打开 WPS，并保留大量图片和媒体文件。

## 日期与列滚动

- 每月1—14日运行：归属本月1日节点。实发取上月全月；退货率取上上月15日至上月15日。不插列，并把 `8月实发（8.15）` 改为 `8月实发`。
- 每月15日至月底运行：归属本月15日节点。实发取本月1—14日；退货率取上月全月。首次插入本月列，标题如 `9月实发（9.15）`。
- 展示顺序始终为：`上月实发 | 同期销量 | 本月实发 | 变化情况 | 退货率`。旧月份实发列保留但隐藏。
- 同期销量为“上月实发 ÷ 上月天数 × 14”；变化情况公式随列位置和月份自动更新。
- 同一节点可重复运行：只刷新数据、公式和报告，不重复新增列。

## 使用

1. 确认 `config.json` 中的 `workbook` 指向要长期维护的主工作簿。相对路径以配置文件所在目录为基准。
2. 关闭该工作簿的 WPS/Excel 编辑窗口，避免文件被锁定。
3. Mac 上双击 `一键同步.command`；Windows 上双击 `run_sync.bat`。Mac 首次如果拦截，请右键文件选择“打开”。

也可以用 Python 3.10+ 命令行运行：

```powershell
python .\erp_excel_sync.py --config .\config.json
```

未设置 `ERP_COOKIE` 时，脚本会提示粘贴ERP登录 Cookie。Cookie过期后重新登录并复制即可。

## 匹配、校验与保存

- 表内所有有货号的款式都会更新；`target_skus.txt` 只是关键货号校验清单，不是过滤器。
- 精确货号和 `sku_aliases.json` 中已确认的别名自动写入；模糊、歧义或缺失数据进入人工审核报告。
- 任一关键货号未同时匹配实发和退货率时，只产生报告，不替换主文件。
- 更新前先在主文件同目录生成临时候选文件，通过 ZIP CRC、工作表、行数、图片、表头和公式校验后才原子替换。
- 原主文件保留为同名 `.xlsx.bak`，只保留最近一份备份。

ERP快照保存在 `snapshots/`，报告保存在 `reports/节点日期/`。

## 安全预览与离线测试

只生成快照和报告，不替换 Excel：

```powershell
python .\erp_excel_sync.py --config .\config.json --dry-run
```

使用两份已保存的ERP响应离线验证：

```powershell
python .\erp_excel_sync.py --config .\config.json `
  --actual-json .\actual.json --return-json .\return.json --dry-run
```

## 测试

```powershell
python -m unittest discover -s tests -v
```

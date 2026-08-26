# ERP 月度实发数量与三类退货率同步

`erp_excel_sync.py` 根据运行日期自动选择每月1日或15日节点，分别请求ERP的实发数量、发货前退货率和总退货率，全量更新“分级总表”。脚本直接补丁 XLSX 内部 XML，不需要打开 WPS，并保留大量图片和媒体文件。

## 日期与列滚动

- 每月1—14日运行：归属本月1日节点。实发取上月全月；退货率取上上月15日至上月15日。不插列，并把 `8月实发（8.15）` 改为 `8月实发`。
- 每月15日至月底运行：归属本月15日节点。实发取本月1—14日；退货率取上月全月。首次插入本月列，标题如 `9月实发（9.15）`。
- 展示顺序始终为：`上月实发 | 同期销量 | 本月实发 | 变化情况 | 发货前退货率 | 发货后退货率 | 退货率`。旧月份实发列保留但隐藏。
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

Mac 会自动打开一个独立的 ERP Chrome 窗口。首次或ERP登录过期时，只需在网页中正常登录；程序会自动检测登录成功并从该浏览器会话请求两组ERP数据，不再要求复制或粘贴 Cookie。专用Chrome使用“恢复上次会话”保留登录状态：运行时自动打开，任务成功后自动关闭，下次运行无需重新登录；发生错误时保留窗口便于检查。

每次查询固定包含售后工单状态 `9、2、12`，并限定订单类型为“平台订单”。退货率按售后类型执行两次独立查询：发货前只选择“未发货仅退款”；总退货率选择“未发货仅退款、已发货仅退款、退货、拒收退货、档口退货、档口换货”，不选择“换货”。发货后退货率的类型尚未启用，该列保持空白。专用 Chrome 启动后只保留一个 ERP 页签，避免恢复会话时重复打开窗口或页签。

日常只查看 ERP 时，双击 `查看ERP.command`。它与一键同步共用登录状态，但不会抓取 API 或修改 Excel；查看结束后正常关闭这个专用 Chrome 即可。

## 匹配、校验与保存

- 表内所有有货号的款式都会更新；`target_skus.txt` 只是关键货号校验清单，不是过滤器。
- 精确货号和 `sku_aliases.json` 中已确认的别名自动写入；模糊、歧义或缺失数据进入人工审核报告。
- 每种退货率分别使用自己查询结果中的销售数量 `itemCount` 判断门槛；销售数量大于等于 50 时，按 `rawRefundMoney ÷ saleMoney` 计算并保留两位小数，否则清空对应单元格。
- 任一关键货号未同时匹配实发、发货前和总退货率查询时，只产生报告，不替换主文件；已匹配但销售数量不足 50 属于正常留空，不阻止更新。
- 更新前先在主文件同目录生成临时候选文件，通过 ZIP CRC、工作表、行数、图片、表头和公式校验后才原子替换。
- 原主文件保留为同名 `.xlsx.bak`，只保留最近一份备份。

ERP快照保存在 `snapshots/`，报告保存在 `reports/节点日期/`。

## 安全预览与离线测试

只生成快照和报告，不替换 Excel：

```powershell
python .\erp_excel_sync.py --config .\config.json --dry-run
```

使用三份已保存的ERP响应离线验证：

```powershell
python .\erp_excel_sync.py --config .\config.json `
  --actual-json .\actual.json `
  --before-return-json .\return-before.json `
  --return-json .\return-overall.json `
  --dry-run
```

三种退货率使用同一个日期窗口，快照分别保存为 `return_before_*.json` 和 `return_overall_*.json`；变化报告分别保存为 `before_return_changes.csv` 和 `overall_return_changes.csv`。

## 测试

```powershell
python -m unittest discover -s tests -v
```

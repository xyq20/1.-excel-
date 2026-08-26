# Mac 一键同步启动器设计

## 目标

在项目根目录新增一个可在 macOS Finder 中直接双击的 `.command` 文件，用于运行现有 ERP 月度 Excel 同步程序。

## 行为

- 启动器先进入自身所在的项目目录，不依赖 Finder 当前路径。
- 依次尝试 `python3.12`、`python3.11`、`python3.10` 和 `python3`，选择第一个版本不低于 3.10 的 Python。
- 运行 `erp_excel_sync.py --config config.json`，不更改配置、Cookie 或 Excel 文件路径。
- 运行前显示配置文件位置；运行后显示成功或失败及退出码。
- 无论成功还是失败，都等待用户按回车后关闭终端窗口，便于查看结果。
- 未设置 ERP Cookie 时，交由现有 Python 程序提示输入；启动器不保存敏感信息。
- 保留现有 Windows `run_sync.bat`，不改变其行为。

## 错误处理

- 找不到 Python 3.10+ 时，显示明确的安装提示，不运行同步程序。
- `config.json` 或 `erp_excel_sync.py` 缺失时，显示缺失文件并以失败退出。
- Python 程序的退出码原样传递给启动器，方便区分成功、关键货号校验失败和其他错误。

## 验证

- 检查文件具有可执行权限。
- 用临时的假 Python 命令验证脚本路径、配置路径和退出码传递，不调用 ERP，不修改 Excel。
- 对启动器执行 shell 语法检查。

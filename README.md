# ShortcutCleaner - 快捷方式清理工具

一款 Windows 桌面工具，用于扫描并清理失效的快捷方式（`.lnk` / `.url`）。

桌面上的快捷方式越来越多？卸载软件、删除游戏、拔掉 U 盘之后，留下的一堆打不开的快捷方式，这个工具帮你一次找出来并安全清理（删除走回收站，随时可还原）。

## 功能特性

### 多位置扫描
- 用户桌面 / 公共桌面 / 开始菜单 / 公共开始菜单 / 快速启动栏 / 任意自定义目录
- 支持包含子目录递归扫描

### 智能失效检测

**`.lnk` 快捷方式**（通过 Windows COM 解析真实目标）：

| 状态 | 说明 |
|---|---|
| 目标存在 | 快捷方式正常 |
| 目标丢失 | 目标文件已被删除，可清理 |
| 驱动器不存在 | 指向 U 盘 / 已卸载分区 |
| 网络目标(未验证) | UNC 路径，默认不探测避免卡顿 |
| 无目标路径 | 命令型 / UWP 应用入口，不算失效 |

**`.url` 网址快捷方式**：

| 状态 | 说明 |
|---|---|
| 可访问 | HTTP 探测正常（HEAD → GET 自动回退） |
| 无法访问 | 返回 4xx/5xx 或连接失败，可清理 |
| 站点启用防护 | WAF/反爬拦截，无法程序化判定 |
| 本地协议已注册 | `steam://` 等自定义协议已关联程序 |
| Steam 游戏未安装 | 读取注册表判断 Steam 游戏是否已安装 |

### 灵活的清理策略
- **异常 .url 处理**三选一：保留 / 按失效项一起删除 / 移动到指定文件夹
- 一键「全选异常项」，删除时分流执行
- 删除均移入**回收站**，可随时还原
- 移动到文件夹时自动处理同名冲突

### 其他
- 后台线程扫描，界面不卡顿，可随时停止
- 右键菜单：打开所在位置（资源管理器定位并选中）/ 打开目标位置 / 单项重新检测
- 内置日志系统（自动轮转），界面一键查看
- 结果表格支持排序、筛选（全部 / 仅失效 / 仅有效）

## 界面预览

> 在这里放一张软件截图（建议 `docs/screenshot.png`）

## 安装使用

### 方式一：直接下载 exe（推荐）

前往 [Releases](../../releases) 下载 `ShortcutCleaner.exe`，双击即可运行，无需安装 Python。

日志文件位于 exe 同目录的 `logs\shortcut_cleaner.log`。

### 方式二：源码运行

```bash
pip install -r requirements.txt
python shortcut_cleaner.py
```

## 自行打包

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name ShortcutCleaner shortcut_cleaner.py
```

打包产物在 `dist\ShortcutCleaner.exe`。

## 依赖

| 包 | 用途 |
|---|---|
| PyQt5 | 图形界面 |
| pywin32 | COM 解析 .lnk 目标 |
| requests | .url 在线校验（可选） |

## 技术说明

- `.lnk` 通过 `WScript.Shell` COM 接口解析真实目标路径
- `steam://rungameid/<AppID>` 通过注册表 `HKCU\Software\Valve\Steam\Apps\<AppID>` 判断游戏安装状态
- 删除使用 `SHFileOperationW`（`FOF_ALLOWUNDO`），等价于资源管理器删除到回收站
- 仅支持 Windows

## 许可证

[MIT License](LICENSE)

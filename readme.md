# fork-WeChatMsg 使用说明

> 本项目仅供个人学习研究使用，请勿用于非法用途。

---

## 1. 项目简介

fork-WeChatMsg 是一个**微信 PC 端聊天记录导出与分析工具**，它可以：

- 从微信进程内存中读取数据库解密密钥
- 解密微信本地 SQLite 加密数据库（MSG.db、MicroMsg.db 等）
- 在图形界面中浏览联系人与聊天记录
- 生成聊天记录报告、词云、年度总结、情感分析等可视化内容
- 将聊天记录导出为 HTML / Word / CSV 等格式

本项目有两个 GUI 入口：

| 入口 | 功能 |
|---|---|
| `python main.py` | 主程序：解密 + 聊天/联系人/分析/导出完整界面 |
| `python decrypt_window.py` | 独立解密工具：只负责解密数据库 |

---

## 2. 环境要求

| 项目 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（必须 Windows，Linux/mac 不可用） |
| Python | 3.8 ～ 3.12 均可；推荐 **3.10 / 3.11**（PyQt5 在 3.13+ 下需要新版 wheel） |
| 微信 PC 端 | **3.2.1 ~ 3.9.8.15**（见第 3 节版本限制） |
| 权限 | 普通用户即可；不需要管理员 |

> Python 3.13 也可运行，但必须使用本仓库当前修改过的 `requirements.txt`（使用 `numpy>=1.26`、`pandas>=2.0` 等兼容 3.13 的版本）。

---

## 3. 微信版本限制 ⚠️（重要）

### 3.1 支持的微信 PC 版本

工具通过**读取 `WeChatWin.dll` 中的固定内存偏移量**来获取密钥。微信每次更新后偏移量都会变化，因此只支持以下版本：

```
3.2.1.154
3.3.0.84   3.3.0.93   3.3.0.115
3.3.5.34   3.3.5.42   3.3.5.46
3.4.0.37   3.4.0.38   3.4.0.50   3.4.0.54
3.4.5.27   3.4.5.45
3.5.0.20   3.5.0.29   3.5.0.33   3.5.0.39   3.5.0.42   3.5.0.44   3.5.0.46
3.6.0.18
3.6.5.7    3.6.5.16
3.7.0.26   3.7.0.29   3.7.0.30
3.7.5.11   3.7.5.23   3.7.5.27   3.7.5.31
3.7.6.24   3.7.6.29   3.7.6.44
3.8.0.31   3.8.0.33   3.8.0.41
3.8.1.26
3.9.0.28
3.9.2.23   3.9.2.26
3.9.5.81   3.9.5.91
3.9.6.19   3.9.6.33
3.9.7.15   3.9.7.25   3.9.7.29
3.9.8.15  ← 推荐版本（最新支持）
```

### 3.2 不支持的版本

- **微信 4.0.x / 4.1.x 及以上的新版**：对数据库加密方式与 `WeChatWin.dll` 内存布局做了改动，原项目维护的偏移量列表中**没有对应记录**。
- 微信 3.1 及以下：太旧，不在列表中。
- 微信 UWP / 应用商店版本 / 小程序版 `WeChatAppEx.exe`：非主程序，不识别。

如果你当前的微信是 4.1.x，运行解密会得到如下错误：

```
[-] WeChat Current Version 4.1.10.29 Is Not Supported
```

### 3.3 如何降级到支持的版本

1. 先通过控制面板卸载当前微信（**聊天记录不会被删除**，只卸载程序）。
2. 搜索下载 `微信 PC 版 3.9.8.15` 安装包（官方历史版本或可信的第三方下载站）。
3. 安装后登录微信，让它把消息同步到本地。
4. 退出微信后再执行解密工具（确保 `WeChat.exe` 或 `Weixin.exe` 进程存在即可）。

### 3.4 自行添加偏移量（进阶）

如果你有逆向基础，可以用 x64dbg / Cheat Engine 在 `WeChatWin.dll` 中定位：

- `name` 偏移（账号昵称）
- `account` 偏移（账号字符串）
- `mobile` 偏移（手机号）
- `mail` 偏移（邮箱）
- `key` 偏移（数据库密钥指针地址）

然后在以下两个文件中各加一行：

- `app/decrypt/version_list.json`
- `app/resources/version_list.json`

```json
"4.1.10.29": [NAME_OFFSET, ACCOUNT_OFFSET, MOBILE_OFFSET, MAIL_OFFSET, KEY_OFFSET]
```

偏移量必须使用**十进制**整数；`0` 代表该字段在当前版本中无法获取。

---

## 4. 安装步骤

### 4.1 克隆项目

```bash
git clone https://github.com/yourname/fork-WeChatMsg.git
cd fork-WeChatMsg
```

### 4.2 安装 Python 依赖

在项目根目录执行：

```bash
pip install -r requirements.txt
```

核心依赖说明：

| 依赖 | 作用 |
|---|---|
| `PyQt5` / `PyQtWebEngine` | GUI 与内嵌浏览器 |
| `pandas` / `numpy` | 数据处理、统计、词频分析 |
| `jieba` / `snownlp` | 中文分词与情感分析 |
| `python-docx` / `docxcompose` | Word 文档导出 |
| `pyecharts` | ECharts 图表（词云、年度报告等） |
| `psutil` / `pymem` | 读取微信进程、读取内存偏移 |
| `pycryptodomex` | 数据库解密算法支持 |
| `pywin32` | Windows API（COM、注册表等） |
| `Pillow` | 头像/图片处理 |
| `xmltodict` | XML 解析 |
| `beautifulsoup4` | HTML 导出辅助 |

### 4.3 验证安装

在项目根目录执行：

```bash
python -c "import PyQt5, pandas, jieba, psutil, pymem, docx, pyecharts; print('OK')"
```

输出 `OK` 即环境就绪。

---

## 5. 使用流程

### 5.1 准备工作

1. 登录并启动 **微信 PC 版**（确保进程 `WeChat.exe` 或 `Weixin.exe` 在运行）。
2. 让微信把消息同步到本地（点击某个会话、等待"加载中…"结束）。
3. 建议退出微信后再解密（避免数据库文件被占用导致解密失败）。

### 5.2 方法一：独立解密工具（推荐先跑这个）

```bash
python decrypt_window.py
```

界面出现后按顺序操作：

1. 点击 **"获取微信信息"** → 若成功，界面会显示 key / wxid / pid / 版本号。
2. 程序会自动探测本机微信数据目录（以 `wxid_xxx` 结尾的目录），或手动点 **"选择数据目录"**。
3. 点击 **"开始解密数据库"** → 进度条跑满，解密完成。
4. 解密后的数据库位于：

   ```
   app/DataBase/Msg/MSG.db
   app/DataBase/Msg/MicroMsg.db
   ...
   ```

然后运行主程序查看聊天记录：

```bash
python main.py
```

### 5.3 方法二：主程序一站式

```bash
python main.py
```

- 若数据库已解密：直接进入聊天/联系人界面。
- 若数据库未解密：先弹出解密窗口，流程与 5.2 相同。

### 5.4 主界面功能

进入主界面后，左侧图标栏依次为：

| 图标 | 功能 |
|---|---|
| 💬 聊天 | 查看好友聊天记录，支持搜索、导出 HTML/Word/CSV |
| 👥 联系人 | 管理好友列表、标签、公众号 |
| 👤 我的信息 | 当前登录账号信息 |
| 📊 数据分析 | 生成词云、聊天时段分布、好友活跃度排行 |
| 🎭 情感分析 | 对聊天内容做情感倾向（正/负面）分析与图表 |
| 📅 年度报告 | 类似"微信数据报告"的单年度可视化总结 |

---

## 6. 常见问题 FAQ

### Q1. 运行 `python decrypt_window.py` 提示 `[-] WeChat No Run`

**原因**：微信进程名匹配不到。有的安装包进程叫 `Weixin.exe`（例如 `D:\Weixin\Weixin.exe`），但原项目只识别 `WeChat.exe`。

**解决**：本仓库已在 `app/decrypt/get_wx_info.py` 中修复，同时识别 `WeChat.exe` 和 `Weixin.exe`。使用当前版本重跑即可。

### Q2. 运行解密提示 `WeChat Current Version x.x.x.x Is Not Supported`

**原因**：当前微信版本不在偏移量支持列表中。

**解决**：
- 推荐：降级到 `3.9.8.15`（见第 3.3 节）。
- 或：自行添加偏移量（见第 3.4 节）。

### Q3. `ModuleNotFoundError: No module named 'pymem' / 'xmltodict' / 'docx'`

**原因**：依赖没装全。

**解决**：

```bash
pip install -r requirements.txt
```

单装某一个也可以：

```bash
pip install pymem xmltodict python-docx
```

### Q4. `qt.svg: Cannot open file ./app/data/icons/logo.svg`

**原因**：原项目图标路径是 `app/data/icons/`，但仓库实际图标放在 `app/resources/icons/`。

**解决**：本仓库已在以下文件修复路径：

- `app/Ui/Icon.py`
- `app/Ui/contact/contactInfoUi.py` + `.ui`
- `app/ui_pc/contact/contactInfoUi.py` + `.ui`
- `app/Ui/decrypt/decrypt.py`
- `app/ImageBox/ui.py`
- `app/DataBase/data.py`
- `app/person.py`

拉取最新代码或重新执行本文件即可。

### Q5. `person.py: SyntaxWarning: invalid escape sequence '\P'`

**原因**：硬编码的 Windows 路径字符串中包含 `\P`、`\W` 被 Python 当成转义序列。

**解决**：已将 `app/person.py` 中两处硬编码路径统一改为通过 `Icon.Default_avatar_path` 获取，不再包含转义字符。

### Q6. 解密进度条跑到一半失败

可能的原因：

1. 微信正在运行且数据库文件被占用 → 先退出微信。
2. MSG.db 文件损坏 → 尝试让微信重新同步消息（切换会话、重启微信）。
3. 版本不匹配 → 再次确认你的微信版本。

### Q7. 解密成功后打开主界面空白

- 检查 `app/DataBase/Msg/` 下是否有 `MSG.db` 等文件（大小应大于 0）。
- 用任意 SQLite 工具（如 DB Browser for SQLite）打开 `MSG.db`，看是否能正常浏览表结构。
- 如果数据库为空/表异常，解密密钥不正确 → 重新检查微信版本与偏移量。

### Q8. 想换个账号 / 清除本地数据

直接删除以下目录（由解密工具生成）：

```
app/DataBase/Msg/
app/data/info.json
```

然后重新执行 `python decrypt_window.py`。

---

## 7. 已修复的问题清单

本 fork 在原项目基础上修复了以下问题：

| # | 问题 | 修复位置 |
|---|---|---|
| 1 | 图标路径 `app/data/icons/` 不存在，Qt 报 SVG 加载失败 | `app/Ui/Icon.py` 及 8 个引用文件 |
| 2 | `ModuleNotFoundError: No module named 'xmltodict'` | `requirements.txt` 中补充 |
| 3 | `ModuleNotFoundError: No module named 'docx'` | `requirements.txt` 中补充 `python-docx` |
| 4 | `get_wx_info.py` 顶部使用 `sys/os/winreg` 但在底部才 import | 把 import 移到文件顶部 |
| 5 | 微信进程名为 `Weixin.exe` 时提示 "No Run" | `read_info()` 同时识别 `WeChat.exe` 与 `Weixin.exe` |
| 6 | 获取信息失败后后续代码仍执行，抛 `string indices must be integers` | `pc_decrypt.py` 中对 result 做类型检查，失败时弹框并 return |
| 7 | `app/person.py` 中有硬编码 Windows 绝对路径，导致 `SyntaxWarning` 与路径错误 | 统一改为使用 `Icon.Default_avatar_path` |
| 8 | `requirements.txt` 锁定 `numpy 1.24.1`、`pandas 1.5.2` 等无法在 Python 3.13 下编译 | 放宽到 `numpy>=1.26`、`pandas>=2.0`、`PyQt5>=5.15` 等 |

---

## 8. 目录结构

```
fork-WeChatMsg/
├── main.py                    # 主程序（解密 + 聊天/分析）
├── decrypt_window.py          # 独立解密工具
├── requirements.txt           # Python 依赖
├── USAGE.md                   # 本文档
├── readme.md                  # 项目简介
└── app/
    ├── Ui/                    # 原 GUI（聊天、联系人、解密、分析）
    │   ├── Icon.py            # 图标路径集中管理
    │   ├── decrypt/
    │   ├── contact/
    │   └── chat/
    ├── ui_pc/                 # PC 版专用 GUI（pc_decrypt 等）
    ├── DataBase/              # 数据库读写与合并
    │   ├── Msg/               # ← 解密后 .db 生成在这里
    │   ├── data.py
    │   ├── micro_msg.py
    │   └── merge.py
    ├── decrypt/               # 核心解密 & 微信信息读取
    │   ├── decrypt.py
    │   ├── get_wx_info.py     # ← 从内存读取偏移量
    │   └── version_list.json  # ← 版本偏移量配置
    ├── resources/             # 图标与资源（qrc / svg / logo）
    ├── ImageBox/              # 图片预览组件
    ├── components/            # 公共组件（气泡、头像等）
    └── util/                  # 工具函数（路径、搜索等）
```

---

## 9. 免责声明

本项目仅用于**个人学习研究**，以及备份本人的聊天记录。使用本项目访问他人的微信数据或用于商业用途均违反相关法律法规，**一切后果由使用者自行承担**。

微信及其数据库是腾讯公司的产品，本项目与腾讯无关。请在法律允许的范围内使用。

---

## 10. 反馈与贡献

如在使用中发现其他问题，可在 issues 中反馈。若你能为新版微信（4.x）提供偏移量，欢迎提交 PR 更新 `version_list.json`。

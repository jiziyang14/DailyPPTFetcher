# 📥 DailyPPTFetcher

> 每天自动从 FTP 服务器下载当天 PPT 到桌面并自动打开。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue)](https://github.com/PowerShell/PowerShell)

[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6)](https://www.microsoft.com/windows)

## ✨ 功能特性

- 🧹 **自动清理**：删除桌面上超过 1 天的旧 PPT（文件名带时间戳）

- 🔍 **智能识别当天文件**：优先从文件名提取日期，备选 FTP 修改时间

- ⬇️ **自动下载**：从 FTP 服务器下载当天 PPT 到桌面，并添加日期戳（避免重名覆盖）

- 💻 **一键打开**：下载成功后自动最小化其他窗口，并打开所有 PPT，逐个全屏播放。

- 🔁 **多路径容错**：支持主路径 + 备用路径（允许为空）

- 🔒 **安全配置**：敏感信息（FTP 密码）通过本地 `config.json` 管理

## 🖥️ 系统要求

- Windows 操作系统（PowerShell 5.1 及以上，Windows 10/11 自带）

- 网络可访问目标 FTP 服务器

- 已安装 Microsoft PowerPoint 或其他可打开 `.ppt`/`.pptx` 的软件

## 📦 安装与配置

### 1. 获取代码

```bash
git clone https://github.com/jiziyang14/DailyPPTFetcher.git

cd DailyPPTFetcher
```

或者直接下载 ZIP 并解压。

### 2. 配置 FTP 连接

复制配置模板并填写真实信息

用文本编辑器（如记事本、VS Code）打开 config.json，按照提示填入：
```json
{

    "ftp_server": "你的网络地址",

    "ftp_username": "你的用户名",

    "ftp_password": "你的密码",

    "primary_remote_path": "/路径1/",

    "secondary_remote_path": "/路径2/"

}
```

⚠️ **注意**

* primary\_remote\_path 必须填写，secondary\_remote\_path 可选（不需要时留空 ""）
* 路径格式与 FTP 服务器上的目录结构一致，以 / 开头并以 / 结尾
* 保存文件编码推荐 **UTF-8**（无 BOM 也可以）

### 3. 运行脚本

* **方式一**：双击 run.bat（推荐，已绕过 PowerShell 执行策略）
* **方式二**：在 PowerShell 中手动执行（需管理员权限设置执行策略或加参数）：

```powershell
powershell -ExecutionPolicy Bypass -File DailyPPTFetcher.ps1
```

## 🗂️ 文件命名与识别规则

### 当天文件判断流程

1. 扫描 FTP 目录下所有 .ppt 或 .pptx 文件
2. 对每个文件：
   * 尝试从文件名中匹配日期（例如 3.21.pptx 或 03.21.pptx）
   * 若匹配成功，用该日期与当天比较
   * 若匹配失败，则获取 FTP 文件的 LastModified 时间与当天比较
3. 满足任一条件即视为当天文件

### 下载到本地的命名

* 格式：原文件名\_YYYYMMDD.pptx
  例如：早读\_20260521.pptx

### 过期清理规则

* 桌面上所有文件名包含 \_YYYYMMDD 格式的 PPT 文件
* 如果该日期早于 **当前日期 1 天**，则自动删除

## ❓ 常见问题

### 1. 提示“未找到配置文件 config.json”

**原因**：没有创建配置文件。
**解决**：将 config.json.example 复制为 config.json 并填写真实信息。

### 2. FTP 连接失败或列表为空

**可能原因**：

* 网络不通或防火墙阻挡
* FTP 地址、用户名、密码错误
* 远程路径不存在或无权限

**解决**：用 FileZilla 等 FTP 客户端测试连接，确认参数正确。

### 3. 下载后的 PPT 没有自动打开

**原因**：系统未关联 .ppt / .pptx 到任何程序。
**解决**：安装 Microsoft PowerPoint 或 WPS，或右键指定默认打开方式。

### 4. 执行脚本时提示“无法加载，因为在此系统上禁止运行脚本”

**原因**：PowerShell 执行策略限制。
**解决**：使用 run.bat 启动（已加 -ExecutionPolicy Bypass），或者管理员运行 PowerShell 并执行：

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 5. 中文显示为乱码或脚本报错

**原因**：脚本文件编码不是 UTF-8 with BOM（旧版 PowerShell 兼容问题）。

**解决**：用记事本打开 DailyPPTFetcher.ps1，另存为，编码选择 **UTF-8 with BOM**。

## 🤝 贡献

欢迎提交 Issue 或 Pull Request。如果你有改进建议，请先通过 Issue 讨论。

### 🛠️ 开发环境准备

* Windows + PowerShell 5.1+
* Git
* 任意文本编辑器（推荐 VS Code）

### 📄 许可证

本项目采用 **MIT 许可证**。
**保留原作者版权信息**：Copyright (c) 嵇子扬（Young）

### 📧 联系与支持

* 报告 Bug 或功能请求：请到 [GitHub Issues](https://github.com/jiziyang14/DailyPPTFetcher/issues)
* 如需私有定制或商业授权，请联系原作者

**如果这个工具对你有帮助，请给一个 ⭐ Star 支持一下！**

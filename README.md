# KMS-Activator (Windows & Office KMS 激活工具)

<p align="left">
  <img src="https://img.shields.io/badge/Platform-Windows%207%20~%2011%20%7C%20Server%202008%20R2%20~%202025-0078D6?logo=windows&logoColor=white" alt="Platform">
  <img src="https://img.shields.io/badge/Office-2010%20~%20LTSC%202024%20(VL)-D83B01" alt="Office">
  <img src="https://img.shields.io/badge/Script-Batch%20%2B%20PowerShell-5391FE?logo=powershell&logoColor=white" alt="PowerShell">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
</p>

单文件、零依赖、双击即用的 **Windows 与 Office KMS 自动激活工具**。打开后通过交互菜单或命令行参数即可选择激活 Windows 或 Office，工具会自动识别当前系统/Office 版本、匹配微软官方 [GVLK 密钥](https://learn.microsoft.com/zh-cn/windows-server/get-started/kms-client-activation-keys) 并连接 KMS 服务器完成激活。

---

## ✨ 功能特性

- **单文件免安装**：采用 `Batch + PowerShell` 混合脚本封装为单个 `KMS-Activator.cmd`，双击自动申请管理员权限运行。
- **智能版本识别**：自动读取系统 Edition / Build 号及本机已安装的 Office 批量授权（Click-to-Run / MSI）组件，精准匹配微软官方 GVLK（含 LTSC 2021 / 2024 独立密钥区分）。
- **安全保护机制**：检测到 Windows 已处于**永久激活**状态时默认跳过（除非显式指定 `/force`），避免误覆盖正版数字权利或 MAK 永久授权。
- **网络预检与诊断**：激活前自动对 KMS 主机端口（默认 `1688`）进行 TCP 连通性检测，并对常见错误码（如 `0xC004F074`、`0xC004F035` 等）提供中文排障说明。
- **无人值守支持**：完整支持命令行静默参数与标准退出码，方便系统封装、批量部署或运维脚本调用。

---

## 🚀 使用方法

### 1. 交互式菜单（推荐）

下载 [`KMS-Activator.cmd`](./KMS-Activator.cmd) 后直接双击运行，同意管理员权限（UAC）后进入主菜单：

```text
  ==================================================================
  |                KMS 激活工具  (Windows & Office)                |
  ==================================================================

      当前系统    : Windows 11 Pro [25H2 (26200.9974)]
      运行权限    : 管理员
      KMS 主机    : kms.cgtsoft.com:1688
  ------------------------------------------------------------------

      [1] 激活 Windows            自动识别系统版本并激活
      [2] 激活 Office             自动识别已安装的 Office 批量版并激活
      [3] 激活 Windows + Office   依次自动激活系统和 Office

      [4] 查看激活状态            查看 Windows 与 Office 详细授权状态
      [5] 仅检测匹配结果          测试版本识别与 GVLK 匹配，不做任何修改
      [6] 切换 KMS 服务器         临时更换本次会话使用的 KMS 主机地址

      [0] 退出程序
```

### 2. 命令行参数（适合无人值守 / 批量部署）

| 参数 | 说明 |
| :--- | :--- |
| （无参数） | 打开交互式主菜单 |
| `/windows` 或 `/win` | 自动识别并激活 Windows |
| `/office` | 自动识别并激活 Office |
| `/all` | 依次自动激活 Windows 和 Office |
| `/check` | 仅检测环境与匹配的 GVLK，不做任何修改（无需管理员权限） |
| `/status` | 查看当前 Windows 与 Office 激活状态 |
| `/server:主机[:端口]` | 指定 KMS 服务器（默认 `kms.cgtsoft.com:1688`） |
| `/silent` 或 `/s` | 静默运行：不询问、结束不暂停（未指定目标时默认 `/all`） |
| `/force` 或 `/f` | 即使 Windows 已永久激活也强制替换为 KMS |
| `/help` 或 `/?` | 显示命令行帮助 |

**调用示例：**

```cmd
:: 静默激活已安装的 Office
KMS-Activator.cmd /office /silent

:: 指定自定义 KMS 服务器并同时激活 Windows 与 Office
KMS-Activator.cmd /all /server:kms.cgtsoft.com:1688

:: 仅检测当前系统与 Office 匹配的官方 GVLK（不修改系统）
KMS-Activator.cmd /check
```

---

## 💻 支持范围

- **Windows 11 / 10**：专业版（Pro）、专业版 N、专业工作站版、专业教育版、教育版、企业版（含 N/G/GN、多会话）
- **Windows 10 / 11 LTSC**：LTSB 2015、LTSB 2016、LTSC 2019 / 2021 / 2024、IoT 企业版 LTSC（含 2021 / 2024 独立密钥匹配）
- **Windows 8.1 / 8 / 7**：专业版、企业版（含 N/E）
- **Windows Server**：2008 R2 / 2012 / 2012 R2 / 2016 / 2019 / 2022 / 2025（含 Essentials、Datacenter Azure Edition、半年频道 SAC）
- **Office 批量授权版（VL）**：Office LTSC 2024 / LTSC 2021 / 2019 / 2016 / 2013 / 2010（ProPlus、Standard、Project Pro/Std、Visio Prem/Pro/Std 及各独立组件，同时支持 Click-to-Run 与 MSI）

---

## 📊 退出码说明

| 退出码 | 含义 |
| :---: | :--- |
| `0` | 激活或检测成功 |
| `1` | 激活过程出现错误或权限不足 |
| `2` | 当前系统/Office 版本不支持 KMS（如家庭版、零售版）或未安装 |
| `3` | 系统为评估版（Evaluation），需先转换版本 |
| `4` | KMS 服务器网络不可达 |
| `5` | 用户取消操作 |

---

## 📄 开源协议

本项目基于 [MIT License](./LICENSE) 开源，所有内置产品密钥均来自微软官方文档公开的 KMS 客户端安装密钥（GVLK）。

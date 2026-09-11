---
name: vb6-on-win10-troubleshoot
description: VB6.0 在 Win10/Win11 启动失败 / 兼容性故障排查。当用户报告 VB6 类程序启动报错（如"VB6EXT.OLB 不能被加载"、意外错误退出、加载项失败、ActiveX 组件错误等）、VB6 工程打不开、VBA 运行时异常、或任何与 Visual Basic 6.0 / Visual Studio 6.0 相关的兼容性问题时使用。触发词：VB6 不能启动、VB6 启动报错、VB6EXT.OLB 不能被加载、VB6 兼容、Win10 运行 VB6、VB6 错误退出、VB6 启动失败、意外错误退出、VBA6 错误、VB6 工程打不开。
agent_created: true
version: 1.0.0
author: 天工创新坊
license: CC BY 4.0
display_name: "VB6 兼容故障排查"
display_name_en: VB6 on Win10/11 Troubleshoot
trigger: ["VB6 不能启动", "VB6 启动报错", "VB6EXT.OLB 不能被加载", "VB6 兼容", "VB6 错误退出"]
description_zh: "VB6 在 Win10/Win11 启动失败 / 兼容故障排查（VB6EXT.OLB 加载错误等）"
description_en: "VB6 startup failure / Win10-11 compatibility troubleshoot"
category: development
---

# VB6 on Win10 / Win11 兼容性故障排查

**根本方法论（铁律）**：普遍问题必有现成解。**第一动作是 web search 报错原文（中英文并行）**，不要写工具 / 抓包 / 查注册表。本技能给出"先搜哪些关键词 → 搜不到才查什么"的高效路径。

---

## 第一步：web search 报错（必做，别跳过）

把错误信息和软件名拼成检索式，**中英文并行搜**：

| 检索式模板 | 说明 |
|---|---|
| `"VB6EXT.OLB 不能被加载" Win10` | 中文核心报错 |
| `"VB6" "意外错误" 退出 解决` | 中文通用报错 |
| `"VB6EXT.OLB could not be loaded" Windows 10 fix` | 英文核心报错 |
| `"Visual Basic 6" "unexpected error" quitting Windows 10` | 英文通用报错 |
| `"vb6.exe" "compatibility" site:superuser.com OR site:stackoverflow.com` | 英文定向 |
| `{显卡型号} VB6 启动错误` | 若用户提到显卡相关（RTX 30/40 魔改卡在 OGL 老路径上有案例） |

**命中前 3 个高赞结果的关键解法 → 直接执行**，不要自己原创修复方案。实测教训：盲目原创分析耗 8 小时，网络搜索现成解只要 10 分钟。

---

## 第二步：web search 指向命中以下任一条 → 按模板直接修

### 高频原因 1（命中率高）：AppCompatFlags 兼容层

`HKCU\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers\<VB6.exe 全路径>` 的值含 `WINXPSP3` / `VISTASP2` / `WIN7RTM` 等兼容模式字符串 → XP/Win7 兼容垫片（shim）干扰 OLB 加载。

```bat
:: 查询当前值
reg query "HKCU\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" /v "D:\Program Files (x86)\Microsoft Visual Studio\VB98\Vb6.exe"

:: 修复：去掉 XP 兼容，保留管理员运行
reg add "HKCU\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" /v "D:\Program Files (x86)\Microsoft Visual Studio\VB98\Vb6.exe" /t REG_SZ /d "~ RUNASADMIN" /f

:: HKLM 路径也要查
reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" /v "D:\Program Files (x86)\Microsoft Visual Studio\VB98\Vb6.exe"
```

**procmon 抓包里出现 AcGenral.dll / AcLayers.dll / shimeng.dll → 必中此条**。

### 高频原因 2：VBADDIN.INI 损坏加载项

```bat
type C:\Windows\VBADDIN.INI
```

`FixItAddin.Connect=1` / `vbscc=1` / `Code Advisor` / 任何 .NET 依赖的加载项 → 改成 `=0`，先备份：

```bat
copy C:\Windows\VBADDIN.INI C:\Windows\VBADDIN.INI.bak
```

### 中频原因 3：TypeLib 注册表 / 文件损坏

正确的 VB6EXT.OLB GUID = `{EF404E00-EDA6-101A-8DAF-00DD010F7EBB}` 版本 **5.3**。

**别记错**：常见的错误 GUID `{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}` 是 Internet Controls（ieframe.dll），**不是** VB6EXT.OLB 的。

检查项：`HKCR\TypeLib\{EF404E00-EDA6-101A-8DAF-00DD010F7EBB}\5.3\0\win32` 默认值必须指向 `…\VBA\VB6EXT.OLB`。

### 中频原因 4：HKCU 配置损坏

`HKCU\Software\Microsoft\Visual Basic\6.0` 整树 → 备份后整树删除，VB6 首次启动会自动重建。

### 低频原因 5：第三方进程注入冲突

- 搜狗 / QQ / 微软拼音等输入法
- WPS / 金山毒霸 / 360 等国产安全软件（事件日志会留 MsiInstaller 撞车记录）
- 安全软件 HIPS / 沙箱

近期装/更新过的，卸载一个重试一次。

---

## 第三步：诊断工具（**仅在 web search 没命中时**用）

按需，不要一上来全跑：

1. **procmon**（Sysinternals）抓 VB6.exe 启动 → 看 `AcGenral.dll` 是否加载、`msvbvm60.dll` 是否加载、注册表访问哪条失败
2. **tlibdump**（自制 C# x86）：导出 OLB 内部接口验证文件没坏
3. **tlibtest**（自制 C# x86）：调 `LoadRegTypeLib` 验证 TypeLib 系统本身正常
4. **Event Viewer** → Application → 过滤 MsiInstaller / .NET Runtime / Application Error 时间窗

参考实现：`tlibtest.cs`、`tlibdump.cs`（自制 C# x86 验证工具）、`runtlibtest.bat`、`evtcheck.bat`。

---

## 修复通用要点

- VB6 是 1998 程序，**不要"以管理员身份运行"** —— UAC 给老程序带来副作用，**双击普通启动**即可（若当前账户已是管理员）
- **不要勾"以 XP SP3 兼容模式运行"** —— 干扰 OLB 加载，与原因 1 撞车
- 装在 `D:\Program Files (x86)\Microsoft Visual Studio\VB98`，**不要默认装 C 盘**

---

## 模型选择（按"国内外大模型分工铁律"）

本技能被调用时：
- **默认先网络搜索**，不下结论前不要追问模型
- 若用户明确"让智能体帮忙分析"，**优先国外模型路线**（CODEX / Claude / GPT / Gemini / Hy3），本地免费国内模型不擅长微软技术栈
- 不要"为省钱用错模型硬上"——微软技术问题用 qwen = 浪费时间和算力还查不到

---

## 沉淀参考

- 配套技能：troubleshoot-first-rule（排错方法论第一铁律）、model-routing-guide（国内外大模型分工）

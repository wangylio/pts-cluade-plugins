---
name: wps
description: 读取 %TSD-Header 加密的 Office 文档（doc/docx/xls/xlsx/ppt/pptx）。当 zipfile、python-docx、openpyxl 报 "not a zip file" / BadZipFile，或文件头为 %TSD-Header-###%，或需要提取公司加密系统保护的文档正文时使用。
---

# TSD 加密文档读取

## 适用判断

文件头是 `%TSD-Header`，完整形式为 `%TSD-Header-###%`，后面紧跟一个 GUID 和整段密文。识别时读文件前 11 个字节比对 `%TSD-Header` 即可，不靠扩展名判断——扩展名是 .docx/.xlsx 的文件同样可能是 TSD 加密的。

再确认全文搜不到 `PK\x03\x04`、`\xd0\xcf\x11\xe0`、`<w:document`——三者都没有，说明是整个文件加密，不是套壳。

此时纯字节读取必然失败，不要换库浪费时间：`zipfile`、`python-docx`、`openpyxl`、`pandoc` 拿到的都是密文。

## 原理

公司加密驱动按**进程白名单**透明解密。让被放行的办公软件去打开文件，读它内存里已经是明文的内容即可。

本机已实测：WPS Office 12.1（`D:\Program Files\Kingsoft\WPS Office\12.1.0.29161`）在白名单内，ProgID 为：

| 应用 | ProgID |
|------|--------|
| WPS 文字 | `KWps.Application` |
| WPS 表格 | `Ket.Application` |
| WPS 演示 | `KWPP.Application` |

注意：本机 `Word.Application` 这个 ProgID 已被 WPS 接管，`New-Object -ComObject Word.Application` 建出来的对象 `Path` 指向 WPS。微软 Word 处于未激活状态且不在白名单，直接用 `WINWORD.EXE` 打不开这类文件——**不要拿 `Word.Application` 能不能创建来判断"Word 可用"**。

## 用法

脚本位于本技能目录的 `scripts/read_tsd.ps1`，用绝对路径调用。

列出目录下所有 TSD 文档（递归，按文件头识别，不靠扩展名）：

```
powershell -NoProfile -ExecutionPolicy Bypass -File "<skill目录>/scripts/read_tsd.ps1" -Root "E:\path\to\folder" -List
```

读取第 N 个：

```
powershell -NoProfile -ExecutionPolicy Bypass -File "<skill目录>/scripts/read_tsd.ps1" -Root "E:\path\to\folder" -Index 3 -OutDir "C:\Users\<用户>\AppData\Local\Temp\tsd_read"
```

输出 UTF-8 txt（`tsd_<N>.txt`），再用 Read 工具读。多个文件就逐个传 `-Index`，每个都起新进程。

全部读完后清理临时输出（默认目录 `%TEMP%\tsd_read`）：

```
powershell -NoProfile -ExecutionPolicy Bypass -File "<skill目录>/scripts/read_tsd.ps1" -Cleanup
```

`-Cleanup` 会删掉 `-OutDir` 整个目录；如果之前用过自定义 `-OutDir`，清理时用同一个值传进去。

## 硬性约束

1. **一个文件一个应用实例**。同一实例连续打开多个加密文件，第二个起会把进程搞崩（`RPC 服务器不可用 0x800706BA`），之后全部报"未将对象引用设置到对象的实例"。这是最容易踩的坑。
2. **脚本保持纯 ASCII**。PowerShell 5.1 把无 BOM 的 .ps1 按 ANSI 解析，脚本里写中文会乱码；中文路径交给脚本内部 `Get-ChildItem` 递归枚举，不要在命令行里拼。
3. **多维数组索引要加括号**：`$arr[($r-1),($c-1)]`。写成 `$arr[$r-1,$c-1]` 会被解析成数组减法，报 `op_Subtraction`。
4. **Excel 取 `Value2`**。WPS 的 COM 上 `Range.Text`、`Text2()` 会返回 `DBNull`。
5. 一律只读打开，不保存、不回写原文件。
6. 单文件卡住时 `taskkill /F /IM wps.exe`（表格是 `et.exe`）后重试。
7. **用完即删**。输出的是解密后的明文，任务结束（含中途放弃）都要执行一次 `-Cleanup`，否则 `%TEMP%\tsd_read` 会一直堆积明文文件。

## 边界

- 只处理使用者有权访问的文档。本技能是"用白名单程序打开"，不是破解加密。
- 不要尝试直接解密字节，不要调用任何解密/导出接口。
- 读不出来就如实报告，不要按文件名猜测内容。
- 提取结果是纯文本：表格结构、图片、页眉页脚会丢失；Excel 日期可能显示为序列号。需要精确版式时人工核对原文件。

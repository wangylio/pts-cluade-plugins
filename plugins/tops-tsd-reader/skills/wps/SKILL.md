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

## 输出格式

每个 txt 第一行是 `=== SOURCE: <路径> (<扩展名>) ===`，之后按文件类型分节，正文顺序与原文档一致。

**Word** 正文按文档顺序输出，再跟两个附加节：

```
=== OUTLINE ===                                  标题大纲，按 OutlineLevel 缩进
=== HEADER / FOOTER / FOOTNOTE / COMMENT ===     页眉、页脚、脚注、尾注、批注
```

正文里的表格转成 TSV：单元格间制表符、行间换行；`[PAGE BREAK]` 标记分页符；`[OBJECT]` 标记内嵌对象（图片、图表、公式等只有占位，没有内容）。

**Excel** 每个工作表一节：

```
=== SHEET: 名称 (hidden) ===     仅隐藏/深度隐藏表带标记，仍会写出
```

单元格之间制表符。日期按单元格的 `NumberFormat` 还原成 `yyyy-MM-dd HH:mm:ss`（纯时间格式输出 `HH:mm:ss`），非日期数字保留原值。含制表符、换行或引号的单元格按 TSV 规范加双引号包裹、内部引号翻倍。

**PowerPoint** 每页一节：

```
=== SLIDE N ===
[TABLE] ... [/TABLE]     表格转 TSV
[NOTES] ...              演讲者备注
```

组合形状（Group）会递归展开，标题、正文、文本框按形状顺序输出。

## 硬性约束

1. **一个文件一个应用实例**。同一实例连续打开多个加密文件，第二个起会把进程搞崩（`RPC 服务器不可用 0x800706BA`），之后全部报"未将对象引用设置到对象的实例"。这是最容易踩的坑。
2. **脚本保持纯 ASCII**。PowerShell 5.1 把无 BOM 的 .ps1 按 ANSI 解析，脚本里写中文会乱码；中文路径交给脚本内部 `Get-ChildItem` 递归枚举，不要在命令行里拼。
3. **WPS 返回的数组下界是 1**。`Range.Value2`、`Range.NumberFormat` 回来的是下界为 1 的 `Object[,]`，用 PowerShell 的 `$arr[$r-1,$c-1]` 会整体错位（还会把索引算成数组减法报 `op_Subtraction`）。脚本用 `GetValue()` 配合 `GetLowerBound()` 取值，不要改回中括号索引。
4. **Excel 取 `Value2`，日期靠 `NumberFormat`**。WPS 的 COM 上 `Range.Text`、`Text2()` 返回 `DBNull`；混合格式区域的 `Range.NumberFormat` 整体读回来是 `null`，必须退回逐格读 `Cells.Item($r,$c).NumberFormat`，否则日期会以序列号输出。
5. 一律只读打开，不保存、不回写原文件。
6. 单文件卡住时 `taskkill /F /IM wps.exe`（表格是 `et.exe`）后重试。
7. **用完即删**。输出的是解密后的明文，任务结束（含中途放弃）都要执行一次 `-Cleanup`，否则 `%TEMP%\tsd_read` 会一直堆积明文文件。

## 边界

- 只处理使用者有权访问的文档。本技能是"用白名单程序打开"，不是破解加密。
- 不要尝试直接解密字节，不要调用任何解密/导出接口。
- 读不出来就如实报告，不要按文件名猜测内容。
- 仍然是**文本提取**，精度到"结构"而不是"版式"。已保留：段落顺序、Word 表格的行列、标题层级、页眉页脚、脚注尾注、批注、Excel 的隐藏表与日期、PPT 的组合形状 / 表格 / 备注。
- 仍会丢失：图片、图表、形状里的公式与 SmartArt 内容（只有 `[OBJECT]` 占位）、Word 的合并单元格边界与单元格内换行、Excel 公式文本（只给计算结果）与单元格批注、PPT 的动画与母版元素。需要精确版式时人工核对原文件。

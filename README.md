# Word–MathType 公式工作流

通过 Word 的 MathType 插件，将文档中明确指定的 TeX 占位符转换为可在 MathType 中双击编辑的公式对象。脚本在独立的隐藏 Word 实例中逐条执行转换，尽量保留输入文档中的段落、编号和引用。

**Windows desktop Word + MathType · PowerShell 7 · 本地测试阶段**

This project converts explicitly marked TeX placeholders in a Word document into editable MathType OLE equations. It automates desktop Word; it is not a standalone TeX renderer or a server-side headless converter.

![虚构示例：行内公式、带自动编号的独立公式、矩阵和交叉引用](docs/example-preview.png)

## 能做什么

- 处理主文档正文中形如 `[[MT:eq1]]` 的指定占位符，生成 `Equation.DSMT4` MathType 嵌入对象。
- 在副本上工作，拒绝覆盖原稿、现有输出和同名报告；检查对象类型、尺寸、部分域代码和书签名称。
- 提供标准库 Python 检查器，检查对象数量并读取非空 `Equation Native` 数据流。
- 提供使用虚构内容的示例生成脚本，以及不依赖 Word/MathType 的结构检查测试。

**自动检查不证明公式语义和排版正确。** 请在 Word 中逐页核对每个公式、上下标、矩阵、字号、裁切、编号和交叉引用，并抽查双击编辑。不要把未经目视检查的输出当作完成品。

## 环境要求

- Windows 用户桌面会话，安装并完成首次启动的桌面 Microsoft Word 与 MathType。此前在 Word 2021 64 位、MathType 7.8.2 上完成过小样本验证。
- PowerShell 7；结构检查与自动测试需要 Python 3.10 或更高版本，使用 Python 标准库。
- MathType 的 Word 插件可用，且 `MTCommand_TeXToggle` 在当前 Word 配置中能够运行。遵循本机宏安全和软件许可设置；本项目不会安装 Word/MathType，也不会修改安全设置。
- 批量脚本运行前请保存并关闭其他 Word 会话。该限制用于避免接管或关闭已有文档。

默认使用 64 位 MathType 的常见模板路径。如果安装位置或 Word 位数不同，使用 `-MathTypeTemplate` 指定与 Word 位数匹配的 `MathType Commands*.dotm`。如果 Word 无法在当前执行环境中启动，先排查 Windows 用户会话和软件本身的提示；隐藏窗口不等于无需桌面会话。

## 快速试用

在仓库根目录打开 PowerShell 7，保存并关闭其他 Word 文档后执行：

```powershell
$repoDir = (Get-Location).Path
$outputDir = Join-Path $repoDir 'outputs'
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

& (Join-Path $repoDir 'examples\create-example.ps1') `
  -OutputDocx (Join-Path $outputDir 'example-input.docx')

& (Join-Path $repoDir 'scripts\insert-mathtype.ps1') `
  -InputDocx (Join-Path $outputDir 'example-input.docx') `
  -Manifest (Join-Path $repoDir 'examples\formulas.json') `
  -OutputDocx (Join-Path $outputDir 'example-output.docx')

python (Join-Path $repoDir 'scripts\verify-mathtype.py') `
  (Join-Path $outputDir 'example-output.docx') --expected-total 5
```

脚本不覆盖已有文件；重试时请选择新名称或手动处理测试文件。打开输出文档更新域并目视检查：独立公式的编号与正文中的交叉引用应一致。输出旁的 `.mathtype.json` 是本地报告，包含公式 TeX、文件路径和哈希；分享前请检查内容。

## 处理自己的文档

1. 在**副本**中需要插入公式的位置放入唯一占位符，标记使用 `[[MT:字母数字_或-]]` 格式。确保它不在自动域、书签、超链接或其他复杂结构内部。
2. 按 [`examples/formulas.json`](examples/formulas.json) 准备 UTF-8 JSON 数组。`tex` 不包含 `$` 定界符和换行；`layout` 为 `inline` 或 `display`。
3. 用 `-InputDocx`、`-Manifest` 和 `-OutputDocx` 指定文件。两种 `layout` 都仅转换占位符；独立公式居中和右编号需在输入文档的公式段落、制表位及域中预先设置。
4. 使用检查器；`--expected-total` 是输出文档中**全部** MathType 对象的数量，包含原有对象。然后在 Word 中更新域、重新分页、逐页检查和抽查双击编辑。

## 适用范围与限制

- 只处理主文档故事范围中的标记；不自动批量转换已有 Word 原生公式。表格、页眉页脚、文本框、修订或受保护文档需要单独验证。
- TeX 能否解析取决于本机 MathType。`layout` 字段不自动控制段落布局，也不保证 MathType 内部样式与其他 TeX 引擎完全一致。
- 对象检查器验证数据流可读取且含非零内容，不解码或证明公式语义；对扩展 DIFAT 等未支持的结构会报错。域代码和书签名称一致也不能证明所有锚点和交叉引用都正确。
- 脚本不提供自动超时；若 Word 或 MathType 对话框阻塞，先检查当前实例，不要结束其他用户的进程。
- 将生成目录、日志、文档及报告视为可能包含敏感内容的本地文件；本仓库不包含实际项目报告、商业软件或许可证。

更多测试范围及已知边界见 [验证记录](docs/validation.md)。本项目与 Microsoft、MathType/WIRIS 无官方关联。

## 作为 Agent Skill 使用

仓库根目录的 [SKILL.md](SKILL.md) 是可选的技能入口。将仓库放在客户端支持的技能目录中时，请保留 `SKILL.md`、`scripts/`、`examples/`、`docs/` 和本说明的相对路径；是否能自动发现取决于所用客户端。技能不会安装或授权 Word 和 MathType。

## 测试

```powershell
python -m unittest discover -s tests -v
```

这些测试使用虚构 OLE/DOCX 样本，不测试 Word、MathType、排版或 GUI 集成。完整转换必须在安装了依赖的 Windows 桌面环境中另行验证。

## 许可与致谢

本项目采用 [MIT 许可证](LICENSE)。转换思路参考了 [tex-to-word-mathtype](https://github.com/klklkl2405/tex-to-word-mathtype)；第三方声明见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。Word 和 MathType 的许可独立于本项目。

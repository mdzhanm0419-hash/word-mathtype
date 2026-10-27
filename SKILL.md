---
name: word-mathtype
description: 在 Windows 桌面 Word 文档中使用 MathType 创建和验证可编辑公式。适用于新插入 MathType 公式、对明确指定的 TeX 占位符批量转换及相关的编号、引用和页面核查；不自动批量转换未经授权的已有公式。
---

# Word–MathType 公式

按用户要求和文档模板，在 Word 中生成真正的 `Equation.DSMT4` MathType 嵌入对象。保留公式数学含义、原文、编号和交叉引用；未经要求不自动转换已有公式。不能完成时如实说明，不以图片、Word 原生公式或 TeX 文本冒充 MathType。

## 操作步骤

1. 核查输入文档和格式约定，在副本的正文或公式段放置唯一 `[[MT:id]]` 标记，标记不得处于域、书签、链接或复杂结构内部。独立公式的制表位、居中和右编号应在文档中设置，不由转换脚本生成。
2. 参考同目录 `examples/formulas.json` 编写 UTF-8 JSON 清单，`tex` 不带美元定界符或换行。`layout` 只记录用途，`inline` 和 `display` 均只替换占位符；段落布局由输入文档承担。
3. 确认桌面 Word、MathType 和插件可用，用户保存并关闭其他 Word 会话，输出路径不存在。运行同目录 `scripts/insert-mathtype.ps1`，以 `-InputDocx`、`-Manifest`、`-OutputDocx` 传入绝对路径；有需要时用 `-MathTypeTemplate` 指定与 Word 位数匹配的模板。脚本为独立隐藏实例，不提供超时；遇到阻塞时检查本次实例，不终止无关用户进程，不更改宏安全设置。
4. 用同目录 `scripts/verify-mathtype.py` 配合 `--expected-total` 检查输出中全部 MathType 对象的数量和非空数据；总数包括原有对象。检查器不证明数学含义。用 Word 更新需要的域，重新分页并逐页查看，核对符号、下标、矩阵、编号、引用和裁切，必要时双击对象核对编辑性。

实际使用方式及已测试环境见 [README.md](README.md) 与 [docs/validation.md](docs/validation.md)。仅在主文档正文范围验证过有限样例；对页眉页脚、表格、修订或受保护文档先做隔离试验。自动检查字段代码和书签名称不证明所有锚点及引用的结果正确。不得把有限验证说成对任意 TeX、所有 Word 版本或无用户登录服务器的保证。

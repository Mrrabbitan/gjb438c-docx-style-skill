# Writing Prompts

Use these prompts as task scaffolds. Fill placeholders with the user's source paths, target document name, and project-specific constraints.

## Generate From Source Materials

```text
使用 $gjb438c-docx-style，基于以下来源资料生成一份 GJB438C 总体技术方案 DOCX：

来源资料：
- <source-path-1>
- <source-path-2>

输出文件：
- <output-docx>

要求：
1. 严格使用模板的 9 个一级章节，不新增、不改名、不换序。
2. 所有内容对齐到二级标题；无法直接对应的内容归并到最接近的二级章节。
3. 正文不得复用模板旧业务内容。
4. 表格、题注、标题、正文必须使用模板样式。
5. 生成后运行审计脚本并说明校验结果。
```

## Generate Standard GJB Document Type

```text
使用 $gjb438c-docx-style，基于以下来源资料生成一份 GJB 438C-2021 标准软件文档：

文档类型：
- <SDP/SIP/STrP/STP/OCD/SSS/IRS/SSDD/IDD/SRS/SDD/DBDD/STD/STR/SPS/SVD/SUM/CPM/FSM/SDSR>

来源资料：
- <source-path-1>
- <source-path-2>

输出文件：
- <output-docx>

要求：
1. 先读取 `references/gjb438c-2021-requirements.md` 判断文档类型和对应附录。
2. 不强行套用总体技术方案 9 章结构。
3. 保留 GJB 通用组成：封面、修改页、目录、正文、附录；缺失内容按裁剪规则说明。
4. 表格、题注、标题、正文必须使用模板样式。
5. 生成后运行审计脚本；标准文档不要启用 `--strict-overall`。
```

## Reformat Existing DOCX

```text
使用 $gjb438c-docx-style，将现有 DOCX 调整为 GJB438C 模板格式：

输入文件：
- <input-docx>

输出文件：
- <output-docx>

要求：
1. 保留输入文档的有效业务内容，不照搬模板旧业务正文。
2. 将内容重排到模板 9 章和二级标题框架下。
3. 所有段落按角色套用模板样式：标题、正文、表格、题注、目录、封面。
4. 对不确定内容标记为待补充，不编造项目事实。
5. 生成后运行审计脚本。
```

## Repair Style Only

```text
使用 $gjb438c-docx-style，只修复以下 DOCX 的格式，不改正文语义：

输入文件：
- <input-docx>

输出文件：
- <output-docx>

要求：
1. 一级标题套用 Heading 1。
2. 二级标题套用 Heading 2。
3. 三级至五级标题套用 Heading 3 到 Heading 5。
4. 正文套用 145正文。
5. 表头套用 145表头，表体套用 145表正文。
6. 图题/表题套用 Caption。
7. 修复后运行审计脚本。
```

## Audit And Repair

```text
使用 $gjb438c-docx-style 审计以下 DOCX 是否符合 GJB438C 模板格式：

输入文件：
- <input-docx>

要求：
1. 检查 9 个一级章节是否完整且顺序正确。
2. 检查二级标题是否对齐模板框架。
3. 检查正文、表格、题注样式。
4. 检查旧模板业务关键词残留。
5. 输出问题清单；如用户要求修复，再生成修复版 DOCX。
```

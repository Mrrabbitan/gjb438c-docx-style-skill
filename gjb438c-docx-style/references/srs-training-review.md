# SRS 培训评审写作配置

本配置把培训中的总分结构、功能分解和图文表对应落实为默认写法，属于**培训写作增强**，不是GJB新增条款。默认`writing_profile=training-review`；用户明确只按标准内容组织时选择`gjb-standard`。两种配置都遵守附录J，均可使用`draft`或`review`。

导航：写法与依据｜模型接口｜合同映射｜检查边界。

## 写法与依据

方法来源为用户提供的《军事训练项目建设过程相关问题研究》培训记录，SHA-256：`fad3f3c1d7c331bd28e0128de6080796f7351491cdbba0e9c846348e2a5c21db`。B编号按DOCX正文块顺序计数，空段也占编号；不是页码。原文件和原文不随公开skill分发。

| 培训定位 | 本配置采用的写法 | 边界 |
|---|---|---|
| B051—054 | 按软件组成递归分解；每个节点有名称、唯一标识、父级和正文位置；展开到可验证功能点 | 不同分支可以不同深度。能力域、模块和CSCI不是同义词；不因分章而新增CSCI |
| B062—065 | 明确用户角色、真实运行方式与适用关系，用泳道、状态或时序表达任务和转换 | 不照搬战时、训练、紧急等示例状态；不假定所有方式互斥 |
| B068—074 | 3.2提供总体概述、组成图、角色用例图、总体活动图、系统功能清单 | 根能力数量由项目设置；本次项目可显式设17，不作为所有项目的固定数量 |
| B076—089 | 每层能力节点提供范围与职责概述、组成图、业务活动图、直属能力/功能清单；叶模块中的需求保留细表 | 叶模块组成图可列其功能点，无须制造空子模块；活动图可以覆盖多项相关需求，但须逐项声明覆盖ID |
| B057、B087—089 | 输入、输出、前置条件、正常/异常判断及责任与流程对应 | 判断节点应有明确分支条件。图里增加的步骤不能无来源地变成新的验收义务 |
| B073、B079、B084 | 功能清单包含标识、名称、描述、研制状态、来源或依据及正文位置 | 新研、改造、沿用须有依据；不依据产品名称、开源属性或技术选型猜定 |
| B094—099 | 外部接口图和表标识双方、方向、对象、协议及联试条件 | 接口可单向；不强制一个数据类型一个接口；平台模拟不等于真实接入 |
| B102—110 | 软件安装、运行、使用和维护条件与实际/测试环境差异可核对 | 不把采购清单自动当最低运行条件；不推定高配环境一定满足性能 |
| B138 | 对合同原文逐项建立可复核映射，同时保留合法的其他来源或派生依据 | 培训不决定合同法律效力；不能把未获采纳的方案或合同建议稿称为批准基线 |

B114之后的类图、类方法、函数调用图属于设计说明主题，不作为SRS必备图。培训中的“不要编号”不能取消正式需求唯一标识；性能简述也不能省略适用的负载、单位、环境与判定条件。

## schema 2.0 扩展接口

扩展字段均保持schema 2.0兼容。旧模型不自动补来源、图、研制状态或合同批准。未指定写作配置的旧模型默认采用培训写法，draft报告缺口；需要保持仅标准内容检查时显式使用`gjb-standard`。

| 位置 | 字段与含义 |
|---|---|
| 根 | `writing_profile`：`training-review`或`gjb-standard`；`capability_overview`：3.2总体概述字符串；`expected_capability_roots`：项目显式确认的入口数量，可省略 |
| `capabilities[]` | 保留`id/name/parent_id/purpose/description`；新增`overview`、`order`（同级排序）、`development_status`（`new/modified/reused/tbd`）、`development_basis`、可选`physical_clause/tbd_ids` |
| `requirements[]` | `clause`仍为附录J类别，如`3.2`；`capability_id`指真实所属节点；`physical_clause`可明确实际放置章号；`exception`或有依据的`exception_applicability=not_applicable`及`exception_basis`；`contract_trace`见下节 |
| `figures[]` | `id/type/title`；图片`path`或受控`reference_id/location`；通过`clause`和/或`capability_ids`指定放置位置，`requirement_ids`声明实际覆盖，`description`说明阅读方式或适用边界 |
| 图类型 | `composition/use_case/activity/swimlane/sequence/state/interface/data/other`。本配置3.2检查`composition/use_case/activity`，各能力节点检查`composition/activity`。泳道作为活动图绘法时使用`type=activity`，description可说明泳道角色 |
| `source_requirements[]` | 可新增`kind=contract/standard/plan/other`及`confirmation=confirmed/pending`；来源性质与文件确认状态分开，不通过REF名称猜测；可用`tbd_ids`显式关联未决文件身份 |

兄弟节点按`order`、输入顺序稳定排序，根从3.2.1编号，按父子树递归生成。Word标题最多9级；超出实际技术限制报错，不能静默压平。目录域覆盖实际标题深度，仍须在办公引擎更新并逐页核验。

`scripts/srs_authoring.py`提供`capability_outline(content)`，返回预序列表`{node, level, physical_clause, requirements}`；异常父级、重复标识、循环和不可达节点抛出异常。`physical_clause`若提供必须匹配计算位置。不得用旧章号掩盖重组后的实际位置。

`canonical_srs_to_sections(content, mode='draft', writing_profile=None)`生成递归章节。每层清单只列直属子能力和直属正式需求，详细需求只输出一次；图按明确关联放置，共享图可以出现在多个有关章节。图源路径仅用于装配，不写进交付正文。路径须可读取；CLI图片相对路径按调用目录解释，跨环境建议在装配前统一解析。

引用登记、来源和能力节点均可使用`tbd_ids`。未决扫描识别同记录中的依据及显式TBD；来源与追踪也可沿实际`reference_id`关联文件身份未决项。不会把根级notes中的任意TBD扩散到所有记录；关联开放项只表示受控，仍不表示已关闭。

## 合同映射和受控保留

`requirements[].contract_trace`包含`status`、`source_ids`、`basis`、`tbd_ids`：

| 状态 | 含义与检查 |
|---|---|
| `mapped` | 合同条文与需求的映射已有证据；来源必须登记为`kind=contract`，也出现在该需求`source_ids`与正常双向追踪中。**不代表合同已签署或范围已批准** |
| `derived` | 有明确派生依据；`basis`说明从哪些上层目标/约束导出，并保留正式来源或`derivation`记录 |
| `pending` | 尚无充分合同依据而受控保留；`basis`解释保留理由，`tbd_ids`指向真实登记项。正文不能表述为合同已确认 |
| `not_applicable` | 有具体依据说明该条不属于合同逐条映射；不删除GJB要求的来源或派生追踪 |

`confirmed`为兼容状态，语义等同映射已核的`mapped`，不表示合同批准。新模型优先用`mapped`。

合同建议稿的条文可能已准确对应，此时`mapped`与来源`confirmation=pending`可以并存。draft保留警告，review指出基线尚未确认。无合同依据的受控需求使用`pending`与TBD，不通过改变状态标签、补造合同号或回填来源全集伪造闭合。

## 检查与结论

结构错误（错误类型、未知关联、重复ID、循环、章号矛盾）在两种模式都是错误；缺概述、空叶、图类型/覆盖不足、研制依据、异常行为或合同处置不足在draft为警告，review为错误。

机器检查模型关联、节点顺序、保存DOCX中的物理归属、功能清单与图对象是否缺失。它不能证明组成关系真实、图节点语义正确、所有异常已覆盖、来源已签署、流程可执行或版面可读。必须分别记录人工语义和最终视觉检查；未执行仍为`not_run`。检查通过不代表项目批准或软件验收。

修改时保留正式需求ID和旧ID去向。重组前后核对需求集合、来源覆盖、合同待定集合及图表关联；凡需求文字改变，应重新核验来源与通过判据，不能把结构重组当成新增承诺的授权。

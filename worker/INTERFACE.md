# Office worker接口与已验收范围

当前工作代码记录：2026年10月2日第20份实现，适用DSH 0.2.0-rc.2；package.json的0.2.2是历史发布编号。以下先列新增接口，后面的0.2.1、0.1.2说明用于兼容旧调用。

## 2026年10月2日新增接口和检查依据

`office_read`允许`source_paths:[]`与`target_path`一起使用，只读取目标。来源有1至8份时仍可只读来源。实际表名及范围在`capabilities.target.table_ranges`；模型收到的字段证据排在大型结构清单之前。结构化值保留全部原文、上下文、限制和引用。

`office_check_output`与新增`office_restore_template`共用完整`contract`或工作区内的`contract_path`。`coverage`与`requirements`列出任务全部需处理的位置；`table_ranges`推荐保存在contract内部，也接受工具顶层对象并自动归入清单；两处同时提供但内容不同，返回OFFICE_TABLE_RANGES_CONFLICT。方法保存继续采用清单内部范围。支持以下新增规则：

- `source_label_refs`按来源顺序指定真实标签格，并检查它与值的位置关系；不同文件的物理位置各自取证。
- `extract`使用唯一的`after`、`before`或`through`文字边界，未使用的边界省略。可统一中文括号，并转成文字或非负整数。例如`72（2024年报）`按`before:"("`得到72；按`after:"("`和`before:"年报"`得到2024。空字符串与重复边界有明确错误。
- `blank`引用真实空标记，例如`source_refs:["s0::资料!B1"]`与`blank_markers:["-"]`。确实没有栏目时，用`source_absence:{source_index:0,sheet:"资料",labels:["企业联系电话"]}`检查当前该表文字格及后续页。真实0保留；这个检查的依据是指定表及标签，图像和未定义的业务含义另由模型处理。
- `formula`接收明确算式树。例`{operator:"ratio",args:[{source:0},{operator:"average",args:[{source:1},{source:2}]}]}`计算第一个来源除以后两个来源的平均值。仅支持有限数值及指定运算；零分母和超限返回错误。
- 任务明确允许文件名作主体线索时，`filename_identity:true`与`s0::$filename`可用于身份文字的copy/extract。财务值、电话和单位换算继续使用文件内容。

`office_restore_template`参数是当前`source_paths`、原`target_path`、候选`output_path`、新的最终`output_name`、上述清单和可选实际`script_path`。程序先检查候选全部填写值和范围；值通过后，把覆盖位置的值写入原模板新副本，同步清单指定的Excel表范围，再独立重开检查。普通Word安全单元格、段落占位与Excel可用。输出已有、路径越界、实际保护、未知位置或错误值按对应原因拒绝。候选和用户原件保留。交接通过时记录真实复用评估事件，代码保存、验证、启用继续分别执行。

检查返回`next_action`说明下一步：值都通过但格式失败时建议同清单交接；来源或值有问题时指出先核查当前引用；完整通过时提示交付及复用评估。该反馈提供操作依据，模型仍负责陌生业务判断。

方法回调新增`calculate_many`：一次接收1至1000项明确计算，返回对应结果；单项`calculate`也支持上述expression。已用300个数值换算加1个平均余额公式验过批量路径。执行资源限制沿用当前配置。保存方法继续使用`def run(inputs)`、宿主读写回调和独立检查，文件/网络/系统操作由宿主控制。

Word目标保护判断读取`w:enforcement`：0、false、off表示保护未启用，保留设置并处理安全位置；1、true、on表示实际启用。Microsoft说明见[DocumentProtection](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentprotection?view=openxml-3.0.1)。Arena当前12份Word为未启用，程序复查得到102个普通安全位置。

候选读取失败时，错误保留原code，并返回file_role=candidate_output及filename。重复坐标错误明确要求修候选文件，帮助模型找到实际写入故障。

程序证据：第20份worker回归129项中128项通过，1项数据库文件链接检查待Windows权限补验；实际DSH工具流程结果见最新进度。模型和桌面验收见项目集中报告及人工核对清单。

## 历史0.2.1结构协议（新增项以上文为准）

stdin仍只接受一个JSON对象、stdout只返回一个JSON对象。单份输入上限10 MiB，适配层全部输入原始字节合计上限20 MiB，请求上限28 MiB，ZIP大小、解压成员、输出大小及执行超时沿用旧检查。没有模型或网络请求。

- inspect_v3：sources需要1至8项，各项含format（docx/xlsx）、bytes_b64和可选Excel scope（sheet、range）；target可省略，省略时只读来源。0.2.1新增requested_fields（1至32个检索词）、source_offset（0至100000整数）、target_mode（fill_blanks默认或review_values）。检索和续页目前只用于Excel来源。
- prepare_v3：同一输入加mappings，每项含source_ref、target_ref；已有值核对更新额外需expected_target_text，必须完全等于本次旧text。返回确定性changes，不写。
- apply_v3：同一输入、mappings、changes、output_name；程序重新推导changes，与提交值逐项一致才写。
- fill_v3：同一输入、mappings、output_name；一次完成推导、检查、最小包写入和重开，模型不能传value。
- records_store：可选保存当前原文记录、文件版本、范围与解析器摘要；提议只含source_ref和explanation，状态始终为未确认提议。
- records_view：只读列出工作区保存历史；snapshot_id可查看原文和提议。历史记录不能直接重放写入或权限。

模型调用office_read后，来源引用为s0::、s1::等加当前结构引用。Excel使用工作表!A1；Word使用包部件与当前XML对象序号，句内空位另有字符范围。序号只对本次内容和解析器版本有效，不能当固定模板坐标。

office_fill需当前read_id、mappings和新文件名，一次执行尝试后不能重放。Word/Excel来源与目标可跨格式，多份来源汇集后生成一份输出；重复目标拒绝，未知引用、过期读取、输入变化、解析器变化、既有输出和缺来源单位均拒绝。普通中文输出名可用；路径、Windows保留名和保留字符拒绝。

模型只看紧凑证据表，原始结构化结果保留物理位置和关系。数据有原文位置不等于字段归属已确认，表头、说明、单位和示例仍可能在原文中，不能强配。Excel可显式范围补读，未读范围和上下文截断有状态；本版Word按完整结构读，超限明确拒绝而非截断。

## 支持与限制

Word遍历正文、表格及表内表，物理格不因合并别名重复输出；正文/页眉页脚中使用w:txbxContent/w:p结构的文本框原文可读，其他形状文字和图片文字不作已读承诺；目标文本框和内容控件不写。空格写入和明确占位符插值保留原有文字片段样式；含复杂域、绘图或控制对象的写入位置拒绝。目标保护、宏和嵌入物逐项显示能力；不能把可读承诺当可写承诺。

0.2.1 Excel来源按实际存储格读取，不依赖dimension声明；保留有限行列上下文、工作表清单及合并范围，关系仍为未确认。每页最多1000个节点，next_offset续页；只有本页位置可用于此次写入，不能把本页当全量。字段检索同时返回同一行附近最多4个实际格、同一列上下最多2个实际格，不保证业务关系成立。超过100000个实际格或32张表明确拒绝。图片、绘图、外链等列为read_ignored_parts，不执行宏或链接，不能声称这些内容已读。读取允许忽略部件不代表目标写入允许改变它们。

公式返回文字及未验证缓存，不能作为直接取值来源。日期数值格式和ISO日期原文显示value_error；读取其他格继续，写入引用拒绝。同为Excel百分号来源与目标保留原始数值，不乘除100；与明确金额单位冲突拒绝。目标公式、保护格和非合并首格仍不写。默认只填空；review_values可列出安全已有值位置，但需要旧原文核验，不自动确认哪些字段应该改。目标复杂部件仍沿用原有拒绝范围。扩行、派生计算和XLSM不在当前固定填写接口范围。

金额只在来源、目标均有明确元/万元/亿元依据时换算。跨格式百分比原始值与显示值尚未验收，拒绝直接照抄。Word文字呈现还须符合当前来源与目标的明确要求，不能只验证数值等价。

SQLite文件固定为工作区office-plugin-records.sqlite3，保存可选且经原生写权限。同名用户文件不接管，越界链接不跟随；原始记录不因提议覆盖，范围或解析器变化生成不同快照。数据库不保存可执行计划或授权。模型不能把未确认解释升成用户确认。

通用方法以插件附带JSON保存，全局用户方法固定在插件library/user目录；宿主CLI只导入草稿、查看或导出，不自动启用或执行任意代码。模型不能任意选择全局宿主写入位置。0.2.2共有7条内置试用方法，按当前条件检索最多3条；完整输出检查只在有当前来源、可读目标及多个目标位置时附加。未定义的匹配条件不默认命中。跨任务自动学习、任意临时脚本的字段／模板最终校验及实际模型效果仍未完成，方法建议不等于这些能力已经实现。

重复治理当前为可验证的同版本只读缓存与离线分类，不启用全会话在线监控。保存、缓存、哈希及建议的成本均须计入后续模型对照。

实际程序证据与未完成项见[实施至对比测试入口记录](../results/实施至对比测试入口记录_20261001.md)。当前不是全部视觉、交互确认或信创验收。

## 0.1.2旧协议历史说明


冻结摘要：2026-09-30，XLSX 与受限 DOCX 同格式闭环；跨格式、DOCX 单位换算、模型、API 和桌面加载不在本 worker 冻结范围。本说明只描述 `worker/office_worker.py` 已实现且已由本地单元测试覆盖的离线行为；不代表 DSH 桌面插件加载、模型、DeepSeek/千问 API 或原生审批已经验证。

## 请求与响应

worker 从 stdin 读取一个 JSON 对象，并只向 stdout 写一个 JSON 对象。stdin 最大为 28 MiB，超过该值在 JSON 解析前返回 `request_too_large`。`worker/xlsx_prepare.py` 是无路径、无 I/O、无模型访问的确定性准备模块；它仅从已打开的工作簿单元格推导值。

必填字段如下：

```json
{
  "operation": "inspect、prepare 或 apply",
  "format": "xlsx",
  "source_b64": "源 XLSX 原始字节的 base64",
  "target_b64": "目标 XLSX 原始字节的 base64"
}
```

`inspect` 返回：

```json
{
  "ok": true,
  "operation": "inspect",
  "source_sha256": "…",
  "target_sha256": "…",
  "source_candidates": [{"source_ref": "工作表!B2", "label": "金额", "value": 15000, "unit": "元", "nearby": {}}],
  "target_slots": [{"target_ref": "工作表!B4", "label": "金额", "unit": "万元", "nearby": {}, "merged_range": null}],
  "hints": {"source": [], "target": []}
}
```

候选只来自有相邻标签的非空、非公式单元格；空位只来自有相邻标签的空白或占位符单元格。模型不在 worker 中运行，也不提交 `value`。工作表名含 `!` 时，首版 `工作表!A1` 引用无法无歧义表达，worker 输出提示而不产生该表候选。

`prepare` 额外要求：

```json
{
  "mappings": [{"source_ref": "工作表!B2", "target_ref": "工作表!B4"}]
}
```

每项 `mappings` 必须且只能包含这两个位置，不能含 `value`。worker 从同一轮已读 source/target 字节重新定位单元格，验证来源非空非公式、目标为空且可写、目标不重复，然后返回不落盘的 `changes`：

```json
{
  "ok": true,
  "operation": "prepare",
  "source_sha256": "…",
  "target_sha256": "…",
  "changes": [{
    "source_ref": "工作表!B2",
    "target_ref": "工作表!B4",
    "value": 1.5,
    "conversion": {"kind": "money_unit_conversion", "source_unit": "元", "target_unit": "万元"}
  }]
}
```

`apply` 额外要求：

```json
{
  "output_name": "filled.xlsx",
  "changes": [{"source_ref": "工作表!B2", "target_ref": "工作表!B4", "value": 1.5}]
}
```

这里的示例值应为数值 `1.5`，不是字符串。适配层必须先验证来源与业务语义；worker 则自行重新推导相同的值和单位结果，并要求提交的 `changes.value` 类型和值完全一致。不同会返回 `value_mismatch`，所以模型或调用方不能把任意文本、公式注入或错误金额带入 `apply`。若提交 `conversion`，它也必须与 worker 的推导完全一致。成功响应包含输入/输出 SHA-256、`applied`、输出路径，以及 `verification.reopened`、`verification.package_members_preserved` 和受限结构检查结果。

单位规则仅为：

- 来源和目标均无单位：保留来源标量原值。
- 来源和目标同为元、万元或亿元：保留来源原值。
- 元、万元、亿元之间：使用已采用的 `vendor_candidates/value_fit.py` 做精确十进制换算，并输出 Excel 数值类型；例如 `15000 元 → 1.5 万元`。
- 来源和目标同为 `%`：仅保留 Excel 原始数值，不乘除 100。公开 q02 样例只验证了 `3.85` 保持为 `3.85` 的情形；它不能证明所有百分比格式、显示格式或文本百分比都可处理。
- 未知单位，以及 `%` 与金额单位、单位与无单位等冲突：返回 `unknown_unit` 或 `unit_conflict`，不猜测。
- 右侧单元格为空或明确表头“单位”时视为无单位。对数值或金额来源，其他非空右侧文本会被视为单位声明；它若不是元、万元、亿元或 `%`，返回 `unknown_unit`，例如“吨 → 千克”不会被默认为无单位。对普通文本来源，右侧业务文本不被猜作单位，避免把横向布局的下一个业务值当作单位。
- 数值或金额来源没有单位、目标却声明金额单位时返回 `unit_conflict`。worker 不会从目标单位、相邻字段、测试答案或其他外部材料推定来源单位。
- 数值不承诺任意精度：整数必须在 JSON 安全范围 `±(2^53-1)` 内；换算小数必须能由 `Decimal(str(float))` 无损复现，否则返回 `numeric_precision`。日期/时间单元格在 `inspect` 阶段返回 `unsupported_value_type`，不会隐式 JSON 序列化。

## 拒绝规则与上限

- 单个 source/target 包最大 10 MiB；ZIP 最多 2,048 项，解压总量最大 50 MiB，单成员压缩比最大 100:1。
- 最多 32 张工作表；每张最多 10,000 行、256 列；所有工作表最多扫描 100,000 格。超限会拒绝，不静默截断后继续。
- 最多 100 项变更；输出最大 15 MiB。
- `output_name` 必须是当前工作目录下的普通 `.xlsx` 文件名，不能含路径、`..` 或冒号；以排他创建防止覆盖已有文件。
- 拒绝公式、已启用工作表保护下的锁定单元格、合并区域非左上单元格、非空目标、重复目标、无效引用和已存在输出。
- 拒绝含外部链接、VBA、嵌入对象、ActiveX、绘图、图表、切片器、线程批注/人员信息或 `calcChain.xml` 的 XLSX。理由是这些部件在 `openpyxl` 往返时可能丢失或改变，首版不应假称保真。
- DOCX 目前统一返回 `unsupported_format`。

## 写后验证与失败恢复

写入只在内存中完成。写前记录目标包 ZIP 成员集合；保存后要求成员集合不变，重开 XLSX 后确认每个目标值。还比较除明确 `target_ref` 的值以外的工作表顺序、合并区域、页面属性、行列尺寸和有意义单元格的内容/样式。

这是一组受限检查，不是对未知 XLSX 特性的完全视觉或二进制保真承诺。输入 `source_b64` 和 `target_b64` 从不被修改；只生成新的输出文件。若排他创建后的写入或同步失败，worker 返回 `output_partial`，并保留可能不完整的输出文件路径及 `recoverable: true`，由调用方显式检查或清理；worker 不会静默删除它。

worker 不包含网络、API 凭据、模型、shell、子进程或文件路径输入功能。

`office_worker.py`、`xlsx_prepare.py` 和 `docx_prepare.py` 会从自身文件位置计算并插入项目根路径。实际 DSH 内置 Python 路径 `D:\deepseekharness preview\resources\runtime\primary-runtime\dependencies\python\python.exe` 使用 `-I -B -X utf8` 运行本地测试时，`openpyxl` 与 `python-docx` 都可用。此前普通 Microsoft Store Python 的隔离模式缺少 `openpyxl`，那只是另一解释器条件，不能据此判断 RC2 运行时缺少依赖。worker 不会自行加入用户 site-packages，也不会安装依赖。

## DOCX 首版范围

DOCX 的 `inspect`、`prepare`、`apply` 使用相同的 base64 协议；`source_format` 是可选字段，默认等于 `format`。首版只接受 `format: "docx"` 且 `source_format: "docx"`，跨格式请求返回 `unsupported_cross_format`。

- 稳定引用为段落 `p:索引`、段落内有证据的值片段 `p:索引:v:序号`，或表格单元格 `t:表索引:r:行索引:c:列索引`。片段候选回传 `label`、`value`、`span_start`、`span_end` 和原段落上下文；它不是调用方自由传入的字符范围。
- 来源候选是明确的“标签：值”段落、单一叙述标签值、句内“标签+人民币数字单位”或“标签为值”结构，或表格中相邻标签/值单元格；目标仅是标签后真正为空的段落，或标签右侧真正为空的表格单元格。说明、备注、提示、签字、审批、日期栏不输出为普通空位。
- `prepare` 仍只收位置映射，不收 `value`；DOCX 当前保留来源文本原值，不进行金额换算。单段落只有一个候选时可继续用 `p:索引`；多个候选必须用存在的 `p:索引:v:序号`。旧 `p:索引` 在多候选场景返回 `ambiguous_source`，不猜测。
- `apply` 会重新推导并比较传入 `value`，用排他创建生成 `.docx`，重开验证目标值；验证非目标段落/单元格文本和 runs 样式、目标原有 runs 样式前缀以及 ZIP 成员集合。
- `inspect` 只从标签括号中的 `元`、`万元`、`亿元`、`%`，或表格值单元格右侧恰为这些单位的单元格识别目标单位；不从字段名称猜单位。来源文本末尾单位与目标显式单位不一致时返回 `unsupported_docx_conversion`，首版不做 Word 金额换算；同单位保留来源原文。
- 含宏、嵌入对象、ActiveX 或文本框的 DOCX 返回 `unsupported_docx_feature`。这不是全部 DOCX 结构或视觉保真承诺。

## 当前测试证据

### 2026年10月2日第16份实现：Word方法、当前工作区和所需来源读取

可执行方法现在支持一个XLSX或一个普通DOCX新文件。Word目标按本次可写物理格及段落占位定位，`read_target()`新增`target_slots`，写入仍用`write_cells`，只收`target_ref`与`value`。`allow_existing_values=true`明确允许更新已有值；否则现有非空Word格子不可写。写入复用原有最小XML替换程序，独立检查从原模板与最终包重新取值，比较段落和文字样式、非目标XML及其他包部件。段落多占位按原固定分隔唯一取值，歧义或超过32处交DSH继续处理。打印分页另行人工核对。

Word数值要求可指定`render:fixed_decimal`或`render:percent_text`，同时给0至10整数`decimals`；检查实际文本的小数位及百分号。金额仍从本次来源按十进制换算、舍入。8项真实Word包与Monty执行检查通过，含零值、缺失、负数舍入、历史旧值、两个段落占位、样式错误和新企业新年度复用。

独立清单检查按明确来源位置读取相邻上下文；这条路径会查到第一页以后的来源位置。按键查找会读完相关页，核对唯一性。`read_source(index,requested_fields?,source_offset?)`支持按标签检索或续读，返回范围和分页情况。记录最多10000项，超限明确交DSH处理。筛选所需字段不表示已解析图片、公式或全部业务事实。

宿主配置新增`workspaceMode:session`；工作区来自真实DSH会话的原生策略，模型参数不能指定目录。默认固定配置兼容原有测试。读取标识和纠错事件按实际会话、工作区隔离，方法库由宿主配置的全局目录共用。宿主写入前仍检查权限和持久化服务。运行文件更新后旧方法要重新验证实际版本；启用或运行时的工具返回给出下一步。

第16份本地回归为108项，107通过；1项数据库符号链接测试因Windows权限待补。单宿主双工作区方法工具流程12项通过，方法库目录链接更新检查5项通过。原生与插件实用样本对照接下来独立执行。


### 2026年10月2日方法工具反馈更新

这是DSH工具层变化，worker协议未变。office_method_list可选query按空格、逗号等拆关键词，候选按启用和命中数排序；limit默认5，接受1至10整数；include_inactive默认false，有启用候选时只展示启用候选，没有时仍展示草稿，true查看草稿及旧版。结果保留total、matched、methods，另有hidden_inactive、lookup_scope。用途摘要最多240字符；last_saved_validation是最新历史验证，verified_for_current_inputs固定false。查找不表示本次适用，执行仍验证完整条件及当前文件。

office_method_run的reused仅在complete和published都为true时为true；method_executed说明程序实际运行。verification_summary说明独立检查出处、覆盖、通过数和未查范围。reuse_review在已有方法成功执行时new_method_needed=false、event_id=null，不产生新增保存事件；失败为null，交DSH分析失败原因。只能用其他真实工具返回的事件标识进行事件评估。以上是反馈澄清，没有减少检查或放宽执行权限。

### 2026年10月1日新增方法基础

以下与后文2026年9月30日历史首版检查分开：`check_output`接收1至8份本次来源、模板、实际产物和独立`contract`。XLSX范围使用`coverage:["企业资料!A2:H5"]`，requirements逐项指定`target_ref`、`source_refs`，来源前缀为`s0::`，操作支持copy、sum、product、difference、ratio、percent。需要金额换算时成对明确source_unit、target_unit及decimals；依据可来自本次文件或已明确用户任务，不猜单位。表对象范围使用对象`table_ranges:{"企业资料!企业资料表":"A1:H5"}`，不能传数组。

多个来源参与算式时，`source_labels`按引用顺序列直接字段名，例如数量和单价；单个`source_label`只用于单来源，不把派生结果名当作各输入字段名。跨批次来源行或文件可能变化时，可用`source_matches:[{sheet,value_column,key_columns:[{column,target_ref}],source_label?}]`；按未修改原模板的本次主体、订单、期间等键在全部当前来源中唯一查找。重复、缺失、公式未算或读取未完时拒绝，不沿用首次来源坐标。保留区另检查内容和格式，不要求把保留表头重复加入填写范围。

DSH工具层可传`contract_path`工作区JSON文件代替展开contract；保存方法可传`metadata_path`文件，内容为`{conditions,contract}`，两种方式各自与内联参数二选一。大清单可由原生脚本循环生成一次，避免反复发送长参数。保存版本不可变；导出仅复制代码，编辑后用parent_id保存新版本，再验证新ID。

方法入口`def run(inputs)`。`read_source(index)`和`read_target()`返回字典`{"nodes":[...],"format":"xlsx","capabilities":{...}}`，遍历的是返回字典的nodes列表，不是整个字典。节点ref为完整工作表名加坐标。inputs由宿主自动加入来源数量、来源相对名称和目标相对名称，仅用于标识，不能直接打开文件。`calculate(operation, values, decimals, source_unit, target_unit)`和`convert_amount`在宿主按十进制计算，返回Excel数值；write_cells只写清单覆盖且明确列为要求的位置。

`ok`表示检查程序正常返回；`complete`表示所列完整要求与所支持格式通过，两者不能混用。遗漏的要求、未读到的来源、未知算式、错误单位/主体、未声明位置变化均不能报完整通过。DOCX通用格式验收尚未支持；打印与全部Office部件不在通用检查承诺内。模型对陌生业务清单理解错，程序无法据此自动纠正所有含义。

固定method_runner通过Monty 1.0.0解释器运行method.py的`run(inputs)`，仅开放JSON值和受信任办公操作。方法源、解释器和办公桥接程序锁定哈希；验证依据由宿主说明文件提供，不由方法程序给自己生成答案。原生执行器中仅暂存文件，完整通过后由宿主排他创建新产物；检查失败不发布，失败稿保留便于原生继续修正。当前方法写入仅XLSX、一次一个文件，内存64MiB、解释器计算5秒、宿主整次60秒；无文件挂载和系统操作，仍不是完整操作系统沙箱。

当前本地回归：97项程序检查中96通过、1因Windows文件链接权限跳过；10项真实DSH方法工具检查通过，包含跨会话/工作区重读新值、只读权限拒绝、未验证拒绝启用、输出已存在拒绝、停用后拒绝、导出修订与事件关联。原有12项结构读写流程及2项权限检查也通过。模型自主入库和效率不能由这些本地检查证明，真实模型试用另行评分。方法执行反馈保留后台完整证据，提示中不展开已通过项目，失败项目仍返回；提示中的精简不减少实际检查。

命令 `D:\deepseekharness preview\resources\runtime\primary-runtime\dependencies\python\python.exe -I -B -X utf8 -m unittest discover -s tests/worker -v` 在 2026-09-30 的实际结果为 24 项通过、0 项失败、0 项错误，框架报告 1.281 秒。除既有安全边界外，新增证据包括：公开开发样例 q01 的无单位/同金额单位准备，q02 的 `12.5 万元 → 125000 元` 与原始百分比 `3.85` 准备，q05 的段落 inspect/prepare/apply，以及 q06 全部 5 个字段的 prepare/apply/reopen：3 个表格来源和同一段落中的 2 个证据片段；另有未知单位、大整数、日期值、实际空白授权目标、非目标内容/样式篡改、“来源数值无单位但目标有单位”的拒绝、以及 DOCX 不同显式单位拒绝/同单位原文保留测试。测试未调用 API、未读取凭据、未修改 `gold/` 或旧目录。

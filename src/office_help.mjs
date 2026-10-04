import { quotePwsh } from './runtime.mjs';
import { contractSchema } from './check_contract.mjs';

export const OFFICE_HELP_TOPICS = ['runtime', 'method_code', 'checklist'];

const topicDescription = '帮助主题：runtime运行命令，method_code方法回调，checklist精简contract说明。一次只返回一个主题。';

const posixQuote = value => `'${String(value).replaceAll("'", "'\\''")}'`;

export function methodRunCommand(config, { platform = process.platform, scriptPath = 'workspace relative.py' } = {}) {
  if (!config?.pythonPath) throw new Error('OFFICE_HELP_PYTHON_PATH_REQUIRED');
  if (typeof scriptPath !== 'string' || !scriptPath.trim() || scriptPath !== scriptPath.trim() ||
      /^[\\/]|^[A-Za-z]:[\\/]/.test(scriptPath) || (platform === 'win32' && scriptPath.includes(':')) ||
      /(^|[\\/])\.\.?($|[\\/])/.test(scriptPath))
    throw new Error('OFFICE_HELP_SCRIPT_PATH_INVALID: 只接受工作区相对脚本路径');
  if (platform === 'win32') return `& ${quotePwsh(config.pythonPath)} -B -X utf8 ${quotePwsh(scriptPath)}`;
  return `${posixQuote(config.pythonPath)} -B -X utf8 ${posixQuote(scriptPath)}`;
}

const runtimeHelp = (config, options) => ({
  topic: 'runtime',
  platform: options.platform ?? process.platform,
  python_path: config.pythonPath,
  script_path: options.scriptPath ?? 'workspace relative.py',
  command: methodRunCommand(config, options),
  command_contract: '命令只运行工作区相对Python脚本；原生DSH宿主仍按自身策略设置超时、输出限制和沙箱。帮助不会执行命令。office_method_validate/run另有固定方法运行时锁。',
  windows: "PowerShell示例为 & '配置的python.exe' -B -X utf8 '工作区相对脚本.py'；路径中的单引号会加倍。",
  posix: "POSIX示例按当前平台生成；路径中的单引号使用shell安全转义。",
  restrictions: ['不得传绝对脚本路径或越出工作区', '不得在方法中导入os、openpyxl、subprocess或网络库', '不得借帮助提高权限、访问密钥或改写原件'],
});

const methodCodeHelp = {
  topic: 'method_code',
  entrypoint: 'def run(inputs)',
  callbacks: {
    read_source: {
      signature: 'read_source(index, requested_fields=None, source_offset=0)',
      returns: {
        format: 'xlsx或docx',
        nodes: [{ ref: '来源表!B2', kind: 'cell', value: '实际值', text: '实际值', row_headers: [], column_headers: [], parent_context: ['来源表'], row_header_refs: [], column_header_refs: [], unit: '', unit_evidence: [], unit_ambiguous: false, formula: false, cached_value: null, number_format: 'General', number_literal: null, value_error: null, physical_path: '来源表!B2', context_truncated: false, relation_status: 'unconfirmed' }],
        capabilities: { read: true, write: false, formula_values: false, scope: null, requested_fields: null, source_offset: 0, next_offset: null, matched_records: 1, requested_refs: null, complete_requested_range: true, complete: true, read_ignored_parts: [], images_read: false, macros_executed: false, sheet_inventory: [], pagination: '只有本页位置可引用；未返回位置不算已读，检索结果不是业务匹配', pages_combined: true, selected_evidence_only: false, all_text_nodes_read: true },
      },
      note: '完整返回字典；按需检索或分页仍返回同样的format、nodes、capabilities形状。',
    },
    read_target: {
      signature: 'read_target()',
      returns: {
        format: 'xlsx或docx',
        nodes: [{ ref: '目标表!B3', kind: 'cell', value: null, text: '', row_headers: ['金额'], column_headers: [], parent_context: ['目标表'], unit: '万元', unit_evidence: ['万元'], unit_ambiguous: false, formula: false, value_error: null, merged_span: null, number_format: 'General', physical_path: '目标表!B3', relation_status: 'unconfirmed', context_truncated: false }],
        target_slots: [{ ref: '目标表!B3', kind: 'cell', value: null, text: '', row_headers: ['金额'], column_headers: [], parent_context: ['目标表'], unit: '万元', unit_evidence: ['万元'], unit_ambiguous: false, formula: false, value_error: null, merged_span: null, number_format: 'General', physical_path: '目标表!B3', relation_status: 'unconfirmed', context_truncated: false, target_ref: '目标表!B3', write_kind: 'cell', requires_expected_text: false }],
        capabilities: { read: true, write: true, formula_values: false, target_mode: 'fill_blanks', unsupported: [], scope: null, complete: true, table_ranges: {}, complete_requested_range: true },
      },
      note: 'target_slots是方法回调的完整可写位置列表；聊天展示可能只显示摘要，不能把聊天compact结果当作回调返回值。',
    },
    write_cells: {
      signature: 'write_cells(changes, output_name, table_ranges=None)',
      changes: [{ target_ref: '目标表!B3', value: 12.5 }],
      note: 'target_ref、写入授权、表范围和已有值规则仍由保存的contract及宿主独立检查决定；方法不能扩大范围。',
    },
    calculations: {
      signatures: ['convert_amount(value, source_unit, target_unit, decimals=2)', 'calculate(operation, values, decimals=2, source_unit=None, target_unit=None, expression=None)', 'calculate_many(calculations)'],
      note: '计算由受信任办公回调完成；单位、精度和算式必须在contract中有依据。',
    },
  },
  minimal_example: `def run(inputs):
    source = read_source(0)
    source_nodes = [node for node in source["nodes"] if node["ref"] == "来源表!B2"]
    if len(source_nodes) != 1:
        raise ValueError("来源表!B2必须唯一")
    target = read_target()
    target_slots = [slot for slot in target["target_slots"] if slot["target_ref"] == "目标表!B3"]
    if len(target_slots) != 1:
        raise ValueError("目标表!B3必须唯一")
    return write_cells([{"target_ref": "目标表!B3", "value": source_nodes[0]["value"]}], inputs["output_name"])
`,
  example_requirements: ['示例只演示回调形状；真实source_ref、target_ref、输出名和table_ranges必须符合本次contract。', 'contract、必填目标、写入授权、原件保护和独立检查仍由office_method_save/validate/run执行。'],
  reuse_candidate: 'office_reuse_pending(event_id)返回已经完整核对的confirmed_contract、当前输入是否仍一致及脚本版本摘要。可以把清单用于metadata_path，不需要重写来源关系。代码仍需适配上述回调；只有实际validate通过再enable，普通临时脚本保存成文件不等于插件可执行方法。',
};

const checklistHelp = {
  topic: 'checklist',
  scope: '这里只解释帮助中展示的compact contract摘要，不替代实际contract。',
  contract_schema: structuredClone(contractSchema),
  compact_contract: {
    coverage: ['目标表!B3:C4'],
    tables: [{
      sheet: '目标表', fields: [
        { column: 'B', source_column: 'B', requirement: { operation: 'copy' } },
        { column: 'C', source_column: 'C', requirement: { operation: 'copy' } },
      ],
      records: [
        { target_row: 3, source_index: 0, source_sheet: '来源表', source_row: 2 },
        { target_row: 4, source_index: 0, source_sheet: '来源表', source_row: 3 },
      ],
    }],
    preserve_template: true,
  },
  irregular_contract: {
    coverage: ['目标表!B3:C3'],
    tables: [{
      sheet: '目标表',
      fields: [{ column: 'B' }, { column: 'C' }],
      records: [{ target_row: 3, source_index: 0, source_sheet: '来源表', source_row: 8,
        bindings: { B: { operation: 'copy', source_refs: ['s0::来源表!D8'] }, C: { operation: 'blank', source_absence: { source_index: 0, sheet: '来源表', labels: ['数量'] } } } }],
    }],
    preserve_template: true,
  },
  rules: ['coverage与requirements必须覆盖本次实际要求。', '只能写入requirements允许的target_ref；已有值、保护单元格和表范围按contract处理。', 'validate/run会重读本次来源、独立检查产物；通过只证明清单范围，不证明未知业务。'],
};

export function officeHelp(topic, config, options = {}) {
  if (!OFFICE_HELP_TOPICS.includes(topic)) throw new Error('OFFICE_HELP_TOPIC_INVALID: 仅支持runtime、method_code、checklist');
  if (topic === 'runtime') return runtimeHelp(config, options);
  if (topic === 'method_code') return structuredClone(methodCodeHelp);
  return structuredClone(checklistHelp);
}

export function installOfficeHelp(register, config) {
  register('office_help', '按需查看固定运行命令、方法回调完整形状或compact contract摘要；只读，不执行代码。', {
    topic: { type: 'string', required: true, enum: OFFICE_HELP_TOPICS, description: topicDescription },
    script_path: { type: 'string', description: 'runtime主题可选的工作区相对Python脚本名；支持空格和单引号，不传则显示workspace relative.py' },
  }, async args => {
    if (args.script_path !== undefined && args.topic !== 'runtime') throw new Error('OFFICE_HELP_SCRIPT_PATH_TOPIC: script_path只用于runtime主题');
    return officeHelp(args.topic, config, { scriptPath: args.script_path });
  });
}


const text=(description,required=false)=>({type:'string',...(required?{required:true}:{}),description});
const matchSchema={type:'object',additionalProperties:false,properties:{
  sheet:text('来源工作表',true),value_column:text('所需值列字母，如D',true),start_row:{type:'number',description:'第一条数据行，默认2'},
  source_label:text('该值的直接列标签，可用于逐项核对；派生结果名不能当作各来源字段名'),
  key_columns:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{
    column:text('来源键列字母，如A',true),target_ref:text('原模板中本次主体/订单/期间值的位置，如结算汇总!A2',true)}}},
}};
export const contractSchema={type:'object',additionalProperties:false,properties:{
  coverage:{type:'array',required:true,description:'任务必须核对的数据位置，和requirements逐项位置一致。例如["企业资料!A2:H5"]；模板表头无需列入，程序另检查其保留。不能由已经写对的项反推范围。',items:{type:'string'}},
  requirements:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{
    target_ref:text('目标工作表!坐标或DOCX物理引用',true),source_refs:{type:'array',items:{type:'string'},description:'明确固定引用，例如["s0::工商资料!B3"]；与source_matches二选一'},
    source_matches:{type:'array',items:matchSchema,description:'来源行/文件顺序会变时，按原模板本次主体、订单、期间等键唯一查找。例sheet=采购明细,value_column=D,key_columns=[{column:A,target_ref:结算汇总!A2},{column:B,target_ref:结算汇总!B2}]。找不到或重复会拒绝；不能沿用首轮来源坐标。'},
    operation:{type:'string',enum:['copy','sum','product','difference','ratio','percent','extract','blank','formula','compose'],description:'copy原值；extract文字拆分；blank有依据留空；formula数据算式；compose分别核对组件后保留模板固定文字'},
    expected_text:text('compose本批模板原文，空格也须完整匹配；不沿用旧批文字'),
    text_template:text('compose显示格式，仅{0}至{15}，每组件恰好一次，如人民币{0}万元'),
    components:{type:'array',items:{type:'object',additionalProperties:true,properties:{}},description:'compose的1至16项独立来源关系，每项使用copy/extract/blank等既有参数，不带target_ref，不嵌套compose或自行声明答案'},
    source_label:text('只用于单来源值的直接标签，不能用于多来源计算；不能借历史相邻标签'),
    source_labels:{type:'array',items:{type:'string'},description:'多来源计算按source_refs/source_matches顺序逐项列字段标签，例如["数量","单价"]；不能把计算结果名称用于所有来源'},
    source_label_refs:{type:'array',items:{type:'string'},description:'标签与值分开时，按来源顺序明确标签物理引用，如值s0::基本信息!D13、标签s0::基本信息!C13；程序核对真实标签及同行/同列上下文，不猜相邻关系'},
    text_extract:{type:'object',additionalProperties:false,properties:{after:text('唯一开始边界，截取其后文字'),before:text('唯一结束边界，截取其前文字'),through:text('唯一结束边界，保留该文字；与before二选一'),cast:{type:'string',enum:['text','integer']},normalize_parentheses:{type:'boolean'}},description:'extract使用。例如72(2024年报)：人数before=(、cast=integer；年报after=(、before=年报)、cast=integer。重复边界拒绝'},
    blank_markers:{type:'array',items:{type:'string'},description:'blank明确列来源空值标记，如["-"]；真实0不能当缺失'},
    source_absence:{type:'object',additionalProperties:false,properties:{source_index:{type:'number',required:true},sheet:text('本次声明检查栏目所在工作表',true),labels:{type:'array',required:true,items:{type:'string'}}},description:'blank无来源值时，声明来源编号、工作表和待查栏目所有别名；程序读完文字格确认这些标签未出现。仅证明列明范围，不证明图片/未知业务缺失'},
    filename_identity:{type:'boolean',description:'任务允许用文件名作身份线索时才为true；source_refs可s0::$filename并extract。禁止由文件名推断金额、联系方式等事实'},
    expression:{type:'object',additionalProperties:true,properties:{},description:'formula数据算式树：{source:0}取第一个来源；{constant:2}明确常数；{operator:ratio,args:[{source:0},{operator:average,args:[{source:1},{source:2}]}]}计算平均余额比。仅和/平均/积/差/比/百分比，无Python执行'},
    source_unit:text('元/千元/万元/亿元'),target_unit:text('元/千元/万元/亿元'),
    decimals:{type:'number',description:'0至10整数，十进制四舍五入'},source_unit_ref:text('来源单位证据，例如s0::表!C1'),
    absolute_tolerance:{type:'number',description:'仅Excel未设decimals的ratio/percent允许明确的绝对误差，须大于0且不超过1e-12；金额仍严格核对'},
    render:{type:'string',enum:['fixed_decimal','percent_text'],description:'Word数值显示要求，需同时给decimals；fixed_decimal固定小数位，percent_text另加百分号'},
    target_unit_source_ref:text('目标单位若由来源规定，可填写来源证据'),target_unit_ref:text('输出中固定单位标签位置'),
    subject_source_ref:text('本次来源主体引用'),subject_target_ref:text('目标该记录的主体位置'),
  }}},
  table_ranges:{type:'object',additionalProperties:true,properties:{},description:'对象，例如{"企业资料!企业资料表":"A1:H5"}。不扩表可以省略；不要传数组。'},
  preserve_template:{type:'boolean',description:'默认true，检查所支持模板属性；可执行方法禁止false'},
  allow_existing_values:{type:'boolean',description:'任务明确要求按来源更新已有值时才为true；默认不改不同的已有值'},
  preserve_only:{type:'boolean',description:'整组交付中此文件本次只保留：coverage及requirements必须为空，候选与模板全部字节必须相同；不计为已填写字段'},
}};
export const conditionsSchema={type:'object',additionalProperties:false,properties:{
  source_count:{type:'number',required:true,description:'适用来源数量，1至8整数'},
  source_formats:{type:'array',required:true,items:{type:'string'},description:'按编号排列，例如["xlsx","docx"]'},
  target_format:text('方法写入支持xlsx及普通docx表格/段落占位；复杂结构继续由DSH处理',true),
  source_guards:{type:'array',required:true,description:'检查必要结构/标签，至少覆盖每份来源；不保存当次客户、金额等事实',items:{type:'object',additionalProperties:false,properties:{
    source_index:{type:'number',required:true},ref:text('来源位置，例如工商资料!A2',true),text:text('必须出现的结构标签，例如企业名称',true)}}},
  target_guards:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{ref:text('模板标签位置，例如企业资料!A1',true),text:text('必须出现的模板标签',true)}}},
}};

// The existing explicit checklist remains available, including Word positions.
// Confirmed repeated table relationships may instead be expanded by the host.
delete contractSchema.properties.requirements.required;
const bindingSchema={type:'object',additionalProperties:true,properties:{},description:'来源及算式字段沿用requirements项，省略target_ref；宿主逐字段验证'};
contractSchema.properties.tables={type:'array',description:'可选：按已确认字段及记录展开完整要求；与requirements可组合。不会猜来源。示例用office_help(topic=checklist)。',items:{type:'object',additionalProperties:false,properties:{
  sheet:text('目标工作表',true),
  fields:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{
    column:text('目标列字母',true),source_column:text('规则表格的来源值列，可选'),source_label_column:text('同行标签列，可选'),requirement:bindingSchema,
  }}},
  records:{type:'array',required:true,items:{type:'object',additionalProperties:false,properties:{
    target_row:{type:'number',required:true},source_index:{type:'number'},source_sheet:text('规则来源工作表'),source_row:{type:'number'},
    bindings:{type:'object',additionalProperties:true,properties:{},description:'不同布局时按目标列声明本条记录的实际来源/缺失依据，覆盖该字段公共要求；宿主验证每项'},
  }}},
}}};

const shortDescriptions={
  target_ref:'目标物理位置，如表!A2或Word引用',source_refs:'本次固定来源引用，如s0::资料!B3；与source_matches二选一',
  source_matches:'按本次模板键唯一查来源，不沿用历史坐标',operation:'复制、明确算式、文字拆分或有依据留空',
  source_label:'单来源值的真实直接标签',source_labels:'按来源顺序逐项直接标签，多来源计算用此项',
  source_label_refs:'按来源顺序列真实标签位置，标签和值分开时使用',
  text_extract:'唯一实际文字边界；不传空字符串，多边界拒绝',blank_markers:'明确来源空标记，如["-"]；0是正常值',
  source_absence:'blank且确实无栏目时，声明来源编号/工作表/所有标签别名，程序读完核实',
  filename_identity:'任务允许文件名仅作主体线索时true',expression:'有限数据算式树，详例用office_help',
  source_unit_ref:'实际来源单位值位置，如s0::表!C1',render:'Word显示方式，需decimals',
  target_unit_source_ref:'目标单位由来源规定时的依据位置',target_unit_ref:'产物真实单位值位置',
};
for(const [key,description] of Object.entries(shortDescriptions))contractSchema.properties.requirements.items.properties[key].description=description;
contractSchema.properties.coverage.description='由任务确认的完整数据位置，如["表!A2:H5"]；与展开后的要求相符。不由已填项反推';
contractSchema.properties.table_ranges.description='对象{"表!表名":"A1:H5"}，不扩表可省略';
contractSchema.properties.allow_existing_values.description='任务明确要求更新已有值时true';
conditionsSchema.properties.source_guards.description='每份来源必要结构/标签，去掉当次客户事实';

export const contractInputSchema={type:'json',description:'完整清单对象：coverage与requirements或tables，可用清单文件。宿主核对类型、覆盖和来源；详细结构和示例用office_help(topic=checklist)按需取得'};

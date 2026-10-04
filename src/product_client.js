window.__ModuleLoader__.load({
  id:'dsh-office-workflows',
  factory(require){
    var module={exports:{}};var exports=module.exports;const React=require('react');
    const FIELDS=[
      ['workspaceRoot','工作区根目录','办公文件只能在此工作区内处理。'],['pythonPath','Python 解释器','使用主程序已核对的解释器。'],
      ['methodDirectory','方法目录','保存已验证的可复用程序。'],['nativeStateDirectory','运行状态目录','保存每次运行的状态，便于核对和恢复。'],
      ['workspaceMode','工作区模式','session 使用当前会话目录；fixed 使用固定目录。'],['nativeScriptMode','普通 Python 程序','默认关闭；个人电脑验证可选兼容模式。'],
      ['learningEnabled','允许学习记录','执行与修正会留下待检查程序；检查通过才入库。'],['captureOrdinaryPython','记录普通 Python','记录普通 Python 程序的结构化运行信息，供后续核对。'],['outputCheckTools','启用输出检查','发布前独立核对实际文件。']
    ];
    const PATHS=FIELDS.slice(0,4).map(x=>x[0]), LABELS=Object.fromEntries(FIELDS.map(([key,label])=>[key,label])), BOOL=new Set(['learningEnabled','captureOrdinaryPython','outputCheckTools']);
    const initial=s=>Object.fromEntries(FIELDS.map(([key])=>[key,BOOL.has(key)?Boolean(s?.value?.[key]):String(s?.value?.[key]??'')]));
    const missing=d=>PATHS.filter(key=>!String(d[key]??'').trim());
    const ops=d=>FIELDS.map(([key])=>({op:'set',path:[key],value:BOOL.has(key)?Boolean(d[key]):String(d[key]??'').trim()}));
    async function saveDraft(form,snapshot,draft){
      const absent=missing(draft);if(absent.length)return {saved:false,error:'请先填写四个路径：'+absent.map(key=>LABELS[key]).join('、')};
      try {const landed=await form.mutate(ops(draft),snapshot.revision);return landed?{saved:true}:{saved:false,error:`保存被拒绝：修订 ${snapshot.revision} 已变化，或主程序未接受这些值。`};}
      catch(error){return {saved:false,error:'保存失败：'+(error?.message||String(error))};}
    }
    function FormComponent(props){
      const form=props.officeForm,view=props.view;
      if(!form||typeof form.getSnapshot!=='function'||typeof form.subscribe!=='function'||typeof form.mutate!=='function')return React.createElement('section',null,React.createElement('p',{role:'alert'},'办公配置表单尚未连接，请重新打开插件配置页。'));
      const snapshot=React.useSyncExternalStore(cb=>form.subscribe(cb),()=>form.getSnapshot(),()=>form.getSnapshot());
      const [draft,setDraft]=React.useState(()=>initial(snapshot)),[baseline,setBaseline]=React.useState(()=>JSON.stringify(initial(snapshot))),[busy,setBusy]=React.useState(false),[notice,setNotice]=React.useState('');
      const current=JSON.stringify(initial(snapshot)),dirty=JSON.stringify(draft)!==baseline,readOnly=!snapshot?.writable;
      React.useEffect(()=>{if(!dirty){setDraft(initial(snapshot));setBaseline(current);}},[current]);
      if(snapshot.status!=='ready'||!snapshot.value)return React.createElement('p',{role:'alert'},'配置正在连接。若持续显示此提示，请返回插件列表后重新打开配置页。');
      if(view!=='page')return React.createElement('span',null,'配置办公工作区、解释器、方法与独立输出检查。');
      const edit=(key,value)=>{setDraft(next=>({...next,[key]:value}));setNotice('');};
      const save=async()=>{if(busy||readOnly)return;setBusy(true);setNotice('');const result=await saveDraft(form,snapshot,draft);if(result.saved){setBaseline(JSON.stringify(draft));setNotice('配置已保存。');}else setNotice(result.error);setBusy(false);};
      const pathMissing=missing(draft),disabled=busy||readOnly,saveDisabled=disabled||pathMissing.length>0;
      const controls=FIELDS.map(([key,label,hint])=>{
        const common={id:'office-workflows-'+key,disabled,onChange:e=>edit(key,BOOL.has(key)?e.target.checked:e.target.value)};
        let input;
        if(BOOL.has(key))input=React.createElement('input',{...common,type:'checkbox',checked:Boolean(draft[key])});
        else if(key==='workspaceMode')input=React.createElement('select',{...common,value:draft[key]||'session'},React.createElement('option',{value:'session'},'session：当前会话'),React.createElement('option',{value:'fixed'},'fixed：固定目录'));
        else if(key==='nativeScriptMode')input=React.createElement('select',{...common,value:draft[key]||'disabled'},React.createElement('option',{value:'disabled'},'disabled：关闭'),React.createElement('option',{value:'personal-compatible'},'personal-compatible：个人电脑兼容验证'));
        else input=React.createElement('input',{...common,type:'text',value:draft[key],placeholder:'绝对路径',style:{width:'100%',boxSizing:'border-box'}});
        return React.createElement('label',{key,style:{display:'block',margin:'0 0 12px'}},React.createElement('strong',null,label),input,React.createElement('small',null,hint));
      });
      return React.createElement('section',null,React.createElement('p',null,pathMissing.length?'办公工具关闭：请填写 '+pathMissing.map(key=>LABELS[key]).join('、')+'。':'路径已填写；仍须由主程序核对后启用。'),controls,React.createElement('button',{type:'button',disabled:saveDisabled,onClick:save},busy?'保存中…':'保存'),readOnly?React.createElement('p',null,'当前配置为只读。'):null,notice?React.createElement('p',{role:'status'},notice):null);
    }
    const inject=['slots','locale','configForms'];
    function apply(ctx){
      const form=ctx.configForms.get('office-workflows');
      ctx.effect(()=>ctx.configForms.whileServed(['office-workflows'],()=>ctx.slots.inject('plugins.row.config',()=>ctx.slots.register({name:'plugins.row.config',key:'dsh-office-workflows#office-workflows',locale:'officeWorkflows',inject:()=>({officeForm:form})},FormComponent))));
    }
    exports.inject=inject;exports.apply=apply;exports.FormComponent=FormComponent;exports.saveDraft=saveDraft;exports.ops=ops;return module.exports;
  }
});

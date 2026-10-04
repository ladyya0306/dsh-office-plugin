import { resolve } from 'node:path';

// Parse one literal argv command. Never evaluate PowerShell expressions.
export function ordinaryPythonCommand(command, pythonPath) {
  if (typeof command !== 'string' || command.length > 16000) return null;
  const trimmed=command.trim().replace(/(?:;|\r?\n)\s*exit\s+\$LASTEXITCODE\s*$/i,'');
  const tokens=[];let offset=0;
  while(offset<trimmed.length){
    if(/\s/.test(trimmed[offset])){offset++;continue;}
    const quote=trimmed[offset];let value='';
    if(quote==="'"||quote==='"'){
      offset++;let closed=false;
      while(offset<trimmed.length){
        const c=trimmed[offset++];
        if(c===quote){if(quote==="'"&&trimmed[offset]==="'"){value+="'";offset++;}else{closed=true;break;}}
        else {if(quote==='"'&&/[$`]/.test(c))return null;value+=c;}
      }
      if(!closed||(offset<trimmed.length&&!/\s/.test(trimmed[offset])))return null;
    }else{
      const start=offset;while(offset<trimmed.length&&!/\s/.test(trimmed[offset]))offset++;
      value=trimmed.slice(start,offset);
      if(value!=='&'&&/[;&|<>$`{}()]/.test(value))return null;
    }
    if(/[\x00\r\n]/.test(value))return null;tokens.push(value);
  }
  if(tokens[0]==='&')tokens.shift();
  if(!tokens.length||resolve(tokens.shift()).toLowerCase()!==resolve(pythonPath).toLowerCase())return null;
  const flags=[];
  while(tokens[0]?.startsWith('-')){
    const flag=tokens.shift();if(['-B','-I','-u'].includes(flag)){flags.push(flag);continue;}
    if(flag==='-X'&&tokens[0]==='utf8'){flags.push(flag,tokens.shift());continue;}
    return null;
  }
  const script=tokens.shift();if(!script?.toLowerCase().endsWith('.py')||tokens.length>64)return null;
  return {script,argv:tokens,flags};
}

// Only classify known read-only operations; never replay writes or authorization.
export function classifyReuse(previous, current) {
  if (!previous || previous.kind !== 'read' || current.kind !== 'read') return {reuse:false,reason:'not_known_read_only'};
  for (const key of ['workspace','parser','version','range','purpose']) {
    if (JSON.stringify(previous[key]) !== JSON.stringify(current[key])) return {reuse:false,reason:key+'_changed'};
  }
  if (!previous.complete || !previous.success) return {reuse:false,reason:'previous_result_incomplete'};
  if (current.corrected || current.afterWrite || current.scriptChanged) return {reuse:false,reason:'new_check_required'};
  return {reuse:true,reason:'same_verified_read'};
}


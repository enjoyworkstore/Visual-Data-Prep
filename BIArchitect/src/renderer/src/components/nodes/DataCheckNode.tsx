import { memo } from 'react';
import NodeInput from './NodeInput';
import { NodeWrap, type NodeStatusTone, useNodeLogic } from './shared';
import { getCheckOperatorLabel } from './dataCheckUtils';

const DataCheckNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const result = fData.checkResult || { count: 0, rows: [], hasMatches: false, isConfigured: false };
  const op = getCheckOperatorLabel(data.checkType);
  const summary = result.isConfigured ? (result.hasMatches ? `NG ${result.count}件` : 'OK 0件') : '';
  const statusTone: NodeStatusTone = result.isConfigured ? (result.hasMatches ? 'red' : 'blue') : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] hover:border-blue-400' : 'bg-white border-gray-300 hover:border-blue-500'}`;
  const titleClass = result.isConfigured
    ? (result.hasMatches ? (isDark ? 'text-rose-400' : 'text-rose-600') : (isDark ? 'text-sky-400' : 'text-sky-600'))
    : (isDark ? 'text-blue-400' : 'text-blue-600');
  const resultCardClass = result.hasMatches
    ? (isDark ? 'bg-rose-950/30 border-rose-500/40' : 'bg-rose-50 border-rose-200')
    : (isDark ? 'bg-sky-950/30 border-sky-500/40' : 'bg-sky-50 border-sky-200');
  const resultTitleClass = result.hasMatches
    ? (isDark ? 'text-rose-300' : 'text-rose-700')
    : (isDark ? 'text-sky-300' : 'text-sky-700');

  return (
    <NodeWrap
      id={id}
      data={data}
      title="Data Check"
      col={titleClass}
      summary={summary}
      statusTone={statusTone}
      helpText="指定した条件に一致するデータが存在するかをチェックします。一致行があれば件数と対象行を表示し、ノードは赤色になります。該当なしなら青色になります。"
    >
      <div className="space-y-2">
        <select className={`${inputClass} ${isDark ? 'text-[#ccc]' : 'text-gray-700'}`} value={data.checkCol || ''} onChange={(e) => onChg('checkCol', e.target.value)}>
          <option value="">Target Column...</option>
          {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}
        </select>
        <div className="flex flex-col gap-2">
          <select className={`${inputClass} ${isDark ? 'text-white' : 'text-gray-900'} font-bold`} value={data.checkType || 'includes'} onChange={(e) => onChg('checkType', e.target.value)}>
            <option value="includes">含む (Includes)</option>
            <option value="exact">完全一致 (=)</option>
            <option value="not">除外 (≠)</option>
            <option value="gt">以上 (&gt;)</option>
            <option value="lt">以下 (&lt;)</option>
          </select>
          <NodeInput
            className={`w-full text-[10px] p-2 border rounded outline-none transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500'}`}
            placeholder="Condition Value..."
            value={data.checkVal || ''}
            onChange={(v) => onChg('checkVal', v)}
          />
        </div>
        {result.isConfigured && (
          <div className={`rounded-lg border p-3 space-y-2 ${resultCardClass}`}>
            <div className={`text-[10px] font-bold flex items-center justify-between ${resultTitleClass}`}>
              <span>{result.hasMatches ? '条件該当あり' : '条件該当なし'}</span>
              <span>{result.count}件</span>
            </div>
            <div className={`text-[9px] ${isDark ? 'text-[#aaa]' : 'text-gray-600'}`}>条件: {data.checkCol} {op} {data.checkVal || ''}</div>
            {result.rows?.length > 0 && (
              <div className="space-y-1 max-h-32 overflow-y-auto custom-scrollbar pr-1">
                {result.rows.map((row: any, idx: number) => (
                  <div key={idx} className={`text-[9px] font-mono rounded border px-2 py-1.5 break-all ${isDark ? 'bg-[#111] border-[#333] text-[#ddd]' : 'bg-white border-gray-200 text-gray-700'}`}>
                    {Object.entries(row).map(([key, cellValue]) => `${key}: ${String(cellValue)}`).join(' | ')}
                  </div>
                ))}
                {result.count > result.rows.length && (
                  <div className={`text-[8px] text-center ${isDark ? 'text-[#666]' : 'text-gray-500'}`}>他 {result.count - result.rows.length} 件</div>
                )}
              </div>
            )}
          </div>
        )}
      </div>
    </NodeWrap>
  );
});

DataCheckNode.displayName = 'DataCheckNode';

export default DataCheckNode;

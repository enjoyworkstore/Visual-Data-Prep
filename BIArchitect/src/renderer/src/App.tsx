import React, { useState, useCallback, useMemo, memo, createContext, useContext, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { ReactFlow, Controls, Background, addEdge, Handle, Position, ReactFlowProvider, useReactFlow, useNodesState, useEdgesState, Panel } from '@xyflow/react';
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import type { Node, Edge } from '@xyflow/react';
import '@xyflow/react/dist/style.css';

// ★ 追加: TypeScriptの厳格なエラーを防ぐため、どんなデータでも持てるカスタムノード型を定義
type CustomNode = Node<Record<string, any>>;

type AppContextType = {
  workbooks: Record<string, XLSX.WorkBook>;
  setWorkbooks: React.Dispatch<React.SetStateAction<Record<string, XLSX.WorkBook>>>;
  setRangeModalNode: React.Dispatch<React.SetStateAction<string | null>>;
  nodeFlowData: Record<string, any>;
  isAutoCameraMove: boolean;
};
const AppContext = createContext<AppContextType>({} as AppContextType);

const GlobalStyle = () => (
  <style>{`
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #1a1a1a; }
    ::-webkit-scrollbar-thumb { background: #444; border-radius: 4px; }
    .react-flow__node { cursor: grab !important; }
    .react-flow__node:active { cursor: grabbing !important; }
    .react-flow__handle { width: 18px !important; height: 18px !important; border: 3px solid #1e1e1e !important; background-color: #38bdf8 !important; transition: transform 0.1s ease; }
    .react-flow__handle:hover { transform: scale(1.5); }
    .custom-scrollbar::-webkit-scrollbar { width: 4px; }
    select {
      appearance: none;
      background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%2338bdf8%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
      background-repeat: no-repeat, repeat;
      background-position: right .7em top 50%, 0 0;
      background-size: .65em auto, 100%;
    }
  `}</style>
);

const IconSvg = ({ children }: { children: React.ReactNode }) => (
  <svg viewBox="0 0 24 24" width="1em" height="1em" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round">{children}</svg>
);

const Icons = {
  Source: <IconSvg><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/></IconSvg>,
  Union: <IconSvg><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></IconSvg>,
  Join: <IconSvg><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></IconSvg>,
  GroupBy: <IconSvg><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></IconSvg>,
  Sort: <IconSvg><path d="m3 16 4 4 4-4"/><path d="M7 20V4"/><path d="m21 8-4-4-4 4"/><path d="M17 4v16"/></IconSvg>,
  Transform: <IconSvg><path d="M20 7h-9"/><path d="M14 17H5"/><circle cx="17" cy="7" r="3"/><circle cx="8" cy="17" r="3"/></IconSvg>,
  Select: <IconSvg><path d="m3 17 2 2 4-4"/><path d="m3 7 2 2 4-4"/><path d="M13 6h8"/><path d="M13 12h8"/><path d="M13 18h8"/></IconSvg>,
  Filter: <IconSvg><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></IconSvg>,
  Chart: <IconSvg><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></IconSvg>,
  Dashboard: <IconSvg><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="21" x2="9" y2="9"/></IconSvg>,
  Warning: <IconSvg><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></IconSvg>,
  Folder: <IconSvg><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></IconSvg>,
  File: <IconSvg><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></IconSvg>,
  Refresh: <IconSvg><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><polyline points="3 3 3 8 8 8"/></IconSvg>,
  Trash: <IconSvg><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></IconSvg>,
  Save: <IconSvg><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></IconSvg>,
  Close: <IconSvg><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></IconSvg>,
  Diamond: <IconSvg><polygon points="12 2 22 8.5 22 15.5 12 22 2 15.5 2 8.5 12 2"/></IconSvg>,
  ArrowLeft: <IconSvg><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></IconSvg>,
  ArrowRight: <IconSvg><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></IconSvg>,
  ChevronDown: <IconSvg><polyline points="6 9 12 15 18 9"/></IconSvg>,
  ChevronUp: <IconSvg><polyline points="18 15 12 9 6 15"/></IconSvg>,
  Code: <IconSvg><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></IconSvg>,
  Copy: <IconSvg><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></IconSvg>,
  Focus: <IconSvg><circle cx="12" cy="12" r="3"/><path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/></IconSvg>
};

const CAMERA_OFFSET_X = 250;
const CAMERA_OFFSET_Y = 100;

// ★ IME入力時の二重入力バグを防ぐため、フォーカスが外れた時のみ確定する専用Inputコンポーネント
const NodeInput = memo(({ value, onChange, placeholder, className }: any) => {
  const [val, setVal] = useState(value || '');
  useEffect(() => { setVal(value || ''); }, [value]);
  return (
    <input 
      type="text" 
      className={`nodrag ${className}`} 
      placeholder={placeholder} 
      value={val} 
      onChange={(e) => setVal(e.target.value)} 
      onBlur={() => onChange(val)} 
      onKeyDown={(e) => { if (e.key === 'Enter') { e.currentTarget.blur(); } }}
    />
  );
});

const calcData = (nId: string, nodes: CustomNode[], edges: Edge[], wbs: any): { data: any[], headers: string[] } => {
  const node = nodes.find(n => n.id === nId);
  if (!node) return { data: [], headers: [] };

  if (node.type === 'dataNode') {
    if (node.data.needsUpload) return { data: [], headers: [] };
    const wb = wbs[node.id];
    if (!wb) return { data: [], headers: [] };
    const ws = wb.Sheets[node.data.currentSheet || wb.SheetNames[0]];
    if (!ws) return { data: [], headers: [] };
    const rngs = node.data.ranges || [];
    const useHdr = node.data.useFirstRowAsHeader !== false;

    try {
      const mat = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: true }) as any[][];
      if (!mat || mat.length === 0) return { data: [], headers: [] };
      let tRngs = rngs.length === 0 && ws['!ref'] ? [XLSX.utils.encode_range(XLSX.utils.decode_range(ws['!ref']))] : rngs;
      let fHdrs: string[] = [], cData: any[] = [];

      tRngs.forEach((rStr: string, idx: number) => {
        const rObj = XLSX.utils.decode_range(rStr);
        const extRows: any[][] = [];
        for (let r = rObj.s.r; r <= rObj.e.r; r++) extRows.push(Array.from({ length: rObj.e.c - rObj.s.c + 1 }, (_, i) => (mat[r] || [])[rObj.s.c + i] ?? ""));
        if (extRows.length === 0) return;

        if (idx === 0) {
          fHdrs = useHdr ? extRows[0].map((h, i) => h ? String(h) : `Col_${i+1}`) : Array.from({length: extRows[0].length}, (_, i) => `Col_${i+1}`);
          for (let i = useHdr ? 1 : 0; i < extRows.length; i++) {
            const obj: any = {}; fHdrs.forEach((h, cIdx) => obj[h] = extRows[i][cIdx]); cData.push(obj);
          }
        } else {
          for (let i = 0; i < extRows.length; i++) {
            const obj: any = {}; fHdrs.forEach((h, cIdx) => obj[h] = extRows[i][cIdx]); cData.push(obj);
          }
        }
      });
      return { data: cData, headers: fHdrs };
    } catch (e) { return { data: [], headers: [] }; }
  }

  if (node.type === 'unionNode' || node.type === 'joinNode') {
    const eA = edges.find(e => e.target === nId && (e as any).targetHandle === 'input-a');
    const eB = edges.find(e => e.target === nId && (e as any).targetHandle === 'input-b');
    if (!eA || !eB) return { data: [], headers: [] };
    const rA = calcData(eA.source, nodes, edges, wbs), rB = calcData(eB.source, nodes, edges, wbs);
    if (node.type === 'unionNode') return { data: [...rA.data, ...rB.data], headers: rA.headers };
    
    const { keyA, keyB, joinType = 'inner' } = node.data;
    if (!keyA || !keyB) return rA;

    let jnd: any[] = [];
    if (joinType === 'inner') {
      rA.data.forEach(a => { const b = rB.data.find(r => String(r[keyB as string]) === String(a[keyA as string])); if (b) jnd.push({ ...a, ...b }); });
    } else if (joinType === 'left') {
      rA.data.forEach(a => { const b = rB.data.find(r => String(r[keyB as string]) === String(a[keyA as string])) || {}; jnd.push({ ...a, ...b }); });
    } else if (joinType === 'right') {
      rB.data.forEach(b => { const a = rA.data.find(r => String(r[keyA as string]) === String(b[keyB as string])) || {}; jnd.push({ ...a, ...b }); });
    }
    const newHeaders = [...new Set([...rA.headers, ...rB.headers])];
    return { data: jnd, headers: newHeaders };
  }

  const inEdge = edges.find(e => e.target === nId);
  if (!inEdge) return { data: [], headers: [] };
  const input = calcData(inEdge.source, nodes, edges, wbs);
  let out = [...input.data], h = [...input.headers];

  if (node.type === 'sortNode') {
    const { sortCol, sortOrder } = node.data;
    if (sortCol) out.sort((a, b) => sortOrder === 'desc' ? String(b[sortCol as string]).localeCompare(String(a[sortCol as string]), undefined, { numeric: true }) : String(a[sortCol as string]).localeCompare(String(b[sortCol as string]), undefined, { numeric: true }));
  }

  if (node.type === 'filterNode') {
    const { filterCol, filterVal, matchType = 'includes' } = node.data;
    if (filterCol && filterVal !== undefined && filterVal !== '') {
      out = out.filter(r => {
        const cVal = String(r[filterCol as string] || '').toLowerCase(), tVal = String(filterVal).toLowerCase();
        const cNum = Number(r[filterCol as string]), tNum = Number(filterVal);
        switch (matchType) {
          case 'exact': return cVal === tVal;
          case 'not': return cVal !== tVal;
          case 'gt': return (!isNaN(cNum) && !isNaN(tNum)) ? cNum > tNum : cVal > tVal;
          case 'lt': return (!isNaN(cNum) && !isNaN(tNum)) ? cNum < tNum : cVal < tVal;
          default: return cVal.includes(tVal);
        }
      });
    }
  }

  if (node.type === 'selectNode') {
    const sel = node.data.selectedColumns || [];
    if (sel.length > 0) { h = sel; out = out.map(r => { const nr: any = {}; sel.forEach((c: string) => nr[c] = r[c]); return nr; }); }
  }

  if (node.type === 'groupByNode') {
    const { groupCol, aggCol, aggType } = node.data;
    if (groupCol && aggCol && out.length > 0) {
      const grps: Record<string, any> = {};
      out.forEach(r => {
        const k = r[groupCol as string];
        if (!grps[k]) grps[k] = { [groupCol as string]: k, _v: 0, _c: 0 };
        grps[k]._v += Number(r[aggCol as string]) || 0; grps[k]._c++;
      });
      out = Object.values(grps).map((g: any) => ({ [groupCol as string]: g[groupCol as string], [aggCol as string]: aggType === 'count' ? g._c : (g._v / (g._c || 1)) }));
      h = [groupCol, aggCol];
    }
  }

  if (node.type === 'transformNode') {
    const { targetCol, command, param0 } = node.data;
    if (targetCol && command) out = out.map(r => {
      let v = String(r[targetCol as string] || "");
      if (command === 'replace') v = v.replace(param0 || "", "");
      if (command === 'math_mul') v = String(Number(v) * Number(param0 || 1));
      if (command === 'add_suffix') v += (param0 || "");
      
      if (command === 'case_when') {
        let match = false;
        const cVal = v.toLowerCase(), tVal = String(node.data.condVal || '').toLowerCase();
        const cNum = Number(v), tNum = Number(node.data.condVal);
        switch (node.data.condOp) {
          case 'exact': match = (cVal === tVal); break;
          case 'not': match = (cVal !== tVal); break;
          case 'gt': match = (!isNaN(cNum) && !isNaN(tNum)) ? cNum > tNum : cVal > tVal; break;
          case 'lt': match = (!isNaN(cNum) && !isNaN(tNum)) ? cNum < tNum : cVal < tVal; break;
          default: match = cVal.includes(tVal);
        }
        v = match ? (node.data.trueVal || '') : (node.data.falseVal || '');
      }

      return { ...r, [targetCol as string]: v };
    });
  }

  return { data: out, headers: h };
};

const useNodeLogic = (id: string) => {
  const { nodeFlowData, isAutoCameraMove } = useContext(AppContext);
  const { updateNodeData, setNodes, setEdges, getNode, getEdges, setCenter } = useReactFlow();
  
  return { 
    fData: nodeFlowData[id] || { incomingHeaders: [], headersA: [], headersB: [] }, 
    onChg: (k: string, v: any) => updateNodeData(id, { [k]: v }), 
    onDel: () => { 
      if (isAutoCameraMove) {
        const edges = getEdges();
        const incomingEdge = edges.find((e: any) => e.target === id);
        if (incomingEdge) {
          const prevNode = getNode(incomingEdge.source) as CustomNode | undefined;
          if (prevNode) {
            setCenter(prevNode.position.x + CAMERA_OFFSET_X, prevNode.position.y + CAMERA_OFFSET_Y, { zoom: 1.1, duration: 600 });
          }
        }
      }
      setNodes((nds: any) => nds.filter((n: any) => n.id !== id)); 
      setEdges((eds: any) => eds.filter((e: any) => e.source !== id && e.target !== id)); 
    } 
  };
};

const TgtHandle = ({ id, style }: any) => <Handle type="target" position={Position.Left} id={id} style={style} className="w-5 h-5 bg-[#252526] border-[3px] border-blue-400 hover:bg-blue-400 hover:scale-125 transition-all cursor-crosshair z-10 -ml-2 nodrag" />;
const SrcHandle = ({ col }: any) => <Handle type="source" position={Position.Right} className={`w-5 h-5 ${col} border-[3px] border-[#1e1e1e] hover:scale-125 transition-transform cursor-crosshair z-10 -mr-2 nodrag`} />;

const NodeWrap = memo(({ id, title, col, children, showTgt = true, multi = false, summary = '' }: any) => {
  const { onDel } = useNodeLogic(id);
  return (
    <div className="bg-[#252526] border border-[#444] rounded-xl shadow-2xl min-w-[260px] pb-1 relative group">
      <button onClick={(e) => { e.stopPropagation(); onDel(); }} className="absolute -top-3 -right-3 bg-blue-600 hover:bg-blue-500 text-white rounded-full w-6 h-6 flex items-center justify-center font-bold text-xs opacity-0 group-hover:opacity-100 transition-opacity shadow-lg z-20 nodrag">
        <span className="flex items-center justify-center w-3 h-3">{Icons.Close}</span>
      </button>
      <div className="bg-[#1a1a1a] p-2 border-b border-[#444] flex justify-between items-center rounded-t-xl select-none">
        <span className={`text-[10px] font-bold tracking-widest uppercase ${col}`}>{title}</span>
        {summary && <span className="text-[9px] bg-[#333] text-[#aaa] px-2 py-0.5 rounded-full max-w-[120px] truncate font-mono" title={summary}>{summary}</span>}
      </div>
      <div className="p-4 relative flex flex-col gap-3">
        {multi ? <><TgtHandle id="input-a" style={{ top: '30%' }} /><TgtHandle id="input-b" style={{ top: '70%' }} /></> : (showTgt && <TgtHandle />)}
        {children}
      </div>
      <SrcHandle col={col.replace('text-', 'bg-')} />
    </div>
  );
});

const DataNode = memo(({ id, data }: any) => {
  const { setWorkbooks, setRangeModalNode } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const onUp = (e: any) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader(); r.onload = (evt: any) => {
      const wb = XLSX.read(evt.target.result, { type: 'binary' });
      setWorkbooks((p: any) => ({ ...p, [id]: wb }));
      updateNodeData(id, { fileName: f.name, sheetNames: wb.SheetNames, currentSheet: wb.SheetNames[0], needsUpload: false });
    }; r.readAsBinaryString(f);
  };
  const summary = data.fileName ? data.fileName : '';
  return (
    <NodeWrap id={id} title="Source" col="text-blue-400" showTgt={false} summary={summary}>
      {data.needsUpload ? (
        <div className="space-y-3">
          <div className="text-[10px] text-white flex items-center gap-2 bg-blue-500/20 p-2 rounded border border-blue-500/50">
            <span className="text-blue-400 flex items-center justify-center">{Icons.Warning}</span> Missing: {data.fileName}
          </div>
          <label className="cursor-pointer text-blue-400 text-[10px] border border-dashed border-blue-500/50 p-3 rounded flex items-center justify-center gap-2 hover:bg-blue-500/20 font-bold uppercase transition-all shadow-lg animate-pulse nodrag">
            <span className="flex items-center justify-center">{Icons.Folder}</span> 再設定 <input type="file" accept=".csv,.xlsx" className="hidden" onChange={onUp} />
          </label>
        </div>
      ) : !data.fileName ? (
        <label className="cursor-pointer text-blue-400 text-[10px] border border-dashed border-blue-500/50 p-4 rounded flex items-center justify-center gap-2 hover:bg-blue-500/10 font-bold uppercase transition-all nodrag">
          <span className="flex items-center justify-center w-4 h-4">{Icons.Folder}</span> Load File <input type="file" accept=".csv,.xlsx" className="hidden" onChange={onUp} />
        </label>
      ) : (
        <div className="space-y-3">
          <div className="flex justify-between items-center bg-[#1a1a1a] p-2 rounded border border-[#333]">
            <div className="text-[10px] text-white font-bold truncate flex items-center gap-2">
              <span className="text-blue-400 flex items-center justify-center">{Icons.File}</span> {data.fileName}
            </div>
            <label className="cursor-pointer text-blue-400 hover:text-white text-[12px] font-bold uppercase transition-colors nodrag" title="Change File">
              <span className="flex items-center justify-center">{Icons.Refresh}</span> <input type="file" className="hidden" onChange={onUp} />
            </label>
          </div>
          <button onClick={() => setRangeModalNode(id)} className="w-full py-2 bg-blue-600/20 text-blue-400 text-[10px] font-bold rounded border border-blue-500/30 hover:bg-blue-600/40 uppercase tracking-widest transition-all nodrag">Visual Range</button>
          <label className="flex items-center gap-2 pt-2 cursor-pointer group"><input type="checkbox" checked={data.useFirstRowAsHeader !== false} onChange={(e) => updateNodeData(id, { useFirstRowAsHeader: e.target.checked })} className="accent-blue-500 w-4 h-4 cursor-pointer nodrag" /><span className="text-[10px] text-[#aaa] group-hover:text-white font-bold uppercase">1st row as Header</span></label>
        </div>
      )}
    </NodeWrap>
  );
});

const UnionNode = memo(({ id }: any) => <NodeWrap id={id} title="Union" col="text-blue-400" multi={true} summary="Append"><div className="text-[10px] text-[#888] text-center italic tracking-widest uppercase py-2">Merge Vertically</div></NodeWrap>);

const JoinNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const summary = data.keyA && data.keyB ? `${data.joinType || 'INNER'} JOIN` : '';
  return (
    <NodeWrap id={id} title="Join" col="text-blue-400" multi={true} summary={summary}>
      <div className="space-y-3">
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white font-bold rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.joinType || 'inner'} onChange={(e) => onChg('joinType', e.target.value)}>
          <option value="inner">INNER JOIN (共通のみ)</option>
          <option value="left">LEFT JOIN (主データを全て残す)</option>
          <option value="right">RIGHT JOIN (副データを全て残す)</option>
        </select>
        <div className="space-y-1">
          <label className="text-[8px] text-[#888] uppercase tracking-widest font-bold">Main Key</label>
          <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.keyA || ''} onChange={(e) => onChg('keyA', e.target.value)}><option value="">Select Column...</option>{fData.headersA?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className="text-[8px] text-[#888] uppercase tracking-widest font-bold">Sub Key</label>
          <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.keyB || ''} onChange={(e) => onChg('keyB', e.target.value)}><option value="">Select Column...</option>{fData.headersB?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
      </div>
    </NodeWrap>
  );
});

const SortNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const summary = data.sortCol ? `${data.sortCol} ${data.sortOrder === 'desc' ? '↓' : '↑'}` : '';
  return (
    <NodeWrap id={id} title="Sort" col="text-blue-400" summary={summary}>
      <div className="space-y-2">
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.sortCol || ''} onChange={(e) => onChg('sortCol', e.target.value)}><option value="">Target Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.sortOrder || 'asc'} onChange={(e) => onChg('sortOrder', e.target.value)}><option value="asc">Ascending ↑</option><option value="desc">Descending ↓</option></select>
      </div>
    </NodeWrap>
  );
});

const TransformNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const summary = data.targetCol ? `${data.command === 'case_when' ? 'CASE WHEN' : (data.command || '...')} on ${data.targetCol}` : '';
  return (
    <NodeWrap id={id} title="Transform" col="text-blue-400" summary={summary}>
      <div className="space-y-2">
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.targetCol || ''} onChange={(e) => onChg('targetCol', e.target.value)}><option value="">Target Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white font-bold rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.command || ''} onChange={(e) => onChg('command', e.target.value)}>
          <option value="">Select Action...</option>
          <option value="replace">不要文字を削除/置換</option>
          <option value="math_mul">数値を掛け算</option>
          <option value="add_suffix">末尾に文字追加</option>
          <option value="case_when">条件分岐 (CASE WHEN)</option>
        </select>
        
        {data.command === 'case_when' && (
          <div className="space-y-2 mt-2 pt-2 border-t border-[#444]">
            <div className="text-[8px] text-[#888] font-bold uppercase">If Condition:</div>
            <div className="flex gap-1">
              <select className="w-1/3 bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-blue-400 font-bold rounded outline-none nodrag hover:border-blue-400 transition-colors" value={data.condOp || 'exact'} onChange={(e) => onChg('condOp', e.target.value)}>
                <option value="exact">=</option><option value="not">≠</option><option value="gt">&gt;</option><option value="lt">&lt;</option><option value="includes">inc</option>
              </select>
              <NodeInput className="w-2/3 bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white rounded outline-none focus:border-blue-400 transition-colors" placeholder="Value..." value={data.condVal || ''} onChange={(v: any) => onChg('condVal', v)} />
            </div>
            <NodeInput className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white rounded outline-none focus:border-blue-400 transition-colors" placeholder="Then (True Value)" value={data.trueVal || ''} onChange={(v: any) => onChg('trueVal', v)} />
            <NodeInput className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#aaa] rounded outline-none focus:border-blue-400 transition-colors" placeholder="Else (False Value)" value={data.falseVal || ''} onChange={(v: any) => onChg('falseVal', v)} />
          </div>
        )}

        {data.command && data.command !== 'case_when' && (
          <NodeInput className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white rounded outline-none focus:border-blue-400 transition-colors" placeholder="Parameter (ex: ',' or '100')" value={data.param0 || ''} onChange={(v: any) => onChg('param0', v)} />
        )}
      </div>
    </NodeWrap>
  );
});

const FilterNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const chgVal = (v: string, t: string) => onChg('filterVal', (t === 'gt' || t === 'lt') ? v.replace(/[０-９．－]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[^0-9.-]/g, '') : v);
  const op = data.matchType === 'gt' ? '>' : data.matchType === 'lt' ? '<' : data.matchType === 'exact' ? '=' : data.matchType === 'not' ? '≠' : 'inc';
  const summary = data.filterCol ? `${data.filterCol} ${op} ${data.filterVal || ''}` : '';
  return (
    <NodeWrap id={id} title="Filter" col="text-blue-400" summary={summary}>
      <div className="space-y-2">
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.filterCol || ''} onChange={(e) => onChg('filterCol', e.target.value)}><option value="">Target Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        <div className="flex flex-col gap-2">
          <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white font-bold rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.matchType || 'includes'} onChange={(e) => { onChg('matchType', e.target.value); if(data.filterVal) chgVal(String(data.filterVal), e.target.value); }}>
            <option value="includes">含む (Includes)</option><option value="exact">完全一致 (=)</option><option value="not">除外 (≠)</option><option value="gt">以上 (&gt;)</option><option value="lt">以下 (&lt;)</option>
          </select>
          <NodeInput className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white rounded outline-none focus:border-blue-400 transition-colors" placeholder="Condition Value..." value={data.filterVal || ''} onChange={(v: any) => chgVal(v, data.matchType || 'includes')} />
        </div>
      </div>
    </NodeWrap>
  );
});

const SelectNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const summary = data.selectedColumns?.length ? `${data.selectedColumns.length} cols selected` : '';
  return (
    <NodeWrap id={id} title="Select" col="text-blue-400" summary={summary}>
      <div className="max-h-48 overflow-y-auto space-y-1 p-1 bg-[#1a1a1a] rounded border border-[#333] custom-scrollbar">
        {fData.incomingHeaders?.length > 0 ? fData.incomingHeaders.map((h: string) => (
          <label key={h} className="flex items-center gap-2 text-[10px] text-[#ccc] hover:bg-[#333] p-1.5 rounded cursor-pointer group"><input type="checkbox" checked={(data.selectedColumns || []).includes(h)} onChange={(e) => { const c = data.selectedColumns || []; onChg('selectedColumns', e.target.checked ? [...c, h] : c.filter((x: string) => x !== h)); }} className="accent-blue-500 w-3 h-3 nodrag" /><span className="truncate group-hover:text-white">{h}</span></label>
        )) : <div className="text-[9px] text-[#555] text-center py-4">Connect to input data</div>}
      </div>
    </NodeWrap>
  );
});

const GroupByNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const summary = data.groupCol ? `By ${data.groupCol}` : '';
  return (
    <NodeWrap id={id} title="Group By" col="text-blue-400" summary={summary}>
      <div className="space-y-3">
        <div className="space-y-1">
          <label className="text-[8px] text-[#888] uppercase tracking-widest font-bold">Group Key</label>
          <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.groupCol || ''} onChange={(e) => onChg('groupCol', e.target.value)}><option value="">Select Key...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className="text-[8px] text-[#888] uppercase tracking-widest font-bold">Aggregation</label>
          <div className="flex gap-2">
            <select className="flex-1 bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.aggCol || ''} onChange={(e) => onChg('aggCol', e.target.value)}><option value="">Value Col...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
            <select className="w-20 bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white font-bold rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.aggType || 'sum'} onChange={(e) => onChg('aggType', e.target.value)}><option value="sum">SUM</option><option value="count">CNT</option></select>
          </div>
        </div>
      </div>
    </NodeWrap>
  );
});

const ChartNode = memo(({ id, data }: any) => {
  const { fData, onChg } = useNodeLogic(id);
  const summary = data.xAxis && data.yAxis ? `${data.chartType} chart` : '';
  return (
    <NodeWrap id={id} title="Visualizer" col="text-blue-400" summary={summary}>
      <div className="space-y-2">
        <select className="w-full bg-[#1a1a1a] text-[10px] p-2 border border-[#444] text-white font-bold outline-none hover:border-blue-400 transition-colors nodrag" value={data.chartType || 'bar'} onChange={(e) => onChg('chartType', e.target.value)}>
          <option value="bar">Bar Chart</option>
          <option value="line">Line Chart</option>
        </select>
        <div className="grid grid-cols-2 gap-2">
          <div className="space-y-1"><label className="text-[8px] text-[#666] font-bold uppercase">X-Axis</label><select className="w-full bg-[#1a1a1a] text-[9px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.xAxis || ''} onChange={(e) => onChg('xAxis', e.target.value)}><option value="">Select...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select></div>
          <div className="space-y-1"><label className="text-[8px] text-[#666] font-bold uppercase">Y-Axis</label><select className="w-full bg-[#1a1a1a] text-[9px] p-2 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors nodrag" value={data.yAxis || ''} onChange={(e) => onChg('yAxis', e.target.value)}><option value="">Select...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select></div>
        </div>
      </div>
    </NodeWrap>
  );
});

const nodeTypesObj = { dataNode: DataNode, unionNode: UnionNode, joinNode: JoinNode, groupByNode: GroupByNode, sortNode: SortNode, transformNode: TransformNode, selectNode: SelectNode, filterNode: FilterNode, chartNode: ChartNode };

const NodeNavigator = ({ tList, nodes }: { tList: any[], nodes: CustomNode[] }) => {
  const { setCenter } = useReactFlow();
  const [isMinimized, setIsMinimized] = useState(false);

  if (isMinimized) {
    return (
      <Panel position="top-right" className="bg-[#252526]/90 backdrop-blur-md border border-[#444] rounded-xl shadow-xl z-50 m-4 mr-6 cursor-pointer hover:bg-[#333] transition-colors" onClick={() => setIsMinimized(false)}>
        <div className="flex items-center gap-2 p-2 px-3 text-[10px] text-[#888] font-bold uppercase tracking-widest">
          <span className="text-blue-400 flex items-center justify-center">{Icons.Diamond}</span>
          <span className="text-white">{nodes.length} Nodes</span>
        </div>
      </Panel>
    );
  }

  return (
    <Panel position="top-right" className="bg-[#252526]/90 backdrop-blur-md border border-[#444] p-3 rounded-xl shadow-xl max-h-[300px] overflow-y-auto custom-scrollbar flex flex-col gap-1.5 w-60 z-50 m-4 mr-6">
      <div className="text-[10px] text-[#888] font-bold uppercase tracking-widest mb-2 px-1 flex items-center justify-between border-b border-[#444] pb-2">
        <span className="flex items-center gap-1.5"><span className="text-blue-400 flex items-center justify-center">{Icons.Diamond}</span> Navigator</span>
        <div className="flex items-center gap-2">
          <span className="bg-[#1a1a1a] text-[#aaa] px-2 py-0.5 rounded-md text-[9px] border border-[#333]">{nodes.length} Nodes</span>
          <button onClick={() => setIsMinimized(true)} className="hover:text-white transition-colors flex items-center justify-center w-5 h-5">{Icons.ChevronUp}</button>
        </div>
      </div>
      {nodes.map(n => {
        const typeInfo = tList.find(t => t.t === n.type);
        const icon = typeInfo ? typeInfo.i : Icons.Diamond;
        const label = typeInfo ? typeInfo.l : 'Node';
        let subText = n.id;
        
        if (n.type === 'dataNode' && n.data.fileName) subText = n.data.fileName;
        else if (n.type === 'filterNode' && n.data.filterCol) subText = `${n.data.filterCol} ${n.data.matchType} ${n.data.filterVal || ''}`;
        else if (n.type === 'chartNode' && n.data.chartType) subText = `${n.data.chartType} chart`;
        else if (n.type === 'transformNode' && n.data.targetCol) subText = `${n.data.command === 'case_when' ? 'CASE WHEN' : (n.data.command || '...')} on ${n.data.targetCol}`;
        else if (n.type === 'sortNode' && n.data.sortCol) subText = `${n.data.sortCol} ${n.data.sortOrder}`;
        else if (n.type === 'groupByNode' && n.data.groupCol) subText = `By ${n.data.groupCol}`;
        else if (n.type === 'selectNode' && n.data.selectedColumns) subText = `${n.data.selectedColumns.length} cols selected`;
        else if (n.type === 'joinNode' || n.type === 'unionNode') subText = "Merge Data";
        else if (n.type === 'dataNode' && n.data.useFirstRowAsHeader) subText = "Setup Required";

        return (
          <button 
            key={n.id} 
            onClick={() => setCenter(n.position.x + CAMERA_OFFSET_X, n.position.y + CAMERA_OFFSET_Y, { zoom: 1.2, duration: 800 })} 
            className="text-left flex items-center gap-3 p-2 hover:bg-[#333] rounded-lg transition-all group border border-transparent hover:border-[#555] active:scale-95"
          >
            <div className={`w-6 h-6 shrink-0 rounded flex items-center justify-center bg-[#1a1a1a] border border-[#444] ${typeInfo?.c || 'text-white'}`}>
              {icon}
            </div>
            <div className="flex flex-col flex-1 min-w-0">
              <span className="text-[10px] font-bold text-[#ccc] group-hover:text-white uppercase tracking-wider truncate transition-colors">{label}</span>
              <span className="text-[9px] text-[#666] group-hover:text-[#aaa] truncate transition-colors">{subText}</span>
            </div>
          </button>
        );
      })}
    </Panel>
  );
};

const FlowBuilder = () => {
  const [workbooks, setWorkbooks] = useState<Record<string, XLSX.WorkBook>>({});
  const [rangeModalNode, setRangeModalNode] = useState<string | null>(null);
  
  const { screenToFlowPosition, setCenter } = useReactFlow();
  
  const [nodes, _setNodes, onNodesChange] = useNodesState<CustomNode>([{ id: 'n-1', type: 'dataNode', position: { x: 50, y: 150 }, data: { useFirstRowAsHeader: true } }]);
  const [edges, _setEdges, onEdgesChange] = useEdgesState<Edge>([]);
  
  const [previewTab, setPreviewTab] = useState<'table' | 'chart' | 'dashboard'>('table');
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [isPreviewOpen, setIsPreviewOpen] = useState(true);
  const [isSaveLoadOpen, setIsSaveLoadOpen] = useState(false);
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [isSqlModalOpen, setIsSqlModalOpen] = useState(false);
  const [savedFlows, setSavedFlows] = useState<any[]>([]);
  const [bottomHeight, setBottomHeight] = useState(300);
  const [isDragging, setIsDragging] = useState(false);
  
  const [isAutoCameraMove, setIsAutoCameraMove] = useState(true);
  const [previewNodeId, setPreviewNodeId] = useState<string | null>(null);

  useEffect(() => {
    const local = localStorage.getItem('bi-architect-flows');
    if (local) setSavedFlows(JSON.parse(local));
  }, []);

  const handleSave = (name: string) => { const up = [...savedFlows, { id: Date.now().toString(), name, updatedAt: new Date().toLocaleString(), flow: { nodes, edges } }]; setSavedFlows(up); localStorage.setItem('bi-architect-flows', JSON.stringify(up)); };
  const handleLoad = (f: any) => { _setNodes(f.flow.nodes.map((n: any) => n.type === 'dataNode' && n.data.fileName ? { ...n, data: { ...n.data, needsUpload: true } } : n)); _setEdges(f.flow.edges || []); setWorkbooks({}); };
  const onEdgeContextMenu = useCallback((e: React.MouseEvent, edge: Edge) => { e.preventDefault(); _setEdges((eds: any) => eds.filter((e: any) => e.id !== edge.id)); }, [_setEdges]);

  const handleReset = () => {
    _setNodes([{ id: 'n-1', type: 'dataNode', position: { x: 50, y: 150 }, data: { useFirstRowAsHeader: true } }]);
    _setEdges([]);
    setWorkbooks({});
    setIsResetModalOpen(false);
  };
  
  const handleDeleteFlow = (id: string) => {
    const updated = savedFlows.filter((f: any) => f.id !== id);
    setSavedFlows(updated);
    localStorage.setItem('bi-architect-flows', JSON.stringify(updated));
  };

  const startResize = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  useEffect(() => {
    const onMouseMove = (e: MouseEvent) => {
      if (!isDragging) return;
      const newHeight = window.innerHeight - e.clientY;
      setBottomHeight(Math.max(100, Math.min(newHeight, window.innerHeight - 100)));
    };
    const onMouseUp = () => setIsDragging(false);
    if (isDragging) { window.addEventListener('mousemove', onMouseMove); window.addEventListener('mouseup', onMouseUp); }
    return () => { window.removeEventListener('mousemove', onMouseMove); window.removeEventListener('mouseup', onMouseUp); };
  }, [isDragging]);

  const sHash = JSON.stringify(nodes.map(n => ({ id: n.id, data: n.data }))) + JSON.stringify(edges);
  const nodeFlowData = useMemo(() => {
    const map: Record<string, any> = {};
    nodes.forEach(n => {
      if (n.type === 'joinNode' || n.type === 'unionNode') {
        const eA = edges.find(e => e.target === n.id && (e as any).targetHandle === 'input-a'), eB = edges.find(e => e.target === n.id && (e as any).targetHandle === 'input-b');
        const hA = eA ? calcData(eA.source, nodes, edges, workbooks).headers : [], hB = eB ? calcData(eB.source, nodes, edges, workbooks).headers : [];
        map[n.id] = { headersA: hA, headersB: hB, incomingHeaders: [...new Set([...hA, ...hB])] };
      } else {
        const inEdge = edges.find(e => e.target === n.id);
        map[n.id] = { incomingHeaders: inEdge ? calcData(inEdge.source, nodes, edges, workbooks).headers : [] };
      }
    }); return map;
  }, [sHash, workbooks]); // eslint-disable-line

  const final = useMemo(() => {
    let targetNodeId = previewNodeId;
    if (!targetNodeId || !nodes.find(n => n.id === targetNodeId)) {
      const term = nodes.find(n => !edges.some(e => e.source === n.id));
      targetNodeId = term?.id || "";
    }
    
    if (!targetNodeId) return { data: [], headers: [], chartConfig: null };

    const result = calcData(targetNodeId, nodes, edges, workbooks);
    const targetNode = nodes.find(n => n.id === targetNodeId);
    return { ...result, chartConfig: targetNode?.type === 'chartNode' ? targetNode.data : null };
  }, [sHash, workbooks, previewNodeId]); // eslint-disable-line

  const dashboardsData = useMemo(() => {
    return nodes.filter(n => n.type === 'chartNode').map(n => {
      const res = calcData(n.id, nodes, edges, workbooks);
      return { id: n.id, config: n.data, data: res.data };
    });
  }, [sHash, workbooks]); // eslint-disable-line

  const handleExport = (format: 'csv' | 'xlsx' | 'json') => {
    if (final.data.length === 0) return;
    if (format === 'csv') { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([[0xEF, 0xBB, 0xBF] as any, Papa.unparse(final.data)], { type: 'text/csv' })); a.download = 'export.csv'; a.click(); }
    else if (format === 'json') { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify(final.data, null, 2)], { type: 'application/json' })); a.download = 'export.json'; a.click(); }
    else if (format === 'xlsx') { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(final.data), "Result"); XLSX.writeFile(wb, "export.xlsx"); }
  };

  const tList = [
    { t: 'dataNode', l: 'Source', i: Icons.Source, c: 'text-blue-400' },
    { t: 'unionNode', l: 'Union', i: Icons.Union, c: 'text-blue-400' },
    { t: 'joinNode', l: 'Join', i: Icons.Join, c: 'text-blue-400' },
    { t: 'groupByNode', l: 'Group By', i: Icons.GroupBy, c: 'text-blue-400' },
    { t: 'sortNode', l: 'Sort', i: Icons.Sort, c: 'text-blue-400' },
    { t: 'transformNode', l: 'Transform', i: Icons.Transform, c: 'text-blue-400' },
    { t: 'selectNode', l: 'Select', i: Icons.Select, c: 'text-blue-400' },
    { t: 'filterNode', l: 'Filter', i: Icons.Filter, c: 'text-blue-400' },
    { t: 'chartNode', l: 'Visualizer', i: Icons.Chart, c: 'text-blue-400' }
  ];

  return (
    <AppContext.Provider value={{ workbooks, setWorkbooks, setRangeModalNode, nodeFlowData, isAutoCameraMove }}>
      <div className="h-screen w-screen bg-[#1a1a1a] flex flex-col font-sans overflow-hidden">
        <GlobalStyle />
        <div className="bg-[#181818] border-b border-[#333] px-6 py-3 flex justify-between items-center z-40 shadow-md gap-4">
          <h1 className="text-[13px] font-bold text-white tracking-[0.5em] uppercase flex items-center gap-3 shrink-0">
            <span className="text-blue-500 w-4 h-4 flex items-center justify-center">{Icons.Diamond}</span>
            BI Architect
          </h1>
          
          <div className="flex-1 flex justify-center">
            <div className="bg-[#1e1e1e] border border-[#333] px-4 py-1.5 rounded-full text-white text-[10px] tracking-widest uppercase font-bold flex items-center gap-2">
              <span className="text-blue-500">{Icons.Diamond}</span> Drag & Drop Nodes to Build Pipeline
            </div>
          </div>

          <div className="flex items-center gap-3 shrink-0">
            <button onClick={() => setIsAutoCameraMove(!isAutoCameraMove)} className={`bg-[#252526] hover:bg-[#333] border border-[#444] ${isAutoCameraMove ? 'text-blue-400' : 'text-[#666]'} text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-2 shadow active:scale-95 transition-colors`} title="Auto Camera Focus">
              <span className="flex items-center justify-center gap-1.5">{Icons.Focus} {isAutoCameraMove ? 'FOCUS: ON' : 'FOCUS: OFF'}</span>
            </button>
            
            <button onClick={() => setIsSqlModalOpen(true)} className="bg-[#252526] hover:bg-blue-900/30 border border-[#444] hover:border-blue-500/50 text-[#aaa] hover:text-blue-400 text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow active:scale-95 transition-colors">
              <span className="flex items-center justify-center gap-1">{Icons.Code} SQL</span>
            </button>
            <button onClick={() => setIsResetModalOpen(true)} className="bg-[#252526] hover:bg-red-900/30 border border-[#444] hover:border-red-500/50 text-[#aaa] hover:text-red-400 text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow active:scale-95 transition-colors">
              <span className="flex items-center justify-center gap-1">{Icons.Trash} RESET</span>
            </button>
            <button onClick={() => setIsSaveLoadOpen(true)} className="bg-[#252526] hover:bg-[#333] border border-[#444] text-white text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow active:scale-95 transition-colors">
              <span className="flex items-center justify-center gap-1">{Icons.Save} / {Icons.Folder} PROJECTS</span>
            </button>
          </div>
        </div>
        <div className="flex-1 flex overflow-hidden relative">
          <aside className={`bg-[#181818] border-r border-[#333] z-20 flex flex-col transition-all duration-300 ease-in-out ${isSidebarOpen ? 'w-56 py-4 pl-4 pr-2' : 'w-16 p-2 items-center'}`}>
            <div className={`flex items-center ${isSidebarOpen ? 'justify-between mb-4 pr-2' : 'justify-center mb-6'} border-b border-[#333] pb-2`}>
              {isSidebarOpen && <div className="text-[10px] font-bold text-white tracking-[0.3em] uppercase">Toolbox</div>}
              <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className="text-[#888] hover:text-white p-1 rounded hover:bg-[#333] transition-colors flex items-center justify-center w-6 h-6">
                {isSidebarOpen ? Icons.ArrowLeft : Icons.ArrowRight}
              </button>
            </div>
            <div className={`flex flex-col ${isSidebarOpen ? 'gap-3 pr-2' : 'gap-4'} overflow-y-auto overflow-x-hidden pb-10 custom-scrollbar`}>
              {tList.map(item => (
                <div key={item.t} className={`bg-[#252526] border border-[#333] rounded-xl hover:border-blue-500 cursor-grab flex items-center transition-all shadow-md active:scale-95 group ${isSidebarOpen ? 'p-3 gap-4' : 'p-3 justify-center w-12 h-12'}`} onDragStart={(e) => e.dataTransfer.setData('application/reactflow', item.t)} draggable title={!isSidebarOpen ? item.l : ''}>
                  <div className={`${item.c} text-lg group-hover:scale-125 transition-transform ${isSidebarOpen ? '' : 'text-xl'}`}>{item.i}</div>
                  {isSidebarOpen && <span className="text-[10px] text-[#888] font-bold uppercase tracking-wider group-hover:text-white truncate">{item.l}</span>}
                </div>
              ))}
            </div>
          </aside>
          <div className="flex-1 relative bg-[#1e1e1e]">
            <ReactFlow 
              nodes={nodes} 
              edges={edges} 
              onNodesChange={onNodesChange} 
              onEdgesChange={onEdgesChange} 
              onConnect={(p) => _setEdges((eds: any) => addEdge({ ...p, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } } as any, eds))} 
              onEdgeContextMenu={onEdgeContextMenu} 
              nodeTypes={nodeTypesObj} 
              onNodeDragStop={(_, node: any) => {
                if (isAutoCameraMove) {
                  setCenter(node.position.x + CAMERA_OFFSET_X, node.position.y + CAMERA_OFFSET_Y, { zoom: 1.1, duration: 600 });
                }
              }}
              onDrop={(e) => { 
                const t = e.dataTransfer.getData('application/reactflow'); 
                if (t) {
                  const position = screenToFlowPosition({ x: e.clientX, y: e.clientY });
                  _setNodes((nds: any) => nds.concat({ id: `n-${Date.now()}`, type: t, position, data: { useFirstRowAsHeader: true } })); 
                  if (isAutoCameraMove) {
                    setCenter(position.x + CAMERA_OFFSET_X, position.y + CAMERA_OFFSET_Y, { zoom: 1.1, duration: 600 });
                  }
                }
              }} 
              onDragOver={(e) => e.preventDefault()} 
              fitView
            >
              <Background color="#333" gap={24} size={1} />
              <Controls />
              <NodeNavigator tList={tList} nodes={nodes} />
            </ReactFlow>
          </div>
        </div>
        
        <div 
          style={{ height: isPreviewOpen ? bottomHeight : 48, transition: isDragging ? 'none' : 'height 0.3s cubic-bezier(0.4, 0, 0.2, 1)' }} 
          className="bg-[#1a1a1a] flex flex-col border-t border-[#333] z-30 shadow-[0_-10px_40px_rgba(0,0,0,0.5)] relative"
        >
          {isPreviewOpen && (
            <div 
              onMouseDown={startResize} 
              className="absolute top-[-3px] left-0 w-full h-2 cursor-row-resize z-50 hover:bg-blue-500/50 transition-colors" 
              title="Drag to resize"
            />
          )}
          <div className="px-6 py-2 bg-[#252526] flex justify-between items-center border-b border-[#333] h-[48px] shrink-0">
            <div className="flex items-center gap-4">
              <button onClick={() => setIsPreviewOpen(!isPreviewOpen)} className="text-[#888] hover:text-white p-1 rounded hover:bg-[#333] transition-colors mr-2 flex items-center justify-center w-6 h-6">
                {isPreviewOpen ? Icons.ChevronDown : Icons.ChevronUp}
              </button>
              {isPreviewOpen && ['table', 'chart', 'dashboard'].map(tab => (
                <button key={tab} onClick={() => setPreviewTab(tab as any)} className={`text-[11px] font-bold uppercase tracking-[0.3em] pb-1 border-b-2 transition-all mt-1 ${previewTab === tab ? 'text-blue-400 border-blue-400' : 'text-[#555] border-transparent hover:text-white'}`}>
                  <span className="flex items-center gap-2">
                    {tab === 'table' && <>{Icons.Select} Data Table</>}
                    {tab === 'chart' && <>{Icons.Chart} Visual Insight</>}
                    {tab === 'dashboard' && <>{Icons.Dashboard} Dashboard</>}
                  </span>
                </button>
              ))}
              {!isPreviewOpen && <span className="text-[10px] text-[#888] font-bold uppercase tracking-[0.2em]">Preview Minimized</span>}
            </div>
            
            <div className="flex items-center gap-3">
              {isPreviewOpen && previewTab !== 'dashboard' && (
                <div className="flex items-center gap-2 mr-2 border-r border-[#444] pr-4">
                  <span className="text-[10px] text-[#888] font-bold uppercase tracking-widest">Preview:</span>
                  <select 
                    className="bg-[#1a1a1a] text-[10px] p-1.5 border border-[#444] text-[#ccc] rounded outline-none hover:border-blue-400 transition-colors cursor-pointer w-40 truncate"
                    value={previewNodeId || ''}
                    onChange={(e) => setPreviewNodeId(e.target.value || null)}
                  >
                    <option value="">Auto (末端のノード)</option>
                    {nodes.map(n => {
                      const typeInfo = tList.find(t => t.t === n.type);
                      const label = typeInfo ? typeInfo.l : n.type;
                      let sub = n.id;
                      if (n.type === 'dataNode' && n.data.fileName) sub = n.data.fileName;
                      return <option key={n.id} value={n.id}>{label} - {sub}</option>
                    })}
                  </select>
                </div>
              )}

              {isPreviewOpen && previewTab !== 'dashboard' && (
                <>
                  <span className="text-[10px] text-white uppercase font-bold tracking-widest mr-2">Export As:</span>
                  <button onClick={() => handleExport('csv')} className="bg-[#333] hover:bg-blue-600 text-white text-[10px] px-3 py-1.5 rounded-lg font-bold uppercase tracking-widest transition-colors shadow">CSV</button>
                  <button onClick={() => handleExport('xlsx')} className="bg-[#333] hover:bg-blue-600 text-white text-[10px] px-3 py-1.5 rounded-lg font-bold uppercase tracking-widest transition-colors shadow">Excel</button>
                </>
              )}
              {previewTab !== 'dashboard' && <span className="text-[10px] text-blue-400 font-bold ml-4">{final.data.length} rows</span>}
            </div>
          </div>
          {isPreviewOpen && (
            <div className="flex-1 overflow-auto bg-[#1e1e1e]">
              {previewTab === 'table' && (
                <table className="w-full text-left text-[11px] whitespace-nowrap border-collapse"><thead className="bg-[#1a1a1a] sticky top-0 border-b border-[#333] z-10 shadow-sm"><tr>{final.headers.map((h, i) => <th key={i} className="px-5 py-3 font-bold text-[#888] border-r border-[#333] uppercase tracking-wider">{h}</th>)}</tr></thead><tbody>{final.data.slice(0, 100).map((row, i) => (<tr key={i} className="hover:bg-[#252526] transition-colors border-b border-[#222]">{final.headers.map((h, j) => <td key={j} className="px-5 py-2 text-[#ccc] border-r border-[#222] font-mono">{row[h]}</td>)}</tr>))}</tbody></table>
              )}
              {previewTab === 'chart' && (
                <div style={{ width: '100%', height: '100%', minHeight: '300px', padding: '24px' }}>
                  {final.chartConfig?.xAxis && final.chartConfig?.yAxis ? (
                    <ResponsiveContainer width="100%" height="100%" minWidth={1} minHeight={1}>
                      {final.chartConfig.chartType === 'line' ? (<LineChart data={final.data}><CartesianGrid strokeDasharray="3 3" stroke="#333" vertical={false} /><XAxis dataKey={final.chartConfig.xAxis} stroke="#555" tick={{ fill: '#ddd', fontSize: 11 }} tickLine={false} axisLine={false} dy={10} /><YAxis stroke="#555" tick={{ fill: '#ddd', fontSize: 11 }} tickLine={false} axisLine={false} dx={-10} /><Tooltip contentStyle={{ backgroundColor: '#1a1a1a', border: '1px solid #444', fontSize: '11px', borderRadius: '8px' }} labelStyle={{ color: '#fff', fontWeight: 'bold', paddingBottom: '4px' }} itemStyle={{ color: '#60a5fa' }} /><Line type="monotone" dataKey={final.chartConfig.yAxis} stroke="#3b82f6" strokeWidth={3} dot={{ r: 4, fill: '#3b82f6', strokeWidth: 0 }} activeDot={{ r: 6 }} /></LineChart>) : (<BarChart data={final.data}><CartesianGrid strokeDasharray="3 3" stroke="#333" vertical={false} /><XAxis dataKey={final.chartConfig.xAxis} stroke="#555" tick={{ fill: '#ddd', fontSize: 11 }} tickLine={false} axisLine={false} dy={10} /><YAxis stroke="#555" tick={{ fill: '#ddd', fontSize: 11 }} tickLine={false} axisLine={false} dx={-10} /><Tooltip cursor={{fill: '#252526'}} contentStyle={{ backgroundColor: '#1a1a1a', border: '1px solid #444', fontSize: '11px', borderRadius: '8px' }} labelStyle={{ color: '#fff', fontWeight: 'bold', paddingBottom: '4px' }} itemStyle={{ color: '#60a5fa' }} /><Bar dataKey={final.chartConfig.yAxis} fill="#3b82f6" radius={[4, 4, 0, 0]} /></BarChart>)}
                    </ResponsiveContainer>
                  ) : <div className="h-full flex items-center justify-center text-[#555] text-[11px] italic tracking-widest uppercase animate-pulse">Visualizerノードを繋ぎ、軸を設定してください</div>}
                </div>
              )}
              {previewTab === 'dashboard' && (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 p-6 h-full overflow-y-auto custom-scrollbar bg-[#111]">
                  {dashboardsData.length === 0 ? (
                     <div className="col-span-full h-full flex items-center justify-center text-[#555] text-[11px] italic tracking-widest uppercase animate-pulse">
                        Canvas上にVisualizerノードを配置してください
                     </div>
                  ) : (
                    dashboardsData.map(d => (
                      <div key={d.id} className="bg-[#1e1e1e] p-4 rounded-2xl border border-[#333] flex flex-col h-[320px] shadow-lg hover:border-[#444] transition-colors">
                        <div className="text-[#888] text-[10px] font-bold uppercase tracking-wider mb-4 border-b border-[#333] pb-2 flex justify-between items-center">
                          <span>{d.config.chartType === 'line' ? '📈 Line Chart' : '📊 Bar Chart'}</span>
                          <span className="text-[#555]">{d.config.yAxis} by {d.config.xAxis}</span>
                        </div>
                        <div className="flex-1 min-h-0">
                          {d.config.xAxis && d.config.yAxis && d.data.length > 0 ? (
                            <ResponsiveContainer width="100%" height="100%" minWidth={1} minHeight={1}>
                              {d.config.chartType === 'line' ? (
                                <LineChart data={d.data}><CartesianGrid strokeDasharray="3 3" stroke="#222" vertical={false} /><XAxis dataKey={d.config.xAxis} stroke="#555" tick={{ fill: '#888', fontSize: 9 }} tickLine={false} axisLine={false} dy={5} /><YAxis stroke="#555" tick={{ fill: '#888', fontSize: 9 }} tickLine={false} axisLine={false} dx={-5} /><Tooltip contentStyle={{ backgroundColor: '#1a1a1a', border: '1px solid #444', fontSize: '11px', borderRadius: '8px' }} /><Line type="monotone" dataKey={d.config.yAxis} stroke="#3b82f6" strokeWidth={2} dot={{ r: 2, fill: '#3b82f6', strokeWidth: 0 }} activeDot={{ r: 4 }} /></LineChart>
                              ) : (
                                <BarChart data={d.data}><CartesianGrid strokeDasharray="3 3" stroke="#222" vertical={false} /><XAxis dataKey={d.config.xAxis} stroke="#555" tick={{ fill: '#888', fontSize: 9 }} tickLine={false} axisLine={false} dy={5} /><YAxis stroke="#555" tick={{ fill: '#888', fontSize: 9 }} tickLine={false} axisLine={false} dx={-5} /><Tooltip cursor={{fill: '#222'}} contentStyle={{ backgroundColor: '#1a1a1a', border: '1px solid #444', fontSize: '11px', borderRadius: '8px' }} /><Bar dataKey={d.config.yAxis} fill="#3b82f6" radius={[2, 2, 0, 0]} /></BarChart>
                              )}
                            </ResponsiveContainer>
                          ) : <div className="h-full flex items-center justify-center text-[#444] text-[10px] italic uppercase">No Data</div>}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              )}
            </div>
          )}
        </div>

        <SqlModal 
          isOpen={isSqlModalOpen} 
          onClose={() => setIsSqlModalOpen(false)} 
          nodes={nodes} 
          edges={edges} 
          onImport={(n: CustomNode[], e: Edge[]) => { _setNodes(n); _setEdges(e); setWorkbooks({}); }} 
        />

        <SaveLoadModal isOpen={isSaveLoadOpen} onClose={() => setIsSaveLoadOpen(false)} onSave={handleSave} onLoad={handleLoad} onDelete={handleDeleteFlow} flows={savedFlows} />

        {isResetModalOpen && (
          <div className="fixed inset-0 bg-black/80 z-[300] flex items-center justify-center p-8 backdrop-blur-sm">
            <div className="bg-[#1e1e1e] border border-[#444] rounded-2xl shadow-2xl w-[320px] p-6 text-center space-y-6">
              <div className="text-white text-4xl w-10 h-10 mx-auto flex items-center justify-center">{Icons.Warning}</div>
              <div className="space-y-2">
                <h3 className="text-white text-sm font-bold tracking-widest">RESET ALL DATA?</h3>
                <p className="text-[#888] text-[10px]">現在のノード構成と読み込んだデータがすべて消去されます。この操作は元に戻せません。</p>
              </div>
              <div className="flex gap-3">
                <button onClick={() => setIsResetModalOpen(false)} className="flex-1 py-3 bg-[#333] hover:bg-[#444] text-white rounded-xl text-[10px] font-bold uppercase tracking-widest transition-colors">Cancel</button>
                <button onClick={handleReset} className="flex-1 py-3 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-[10px] font-bold uppercase tracking-widest shadow-xl transition-all">Reset</button>
              </div>
            </div>
          </div>
        )}

        {rangeModalNode && <RangeSelectorModal isOpen={true} onClose={() => setRangeModalNode(null)} workbook={workbooks[rangeModalNode]} currentSheet={nodes.find(n => n.id === rangeModalNode)?.data.currentSheet} initialRanges={nodes.find(n => n.id === rangeModalNode)?.data.ranges} initialUseHeader={nodes.find(n => n.id === rangeModalNode)?.data.useFirstRowAsHeader} onRangesConfirm={(r: string[], h: boolean) => { _setNodes((nds: any[]) => nds.map((n: any) => n.id === rangeModalNode ? { ...n, data: { ...n.data, ranges: r, useFirstRowAsHeader: h } } : n)); setRangeModalNode(null); }} />}
      </div>
    </AppContext.Provider>
  );
};

const SqlModal = ({ isOpen, onClose, nodes, edges, onImport }: any) => {
  const [tab, setTab] = useState<'export' | 'import'>('export'); 
  const [sqlInput, setSqlInput] = useState('');
  const [copyMsg, setCopyMsg] = useState('');

  const generatedSql = useMemo(() => {
    if (tab !== 'export' || !isOpen) return '';
    const terminals = nodes.filter((n: any) => !edges.some((e: any) => e.source === n.id));
    if (terminals.length === 0) return "-- ノードが接続されていません";
    
    let curr = terminals[0].id;
    const path: any[] = [];
    while (curr) {
      const node = nodes.find((n: any) => n.id === curr);
      if (node) path.unshift(node);
      const inEdge = edges.find((e: any) => e.target === curr);
      curr = inEdge ? inEdge.source : null;
    }

    let select = "*", from = "source_data", joinStr = "", wheres: string[] = [], groupBy = "", orderBy = "";
    
    path.forEach((n: any) => {
      if (n.type === 'dataNode') {
        from = n.data.fileName ? `\`${n.data.fileName}\`` : "source_data";
        if (n.data.ranges && n.data.ranges.length > 0) {
          from += ` /* Range: ${n.data.ranges.join(', ')} */`;
        }
      }
      if (n.type === 'joinNode' && n.data.keyA && n.data.keyB) {
        const jType = n.data.joinType === 'left' ? 'LEFT' : n.data.joinType === 'right' ? 'RIGHT' : 'INNER';
        joinStr = `\n${jType} JOIN sub_data ON \`${n.data.keyA}\` = sub_data.\`${n.data.keyB}\``;
      }
      if (n.type === 'selectNode' && n.data.selectedColumns?.length) select = n.data.selectedColumns.map((c: string) => `\`${c}\``).join(", ");
      
      if (n.type === 'transformNode' && n.data.command === 'case_when' && n.data.targetCol) {
        const op = n.data.condOp === 'gt' ? '>' : n.data.condOp === 'lt' ? '<' : n.data.condOp === 'not' ? '!=' : '=';
        const cWhen = `CASE WHEN \`${n.data.targetCol}\` ${op} '${n.data.condVal || ''}' THEN '${n.data.trueVal || ''}' ELSE '${n.data.falseVal || ''}' END AS \`${n.data.targetCol}\``;
        select = select === "*" ? `*, ${cWhen}` : `${select}, ${cWhen}`;
      }

      if (n.type === 'filterNode' && n.data.filterCol && n.data.filterVal) {
        const op = n.data.matchType === 'gt' ? '>' : n.data.matchType === 'lt' ? '<' : n.data.matchType === 'exact' ? '=' : n.data.matchType === 'not' ? '!=' : 'LIKE';
        const val = n.data.matchType === 'includes' ? `'%${n.data.filterVal}%'` : `'${n.data.filterVal}'`;
        wheres.push(`\`${n.data.filterCol}\` ${op} ${val}`);
      }
      if (n.type === 'groupByNode' && n.data.groupCol && n.data.aggCol) {
        groupBy = `\`${n.data.groupCol}\``;
        select = `\`${n.data.groupCol}\`, ${n.data.aggType.toUpperCase()}(\`${n.data.aggCol}\`)`;
      }
      if (n.type === 'sortNode' && n.data.sortCol) {
        orderBy = `\`${n.data.sortCol}\` ${n.data.sortOrder === 'desc' ? 'DESC' : 'ASC'}`;
      }
    });

    let sql = `SELECT ${select}\nFROM ${from}`;
    if (joinStr) sql += joinStr;
    if (wheres.length > 0) sql += `\nWHERE ${wheres.join(" AND ")}`;
    if (groupBy) sql += `\nGROUP BY ${groupBy}`;
    if (orderBy) sql += `\nORDER BY ${orderBy}`;
    return sql;
  }, [nodes, edges, tab, isOpen]);

  const handleImport = () => {
    const sql = sqlInput.replace(/\n/g, ' ');
    const n: CustomNode[] = [], e: Edge[] = [];
    let x = 50, y = 150, id = 1;
    const getId = () => `n-sql-${Date.now()}-${id++}`;
    
    let ranges: string[] = [];
    const fMatch = sql.match(/FROM\s+(.*?)(?:\sWHERE|\sGROUP BY|\sORDER BY|$)/i);
    if (fMatch) {
      const rangeMatch = fMatch[1].match(/\/\*\s*Range:\s*(.*?)\s*\*\//i);
      if (rangeMatch) ranges = rangeMatch[1].split(',').map((r: string) => r.trim());
    }

    let prevId = getId();
    n.push({ id: prevId, type: 'dataNode', position: { x, y }, data: { useFirstRowAsHeader: true, ranges } });
    x += 280;

    const wMatch = sql.match(/WHERE\s+(.*?)(?:GROUP BY|ORDER BY|$)/i);
    if (wMatch) {
      const cond = wMatch[1].trim();
      const parts = cond.split(/(?:=|!=|>|<|LIKE)/i);
      if (parts.length >= 2) {
        const opStr = cond.match(/(?:=|!=|>|<|LIKE)/i)?.[0].toUpperCase();
        const mt = opStr === '=' ? 'exact' : opStr === '!=' ? 'not' : opStr === '>' ? 'gt' : opStr === '<' ? 'lt' : 'includes';
        const newId = getId();
        n.push({ id: newId, type: 'filterNode', position: { x, y }, data: { filterCol: parts[0].trim().replace(/`/g, ''), filterVal: parts[1].trim().replace(/'|%/g, ''), matchType: mt } });
        e.push({ id: `e-${prevId}-${newId}`, source: prevId, target: newId, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } });
        prevId = newId; x += 280;
      }
    }
    
    const gMatch = sql.match(/GROUP BY\s+(.*?)(?:ORDER BY|$)/i);
    const sMatch = sql.match(/SELECT\s+(.*?)\s+FROM/i);
    if (gMatch) {
      const gCol = gMatch[1].trim().replace(/`/g, '');
      let aCol = '', aTyp = 'sum';
      if (sMatch) {
        const sm = sMatch[1].match(/SUM\((.*?)\)/i), cm = sMatch[1].match(/COUNT\((.*?)\)/i);
        if (sm) { aCol = sm[1].trim().replace(/`/g, ''); aTyp = 'sum'; }
        if (cm) { aCol = cm[1].trim().replace(/`/g, ''); aTyp = 'count'; }
      }
      const newId = getId();
      n.push({ id: newId, type: 'groupByNode', position: { x, y }, data: { groupCol: gCol, aggCol: aCol, aggType: aTyp } });
      e.push({ id: `e-${prevId}-${newId}`, source: prevId, target: newId, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } });
      prevId = newId; x += 280;
    } else if (sMatch && sMatch[1].trim() !== '*') {
      const colsStr = sMatch[1].replace(/CASE WHEN.*?END AS `.*?`/gi, '');
      const cols = colsStr.split(',').map((c: string) => c.trim().replace(/`/g, '')).filter((c: string) => c);
      if(cols.length > 0) {
        const newId = getId();
        n.push({ id: newId, type: 'selectNode', position: { x, y }, data: { selectedColumns: cols } });
        e.push({ id: `e-${prevId}-${newId}`, source: prevId, target: newId, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } });
        prevId = newId; x += 280;
      }
    }

    const oMatch = sql.match(/ORDER BY\s+(.*?)$/i);
    if (oMatch) {
      const p = oMatch[1].trim().split(/\s+/);
      const newId = getId();
      n.push({ id: newId, type: 'sortNode', position: { x, y }, data: { sortCol: p[0].replace(/`/g, ''), sortOrder: p[1]?.toUpperCase() === 'DESC' ? 'desc' : 'asc' } });
      e.push({ id: `e-${prevId}-${newId}`, source: prevId, target: newId, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } });
    }
    
    onImport(n, e);
    onClose();
    setSqlInput('');
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(generatedSql);
    setCopyMsg('Copied!');
    setTimeout(() => setCopyMsg(''), 2000);
  };

  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-black/90 z-[200] flex items-center justify-center p-8 backdrop-blur-md">
      <div className="bg-[#1e1e1e] border border-[#444] rounded-2xl shadow-2xl w-[600px] overflow-hidden flex flex-col">
        <div className="flex border-b border-[#444]">
          <button onClick={() => setTab('export')} className={`flex-1 p-4 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'export' ? 'bg-[#252526] text-blue-400 border-b-2 border-blue-400' : 'text-[#666] hover:bg-[#252526]'}`}><span>{Icons.Code}</span> Flow to SQL</button>
          <button onClick={() => setTab('import')} className={`flex-1 p-4 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'import' ? 'bg-[#252526] text-green-400 border-b-2 border-green-400' : 'text-[#666] hover:bg-[#252526]'}`}><span>{Icons.Diamond}</span> SQL to Flow</button>
        </div>
        <div className="p-6 h-[350px] bg-[#1a1a1a] flex flex-col">
          {tab === 'export' ? (
            <div className="flex-1 flex flex-col space-y-4">
              <p className="text-[10px] text-[#888]">現在のノード構成（単一パス）を解析して、対応するSQL文を自動生成します。</p>
              <div className="flex-1 relative">
                <textarea readOnly value={generatedSql} className="w-full h-full bg-[#1e1e1e] text-blue-300 font-mono text-[12px] p-4 border border-[#444] rounded-xl outline-none resize-none custom-scrollbar" />
                <button onClick={handleCopy} className="absolute top-3 right-3 bg-[#333] hover:bg-blue-600 text-white text-[10px] px-3 py-1.5 rounded flex items-center gap-1 transition-colors">
                  {copyMsg || <><span className="w-3 h-3 flex">{Icons.Copy}</span> Copy</>}
                </button>
              </div>
            </div>
          ) : (
            <div className="flex-1 flex flex-col space-y-4">
              <p className="text-[10px] text-[#888]">SELECT, FROM, WHERE, GROUP BY, ORDER BY 構文を解析してノードを配置します。</p>
              <textarea value={sqlInput} onChange={(e) => setSqlInput(e.target.value)} placeholder="SELECT col1, SUM(col2)&#10;FROM data /* Range: A1:D10 */&#10;WHERE col1 = 'value'&#10;GROUP BY col1&#10;ORDER BY col1 DESC" className="flex-1 w-full bg-[#1e1e1e] text-green-300 font-mono text-[12px] p-4 border border-[#444] rounded-xl outline-none focus:border-green-500 resize-none transition-colors custom-scrollbar" />
              <button disabled={!sqlInput} onClick={handleImport} className="w-full bg-green-600 hover:bg-green-500 disabled:opacity-30 text-white py-3 rounded-xl text-[11px] font-bold uppercase shadow-xl active:scale-95 transition-all">Build Flow from SQL</button>
            </div>
          )}
        </div>
        <button onClick={onClose} className="w-full p-4 bg-[#252526] text-white hover:bg-[#333] text-[10px] font-bold uppercase border-t border-[#444] transition-colors">Close</button>
      </div>
    </div>
  );
};

const SaveLoadModal = ({ isOpen, onClose, onSave, onLoad, onDelete, flows }: any) => {
  const [tab, setTab] = useState<'load' | 'save'>('load'); const [sName, setSName] = useState('');
  const sortedFlows = useMemo(() => [...flows].sort((a, b) => Number(b.id) - Number(a.id)), [flows]);

  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-black/90 z-[200] flex items-center justify-center p-8 backdrop-blur-md">
      <div className="bg-[#1e1e1e] border border-[#444] rounded-2xl shadow-2xl w-[500px] overflow-hidden">
        <div className="flex border-b border-[#444]"><button onClick={() => setTab('load')} className={`flex-1 p-3 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'load' ? 'bg-[#252526] text-blue-400 border-b-2 border-blue-400' : 'text-[#666] hover:bg-[#252526]'}`}><span>{Icons.Folder}</span> Load</button><button onClick={() => setTab('save')} className={`flex-1 p-3 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'save' ? 'bg-[#252526] text-white border-b-2 border-white' : 'text-[#666] hover:bg-[#252526]'}`}><span>{Icons.Save}</span> Save</button></div>
        <div className="p-6 h-[300px] overflow-y-auto bg-[#1a1a1a] custom-scrollbar">
          {tab === 'load' ? (
            <div className="space-y-3">
              {sortedFlows.length === 0 && <div className="text-[10px] text-[#555] text-center mt-10">No saved projects</div>}
              {sortedFlows.map((f: any) => (
                <div key={f.id} className="bg-[#252526] border border-[#444] p-4 rounded-xl flex justify-between items-center hover:border-blue-500/50 cursor-pointer group transition-colors" onClick={() => { onLoad(f); onClose(); }}>
                  <div>
                    <div className="text-[12px] text-white font-bold group-hover:text-blue-400 transition-colors">{f.name}</div>
                    <div className="text-[9px] text-[#666] mt-1">{f.updatedAt}</div>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={(e) => { e.stopPropagation(); onDelete(f.id); }} className="bg-[#333] text-[#aaa] border border-[#555] px-3 py-2 rounded-lg text-[10px] font-bold shadow hover:bg-blue-600 hover:text-white hover:border-blue-500 transition-all flex items-center justify-center w-8" title="Delete Project">
                      <span>{Icons.Close}</span>
                    </button>
                    <button className="bg-blue-600/20 text-blue-400 border border-blue-500/30 px-5 py-2 rounded-lg text-[10px] font-bold shadow group-hover:bg-blue-600 group-hover:text-white transition-all">Load</button>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="space-y-4 pt-8">
              <div className="space-y-2">
                <label className="text-[10px] text-white font-bold uppercase">Project Name</label>
                <input type="text" placeholder="e.g. Sales Report" value={sName} onChange={(e) => setSName(e.target.value)} className="w-full bg-[#252526] border border-[#444] text-white p-3 rounded-xl outline-none focus:border-blue-500 text-sm transition-colors" />
              </div>
              <button disabled={!sName} onClick={() => { onSave(sName); setSName(''); onClose(); }} className="w-full bg-blue-600 hover:bg-blue-500 disabled:opacity-30 text-white py-3 rounded-xl text-[11px] font-bold uppercase shadow-xl active:scale-95 transition-all">Save Now</button>
            </div>
          )}
        </div>
        <button onClick={onClose} className="w-full p-3 bg-[#252526] text-white hover:bg-[#333] text-[10px] font-bold uppercase border-t border-[#444] transition-colors">Close</button>
      </div>
    </div>
  );
};

const RangeSelectorModal = ({ isOpen, onClose, workbook, currentSheet, onRangesConfirm, initialRanges, initialUseHeader }: any) => {
  const [sRanges, setSRanges] = useState<string[]>(initialRanges || []);
  const [uHead, setUHead] = useState(initialUseHeader !== false);
  const [cWidth, setCWidth] = useState(120);
  const [dStart, setDStart] = useState<{ c: number, r: number } | null>(null);
  const [dEnd, setDEnd] = useState<{ c: number, r: number } | null>(null);
  const sData = useMemo(() => (!workbook || !currentSheet) ? [] : XLSX.utils.sheet_to_json(workbook.Sheets[currentSheet], { header: 1, defval: "", blankrows: true }).slice(0, 100) as any[][], [workbook, currentSheet]);
  const cols = useMemo(() => sData.length > 0 ? Array.from({ length: Math.max(...sData.map(r => r.length)) }, (_, i) => XLSX.utils.encode_col(i)) : [], [sData]);
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-black/95 z-[100] flex items-center justify-center p-12 backdrop-blur-md">
      <div className="bg-[#1e1e1e] border border-[#444] rounded-3xl shadow-2xl w-full h-full flex flex-col overflow-hidden ring-1 ring-white/10">
        <div className="p-4 bg-[#1a1a1a] border-b border-[#444] flex flex-col font-sans">
          <div className="flex justify-between items-center mb-2">
            <h2 className="text-[12px] font-bold text-white uppercase tracking-[0.4em] flex items-center gap-2">
              <span className="text-blue-500">{Icons.Select}</span> Visual Range Selection
            </h2>
            <div className="flex items-center gap-8">
              <div className="flex items-center gap-3">
                <span className="text-[9px] text-[#555] font-bold uppercase tracking-widest">Width</span>
                <input type="range" min="60" max="400" value={cWidth} onChange={(e) => setCWidth(Number(e.target.value))} className="w-32 accent-blue-500" />
              </div>
              <button onClick={onClose} className="text-[#666] hover:text-white transition-colors text-xl flex items-center justify-center w-6 h-6">
                <span className="w-4 h-4 block flex items-center justify-center">{Icons.Close}</span>
              </button>
            </div>
          </div>
          <p className="text-[#888] text-[10px] leading-relaxed">
            🖱️ 抽出したいデータの範囲を<strong>マウスでドラッグして選択</strong>してください。（複数の範囲を選択することも可能です）<br/>
            ✅ 選択後、右側のパネルで<strong>「1行目をヘッダーとして使用する (First row as Header)」</strong>オプションをオンにすると、最初の行が列名として認識されます。
          </p>
        </div>
        <div className="flex-1 flex overflow-hidden">
          <div className="flex-1 overflow-auto bg-[#1a1a1a] custom-scrollbar"><table className="border-collapse table-fixed"><thead className="sticky top-0 z-10 bg-[#252526]"><tr><th className="w-12 border-b border-[#444]"></th>{cols.map(c => <th key={c} style={{ width: cWidth }} className="px-3 py-1.5 border-r border-b border-[#444] text-[#888] text-[10px] font-bold">{c}</th>)}</tr></thead><tbody>{sData.map((row, r) => (<tr key={r}><td className="w-12 bg-[#252526] border-r border-b border-[#444] text-[#666] text-center text-[10px] font-mono">{r+1}</td>{cols.map((_, c) => { const isSel = dStart && dEnd && c >= Math.min(dStart.c, dEnd.c) && c <= Math.max(dStart.c, dEnd.c) && r >= Math.min(dStart.r, dEnd.r) && r <= Math.max(dStart.r, dEnd.r); return <td key={c} onMouseDown={() => {setDStart({c, r}); setDEnd({c, r});}} onMouseOver={() => dStart && setDEnd({c, r})} onMouseUp={() => { if(dStart && dEnd){ const range = XLSX.utils.encode_range({s:{c:Math.min(dStart.c, dEnd.c),r:Math.min(dStart.r, dEnd.r)},e:{c:Math.max(dStart.c, dEnd.c),r:Math.max(dStart.r, dEnd.r)}}); setSRanges(p => [...p, range]); setDStart(null); setDEnd(null); }}} style={{ width: cWidth }} className={`px-2 py-1.5 border-r border-b border-[#2a2a2a] truncate text-[11px] select-none cursor-crosshair ${isSel ? 'bg-blue-500/50 ring-1 ring-blue-400 ring-inset text-white' : 'text-[#aaa] hover:bg-[#222]'}`}>{row[c]}</td> })}</tr>))}</tbody></table></div>
          <div className="w-72 bg-[#252526] border-l border-[#444] p-6 flex flex-col gap-6 font-sans">
            <h3 className="text-[10px] font-bold text-[#888] uppercase border-b border-[#444] pb-3">Selection Settings</h3>
            <label className="flex items-center gap-3 cursor-pointer group"><input type="checkbox" checked={uHead} onChange={(e) => setUHead(e.target.checked)} className="w-4 h-4 accent-blue-500" /><span className="text-[11px] text-[#ccc] group-hover:text-white uppercase font-bold transition-colors">First row as Header</span></label>
            <div className="flex-1 overflow-auto space-y-3 pr-2 custom-scrollbar">{sRanges.map((rs, i) => (<div key={i} className="bg-[#1a1a1a] p-3 rounded-xl border border-[#444] flex justify-between items-center group hover:border-blue-500/50 transition-colors"><div className="flex flex-col"><span className="text-[8px] text-[#555] font-bold uppercase tracking-widest">Range {i+1}</span><span className="text-[11px] text-blue-400 font-mono font-bold">{rs}</span></div><button onClick={() => setSRanges(p => p.filter((_, idx) => idx !== i))} className="text-[#555] group-hover:text-white transition-colors flex items-center justify-center w-5 h-5"><span>{Icons.Close}</span></button></div>))}</div>
            <button onClick={() => {onRangesConfirm(sRanges, uHead); onClose();}} className="bg-blue-600 hover:bg-blue-500 py-3.5 rounded-xl text-[11px] font-bold text-white uppercase shadow-xl active:scale-95 transition-all">Apply Selection</button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default function App() { return <ReactFlowProvider><FlowBuilder /></ReactFlowProvider>; }
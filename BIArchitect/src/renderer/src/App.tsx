import React, { useState, useCallback, useMemo, memo, createContext, useContext, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { ReactFlow, Controls, Background, addEdge, Handle, Position, ReactFlowProvider, useReactFlow, useNodesState, useEdgesState, Panel } from '@xyflow/react';
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import type { Node, Edge } from '@xyflow/react';
import '@xyflow/react/dist/style.css';

type CustomNode = Node<Record<string, any>>;

type AppContextType = {
  workbooks: Record<string, XLSX.WorkBook>;
  setWorkbooks: React.Dispatch<React.SetStateAction<Record<string, XLSX.WorkBook>>>;
  setRangeModalNode: React.Dispatch<React.SetStateAction<string | null>>;
  nodeFlowData: Record<string, any>;
  isAutoCameraMove: boolean;
  focusNode: (id: string, force?: boolean, isDragging?: boolean) => void;
  theme: 'light' | 'dark';
  activePreviewId: string | null;
};
const AppContext = createContext<AppContextType>({} as AppContextType);

const GlobalStyle = () => (
  <style>{`
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    .dark ::-webkit-scrollbar-track { background: #1a1a1a; }
    .dark ::-webkit-scrollbar-thumb { background: #444; border-radius: 4px; }
    ::-webkit-scrollbar-track { background: #f3f4f6; }
    ::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 4px; }
    
    .react-flow__node { cursor: grab !important; }
    .react-flow__node:active { cursor: grabbing !important; }
    .react-flow__handle { width: 18px !important; height: 18px !important; border: 3px solid #fff !important; background-color: #3b82f6 !important; transition: transform 0.1s ease; }
    .dark .react-flow__handle { border: 3px solid #1e1e1e !important; background-color: #38bdf8 !important; }
    .react-flow__handle:hover { transform: scale(1.5); }
    .custom-scrollbar::-webkit-scrollbar { width: 4px; }
    
    /* Tailwindの設定に依存せず、Controlsのダークモードを強制適用 */
    .dark .react-flow__controls-button {
      background-color: #252526 !important;
      border-bottom: 1px solid #444 !important;
      fill: #ccc !important;
    }
    .dark .react-flow__controls-button:hover {
      background-color: #333 !important;
    }
    
    select {
      appearance: none;
      background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%233b82f6%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
      background-repeat: no-repeat, repeat;
      background-position: right .7em top 50%, 0 0;
      background-size: .65em auto, 100%;
    }
    .dark select {
      background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%2338bdf8%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.8%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
    }
    
    /* ★ PDF印刷用のスタイル */
    @media print {
      body { background: white; }
      .no-print { display: none !important; }
      .print-preview-area { position: absolute; left: 0; top: 0; width: 100%; height: auto; overflow: visible !important; background: white !important; }
      .print-preview-area * { color: black !important; }
      .print-preview-area table { width: 100%; border-collapse: collapse; }
      .print-preview-area th, .print-preview-area td { border: 1px solid #ccc !important; padding: 4px; font-size: 10px; }
      .print-preview-area thead { background-color: #eee !important; }
      .recharts-text { fill: black !important; }
      .recharts-cartesian-grid line { stroke: #eee !important; }
    }
  `}</style>
);

const IconSvg = ({ children }: { children: React.ReactNode }) => (
  <svg viewBox="0 0 24 24" width="1em" height="1em" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round">{children}</svg>
);

const Icons = {
  Sun: <IconSvg><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></IconSvg>,
  Moon: <IconSvg><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></IconSvg>,
  Help: <IconSvg><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></IconSvg>,
  Source: <IconSvg><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/></IconSvg>,
  FolderAuto: <IconSvg><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/><path d="M12 11v6"/><polyline points="9 14 12 17 15 14"/></IconSvg>,
  Web: <IconSvg><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></IconSvg>,
  Union: <IconSvg><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></IconSvg>,
  Join: <IconSvg><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></IconSvg>,
  Vlookup: <IconSvg><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></IconSvg>,
  Minus: <IconSvg><line x1="5" y1="12" x2="19" y2="12"/></IconSvg>,
  GroupBy: <IconSvg><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></IconSvg>,
  Sort: <IconSvg><path d="m3 16 4 4 4-4"/><path d="M7 20V4"/><path d="m21 8-4-4-4 4"/><path d="M17 4v16"/></IconSvg>,
  Transform: <IconSvg><path d="M20 7h-9"/><path d="M14 17H5"/><circle cx="17" cy="7" r="3"/><circle cx="8" cy="17" r="3"/></IconSvg>,
  Calculate: <IconSvg><rect x="4" y="2" width="16" height="20" rx="2" ry="2"/><line x1="8" y1="6" x2="16" y2="6"/><line x1="16" y1="14" x2="16" y2="18"/><path d="M16 10h.01"/><path d="M12 10h.01"/><path d="M8 10h.01"/><path d="M12 14h.01"/><path d="M8 14h.01"/><path d="M12 18h.01"/><path d="M8 18h.01"/></IconSvg>,
  Select: <IconSvg><path d="m3 17 2 2 4-4"/><path d="m3 7 2 2 4-4"/><path d="M13 6h8"/><path d="M13 12h8"/><path d="M13 18h8"/></IconSvg>,
  Filter: <IconSvg><polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/></IconSvg>,
  Chart: <IconSvg><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></IconSvg>,
  Dashboard: <IconSvg><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="21" x2="9" y2="9"/></IconSvg>,
  Warning: <IconSvg><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></IconSvg>,
  Folder: <IconSvg><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></IconSvg>,
  File: <IconSvg><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></IconSvg>,
  Refresh: <IconSvg><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><polyline points="3 3 3 8 8 8"/></IconSvg>,
  Paste: <IconSvg><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1" ry="1"/></IconSvg>,
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

  if (node.type === 'webSourceNode') {
    if (!node.data.rawData) return { data: [], headers: [] };
    try {
      let cData: any[] = [];
      let fHdrs: string[] = [];
      const type = node.data.dataType || 'auto';
      const raw = node.data.rawData;

      if (type === 'json' || (type === 'auto' && (raw.trim().startsWith('[') || raw.trim().startsWith('{')))) {
        const parsed = JSON.parse(raw);
        cData = Array.isArray(parsed) ? parsed : [parsed];
        if (cData.length > 0) fHdrs = Object.keys(cData[0]);
      } else if (type === 'html' || (type === 'auto' && raw.toLowerCase().includes('<table'))) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(raw, "text/html");
        
        const tables = Array.from(doc.querySelectorAll('table'));
        let targetTable = tables[0];
        let maxRows = 0;
        tables.forEach(t => {
          const rows = t.querySelectorAll('tr').length;
          if (rows > maxRows) {
            maxRows = rows;
            targetTable = t;
          }
        });

        if (targetTable) {
          const rows = Array.from(targetTable.querySelectorAll('tr'));
          let headerIdx = 0;
          for (let i = 0; i < rows.length; i++) {
            if (rows[i].querySelector('th')) { headerIdx = i; break; }
          }

          let ths = Array.from(rows[headerIdx]?.querySelectorAll('th, td') || []).map(th => th.textContent?.replace(/\s+/g, ' ').trim() || '');
          const seen = new Set();
          fHdrs = (ths.length > 0 ? ths : Array.from({length: rows[0]?.querySelectorAll('td').length || 0}, (_, i) => `Col_${i+1}`)).map((h, i) => {
            let newH = h === '' ? `Col_${i+1}` : h;
            let counter = 1;
            while (seen.has(newH)) { newH = `${h}_${counter}`; counter++; }
            seen.add(newH);
            return newH;
          });
          
          for (let i = headerIdx + 1; i < rows.length; i++) {
            const tds = Array.from(rows[i].querySelectorAll('td')).map(td => td.textContent?.replace(/\s+/g, ' ').trim() || '');
            if (tds.length > 0 && tds.some(t => t !== '')) {
              const obj: any = {};
              fHdrs.forEach((h, cIdx) => obj[h] = tds[cIdx] || '');
              cData.push(obj);
            }
          }
        }
      } else {
        const parsed = Papa.parse(raw, { header: true, skipEmptyLines: true });
        cData = parsed.data;
        fHdrs = parsed.meta.fields || [];
      }
      return { data: cData, headers: fHdrs };
    } catch (e) { return { data: [], headers: [] }; }
  }

  if (node.type === 'pasteNode') {
    if (!node.data.rawData) return { data: [], headers: [] };
    try {
      const parsed = Papa.parse(node.data.rawData.trim(), { header: true, skipEmptyLines: true });
      return { data: parsed.data, headers: parsed.meta.fields || [] };
    } catch (e) { return { data: [], headers: [] }; }
  }

  if (node.type === 'dataNode' || node.type === 'folderSourceNode') {
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

  if (node.type === 'unionNode' || node.type === 'joinNode' || node.type === 'minusNode' || node.type === 'vlookupNode') {
    const eA = edges.find(e => e.target === nId && (e as any).targetHandle === 'input-a');
    const eB = edges.find(e => e.target === nId && (e as any).targetHandle === 'input-b');
    if (!eA || !eB) return { data: [], headers: [] };
    const rA = calcData(eA.source, nodes, edges, wbs), rB = calcData(eB.source, nodes, edges, wbs);
    
    if (node.type === 'unionNode') return { data: [...rA.data, ...rB.data], headers: rA.headers };
    
    if (node.type === 'minusNode') {
      const { keyA, keyB } = node.data;
      if (!keyA || !keyB) return rA;
      const bKeys = new Set(rB.data.map(b => String(b[keyB as string])));
      const mData = rA.data.filter(a => !bKeys.has(String(a[keyA as string])));
      return { data: mData, headers: rA.headers };
    }

    if (node.type === 'vlookupNode') {
      const { keyA, keyB, fetchCol, targetCol } = node.data;
      if (!keyA || !keyB || !fetchCol || !targetCol) return rA;

      const newColName = targetCol;
      const bMap = new Map();
      rB.data.forEach(b => {
        bMap.set(String(b[keyB as string]), b[fetchCol as string]);
      });

      const vData = rA.data.map(a => {
        const key = String(a[keyA as string]);
        const val = bMap.has(key) ? bMap.get(key) : null;
        return { ...a, [newColName]: val };
      });
      return { data: vData, headers: [...rA.headers, newColName] };
    }

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
    const { targetCol, command, param0, applyCond, condCol, condOp, condVal } = node.data;
    if (command === 'remove_duplicates') {
      const seen = new Set();
      out = out.filter(r => {
        const key = targetCol ? String(r[targetCol as string]) : JSON.stringify(r);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    } else if (targetCol && command) {
      out = out.map(r => {
        if (applyCond && condCol && condOp) {
          const cValCheck = String(r[condCol as string] || '').toLowerCase();
          const tValCheck = String(condVal || '').toLowerCase();
          const cNumCheck = Number(r[condCol as string]), tNumCheck = Number(condVal);
          let isMatch = false;
          switch (condOp) {
            case 'exact': isMatch = (cValCheck === tValCheck); break;
            case 'not': isMatch = (cValCheck !== tValCheck); break;
            case 'gt': isMatch = (!isNaN(cNumCheck) && !isNaN(tNumCheck)) ? cNumCheck > tNumCheck : cValCheck > tValCheck; break;
            case 'lt': isMatch = (!isNaN(cNumCheck) && !isNaN(tNumCheck)) ? cNumCheck < tNumCheck : cValCheck < tValCheck; break;
            default: isMatch = cValCheck.includes(tValCheck);
          }
          if (!isMatch) return r; 
        }

        let v = r[targetCol as string];
        let vStr = v === null || v === undefined ? "" : String(v);

        if (command === 'replace') v = vStr.replace(param0 || "", "");
        else if (command === 'math_mul') v = Number(vStr) * Number(param0 || 1);
        else if (command === 'add_suffix') v = vStr + (param0 || "");
        else if (command === 'concat') v = vStr + String(param0 || '');
        else if (command === 'to_string') v = vStr;
        else if (command === 'to_number') {
            const num = Number(vStr.replace(/[^0-9.-]/g, ''));
            v = isNaN(num) ? null : num;
        }
        else if (command === 'fill_zero') {
            if (vStr.trim() === '') v = 0;
        }
        else if (command === 'zero_padding') {
            const len = Number(param0) || 1;
            v = vStr.padStart(len, '0');
        }
        else if (command === 'round') {
            const d = Number(param0) || 0;
            const m = Math.pow(10, d);
            v = Math.round(Number(vStr) * m) / m;
        }
        else if (command === 'mod') {
            const denom = Number(param0) || 1;
            v = Number(vStr) % denom;
            if (isNaN(v)) v = null;
        }
        else if (command === 'substring') {
            const params = String(param0 || '1').split(',').map(s => Number(s.trim()));
            const start = params[0] || 1;
            const len = params.length > 1 ? params[1] : vStr.length;
            const sIdx = Math.max(0, start - 1);
            v = vStr.slice(sIdx, sIdx + len);
        }
        else if (command === 'case_when') {
          let match = false;
          const cwValCheck = vStr.toLowerCase(), cwTVal = String(node.data.cwCondVal || '').toLowerCase();
          const cwNum = Number(v), cwTNum = Number(node.data.cwCondVal);
          switch (node.data.cwCondOp) {
            case 'exact': match = (cwValCheck === cwTVal); break;
            case 'not': match = (cwValCheck !== cwTVal); break;
            case 'gt': match = (!isNaN(cwNum) && !isNaN(cwTNum)) ? cwNum > cwTNum : cwValCheck > cwTVal; break;
            case 'lt': match = (!isNaN(cwNum) && !isNaN(cwTNum)) ? cwNum < cwTNum : cwValCheck < cwTVal; break;
            default: match = cwValCheck.includes(cwTVal);
          }
          v = match ? (node.data.trueVal || '') : (node.data.falseVal || '');
        }

        return { ...r, [targetCol as string]: v };
      });
    }
  }

  if (node.type === 'calculateNode') {
    const { colA, colB, operator, newColName } = node.data;
    if (colA && colB && newColName) {
      out = out.map(r => {
        let valA = r[colA];
        let valB = r[colB];
        let result: any = null;

        if (operator === 'concat') {
          result = String(valA || '') + String(valB || '');
        } else {
          const numA = Number(valA) || 0;
          const numB = Number(valB) || 0;
          if (operator === 'add') result = numA + numB;
          else if (operator === 'sub') result = numA - numB;
          else if (operator === 'mul') result = numA * numB;
          else if (operator === 'div') result = numB !== 0 ? numA / numB : null;
        }

        return { ...r, [newColName]: result };
      });
      h = [...h, newColName];
    }
  }

  return { data: out, headers: h };
};

const useNodeLogic = (id: string) => {
  const { nodeFlowData, isAutoCameraMove, focusNode, theme, activePreviewId } = useContext(AppContext);
  const { updateNodeData, setNodes, setEdges, getEdges } = useReactFlow();
  const isDark = theme === 'dark';
  
  return { 
    fData: nodeFlowData[id] || { incomingHeaders: [], headersA: [], headersB: [] }, 
    isDark,
    activePreviewId,
    onChg: (k: string, v: any) => {
      updateNodeData(id, { [k]: v });
      if (['command', 'joinType', 'chartType', 'matchType', 'aggType', 'groupCol', 'sortCol', 'targetCol', 'filterCol', 'xAxis', 'yAxis', 'applyCond', 'dataType', 'fetchCol', 'colA', 'colB', 'operator', 'newColName'].includes(k)) {
        focusNode(id);
      }
    }, 
    onDel: () => { 
      if (isAutoCameraMove) {
        const edges = getEdges();
        const incomingEdge = edges.find((e: any) => e.target === id);
        if (incomingEdge) {
          focusNode(incomingEdge.source);
        }
      }
      setNodes((nds: any) => nds.filter((n: any) => n.id !== id)); 
      setEdges((eds: any) => eds.filter((e: any) => e.source !== id && e.target !== id)); 
    } 
  };
};

const TgtHandle = ({ id, style }: any) => <Handle type="target" position={Position.Left} id={id} style={style} className="react-flow__handle z-10 -ml-2 nodrag" />;
const SrcHandle = () => <Handle type="source" position={Position.Right} className="react-flow__handle z-10 -mr-2 nodrag" />;

const NodeWrap = memo(({ id, title, col, children, showTgt = true, multi = false, summary = '', helpText = '' }: any) => {
  const { onDel, isDark, activePreviewId } = useNodeLogic(id);
  const [showHelp, setShowHelp] = useState(false);
  const [isCollapsed, setIsCollapsed] = useState(false);
  
  const isPreview = activePreviewId === id;
  const borderClass = isPreview 
    ? `border-blue-500 ring-2 ring-blue-500 shadow-[0_0_15px_rgba(59,130,246,0.5)] ${isDark ? 'ring-offset-[#1a1a1a]' : 'ring-offset-gray-50'} ring-offset-2`
    : (isDark ? 'border-[#444]' : 'border-gray-300');

  return (
    <div className={`${isDark ? 'bg-[#252526]' : 'bg-white'} border ${borderClass} rounded-xl shadow-2xl min-w-[260px] pb-1 relative group/node transition-colors`}>
      <button onClick={(e) => { e.stopPropagation(); onDel(); }} className="absolute -top-3 -right-3 bg-blue-600 hover:bg-blue-500 text-white rounded-full w-6 h-6 flex items-center justify-center font-bold text-xs opacity-0 group-hover/node:opacity-100 transition-opacity shadow-lg z-20 nodrag">
        <span className="flex items-center justify-center w-3 h-3">{Icons.Close}</span>
      </button>
      <div 
        className={`${isDark ? 'bg-[#1a1a1a] border-[#444]' : 'bg-gray-50 border-gray-300'} p-2 border-b flex justify-between items-center rounded-t-xl select-none group/header cursor-pointer transition-colors`}
        onDoubleClick={() => setIsCollapsed(!isCollapsed)}
      >
        <div className="flex items-center gap-2">
          <button 
            onClick={(e) => { e.stopPropagation(); setIsCollapsed(!isCollapsed); }}
            className={`${isDark ? 'text-[#888] hover:text-white' : 'text-gray-500 hover:text-gray-800'} transition-colors flex items-center justify-center w-4 h-4 nodrag`}
            title={isCollapsed ? "展開する" : "最小化する"}
          >
            {isCollapsed ? Icons.ChevronDown : Icons.ChevronUp}
          </button>
          <span className={`text-[10px] font-bold tracking-widest uppercase ${col}`}>{title}</span>
          {helpText && (
            <div className="relative flex items-center">
              <button 
                onClick={(e) => { e.preventDefault(); e.stopPropagation(); setShowHelp(!showHelp); }}
                className={`text-[10px] flex items-center justify-center w-4 h-4 rounded-full border nodrag transition-colors ${showHelp ? 'bg-blue-500 text-white border-blue-500' : (isDark ? 'text-[#888] hover:text-white border-[#555] bg-[#222]' : 'text-gray-500 hover:text-gray-800 border-gray-300 bg-gray-100')}`}
              >
                ?
              </button>
              {showHelp && (
                <div className={`absolute left-6 top-1/2 -translate-y-1/2 w-56 ${isDark ? 'bg-[#111] text-[#ccc] border-[#444]' : 'bg-white text-gray-700 border-gray-300'} text-[11px] p-3 rounded-lg border z-50 shadow-2xl normal-case tracking-normal leading-relaxed`}>
                  {helpText}
                </div>
              )}
            </div>
          )}
        </div>
        {summary && <span className={`text-[9px] ${isDark ? 'bg-[#333] text-[#aaa]' : 'bg-gray-200 text-gray-600'} px-2 py-0.5 rounded-full max-w-[100px] truncate font-mono`} title={summary}>{summary}</span>}
      </div>
      <div className={`relative transition-all ${isCollapsed ? 'h-8' : 'p-4 flex flex-col gap-3'}`}>
        {multi ? <><TgtHandle id="input-a" style={{ top: '30%' }} /><TgtHandle id="input-b" style={{ top: '70%' }} /></> : (showTgt && <TgtHandle />)}
        <div className={isCollapsed ? 'hidden' : 'contents'}>
          {children}
        </div>
      </div>
      <SrcHandle />
    </div>
  );
});

const DataNode = memo(({ id, data }: any) => {
  const { setWorkbooks, setRangeModalNode, focusNode, theme } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const isDark = theme === 'dark';

  const onUp = (e: any) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader(); r.onload = (evt: any) => {
      const wb = XLSX.read(evt.target.result, { type: 'binary' });
      setWorkbooks((p: any) => ({ ...p, [id]: wb }));
      updateNodeData(id, { fileName: f.name, sheetNames: wb.SheetNames, currentSheet: wb.SheetNames[0], needsUpload: false });
      focusNode(id);
    }; r.readAsBinaryString(f);
  };
  const summary = data.fileName ? data.fileName : '';
  return (
    <NodeWrap id={id} title="Source" col={isDark ? "text-blue-400" : "text-blue-600"} showTgt={false} summary={summary} helpText="ローカルのCSVやExcelファイルを選択して読み込みます。パネルから抽出範囲やヘッダーの設定が可能です。">
      {data.needsUpload ? (
        <div className="space-y-3">
          <div className={`text-[10px] ${isDark ? 'text-white bg-blue-500/20 border-blue-500/50' : 'text-gray-800 bg-blue-50 border-blue-200'} flex items-center gap-2 p-2 rounded border`}>
            <span className={`${isDark ? 'text-blue-400' : 'text-blue-600'} flex items-center justify-center`}>{Icons.Warning}</span> Missing: {data.fileName}
          </div>
          <label className={`cursor-pointer ${isDark ? 'text-blue-400 border-blue-500/50 hover:bg-blue-500/20' : 'text-blue-600 border-blue-300 hover:bg-blue-50'} text-[10px] border border-dashed p-3 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors shadow-sm animate-pulse nodrag`}>
            <span className="flex items-center justify-center">{Icons.Folder}</span> 再設定 <input type="file" accept=".csv,.xlsx" className="hidden" onChange={onUp} />
          </label>
        </div>
      ) : !data.fileName ? (
        <label className={`cursor-pointer ${isDark ? 'text-blue-400 border-blue-500/50 hover:bg-blue-500/10' : 'text-blue-600 border-blue-300 hover:bg-blue-50'} text-[10px] border border-dashed p-4 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors nodrag`}>
          <span className="flex items-center justify-center w-4 h-4">{Icons.Folder}</span> Load File <input type="file" accept=".csv,.xlsx" className="hidden" onChange={onUp} />
        </label>
      ) : (
        <div className="space-y-3">
          <div className={`flex justify-between items-center ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'} p-2 rounded border transition-colors`}>
            <div className={`text-[10px] ${isDark ? 'text-white' : 'text-gray-800'} font-bold truncate flex items-center gap-2`}>
              <span className={`${isDark ? 'text-blue-400' : 'text-blue-600'} flex items-center justify-center`}>{Icons.File}</span> {data.fileName}
            </div>
            <label className={`cursor-pointer ${isDark ? 'text-blue-400 hover:text-white' : 'text-blue-600 hover:text-gray-900'} text-[12px] font-bold uppercase transition-colors nodrag`} title="Change File">
              <span className="flex items-center justify-center">{Icons.Refresh}</span> <input type="file" className="hidden" onChange={onUp} />
            </label>
          </div>
          <button onClick={() => setRangeModalNode(id)} className={`w-full py-2 ${isDark ? 'bg-blue-600/20 text-blue-400 border-blue-500/30 hover:bg-blue-600/40' : 'bg-blue-50 text-blue-600 border-blue-200 hover:bg-blue-100'} text-[10px] font-bold rounded border uppercase tracking-widest transition-colors nodrag`}>範囲選択</button>
          <label className="flex items-center gap-2 pt-2 cursor-pointer group"><input type="checkbox" checked={data.useFirstRowAsHeader !== false} onChange={(e) => updateNodeData(id, { useFirstRowAsHeader: e.target.checked })} className="accent-blue-500 w-4 h-4 cursor-pointer nodrag" /><span className={`text-[10px] ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-600 group-hover:text-gray-900'} font-bold uppercase transition-colors`}>1行目をヘッダーにする</span></label>
        </div>
      )}
    </NodeWrap>
  );
});

const FolderSourceNode = memo(({ id, data }: any) => {
  const { setWorkbooks, setRangeModalNode, focusNode, theme } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const isDark = theme === 'dark';
  
  const [isLoading, setIsLoading] = useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);

  const handleFolderSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsLoading(true);

    try {
      let latestFile: File | null = null;
      let latestTime = 0;

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const lowerName = file.name.toLowerCase();
        if (lowerName.endsWith('.csv') || lowerName.endsWith('.xlsx')) {
          if (file.lastModified > latestTime) {
            latestTime = file.lastModified;
            latestFile = file;
          }
        }
      }

      if (latestFile) {
        const r = new FileReader();
        r.onload = (evt: any) => {
          try {
            const wb = XLSX.read(evt.target.result, { type: 'binary' });
            setWorkbooks((p: any) => ({ ...p, [id]: wb }));
            
            const folderPath = latestFile!.webkitRelativePath;
            const folderName = folderPath ? folderPath.split('/')[0] : 'Selected Folder';

            updateNodeData(id, { 
                folderName: folderName, 
                fileName: latestFile!.name, 
                sheetNames: wb.SheetNames, 
                currentSheet: wb.SheetNames[0], 
                needsUpload: false 
            });
            focusNode(id);
          } catch(err) {
            alert('ファイルの解析に失敗しました。対応していない形式か、ファイルが破損しています。');
          } finally {
            setIsLoading(false);
          }
        };
        r.onerror = () => {
          alert('ファイルの読み込み中にエラーが発生しました。');
          setIsLoading(false);
        };
        r.readAsBinaryString(latestFile);
      } else {
        alert("選択されたフォルダ内に、CSVまたはExcelファイル（.csv, .xlsx）が見つかりませんでした。");
        setIsLoading(false);
      }
    } catch (err) {
      console.error("フォルダ読み込みエラー:", err);
      setIsLoading(false);
    }
    
    e.target.value = '';
  };

  const triggerClick = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!isLoading && inputRef.current) {
      inputRef.current.click();
    }
  };

  const summary = data.folderName ? `[${data.folderName}] ${data.fileName}` : '';
  return (
    <NodeWrap id={id} title="Auto Folder" col={isDark ? "text-indigo-400" : "text-indigo-600"} showTgt={false} summary={summary} helpText="指定したフォルダを監視し、その中の『最も新しく更新されたファイル』を自動で読み込みます。毎月の売上データ追加など、定期的な更新作業に便利です。">
      {data.needsUpload ? (
        <div className="space-y-3">
          <div className={`text-[10px] ${isDark ? 'text-white bg-indigo-500/20 border-indigo-500/50' : 'text-gray-800 bg-indigo-50 border-indigo-200'} flex items-center gap-2 p-2 rounded border`}>
            <span className={`${isDark ? 'text-indigo-400' : 'text-indigo-500'} flex items-center justify-center`}>{Icons.Warning}</span> Missing: {data.folderName}
          </div>
          <button onClick={triggerClick} disabled={isLoading} className={`w-full ${isDark ? 'text-indigo-400 border-indigo-500/50 hover:bg-indigo-500/20' : 'text-indigo-600 border-indigo-300 hover:bg-indigo-50'} text-[10px] border border-dashed p-3 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors shadow-sm animate-pulse disabled:opacity-50 nodrag`}>
            <span className="flex items-center justify-center">{Icons.FolderAuto}</span> {isLoading ? '読込中...' : 'フォルダを再選択'}
          </button>
        </div>
      ) : !data.folderName ? (
        <button onClick={triggerClick} disabled={isLoading} className={`w-full ${isDark ? 'text-indigo-400 border-indigo-500/50 hover:bg-indigo-500/10' : 'text-indigo-600 border-indigo-300 hover:bg-indigo-50'} text-[10px] border border-dashed p-4 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors disabled:opacity-50 nodrag`}>
          <span className="flex items-center justify-center w-4 h-4">{Icons.FolderAuto}</span> {isLoading ? '読込中...' : 'Select Folder'}
        </button>
      ) : (
        <div className="space-y-3">
          <div className={`flex flex-col ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'} p-2 rounded border gap-1 transition-colors`}>
            <div className={`text-[9px] ${isDark ? 'text-[#888]' : 'text-gray-500'} font-bold truncate flex items-center gap-1`}>
              <span className="flex items-center justify-center w-3 h-3">{Icons.Folder}</span> {data.folderName}
            </div>
            <div className={`text-[10px] ${isDark ? 'text-white' : 'text-gray-800'} font-bold truncate flex items-center gap-2 justify-between`}>
              <div className="flex items-center gap-1 truncate"><span className={`${isDark ? 'text-indigo-400' : 'text-indigo-500'} flex items-center justify-center w-3 h-3`}>{Icons.File}</span> <span className="truncate">{data.fileName}</span></div>
              <button onClick={triggerClick} disabled={isLoading} className={`cursor-pointer ${isDark ? 'text-indigo-400 hover:text-white' : 'text-indigo-500 hover:text-gray-900'} text-[12px] font-bold uppercase transition-colors disabled:opacity-50 nodrag shrink-0`} title="Rescan Folder">
                <span className="flex items-center justify-center">{Icons.Refresh}</span>
              </button>
            </div>
          </div>
          <button onClick={() => setRangeModalNode(id)} className={`w-full py-2 ${isDark ? 'bg-indigo-600/20 text-indigo-400 border-indigo-500/30 hover:bg-indigo-600/40' : 'bg-indigo-50 text-indigo-600 border-indigo-200 hover:bg-indigo-100'} text-[10px] font-bold rounded border uppercase tracking-widest transition-colors nodrag`}>範囲選択</button>
          <label className="flex items-center gap-2 pt-2 cursor-pointer group"><input type="checkbox" checked={data.useFirstRowAsHeader !== false} onChange={(e) => updateNodeData(id, { useFirstRowAsHeader: e.target.checked })} className="accent-indigo-500 w-4 h-4 cursor-pointer nodrag" /><span className={`text-[10px] ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-600 group-hover:text-gray-900'} font-bold uppercase transition-colors`}>1行目をヘッダーにする</span></label>
        </div>
      )}
      <input type="file" ref={inputRef} className="hidden" onChange={handleFolderSelect} {...{ webkitdirectory: "true", directory: "true" } as any} />
    </NodeWrap>
  );
});

const WebSourceNode = memo(({ id, data }: any) => {
  const { focusNode, theme } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const isDark = theme === 'dark';
  const [url, setUrl] = useState(data.url || '');
  const [loading, setLoading] = useState(false);

  const handleFetch = async () => {
    setLoading(true);
    try {
      const proxyUrl = `https://api.allorigins.win/get?url=${encodeURIComponent(url)}`;
      const res = await fetch(proxyUrl);
      if (!res.ok) throw new Error('Network response was not ok');
      const json = await res.json();
      if (!json.contents) throw new Error('No contents');
      const text = json.contents;
      
      updateNodeData(id, { url, fetchedUrl: url, rawData: text });
      focusNode(id);
    } catch(e) {
      console.error(e);
      alert("データの取得に失敗しました。URLが間違っているか、アクセスが拒否されました。");
    } finally {
      setLoading(false);
    }
  };

  const summary = data.fetchedUrl ? `Loaded` : '';

  return (
    <NodeWrap id={id} title="Web Source" col={isDark ? "text-emerald-400" : "text-emerald-600"} showTgt={false} summary={summary} helpText="指定したURLからデータを取得します。API(JSON)やCSV、またはWebページの表(HTML Table)を自動的に抽出します。">
      <div className="space-y-3">
         <div className="flex flex-col gap-2">
           <input type="text" className={`w-full text-[10px] p-2 border rounded outline-none nodrag transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-emerald-500' : 'bg-white border-gray-300 text-gray-800 focus:border-emerald-500'}`} placeholder="https://example.com/api/data" value={url} onChange={e=>setUrl(e.target.value)} />
           <button onClick={handleFetch} disabled={loading || !url} className="w-full bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white text-[10px] font-bold py-2 rounded nodrag shadow-sm transition-all active:scale-95 flex items-center justify-center gap-1.5">
             {loading ? '取得中...' : <><span className="flex items-center justify-center">{Icons.Web}</span> データ取得</>}
           </button>
         </div>
         <div className="space-y-1">
           <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} font-bold uppercase tracking-widest`}>Data Type</label>
           <select className={`w-full text-[10px] p-1.5 border rounded outline-none nodrag transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-emerald-400' : 'bg-white border-gray-300 text-gray-700 hover:border-emerald-500'}`} value={data.dataType || 'auto'} onChange={(e) => updateNodeData(id, { dataType: e.target.value })}>
             <option value="auto">自動判別 (Auto)</option>
             <option value="json">JSON API</option>
             <option value="csv">CSVファイル</option>
             <option value="html">Webページ (表抽出)</option>
           </select>
         </div>
         {data.fetchedUrl && <div className={`text-[8px] truncate mt-2 p-1.5 rounded border ${isDark ? 'text-emerald-400 bg-emerald-900/20 border-emerald-500/30' : 'text-emerald-700 bg-emerald-50 border-emerald-200'}`}>Loaded: {data.fetchedUrl}</div>}
      </div>
    </NodeWrap>
  )
});

const PasteNode = memo(({ id, data }: any) => {
  const { focusNode, theme } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const isDark = theme === 'dark';
  const [text, setText] = useState(data.rawData || '');

  const handleApply = () => {
    updateNodeData(id, { rawData: text });
    focusNode(id);
  };

  const summary = data.rawData ? `Loaded` : '';

  return (
    <NodeWrap id={id} title="Paste Data" col={isDark ? "text-orange-400" : "text-orange-600"} showTgt={false} summary={summary} helpText="Excelやスプレッドシート、CSVなどのデータを直接ここに貼り付けてデータソースとして使用します。">
      <div className="space-y-3">
         <textarea 
           className={`w-full text-[10px] p-2 border rounded outline-none nodrag custom-scrollbar transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] focus:border-orange-400' : 'bg-white border-gray-300 text-gray-800 focus:border-orange-500'} h-24 whitespace-pre`} 
           placeholder="タブ区切り(Excelコピペ等)やカンマ区切りのテキストを貼り付けてください..." 
           value={text} 
           onChange={e=>setText(e.target.value)} 
         />
         <button onClick={handleApply} disabled={!text} className="w-full bg-orange-600 hover:bg-orange-500 disabled:opacity-50 text-white text-[10px] font-bold py-2 rounded nodrag shadow-sm transition-all active:scale-95 flex items-center justify-center gap-1.5">
           <span className="flex items-center justify-center">{Icons.Paste}</span> データを適用
         </button>
      </div>
    </NodeWrap>
  )
});

const UnionNode = memo(({ id }: any) => {
  const { isDark } = useNodeLogic(id);
  return <NodeWrap id={id} title="Union" col={isDark ? "text-blue-400" : "text-blue-600"} multi={true} summary="Append" helpText="2つのデータを「縦」に繋ぎ合わせます。（例: 1月のデータと2月のデータを1つの表にする）"><div className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'} text-center italic tracking-widest uppercase py-2`}>Merge Vertically</div></NodeWrap>
});

const JoinNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.keyA && data.keyB ? `${data.joinType || 'INNER'} JOIN` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] hover:border-blue-400' : 'bg-white border-gray-300 hover:border-blue-500'}`;
  
  return (
    <NodeWrap id={id} title="Join" col={isDark ? "text-blue-400" : "text-blue-600"} multi={true} summary={summary} helpText="2つのデータを共通の「キー（列）」を使って「横」に繋ぎ合わせます。">
      <div className="space-y-3">
        <select className={`${inputClass} ${isDark ? 'text-white' : 'text-gray-800'} font-bold`} value={data.joinType || 'inner'} onChange={(e) => onChg('joinType', e.target.value)}>
          <option value="inner">INNER JOIN (共通のみ)</option>
          <option value="left">LEFT JOIN (主データを全て残す)</option>
          <option value="right">RIGHT JOIN (副データを全て残す)</option>
        </select>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Main Key (A)</label>
          <select className={`${inputClass} ${isDark ? 'text-[#ccc]' : 'text-gray-700'}`} value={data.keyA || ''} onChange={(e) => onChg('keyA', e.target.value)}><option value="">Select Column...</option>{fData.headersA?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Sub Key (B)</label>
          <select className={`${inputClass} ${isDark ? 'text-[#ccc]' : 'text-gray-700'}`} value={data.keyB || ''} onChange={(e) => onChg('keyB', e.target.value)}><option value="">Select Column...</option>{fData.headersB?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
      </div>
    </NodeWrap>
  );
});

const VlookupNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.targetCol && data.fetchCol ? `Add ${data.targetCol}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-pink-400 focus:border-pink-400' : 'bg-white border-gray-300 text-gray-700 hover:border-pink-500 focus:border-pink-500'}`;

  return (
    <NodeWrap id={id} title="VLOOKUP" col={isDark ? "text-pink-400" : "text-pink-600"} multi={true} summary={summary} helpText="上(A)のデータの指定列をキーとして、下(B)のマスタデータを検索し、一致する行の特定列の値を新しい列として追加します。">
      <div className="space-y-3">
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Search Key (A)</label>
          <select className={inputClass} value={data.keyA || ''} onChange={(e) => onChg('keyA', e.target.value)}><option value="">Select Column...</option>{fData.headersA?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Master Key (B)</label>
          <select className={inputClass} value={data.keyB || ''} onChange={(e) => onChg('keyB', e.target.value)}><option value="">Select Column...</option>{fData.headersB?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Column to fetch (B)</label>
          <select className={`${inputClass} ${isDark ? 'text-white' : 'text-gray-900'} font-bold`} value={data.fetchCol || ''} onChange={(e) => onChg('fetchCol', e.target.value)}><option value="">Select Column...</option>{fData.headersB?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>New Column Name</label>
          <NodeInput className={`${inputClass} ${isDark ? 'text-white' : 'text-gray-900'}`} placeholder="e.g. Price" value={data.targetCol || ''} onChange={(v: any) => onChg('targetCol', v)} />
        </div>
      </div>
    </NodeWrap>
  );
});

const MinusNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.keyA && data.keyB ? `Minus` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-rose-400' : 'bg-white border-gray-300 text-gray-700 hover:border-rose-500'}`;

  return (
    <NodeWrap id={id} title="Minus" col={isDark ? "text-rose-500" : "text-rose-600"} multi={true} summary={summary} helpText="上(A)のデータから、下(B)のデータに存在するレコードを差し引いて残りのデータを抽出します。">
      <div className="space-y-3">
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Target Key (A)</label>
          <select className={inputClass} value={data.keyA || ''} onChange={(e) => onChg('keyA', e.target.value)}><option value="">Select Column...</option>{fData.headersA?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Subtract Key (B)</label>
          <select className={inputClass} value={data.keyB || ''} onChange={(e) => onChg('keyB', e.target.value)}><option value="">Select Column...</option>{fData.headersB?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
      </div>
    </NodeWrap>
  );
});

// ★ Calculate ノード (四則演算と文字結合)
const CalculateNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.newColName ? `Add ${data.newColName}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-teal-400' : 'bg-white border-gray-300 text-gray-700 hover:border-teal-500'}`;

  return (
    <NodeWrap id={id} title="Calculate" col={isDark ? "text-teal-400" : "text-teal-600"} summary={summary} helpText="2つの列を使って計算（足し算や文字列結合など）を行い、新しい列として追加します。">
      <div className="space-y-2">
        <select className={inputClass} value={data.colA || ''} onChange={(e) => onChg('colA', e.target.value)}>
          <option value="">Column A...</option>
          {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}
        </select>
        <select className={`w-full text-[10px] p-2 border rounded outline-none font-bold transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white hover:border-teal-400' : 'bg-white border-gray-300 text-gray-900 hover:border-teal-500'}`} value={data.operator || 'add'} onChange={(e) => onChg('operator', e.target.value)}>
          <option value="add">➕ 足し算 (Add)</option>
          <option value="sub">➖ 引き算 (Subtract)</option>
          <option value="mul">✖️ 掛け算 (Multiply)</option>
          <option value="div">➗ 割り算 (Divide)</option>
          <option value="concat">🔗 文字結合 (Concat)</option>
        </select>
        <select className={inputClass} value={data.colB || ''} onChange={(e) => onChg('colB', e.target.value)}>
          <option value="">Column B...</option>
          {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}
        </select>
        <div className={`space-y-1 mt-2 pt-2 border-t ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
          <label className={`text-[8px] uppercase tracking-widest font-bold ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>New Column Name</label>
          <NodeInput className={`w-full text-[10px] p-2 border rounded outline-none transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-teal-400' : 'bg-white border-gray-300 text-gray-900 focus:border-teal-500'}`} placeholder="e.g. Total Price" value={data.newColName || ''} onChange={(v: any) => onChg('newColName', v)} />
        </div>
      </div>
    </NodeWrap>
  );
});

const SortNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.sortCol ? `${data.sortCol} ${data.sortOrder === 'desc' ? '↓' : '↑'}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} title="Sort" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="指定した列の値を使って、データを昇順（小さい順）または降順（大きい順）に並び替えます。">
      <div className="space-y-2">
        <select className={inputClass} value={data.sortCol || ''} onChange={(e) => onChg('sortCol', e.target.value)}><option value="">Target Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        <select className={inputClass} value={data.sortOrder || 'asc'} onChange={(e) => onChg('sortOrder', e.target.value)}><option value="asc">Ascending ↑</option><option value="desc">Descending ↓</option></select>
      </div>
    </NodeWrap>
  );
});

const TransformNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.targetCol || data.command === 'remove_duplicates' ? `${data.command === 'case_when' ? 'CASE WHEN' : (data.command || '...')} on ${data.targetCol || 'All'}` : '';
  
  let ph = "Parameter (ex: ',' or '100')";
  if (data.command === 'zero_padding') ph = "桁数を入力 (例: 3)";
  else if (data.command === 'substring') ph = "開始位置, 文字数 (例: 1, 3)";
  else if (data.command === 'round') ph = "小数点以下の桁数 (例: 0)";
  else if (data.command === 'mod') ph = "割る数 (例: 2)";

  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] focus:border-blue-400 hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 focus:border-blue-500 hover:border-blue-500'}`;
  const inputClassWhite = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400 hover:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} title="Transform" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="データの内容を書き換えたり、型を変換したり、欠損値を補填したりする強力なクレンジングノードです。">
      <div className="space-y-2">
        <select className={inputClass} value={data.targetCol || ''} onChange={(e) => onChg('targetCol', e.target.value)}>
          <option value="">{data.command === 'remove_duplicates' ? '全体で重複判定 (All Columns)' : 'Target Column...'}</option>
          {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}
        </select>
        <select className={`${inputClassWhite} font-bold`} value={data.command || ''} onChange={(e) => onChg('command', e.target.value)}>
          <option value="">Select Action...</option>
          <option value="replace">不要文字を削除/置換</option>
          <option value="math_mul">数値を掛け算</option>
          <option value="add_suffix">末尾に文字追加</option>
          <option value="concat">文字結合 (CONCAT)</option>
          <option value="substring">文字抽出 (SUBSTRING)</option>
          <option value="case_when">条件分岐 (CASE WHEN)</option>
          <option value="to_string">文字列に変換 (ToString)</option>
          <option value="to_number">数値に変換 (ToNumber)</option>
          <option value="round">四捨五入 (ROUND)</option>
          <option value="mod">剰余/余り (MOD)</option>
          <option value="fill_zero">空白/nullを0で補填</option>
          <option value="zero_padding">指定桁数で0埋め (Zero Padding)</option>
          <option value="remove_duplicates">重複行を削除 (Remove Duplicates)</option>
        </select>

        {data.command && data.command !== 'remove_duplicates' && data.command !== 'case_when' && (
          <div className={`pt-2 border-t ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
            <label className="flex items-center gap-2 cursor-pointer group mb-2">
              <input type="checkbox" checked={data.applyCond || false} onChange={(e) => onChg('applyCond', e.target.checked)} className="accent-blue-500 w-3 h-3 cursor-pointer nodrag" />
              <span className={`text-[9px] font-bold transition-colors ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-500 group-hover:text-gray-900'}`}>特定の条件の時だけ適用する</span>
            </label>
            {data.applyCond && (
              <div className={`flex flex-col gap-1.5 p-2 rounded border ${isDark ? 'bg-[#111] border-[#333]' : 'bg-gray-50 border-gray-200'}`}>
                <select className={`w-full text-[9px] p-1.5 border rounded outline-none nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc]' : 'bg-white border-gray-300 text-gray-700'}`} value={data.condCol || ''} onChange={(e) => onChg('condCol', e.target.value)}>
                  <option value="">If Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}
                </select>
                <div className="flex gap-1">
                  <select className={`w-1/3 text-[9px] p-1.5 border rounded outline-none nodrag font-bold ${isDark ? 'bg-[#1a1a1a] border-[#444] text-blue-400' : 'bg-white border-gray-300 text-blue-600'}`} value={data.condOp || 'exact'} onChange={(e) => onChg('condOp', e.target.value)}>
                    <option value="exact">=</option><option value="not">≠</option><option value="gt">&gt;</option><option value="lt">&lt;</option><option value="includes">inc</option>
                  </select>
                  <NodeInput className={`w-2/3 text-[9px] p-1.5 border rounded outline-none transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500'}`} placeholder="Value..." value={data.condVal || ''} onChange={(v: any) => onChg('condVal', v)} />
                </div>
              </div>
            )}
          </div>
        )}
        
        {data.command === 'case_when' && (
          <div className={`space-y-2 mt-2 pt-2 border-t ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
            <div className={`text-[8px] font-bold uppercase ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>If Condition:</div>
            <div className="flex gap-1">
              <select className={`w-1/3 text-[10px] p-2 border rounded outline-none nodrag font-bold transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-blue-400 hover:border-blue-400' : 'bg-white border-gray-300 text-blue-600 hover:border-blue-500'}`} value={data.cwCondOp || 'exact'} onChange={(e) => onChg('cwCondOp', e.target.value)}>
                <option value="exact">=</option><option value="not">≠</option><option value="gt">&gt;</option><option value="lt">&lt;</option><option value="includes">inc</option>
              </select>
              <NodeInput className={`w-2/3 text-[10px] p-2 border rounded outline-none transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500'}`} placeholder="Value..." value={data.cwCondVal || ''} onChange={(v: any) => onChg('cwCondVal', v)} />
            </div>
            <NodeInput className={inputClassWhite} placeholder="Then (True Value)" value={data.trueVal || ''} onChange={(v: any) => onChg('trueVal', v)} />
            <NodeInput className={inputClass} placeholder="Else (False Value)" value={data.falseVal || ''} onChange={(v: any) => onChg('falseVal', v)} />
          </div>
        )}

        {data.command && !['case_when', 'to_string', 'to_number', 'fill_zero', 'remove_duplicates'].includes(data.command) && (
          <NodeInput 
            className={`${inputClassWhite} mt-2`} 
            placeholder={ph} 
            value={data.param0 || ''} 
            onChange={(v: any) => onChg('param0', v)} 
          />
        )}
      </div>
    </NodeWrap>
  );
});

const FilterNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const chgVal = (v: string, t: string) => onChg('filterVal', (t === 'gt' || t === 'lt') ? v.replace(/[０-９．－]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[^0-9.-]/g, '') : v);
  const op = data.matchType === 'gt' ? '>' : data.matchType === 'lt' ? '<' : data.matchType === 'exact' ? '=' : data.matchType === 'not' ? '≠' : 'inc';
  const summary = data.filterCol ? `${data.filterCol} ${op} ${data.filterVal || ''}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] hover:border-blue-400' : 'bg-white border-gray-300 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} title="Filter" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="指定した列の値が、入力した条件に一致する行だけを抽出して残します。（例: 売上が1000以上、など）">
      <div className="space-y-2">
        <select className={`${inputClass} ${isDark ? 'text-[#ccc]' : 'text-gray-700'}`} value={data.filterCol || ''} onChange={(e) => onChg('filterCol', e.target.value)}><option value="">Target Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        <div className="flex flex-col gap-2">
          <select className={`${inputClass} ${isDark ? 'text-white' : 'text-gray-900'} font-bold`} value={data.matchType || 'includes'} onChange={(e) => { onChg('matchType', e.target.value); if(data.filterVal) chgVal(String(data.filterVal), e.target.value); }}>
            <option value="includes">含む (Includes)</option><option value="exact">完全一致 (=)</option><option value="not">除外 (≠)</option><option value="gt">以上 (&gt;)</option><option value="lt">以下 (&lt;)</option>
          </select>
          <NodeInput className={`w-full text-[10px] p-2 border rounded outline-none transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500'}`} placeholder="Condition Value..." value={data.filterVal || ''} onChange={(v: any) => chgVal(v, data.matchType || 'includes')} />
        </div>
      </div>
    </NodeWrap>
  );
});

const SelectNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.selectedColumns?.length ? `${data.selectedColumns.length} cols selected` : '';
  return (
    <NodeWrap id={id} title="Select" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="データの中で「必要な列（カラム）」だけを選んで残し、不要な列を削除します。">
      <div className={`max-h-48 overflow-y-auto space-y-1 p-1 rounded border custom-scrollbar ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'}`}>
        {fData.incomingHeaders?.length > 0 ? fData.incomingHeaders.map((h: string) => (
          <label key={h} className={`flex items-center gap-2 text-[10px] p-1.5 rounded cursor-pointer group ${isDark ? 'text-[#ccc] hover:bg-[#333]' : 'text-gray-700 hover:bg-gray-200'}`}><input type="checkbox" checked={(data.selectedColumns || []).includes(h)} onChange={(e) => { const c = data.selectedColumns || []; onChg('selectedColumns', e.target.checked ? [...c, h] : c.filter((x: string) => x !== h)); }} className="accent-blue-500 w-3 h-3 nodrag" /><span className={`truncate ${isDark ? 'group-hover:text-white' : 'group-hover:text-gray-900'}`}>{h}</span></label>
        )) : <div className={`text-[9px] text-center py-4 ${isDark ? 'text-[#555]' : 'text-gray-500'}`}>Connect to input data</div>}
      </div>
    </NodeWrap>
  );
});

const GroupByNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.groupCol ? `By ${data.groupCol}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} title="Group By" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="指定したキーでデータをグループ化し、数値を合計(SUM)したり件数をカウント(CNT)したりして集計します。">
      <div className="space-y-3">
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Group Key</label>
          <select className={inputClass} value={data.groupCol || ''} onChange={(e) => onChg('groupCol', e.target.value)}><option value="">Select Key...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        </div>
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Aggregation</label>
          <div className="flex gap-2">
            <select className={`flex-1 text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 hover:border-blue-500'}`} value={data.aggCol || ''} onChange={(e) => onChg('aggCol', e.target.value)}><option value="">Value Col...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
            <select className={`w-20 text-[10px] p-2 border rounded outline-none font-bold transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white hover:border-blue-400' : 'bg-white border-gray-300 text-gray-900 hover:border-blue-500'}`} value={data.aggType || 'sum'} onChange={(e) => onChg('aggType', e.target.value)}><option value="sum">SUM</option><option value="count">CNT</option></select>
          </div>
        </div>
      </div>
    </NodeWrap>
  );
});

const ChartNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.xAxis && data.yAxis ? `${data.chartType} chart` : '';
  const inputClass = `w-full text-[9px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} title="Visualizer" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="データをグラフとして描画します。配置したグラフは下部の「Dashboard」タブで一覧表示できます。">
      <div className="space-y-2">
        <select className={`w-full text-[10px] p-2 border rounded outline-none font-bold transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white hover:border-blue-400' : 'bg-white border-gray-300 text-gray-900 hover:border-blue-500'}`} value={data.chartType || 'bar'} onChange={(e) => onChg('chartType', e.target.value)}>
          <option value="bar">Bar Chart</option>
          <option value="line">Line Chart</option>
        </select>
        <div className="grid grid-cols-2 gap-2">
          <div className="space-y-1"><label className={`text-[8px] ${isDark ? 'text-[#666]' : 'text-gray-500'} font-bold uppercase`}>X-Axis</label><select className={inputClass} value={data.xAxis || ''} onChange={(e) => onChg('xAxis', e.target.value)}><option value="">Select...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select></div>
          <div className="space-y-1"><label className={`text-[8px] ${isDark ? 'text-[#666]' : 'text-gray-500'} font-bold uppercase`}>Y-Axis</label><select className={inputClass} value={data.yAxis || ''} onChange={(e) => onChg('yAxis', e.target.value)}><option value="">Select...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select></div>
        </div>
      </div>
    </NodeWrap>
  );
});

// ★ Calculate Nodeを追加
const nodeTypesObj = { dataNode: DataNode, folderSourceNode: FolderSourceNode, webSourceNode: WebSourceNode, pasteNode: PasteNode, unionNode: UnionNode, joinNode: JoinNode, vlookupNode: VlookupNode, minusNode: MinusNode, groupByNode: GroupByNode, sortNode: SortNode, transformNode: TransformNode, calculateNode: CalculateNode, selectNode: SelectNode, filterNode: FilterNode, chartNode: ChartNode };

const NodeNavigator = ({ tList, nodes }: { tList: any[], nodes: CustomNode[] }) => {
  const { focusNode, theme } = useContext(AppContext);
  const [isMinimized, setIsMinimized] = useState(false);
  const isDark = theme === 'dark';

  if (isMinimized) {
    return (
      <Panel position="top-right" className={`${isDark ? 'bg-[#252526]/90 border-[#444] hover:bg-[#333]' : 'bg-white/90 border-gray-200 hover:bg-gray-50'} backdrop-blur-md border rounded-xl shadow-xl z-50 m-4 mr-6 cursor-pointer transition-colors no-print`} onClick={() => setIsMinimized(false)}>
        <div className={`flex items-center gap-2 p-2 px-3 text-[10px] font-bold uppercase tracking-widest ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>
          <span className={`${isDark ? 'text-blue-400' : 'text-blue-500'} flex items-center justify-center`}>{Icons.Diamond}</span>
          <span className={isDark ? 'text-white' : 'text-gray-800'}>{nodes.length} Nodes</span>
        </div>
      </Panel>
    );
  }

  return (
    <Panel position="top-right" className={`${isDark ? 'bg-[#252526]/90 border-[#444]' : 'bg-white/90 border-gray-200'} backdrop-blur-md border p-3 rounded-xl shadow-xl max-h-[300px] overflow-y-auto custom-scrollbar flex flex-col gap-1.5 w-60 z-50 m-4 mr-6 no-print transition-colors`}>
      <div className={`text-[10px] font-bold uppercase tracking-widest mb-2 px-1 flex items-center justify-between border-b pb-2 ${isDark ? 'text-[#888] border-[#444]' : 'text-gray-500 border-gray-200'}`}>
        <span className="flex items-center gap-1.5"><span className={`${isDark ? 'text-blue-400' : 'text-blue-500'} flex items-center justify-center`}>{Icons.Diamond}</span> Navigator</span>
        <div className="flex items-center gap-2">
          <span className={`${isDark ? 'bg-[#1a1a1a] text-[#aaa] border-[#333]' : 'bg-gray-100 text-gray-600 border-gray-200'} px-2 py-0.5 rounded-md text-[9px] border`}>{nodes.length} Nodes</span>
          <button onClick={() => setIsMinimized(true)} className={`${isDark ? 'hover:text-white' : 'hover:text-gray-800'} transition-colors flex items-center justify-center w-5 h-5`}>{Icons.ChevronUp}</button>
        </div>
      </div>
      {nodes.map(n => {
        const typeInfo = tList.find(t => t.t === n.type);
        const icon = typeInfo ? typeInfo.i : Icons.Diamond;
        const label = typeInfo ? typeInfo.l : 'Node';
        let subText = n.id;
        
        if ((n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.fileName) subText = n.data.fileName;
        else if (n.type === 'webSourceNode' && n.data.fetchedUrl) subText = "Loaded";
        else if (n.type === 'pasteNode' && n.data.rawData) subText = "Loaded";
        else if (n.type === 'filterNode' && n.data.filterCol) subText = `${n.data.filterCol} ${n.data.matchType} ${n.data.filterVal || ''}`;
        else if (n.type === 'chartNode' && n.data.chartType) subText = `${n.data.chartType} chart`;
        else if (n.type === 'transformNode' && (n.data.targetCol || n.data.command === 'remove_duplicates')) subText = `${n.data.command === 'case_when' ? 'CASE WHEN' : (n.data.command || '...')} on ${n.data.targetCol || 'All'}`;
        else if (n.type === 'calculateNode' && n.data.newColName) subText = `Add ${n.data.newColName}`;
        else if (n.type === 'sortNode' && n.data.sortCol) subText = `${n.data.sortCol} ${n.data.sortOrder}`;
        else if (n.type === 'groupByNode' && n.data.groupCol) subText = `By ${n.data.groupCol}`;
        else if (n.type === 'selectNode' && n.data.selectedColumns) subText = `${n.data.selectedColumns.length} cols selected`;
        else if (n.type === 'joinNode' || n.type === 'unionNode' || n.type === 'minusNode') subText = "Merge Data";
        else if (n.type === 'vlookupNode' && n.data.targetCol) subText = `Add ${n.data.targetCol}`;
        else if ((n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.useFirstRowAsHeader) subText = "Setup Required";

        return (
          <button 
            key={n.id} 
            onClick={() => focusNode(n.id, true)} 
            className={`text-left flex items-center gap-3 p-2 rounded-lg transition-all group border border-transparent active:scale-95 ${isDark ? 'hover:bg-[#333] hover:border-[#555]' : 'hover:bg-gray-100 hover:border-gray-300'}`}
          >
            <div className={`w-6 h-6 shrink-0 rounded flex items-center justify-center border ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white' : 'bg-gray-50 border-gray-200 text-gray-800'}`}>
              {icon}
            </div>
            <div className="flex flex-col flex-1 min-w-0">
              <span className={`text-[10px] font-bold uppercase tracking-wider truncate transition-colors ${isDark ? 'text-[#ccc] group-hover:text-white' : 'text-gray-700 group-hover:text-gray-900'}`}>{label}</span>
              <span className={`text-[9px] truncate transition-colors ${isDark ? 'text-[#666] group-hover:text-[#aaa]' : 'text-gray-500 group-hover:text-gray-700'}`}>{subText}</span>
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
  
  const { screenToFlowPosition, setCenter, getZoom, getNode } = useReactFlow();
  
  const [nodes, _setNodes, onNodesChange] = useNodesState<CustomNode>([{ id: 'n-1', type: 'dataNode', position: { x: 50, y: 150 }, data: { useFirstRowAsHeader: true } }]);
  const [edges, _setEdges, onEdgesChange] = useEdgesState<Edge>([]);
  
  const [previewTab, setPreviewTab] = useState<'table' | 'chart' | 'dashboard'>('table');
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [isPreviewOpen, setIsPreviewOpen] = useState(true);
  const [isSaveLoadOpen, setIsSaveLoadOpen] = useState(false);
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [isSqlModalOpen, setIsSqlModalOpen] = useState(false);
  
  const [showTutorial, setShowTutorial] = useState(() => !localStorage.getItem('bi-architect-visited'));
  
  const [savedFlows, setSavedFlows] = useState<any[]>([]);
  const [bottomHeight, setBottomHeight] = useState(300);
  const [isDragging, setIsDragging] = useState(false);
  
  const [isAutoCameraMove, setIsAutoCameraMove] = useState(true);
  const [previewNodeId, setPreviewNodeId] = useState<string | null>(null);

  const [theme, setTheme] = useState<'light' | 'dark'>('dark');
  
  useEffect(() => {
    const savedTheme = localStorage.getItem('bi-architect-theme') as 'light' | 'dark';
    if (savedTheme === 'light' || savedTheme === 'dark') setTheme(savedTheme);
    const localFlows = localStorage.getItem('bi-architect-flows');
    if (localFlows) setSavedFlows(JSON.parse(localFlows));
  }, []);

  useEffect(() => {
    if (theme === 'dark') {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [theme]);

  const toggleTheme = useCallback(() => {
    setTheme(t => {
      const next = t === 'dark' ? 'light' : 'dark';
      localStorage.setItem('bi-architect-theme', next);
      return next;
    });
  }, []);

  const closeTutorial = () => {
    setShowTutorial(false);
    localStorage.setItem('bi-architect-visited', '1');
  };

  const focusNode = useCallback((nodeId: string, force: boolean = false) => {
    if (!isAutoCameraMove && !force) return;
    
    setTimeout(() => {
      const n = getNode(nodeId) as CustomNode | undefined;
      if (n) {
        const h = n.measured?.height || 150;
        setCenter(n.position.x + CAMERA_OFFSET_X, n.position.y + h / 2, { zoom: getZoom(), duration: 1200 });
      }
    }, 50);
  }, [isAutoCameraMove, getNode, setCenter, getZoom]);

  const handleSave = (name: string) => { const up = [...savedFlows, { id: Date.now().toString(), name, updatedAt: new Date().toLocaleString(), flow: { nodes, edges } }]; setSavedFlows(up); localStorage.setItem('bi-architect-flows', JSON.stringify(up)); };
  const handleLoad = (f: any) => { _setNodes(f.flow.nodes.map((n: any) => (n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.fileName ? { ...n, data: { ...n.data, needsUpload: true } } : n)); _setEdges(f.flow.edges || []); setWorkbooks({}); };
  const onEdgeContextMenu = useCallback((e: React.MouseEvent, edge: Edge) => { e.preventDefault(); _setEdges((eds: any) => eds.filter((e: any) => e.id !== edge.id)); }, [_setEdges]);

  const handleReset = () => {
    _setNodes([{ id: 'n-1', type: 'dataNode', position: { x: 50, y: 150 }, data: { useFirstRowAsHeader: true } }]);
    _setEdges([]);
    setWorkbooks({});
    setIsResetModalOpen(false);
    setCenter(50 + CAMERA_OFFSET_X, 150 + 75, { zoom: 0.9, duration: 1200 });
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
      if (n.type === 'joinNode' || n.type === 'unionNode' || n.type === 'minusNode' || n.type === 'vlookupNode') {
        const eA = edges.find(e => e.target === n.id && (e as any).targetHandle === 'input-a'), eB = edges.find(e => e.target === n.id && (e as any).targetHandle === 'input-b');
        const hA = eA ? calcData(eA.source, nodes, edges, workbooks).headers : [], hB = eB ? calcData(eB.source, nodes, edges, workbooks).headers : [];
        map[n.id] = { headersA: hA, headersB: hB, incomingHeaders: [...new Set([...hA, ...hB])] };
      } else {
        const inEdge = edges.find(e => e.target === n.id);
        map[n.id] = { incomingHeaders: inEdge ? calcData(inEdge.source, nodes, edges, workbooks).headers : [] };
      }
    }); return map;
  }, [sHash, workbooks]); // eslint-disable-line

  const activePreviewId = useMemo(() => {
    if (previewNodeId && nodes.find(n => n.id === previewNodeId)) return previewNodeId;
    const term = nodes.find(n => !edges.some(e => e.source === n.id));
    return term?.id || null;
  }, [previewNodeId, nodes, edges]);

  const final = useMemo(() => {
    if (!activePreviewId) return { data: [], headers: [], chartConfig: null };

    const result = calcData(activePreviewId, nodes, edges, workbooks);
    const targetNode = nodes.find(n => n.id === activePreviewId);
    return { ...result, chartConfig: targetNode?.type === 'chartNode' ? targetNode.data : null };
  }, [sHash, workbooks, activePreviewId]); // eslint-disable-line

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
    { t: 'dataNode', l: 'Source', i: Icons.Source, c: 'text-blue-500 dark:text-blue-400', desc: 'CSVやExcelファイルを選択して読み込みます。' },
    { t: 'folderSourceNode', l: 'Auto Folder', i: Icons.FolderAuto, c: 'text-indigo-500 dark:text-indigo-400', desc: '指定フォルダを監視し、中の最新ファイルを自動で読み込みます。' },
    { t: 'webSourceNode', l: 'Web Source', i: Icons.Web, c: 'text-emerald-500 dark:text-emerald-400', desc: 'URLからAPI(JSON)やCSV、Webページの表(HTML)を取得します。' },
    { t: 'pasteNode', l: 'Paste Data', i: Icons.Paste, c: 'text-orange-500 dark:text-orange-400', desc: 'Excel等からコピーしたテキストデータを直接貼り付けて読み込みます。' },
    { t: 'unionNode', l: 'Union', i: Icons.Union, c: 'text-blue-500 dark:text-blue-400', desc: '2つのデータを「縦」に繋ぎ合わせます。(データの追加)' },
    { t: 'joinNode', l: 'Join', i: Icons.Join, c: 'text-blue-500 dark:text-blue-400', desc: '2つのデータを共通のキーで「横」に繋ぎます。' },
    { t: 'vlookupNode', l: 'VLOOKUP', i: Icons.Vlookup, c: 'text-pink-500 dark:text-pink-400', desc: '別データから一致するキーを検索し、特定の列の値を新しい列として追加します。' },
    { t: 'minusNode', l: 'Minus', i: Icons.Minus, c: 'text-rose-600 dark:text-rose-500', desc: '上のデータから下のデータに存在するレコードを差し引きます。' },
    { t: 'groupByNode', l: 'Group By', i: Icons.GroupBy, c: 'text-blue-500 dark:text-blue-400', desc: '指定したキーでデータをグループ化し、合計や件数を集計します。' },
    { t: 'sortNode', l: 'Sort', i: Icons.Sort, c: 'text-blue-500 dark:text-blue-400', desc: '指定した列を基準に、データを昇順・降順に並び替えます。' },
    { t: 'transformNode', l: 'Transform', i: Icons.Transform, c: 'text-blue-500 dark:text-blue-400', desc: 'データの内容を書き換えたり、型変換や0埋めなどを行うクレンジングノードです。' },
    { t: 'calculateNode', l: 'Calculate', i: Icons.Calculate, c: 'text-teal-500 dark:text-teal-400', desc: '2つの列の値を計算（足し算、文字結合など）し、新しい列として追加します。' },
    { t: 'selectNode', l: 'Select', i: Icons.Select, c: 'text-blue-500 dark:text-blue-400', desc: '必要な列(カラム)だけを選んで残し、不要な列を削除します。' },
    { t: 'filterNode', l: 'Filter', i: Icons.Filter, c: 'text-blue-500 dark:text-blue-400', desc: '条件に一致する行だけを抽出します。(例: 売上1000以上)' },
    { t: 'chartNode', l: 'Visualizer', i: Icons.Chart, c: 'text-blue-500 dark:text-blue-400', desc: 'データをグラフ化します。Dashboardタブで一覧表示できます。' }
  ];

  const isDark = theme === 'dark';
  
  const btnClasses = isSidebarOpen ? 'p-3 gap-4' : 'p-2 justify-center w-10 h-10';

  return (
    <AppContext.Provider value={{ workbooks, setWorkbooks, setRangeModalNode, nodeFlowData, isAutoCameraMove, focusNode, theme, activePreviewId: activePreviewId as string | null }}>
      <div className={`h-screen w-screen flex flex-col font-sans overflow-hidden transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
        <GlobalStyle />
        <div className={`border-b px-6 py-3 flex justify-between items-center z-40 gap-4 no-print transition-colors ${isDark ? 'bg-[#181818] border-[#333] shadow-md' : 'bg-white border-gray-200 shadow-sm'}`}>
          <h1 className={`text-[13px] font-bold tracking-[0.5em] uppercase flex items-center gap-3 shrink-0 ${isDark ? 'text-white' : 'text-gray-800'}`}>
            <span className="text-blue-500 w-4 h-4 flex items-center justify-center">{Icons.Diamond}</span>
            Visual Data Prep
          </h1>
          
          <div className="flex-1 flex justify-center">
            <div className={`px-4 py-1.5 rounded-full text-[10px] tracking-widest font-bold flex items-center gap-2 border ${isDark ? 'bg-[#1e1e1e] border-[#333] text-[#aaa]' : 'bg-gray-50 border-gray-200 text-gray-500'}`}>
              <span className="text-blue-500">{Icons.Diamond}</span> ノードを繋いで構築 <span className={`mx-1 ${isDark ? 'text-[#555]' : 'text-gray-400'}`}>|</span> 接続線は右クリックで削除
            </div>
          </div>

          <div className="flex items-center gap-3 shrink-0">
            <button onClick={() => setShowTutorial(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-2 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-[#666]' : 'bg-gray-50 hover:bg-gray-100 border-gray-200 text-gray-600'}`} title="Tutorial">
              <span className="flex items-center justify-center text-lg">{Icons.Help}</span>
            </button>

            <button onClick={toggleTheme} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-2 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-[#666]' : 'bg-gray-50 hover:bg-gray-100 border-gray-200 text-gray-600'}`} title="Toggle Theme">
              <span className="flex items-center justify-center text-lg">{isDark ? Icons.Sun : Icons.Moon}</span>
            </button>

            <button onClick={() => setIsAutoCameraMove(!isAutoCameraMove)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444]' : 'bg-gray-50 hover:bg-gray-100 border-gray-200'} ${isAutoCameraMove ? (isDark ? 'text-blue-400' : 'text-blue-500') : (isDark ? 'text-[#666]' : 'text-gray-500')}`} title="Auto Camera Focus">
              <span className="flex items-center justify-center gap-1">{Icons.Focus}</span> {isAutoCameraMove ? 'CAMERA FOCUS ON' : 'CAMERA FOCUS OFF'}
            </button>
            
            <button onClick={() => setIsSqlModalOpen(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border hidden md:flex ${isDark ? 'bg-[#252526] hover:bg-blue-900/30 border-[#444] hover:border-blue-500/50 text-[#aaa] hover:text-blue-400' : 'bg-gray-50 hover:bg-blue-50 border-gray-200 hover:border-blue-300 text-gray-600 hover:text-blue-600'}`}>
              <span className="flex items-center justify-center gap-1">{Icons.Code}</span> SQL
            </button>
            <button onClick={() => setIsResetModalOpen(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border hidden md:flex ${isDark ? 'bg-[#252526] hover:bg-red-900/30 border-[#444] hover:border-red-500/50 text-[#aaa] hover:text-red-400' : 'bg-gray-50 hover:bg-red-50 border-gray-200 hover:border-red-300 text-gray-600 hover:text-red-600'}`}>
              <span className="flex items-center justify-center gap-1">{Icons.Trash}</span> RESET
            </button>
            <button onClick={() => setIsSaveLoadOpen(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-white' : 'bg-gray-800 hover:bg-gray-700 border-gray-700 text-white'}`}>
              <span className="flex items-center justify-center gap-1">{Icons.Save}</span> / <span className="flex items-center justify-center gap-1">{Icons.Folder}</span> PROJECTS
            </button>
          </div>
        </div>
        <div className="flex-1 flex overflow-hidden relative no-print">
          <aside className={`border-r z-20 flex flex-col transition-all duration-300 ease-in-out ${isDark ? 'bg-[#181818] border-[#333]' : 'bg-white border-gray-200'} ${isSidebarOpen ? 'w-64 py-4 pl-4 pr-2' : 'w-16 py-4 px-2 items-center'}`}>
            <div className={`flex items-center ${isSidebarOpen ? 'justify-between mb-4 pr-2' : 'justify-center mb-6'} border-b pb-2 ${isDark ? 'border-[#333]' : 'border-gray-200'}`}>
              {isSidebarOpen && <div className={`text-[10px] font-bold tracking-[0.3em] uppercase ${isDark ? 'text-white' : 'text-gray-800'}`}>Toolbox</div>}
              <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className={`p-1 rounded transition-colors flex items-center justify-center w-6 h-6 ${isDark ? 'text-[#888] hover:text-white hover:bg-[#333]' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-100'}`}>
                {isSidebarOpen ? Icons.ArrowLeft : Icons.ArrowRight}
              </button>
            </div>
            <div className={`flex flex-col ${isSidebarOpen ? 'gap-3 pr-2' : 'gap-4 pr-1 w-full items-center'} overflow-y-auto overflow-x-hidden pb-10 custom-scrollbar`}>
              {tList.map(item => {
                const hoverBorderClass = isDark ? (item.t === 'calculateNode' ? 'hover:border-teal-500' : item.t === 'vlookupNode' ? 'hover:border-pink-500' : item.t === 'minusNode' ? 'hover:border-rose-500' : item.t === 'webSourceNode' ? 'hover:border-emerald-500' : item.t === 'folderSourceNode' ? 'hover:border-indigo-500' : item.t === 'pasteNode' ? 'hover:border-orange-500' : 'hover:border-blue-500') 
                                              : (item.t === 'calculateNode' ? 'hover:border-teal-400' : item.t === 'vlookupNode' ? 'hover:border-pink-400' : item.t === 'minusNode' ? 'hover:border-rose-400' : item.t === 'webSourceNode' ? 'hover:border-emerald-400' : item.t === 'folderSourceNode' ? 'hover:border-indigo-400' : item.t === 'pasteNode' ? 'hover:border-orange-400' : 'hover:border-blue-400');
                return (
                <div key={item.t} className={`relative group/btn rounded-xl cursor-grab flex items-center transition-all shadow-sm active:scale-95 border ${isDark ? 'bg-[#252526] border-[#333]' : 'bg-gray-50 border-gray-200'} ${hoverBorderClass} ${btnClasses}`} onDragStart={(e) => e.dataTransfer.setData('application/reactflow', item.t)} draggable>
                  <div className={`${item.c} text-lg group-hover/btn:scale-125 transition-transform flex items-center justify-center ${isSidebarOpen ? '' : 'text-xl'}`}>{item.i}</div>
                  {isSidebarOpen && <span className={`text-[10px] font-bold uppercase tracking-wider truncate ${isDark ? 'text-[#888] group-hover/btn:text-white' : 'text-gray-600 group-hover/btn:text-gray-900'}`}>{item.l}</span>}
                  <div className={`absolute left-full ml-4 top-1/2 -translate-y-1/2 w-56 text-[11px] p-3 rounded-lg border opacity-0 group-hover/btn:opacity-100 pointer-events-none transition-opacity z-50 shadow-xl hidden md:block normal-case leading-relaxed ${isDark ? 'bg-[#111] text-[#ccc] border-[#444]' : 'bg-white text-gray-700 border-gray-200'}`}>
                    <div className={`font-bold mb-1 tracking-widest ${isDark ? 'text-white' : 'text-gray-900'}`}>{item.l}</div>
                    {item.desc}
                  </div>
                </div>
              )})}
            </div>
          </aside>
          <div className="flex-1 relative transition-colors">
            <ReactFlow 
              nodes={nodes} 
              edges={edges} 
              onNodesChange={onNodesChange} 
              onEdgesChange={onEdgesChange} 
              onConnect={(p) => _setEdges((eds: any) => addEdge({ ...p, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } } as any, eds))} 
              onEdgeContextMenu={onEdgeContextMenu} 
              nodeTypes={nodeTypesObj} 
              onNodeDragStop={(_, node: any) => {
                focusNode(node.id);
              }}
              onDrop={(e) => { 
                const t = e.dataTransfer.getData('application/reactflow'); 
                if (t) {
                  const position = screenToFlowPosition({ x: e.clientX, y: e.clientY });
                  const id = `n-${Date.now()}`;
                  _setNodes((nds: any) => nds.concat({ id, type: t, position, data: { useFirstRowAsHeader: true } })); 
                  focusNode(id);
                }
              }} 
              onDragOver={(e) => e.preventDefault()} 
              defaultViewport={{ x: 250, y: 100, zoom: 0.9 }}
            >
              <Background color={isDark ? "#333" : "#d1d5db"} gap={24} size={1} />
              <Controls className={`border fill-gray-600 ${isDark ? 'bg-[#252526] border-[#444]' : 'bg-white border-gray-200'}`} />
              <NodeNavigator tList={tList} nodes={nodes} />
            </ReactFlow>
          </div>
        </div>
        
        <div 
          style={{ height: isPreviewOpen ? bottomHeight : 48, transition: isDragging ? 'none' : 'height 0.3s cubic-bezier(0.4, 0, 0.2, 1)' }} 
          className={`flex flex-col border-t z-30 relative transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#333] shadow-[0_-10px_40px_rgba(0,0,0,0.5)]' : 'bg-gray-50 border-gray-200 shadow-[0_-10px_40px_rgba(0,0,0,0.1)]'}`}
        >
          {isPreviewOpen && (
            <div 
              onMouseDown={startResize} 
              className={`absolute top-[-3px] left-0 w-full h-2 cursor-row-resize z-50 transition-colors no-print ${isDark ? 'hover:bg-blue-500/50' : 'hover:bg-blue-400/50'}`} 
              title="Drag to resize"
            />
          )}
          <div className={`px-6 py-2 flex justify-between items-center border-b h-[48px] shrink-0 no-print transition-colors ${isDark ? 'bg-[#252526] border-[#333]' : 'bg-white border-gray-200'}`}>
            <div className="flex items-center gap-4">
              <button onClick={() => setIsPreviewOpen(!isPreviewOpen)} className={`p-1 rounded transition-colors mr-2 flex items-center justify-center w-6 h-6 ${isDark ? 'text-[#888] hover:text-white hover:bg-[#333]' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-100'}`}>
                {isPreviewOpen ? Icons.ChevronDown : Icons.ChevronUp}
              </button>
              {isPreviewOpen && ['table', 'chart', 'dashboard'].map(tab => (
                <button key={tab} onClick={() => setPreviewTab(tab as any)} className={`text-[11px] font-bold uppercase tracking-[0.3em] pb-1 border-b-2 transition-all mt-1 ${previewTab === tab ? (isDark ? 'text-blue-400 border-blue-400' : 'text-blue-600 border-blue-600') : (isDark ? 'text-[#555] border-transparent hover:text-white' : 'text-gray-400 border-transparent hover:text-gray-800')}`}>
                  <span className="flex items-center gap-2">
                    {tab === 'table' && <><span className="flex items-center justify-center">{Icons.Select}</span> Data Table</>}
                    {tab === 'chart' && <><span className="flex items-center justify-center">{Icons.Chart}</span> Visual Insight</>}
                    {tab === 'dashboard' && <><span className="flex items-center justify-center">{Icons.Dashboard}</span> Dashboard</>}
                  </span>
                </button>
              ))}
              {!isPreviewOpen && <span className={`text-[10px] font-bold uppercase tracking-[0.2em] ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Preview Minimized</span>}
            </div>
            
            <div className="flex items-center gap-3">
              {isPreviewOpen && previewTab !== 'dashboard' && (
                <div className={`flex items-center gap-2 mr-2 border-r pr-4 ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
                  <span className={`text-[10px] font-bold uppercase tracking-widest ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Preview:</span>
                  <select 
                    className={`text-[10px] p-1.5 border rounded outline-none transition-colors cursor-pointer w-40 truncate ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-200 text-gray-700 hover:border-blue-500'}`}
                    value={previewNodeId || ''}
                    onChange={(e) => setPreviewNodeId(e.target.value || null)}
                  >
                    <option value="">Auto (末端のノード)</option>
                    {nodes.map(n => {
                      const typeInfo = tList.find(t => t.t === n.type);
                      const label = typeInfo ? typeInfo.l : n.type;
                      let sub = n.id;
                      if ((n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.fileName) sub = n.data.fileName;
                      else if (n.type === 'webSourceNode' && n.data.fetchedUrl) sub = "Loaded";
                      else if (n.type === 'pasteNode' && n.data.rawData) sub = "Loaded";
                      return <option key={n.id} value={n.id}>{label} - {sub}</option>
                    })}
                  </select>
                </div>
              )}

              {isPreviewOpen && previewTab !== 'dashboard' && (
                <>
                  <span className={`text-[10px] uppercase font-bold tracking-widest mr-2 ${isDark ? 'text-white' : 'text-gray-800'}`}>Export As:</span>
                  <button onClick={() => handleExport('csv')} className={`text-white text-[10px] px-3 py-1.5 rounded-lg font-bold uppercase tracking-widest transition-colors shadow-sm ${isDark ? 'bg-[#333] hover:bg-blue-600' : 'bg-gray-800 hover:bg-blue-600'}`}>CSV</button>
                  <button onClick={() => handleExport('xlsx')} className={`text-white text-[10px] px-3 py-1.5 rounded-lg font-bold uppercase tracking-widest transition-colors shadow-sm ${isDark ? 'bg-[#333] hover:bg-green-600' : 'bg-gray-800 hover:bg-green-600'}`}>Excel</button>
                </>
              )}
              {previewTab !== 'dashboard' && <span className={`text-[10px] font-bold ml-4 ${isDark ? 'text-blue-400' : 'text-blue-600'}`}>{final.data.length} rows</span>}
            </div>
          </div>
          {isPreviewOpen && (
            <div className={`flex-1 overflow-auto print-preview-area transition-colors ${isDark ? 'bg-[#1e1e1e]' : 'bg-gray-50'}`}>
              {previewTab === 'table' && (
                <table className="w-full text-left text-[11px] whitespace-nowrap border-collapse">
                  <thead className={`sticky top-0 border-b z-10 shadow-sm transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-100 border-gray-200'}`}>
                    <tr>{final.headers.map((h, i) => <th key={i} className={`px-5 py-3 font-bold border-r uppercase tracking-wider ${isDark ? 'text-[#888] border-[#333]' : 'text-gray-600 border-gray-200'}`}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {final.data.slice(0, 100).map((row, i) => (
                      <tr key={i} className={`transition-colors border-b ${isDark ? 'hover:bg-[#252526] border-[#222]' : 'hover:bg-white border-gray-200'}`}>
                        {final.headers.map((h, j) => <td key={j} className={`px-5 py-2 border-r font-mono ${isDark ? 'text-[#ccc] border-[#222]' : 'text-gray-800 border-gray-200'}`}>{row[h]}</td>)}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              {previewTab === 'chart' && (
                <div style={{ width: '100%', height: '100%', minHeight: '300px', padding: '24px' }}>
                  {final.chartConfig?.xAxis && final.chartConfig?.yAxis ? (
                    <ResponsiveContainer width="100%" height="100%" minWidth={10} minHeight={10}>
                      {final.chartConfig.chartType === 'line' ? (
                        <LineChart data={final.data}>
                          <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#333" : "#e5e7eb"} vertical={false} />
                          <XAxis dataKey={final.chartConfig.xAxis} stroke={isDark ? "#555" : "#9ca3af"} tick={{ fill: isDark ? '#888' : '#6b7280', fontSize: 11 }} tickLine={false} axisLine={false} dy={10} />
                          <YAxis stroke={isDark ? "#555" : "#9ca3af"} tick={{ fill: isDark ? '#888' : '#6b7280', fontSize: 11 }} tickLine={false} axisLine={false} dx={-10} />
                          <Tooltip contentStyle={{ backgroundColor: isDark ? '#1a1a1a' : '#fff', border: isDark ? '1px solid #444' : '1px solid #e5e7eb', fontSize: '11px', borderRadius: '8px' }} labelStyle={{ color: isDark ? '#fff' : '#1f2937', fontWeight: 'bold', paddingBottom: '4px' }} itemStyle={{ color: '#60a5fa' }} />
                          <Line type="monotone" dataKey={final.chartConfig.yAxis} stroke="#3b82f6" strokeWidth={3} dot={{ r: 4, fill: '#3b82f6', strokeWidth: 0 }} activeDot={{ r: 6 }} />
                        </LineChart>
                      ) : (
                        <BarChart data={final.data}>
                          <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#333" : "#e5e7eb"} vertical={false} />
                          <XAxis dataKey={final.chartConfig.xAxis} stroke={isDark ? "#555" : "#9ca3af"} tick={{ fill: isDark ? '#888' : '#6b7280', fontSize: 11 }} tickLine={false} axisLine={false} dy={10} />
                          <YAxis stroke={isDark ? "#555" : "#9ca3af"} tick={{ fill: isDark ? '#888' : '#6b7280', fontSize: 11 }} tickLine={false} axisLine={false} dx={-10} />
                          <Tooltip cursor={{fill: isDark ? '#252526' : '#f3f4f6'}} contentStyle={{ backgroundColor: isDark ? '#1a1a1a' : '#fff', border: isDark ? '1px solid #444' : '1px solid #e5e7eb', fontSize: '11px', borderRadius: '8px' }} labelStyle={{ color: isDark ? '#fff' : '#1f2937', fontWeight: 'bold', paddingBottom: '4px' }} itemStyle={{ color: '#3b82f6' }} />
                          <Bar dataKey={final.chartConfig.yAxis} fill="#3b82f6" radius={[4, 4, 0, 0]} />
                        </BarChart>
                      )}
                    </ResponsiveContainer>
                  ) : <div className={`h-full flex items-center justify-center text-[11px] italic tracking-widest uppercase animate-pulse ${isDark ? 'text-[#555]' : 'text-gray-400'}`}>Visualizerノードを繋ぎ、軸を設定してください</div>}
                </div>
              )}
              {previewTab === 'dashboard' && (
                <div className={`grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 p-6 h-full overflow-y-auto custom-scrollbar ${isDark ? 'bg-[#111]' : 'bg-gray-100'}`}>
                  {dashboardsData.length === 0 ? (
                     <div className={`col-span-full h-full flex items-center justify-center text-[11px] italic tracking-widest uppercase animate-pulse ${isDark ? 'text-[#555]' : 'text-gray-400'}`}>
                        Canvas上にVisualizerノードを配置してください
                     </div>
                  ) : (
                    dashboardsData.map(d => (
                      <div key={d.id} className={`p-4 rounded-2xl border flex flex-col h-[320px] transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#333] shadow-lg hover:border-[#444]' : 'bg-white border-gray-200 shadow-md hover:border-gray-300'}`}>
                        <div className={`text-[10px] font-bold uppercase tracking-wider mb-4 border-b pb-2 flex justify-between items-center ${isDark ? 'text-[#888] border-[#333]' : 'text-gray-500 border-gray-100'}`}>
                          <span className="flex items-center gap-1.5"><span className="flex items-center justify-center w-4 h-4">{Icons.Chart}</span> {d.config.chartType === 'line' ? 'Line Chart' : 'Bar Chart'}</span>
                          <span className={isDark ? 'text-[#555]' : 'text-gray-400'}>{d.config.yAxis} by {d.config.xAxis}</span>
                        </div>
                        <div className="flex-1 min-h-0">
                          {d.config.xAxis && d.config.yAxis && d.data.length > 0 ? (
                            <ResponsiveContainer width="100%" height="100%" minWidth={10} minHeight={10}>
                              {d.config.chartType === 'line' ? (
                                <LineChart data={d.data}>
                                  <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#222" : "#f3f4f6"} vertical={false} />
                                  <XAxis dataKey={d.config.xAxis} stroke={isDark ? "#555" : "#d1d5db"} tick={{ fill: isDark ? '#888' : '#9ca3af', fontSize: 9 }} tickLine={false} axisLine={false} dy={5} />
                                  <YAxis stroke={isDark ? "#555" : "#d1d5db"} tick={{ fill: isDark ? '#888' : '#9ca3af', fontSize: 9 }} tickLine={false} axisLine={false} dx={-5} />
                                  <Tooltip contentStyle={{ backgroundColor: isDark ? '#1a1a1a' : '#fff', border: isDark ? '1px solid #444' : '1px solid #e5e7eb', fontSize: '11px', borderRadius: '8px' }} />
                                  <Line type="monotone" dataKey={d.config.yAxis} stroke="#3b82f6" strokeWidth={2} dot={{ r: 2, fill: '#3b82f6', strokeWidth: 0 }} activeDot={{ r: 4 }} />
                                </LineChart>
                              ) : (
                                <BarChart data={d.data}>
                                  <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#222" : "#f3f4f6"} vertical={false} />
                                  <XAxis dataKey={d.config.xAxis} stroke={isDark ? "#555" : "#d1d5db"} tick={{ fill: isDark ? '#888' : '#9ca3af', fontSize: 9 }} tickLine={false} axisLine={false} dy={5} />
                                  <YAxis stroke={isDark ? "#555" : "#d1d5db"} tick={{ fill: isDark ? '#888' : '#9ca3af', fontSize: 9 }} tickLine={false} axisLine={false} dx={-5} />
                                  <Tooltip cursor={{fill: isDark ? '#222' : '#f3f4f6'}} contentStyle={{ backgroundColor: isDark ? '#1a1a1a' : '#fff', border: isDark ? '1px solid #444' : '1px solid #e5e7eb', fontSize: '11px', borderRadius: '8px' }} />
                                  <Bar dataKey={d.config.yAxis} fill="#3b82f6" radius={[2, 2, 0, 0]} />
                                </BarChart>
                              )}
                            </ResponsiveContainer>
                          ) : <div className={`h-full flex items-center justify-center text-[10px] italic uppercase ${isDark ? 'text-[#444]' : 'text-gray-400'}`}>No Data</div>}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              )}
            </div>
          )}
        </div>

        {/* ★ チュートリアルモーダル */}
        {showTutorial && (
          <div className={`fixed inset-0 z-[400] flex items-center justify-center p-8 backdrop-blur-md no-print ${isDark ? 'bg-black/90' : 'bg-gray-900/50'}`}>
            <div className={`border rounded-2xl shadow-2xl w-[600px] overflow-hidden flex flex-col transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
              <div className={`p-4 border-b flex justify-between items-center transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444]' : 'bg-gray-50 border-gray-200'}`}>
                <h2 className={`text-[12px] font-bold uppercase tracking-[0.4em] flex items-center gap-2 ${isDark ? 'text-white' : 'text-gray-800'}`}>
                  <span className="text-blue-500 w-4 h-4 flex items-center justify-center">{Icons.Help}</span> 
                  Visual Data Prep
                </h2>
                <button onClick={closeTutorial} className={`transition-colors text-xl flex items-center justify-center w-6 h-6 ${isDark ? 'text-[#666] hover:text-white' : 'text-gray-500 hover:text-gray-800'}`}>
                  <span className="w-4 h-4 block flex items-center justify-center">{Icons.Close}</span>
                </button>
              </div>

              <div className={`p-8 flex flex-col items-center space-y-6 transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
                <div className={`text-[11px] leading-relaxed text-center ${isDark ? 'text-[#aaa]' : 'text-gray-600'}`}>
                  ”ノード”を繋いで視覚的にデータを作る、データ整形ツールです。<br/><br/>基本的な使い方は以下の3ステップです。
                </div>
                
                <div className="w-full space-y-4 text-left">
                  <div className={`p-4 rounded-xl border flex items-center gap-4 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] hover:border-blue-500/50' : 'bg-white border-gray-200 hover:border-blue-300'}`}>
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg shrink-0 border transition-colors ${isDark ? 'bg-[#252526] text-blue-400 border-[#444]' : 'bg-gray-50 text-blue-600 border-gray-200 shadow-sm'}`}>
                      <span className="w-5 h-5 flex items-center justify-center">{Icons.Source}</span>
                    </div>
                    <div className="flex-1">
                      <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-white' : 'text-gray-800'}`}>Step 1: Add Nodes</div>
                      <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>左の<strong>Toolbox</strong>から、SOURCE(ファイル読み込み)やUNION(結合)などの「ノード」を画面へドラッグ＆ドロップします。</p>
                    </div>
                  </div>

                  <div className={`p-4 rounded-xl border flex items-center gap-4 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] hover:border-pink-500/50' : 'bg-white border-gray-200 hover:border-pink-300'}`}>
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg shrink-0 border transition-colors ${isDark ? 'bg-[#252526] text-pink-400 border-[#444]' : 'bg-gray-50 text-pink-600 border-gray-200 shadow-sm'}`}>
                      <span className="w-5 h-5 flex items-center justify-center">{Icons.Join}</span>
                    </div>
                    <div className="flex-1">
                      <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-white' : 'text-gray-800'}`}>Step 2: Connect Flow</div>
                      <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>ノード同士の<strong>青い○（ハンドル）</strong>をマウスで繋ぐと、データが左から右へと流れて処理されます。</p>
                    </div>
                  </div>

                  <div className={`p-4 rounded-xl border flex items-center gap-4 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] hover:border-emerald-500/50' : 'bg-white border-gray-200 hover:border-emerald-300'}`}>
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg shrink-0 border transition-colors ${isDark ? 'bg-[#252526] text-emerald-400 border-[#444]' : 'bg-gray-50 text-emerald-600 border-gray-200 shadow-sm'}`}>
                      <span className="w-5 h-5 flex items-center justify-center">{Icons.Dashboard}</span>
                    </div>
                    <div className="flex-1">
                      <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-white' : 'text-gray-800'}`}>Step 3: Preview & Export</div>
                      <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>画面下部の<strong>Preview</strong>に処理結果がリアルタイムで表示されます。Excel出力やSQL変換も可能です。</p>
                    </div>
                  </div>
                </div>
              </div>
              <button onClick={closeTutorial} className={`w-full p-4 text-[11px] font-bold uppercase tracking-widest flex items-center justify-center gap-2 transition-colors border-t ${isDark ? 'bg-[#252526] text-blue-400 border-[#444] hover:bg-[#333]' : 'bg-gray-50 text-blue-600 border-gray-200 hover:bg-gray-100'}`}>
                <span className="w-4 h-4 flex items-center justify-center"></span> 使ってみる
              </button>
            </div>
          </div>
        )}

        <SqlModal 
          isOpen={isSqlModalOpen} 
          onClose={() => setIsSqlModalOpen(false)} 
          nodes={nodes} 
          edges={edges} 
          onImport={(n: CustomNode[], e: Edge[]) => { _setNodes(n); _setEdges(e); setWorkbooks({}); }} 
        />

        <SaveLoadModal isOpen={isSaveLoadOpen} onClose={() => setIsSaveLoadOpen(false)} onSave={handleSave} onLoad={handleLoad} onDelete={handleDeleteFlow} flows={savedFlows} />

        {isResetModalOpen && (
          <div className={`fixed inset-0 z-[300] flex items-center justify-center p-8 backdrop-blur-sm no-print ${isDark ? 'bg-black/80' : 'bg-gray-900/50'}`}>
            <div className={`border rounded-2xl shadow-2xl w-[320px] p-6 text-center space-y-6 ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
              <div className={`text-4xl w-10 h-10 mx-auto flex items-center justify-center ${isDark ? 'text-white' : 'text-gray-800'}`}>{Icons.Warning}</div>
              <div className="space-y-2">
                <h3 className={`text-sm font-bold tracking-widest ${isDark ? 'text-white' : 'text-gray-900'}`}>RESET ALL DATA?</h3>
                <p className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>現在のノード構成と読み込んだデータがすべて消去されます。この操作は元に戻せません。</p>
              </div>
              <div className="flex gap-3">
                <button onClick={() => setIsResetModalOpen(false)} className={`flex-1 py-3 rounded-xl text-[10px] font-bold uppercase tracking-widest transition-colors ${isDark ? 'bg-[#333] hover:bg-[#444] text-white' : 'bg-gray-200 hover:bg-gray-300 text-gray-800'}`}>Cancel</button>
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
  const { theme } = useContext(AppContext);
  const isDark = theme === 'dark';
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
      if (n.type === 'dataNode' || n.type === 'folderSourceNode') {
        from = n.data.fileName ? `\`${n.data.fileName}\`` : "source_data";
        if (n.data.ranges && n.data.ranges.length > 0) {
          from += ` /* Range: ${n.data.ranges.join(', ')} */`;
        }
      } else if (n.type === 'webSourceNode') {
        from = n.data.fetchedUrl ? `\`${n.data.fetchedUrl}\`` : "web_data";
      } else if (n.type === 'pasteNode') {
        from = "pasted_data";
      }
      
      if (n.type === 'joinNode' && n.data.keyA && n.data.keyB) {
        const jType = n.data.joinType === 'left' ? 'LEFT' : n.data.joinType === 'right' ? 'RIGHT' : 'INNER';
        joinStr += `\n${jType} JOIN sub_data ON \`${n.data.keyA}\` = sub_data.\`${n.data.keyB}\``;
      }

      if (n.type === 'vlookupNode' && n.data.keyA && n.data.keyB && n.data.fetchCol && n.data.targetCol) {
        joinStr += `\nLEFT JOIN sub_data ON \`${n.data.keyA}\` = sub_data.\`${n.data.keyB}\``;
        select = select === "*" ? `*, sub_data.\`${n.data.fetchCol}\` AS \`${n.data.targetCol}\`` : `${select}, sub_data.\`${n.data.fetchCol}\` AS \`${n.data.targetCol}\``;
      }
      
      if (n.type === 'minusNode' && n.data.keyA && n.data.keyB) {
        joinStr += `\nLEFT JOIN sub_data ON \`${n.data.keyA}\` = sub_data.\`${n.data.keyB}\``;
        wheres.push(`sub_data.\`${n.data.keyB}\` IS NULL`);
      }

      if (n.type === 'selectNode' && n.data.selectedColumns?.length) select = n.data.selectedColumns.map((c: string) => `\`${c}\``).join(", ");
      
      if (n.type === 'transformNode' && n.data.command) {
        if (n.data.command === 'case_when' && n.data.targetCol) {
          const op = n.data.condOp === 'gt' ? '>' : n.data.condOp === 'lt' ? '<' : n.data.condOp === 'not' ? '!=' : '=';
          const cWhen = `CASE WHEN \`${n.data.targetCol}\` ${op} '${n.data.condVal || ''}' THEN '${n.data.trueVal || ''}' ELSE '${n.data.falseVal || ''}' END AS \`${n.data.targetCol}\``;
          select = select === "*" ? `*, ${cWhen}` : `${select}, ${cWhen}`;
        } else if (n.data.command === 'remove_duplicates') {
          select = `DISTINCT ${select}`;
        } else if (n.data.targetCol) {
          let func = '';
          if (n.data.command === 'to_string') func = `CAST(\`${n.data.targetCol}\` AS CHAR)`;
          if (n.data.command === 'to_number') func = `CAST(\`${n.data.targetCol}\` AS DECIMAL)`;
          if (n.data.command === 'fill_zero') func = `COALESCE(\`${n.data.targetCol}\`, 0)`;
          if (n.data.command === 'zero_padding') func = `LPAD(\`${n.data.targetCol}\`, ${n.data.param0 || 1}, '0')`;
          if (n.data.command === 'replace') func = `REPLACE(\`${n.data.targetCol}\`, '${n.data.param0 || ""}', '')`;
          if (n.data.command === 'math_mul') func = `\`${n.data.targetCol}\` * ${n.data.param0 || 1}`;
          if (n.data.command === 'add_suffix') func = `CONCAT(\`${n.data.targetCol}\`, '${n.data.param0 || ""}')`;
          if (n.data.command === 'concat') func = `CONCAT(\`${n.data.targetCol}\`, '${n.data.param0 || ""}')`;
          if (n.data.command === 'round') func = `ROUND(\`${n.data.targetCol}\`, ${n.data.param0 || 0})`;
          if (n.data.command === 'mod') func = `MOD(\`${n.data.targetCol}\`, ${n.data.param0 || 1})`;
          if (n.data.command === 'substring') {
            const params = String(n.data.param0 || '1').split(',').map(s => s.trim());
            const start = params[0] || '1';
            const len = params.length > 1 ? params[1] : '255';
            func = `SUBSTRING(\`${n.data.targetCol}\`, ${start}, ${len})`;
          }

          if (func) {
             if (n.data.applyCond && n.data.condCol && n.data.condOp) {
                 const op = n.data.condOp === 'gt' ? '>' : n.data.condOp === 'lt' ? '<' : n.data.condOp === 'not' ? '!=' : '=';
                 const val = n.data.condOp === 'includes' ? `'%${n.data.condVal || ''}%'` : `'${n.data.condVal || ''}'`;
                 const actualOp = n.data.condOp === 'includes' ? 'LIKE' : op;
                 func = `CASE WHEN \`${n.data.condCol}\` ${actualOp} ${val} THEN ${func} ELSE \`${n.data.targetCol}\` END`;
             }
             select = select === "*" ? `*, ${func} AS \`${n.data.targetCol}\`` : `${select}, ${func} AS \`${n.data.targetCol}\``;
          }
        }
      }
      
      if (n.type === 'calculateNode' && n.data.colA && n.data.colB && n.data.newColName) {
        let func = '';
        if (n.data.operator === 'add') func = `(COALESCE(\`${n.data.colA}\`, 0) + COALESCE(\`${n.data.colB}\`, 0))`;
        else if (n.data.operator === 'sub') func = `(COALESCE(\`${n.data.colA}\`, 0) - COALESCE(\`${n.data.colB}\`, 0))`;
        else if (n.data.operator === 'mul') func = `(COALESCE(\`${n.data.colA}\`, 0) * COALESCE(\`${n.data.colB}\`, 0))`;
        else if (n.data.operator === 'div') func = `(COALESCE(\`${n.data.colA}\`, 0) / NULLIF(COALESCE(\`${n.data.colB}\`, 0), 0))`;
        else if (n.data.operator === 'concat') func = `CONCAT(COALESCE(\`${n.data.colA}\`, ''), COALESCE(\`${n.data.colB}\`, ''))`;
        
        if (func) {
          select = select === "*" ? `*, ${func} AS \`${n.data.newColName}\`` : `${select}, ${func} AS \`${n.data.newColName}\``;
        }
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
    <div className={`fixed inset-0 z-[200] flex items-center justify-center p-8 backdrop-blur-md no-print ${isDark ? 'bg-black/90' : 'bg-gray-900/50'}`}>
      <div className={`border rounded-2xl shadow-2xl w-[600px] overflow-hidden flex flex-col transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
        <div className={`flex border-b ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
          <button onClick={() => setTab('export')} className={`flex-1 p-4 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'export' ? (isDark ? 'bg-[#252526] text-blue-400 border-b-2 border-blue-400' : 'bg-gray-50 text-blue-600 border-b-2 border-blue-600') : (isDark ? 'text-[#666] hover:bg-[#252526]' : 'text-gray-500 hover:bg-gray-50')}`}><span>{Icons.Code}</span> Flow to SQL</button>
          <button onClick={() => setTab('import')} className={`flex-1 p-4 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'import' ? (isDark ? 'bg-[#252526] text-green-400 border-b-2 border-green-400' : 'bg-gray-50 text-green-600 border-b-2 border-green-600') : (isDark ? 'text-[#666] hover:bg-[#252526]' : 'text-gray-500 hover:bg-gray-50')}`}><span>{Icons.Diamond}</span> SQL to Flow</button>
        </div>
        <div className={`p-6 h-[350px] flex flex-col transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
          {tab === 'export' ? (
            <div className="flex-1 flex flex-col space-y-4">
              <p className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>現在のノード構成（単一パス）を解析して、対応するSQL文を自動生成します。</p>
              <div className="flex-1 relative">
                <textarea readOnly value={generatedSql} className={`w-full h-full font-mono text-[12px] p-4 border rounded-xl outline-none resize-none custom-scrollbar transition-colors ${isDark ? 'bg-[#1e1e1e] text-blue-300 border-[#444]' : 'bg-white text-blue-600 border-gray-300'}`} />
                <button onClick={handleCopy} className={`absolute top-3 right-3 text-white text-[10px] px-3 py-1.5 rounded flex items-center gap-1 transition-colors shadow-sm ${isDark ? 'bg-[#333] hover:bg-blue-600' : 'bg-gray-800 hover:bg-blue-600'}`}>
                  {copyMsg || <><span className="w-3 h-3 flex">{Icons.Copy}</span> Copy</>}
                </button>
              </div>
            </div>
          ) : (
            <div className="flex-1 flex flex-col space-y-4">
              <p className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>SELECT, FROM, WHERE, GROUP BY, ORDER BY 構文を解析してノードを配置します。</p>
              <textarea value={sqlInput} onChange={(e) => setSqlInput(e.target.value)} placeholder="SELECT col1, SUM(col2)&#10;FROM data /* Range: A1:D10 */&#10;WHERE col1 = 'value'&#10;GROUP BY col1&#10;ORDER BY col1 DESC" className={`flex-1 w-full font-mono text-[12px] p-4 border rounded-xl outline-none focus:border-green-500 resize-none transition-colors custom-scrollbar ${isDark ? 'bg-[#1e1e1e] text-green-300 border-[#444]' : 'bg-white text-green-600 border-gray-300'}`} />
              <button disabled={!sqlInput} onClick={handleImport} className="w-full bg-green-600 hover:bg-green-500 disabled:opacity-30 text-white py-3 rounded-xl text-[11px] font-bold uppercase shadow-xl active:scale-95 transition-all">Build Flow from SQL</button>
            </div>
          )}
        </div>
        <button onClick={onClose} className={`w-full p-4 text-[10px] font-bold uppercase border-t transition-colors ${isDark ? 'bg-[#252526] text-white hover:bg-[#333] border-[#444]' : 'bg-white text-gray-800 hover:bg-gray-100 border-gray-200'}`}>Close</button>
      </div>
    </div>
  );
};

const SaveLoadModal = ({ isOpen, onClose, onSave, onLoad, onDelete, flows }: any) => {
  const { theme } = useContext(AppContext);
  const isDark = theme === 'dark';
  const [tab, setTab] = useState<'load' | 'save'>('load'); const [sName, setSName] = useState('');
  const sortedFlows = useMemo(() => [...flows].sort((a, b) => Number(b.id) - Number(a.id)), [flows]);

  if (!isOpen) return null;
  return (
    <div className={`fixed inset-0 z-[200] flex items-center justify-center p-8 backdrop-blur-md no-print ${isDark ? 'bg-black/90' : 'bg-gray-900/50'}`}>
      <div className={`border rounded-2xl shadow-2xl w-[500px] overflow-hidden transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
        <div className={`flex border-b ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
          <button onClick={() => setTab('load')} className={`flex-1 p-3 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'load' ? (isDark ? 'bg-[#252526] text-blue-400 border-b-2 border-blue-400' : 'bg-gray-50 text-blue-600 border-b-2 border-blue-600') : (isDark ? 'text-[#666] hover:bg-[#252526]' : 'text-gray-500 hover:bg-gray-50')}`}><span>{Icons.Folder}</span> Load</button>
          <button onClick={() => setTab('save')} className={`flex-1 p-3 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'save' ? (isDark ? 'bg-[#252526] text-white border-b-2 border-white' : 'bg-gray-50 text-gray-900 border-b-2 border-gray-900') : (isDark ? 'text-[#666] hover:bg-[#252526]' : 'text-gray-500 hover:bg-gray-50')}`}><span>{Icons.Save}</span> Save</button>
        </div>
        <div className={`p-6 h-[300px] overflow-y-auto custom-scrollbar transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
          {tab === 'load' ? (
            <div className="space-y-3">
              {sortedFlows.length === 0 && <div className={`text-[10px] text-center mt-10 ${isDark ? 'text-[#555]' : 'text-gray-400'}`}>No saved projects</div>}
              {sortedFlows.map((f: any) => (
                <div key={f.id} className={`p-4 rounded-xl flex justify-between items-center cursor-pointer group transition-colors border ${isDark ? 'bg-[#252526] border-[#444] hover:border-blue-500/50' : 'bg-white border-gray-200 hover:border-blue-300'}`} onClick={() => { onLoad(f); onClose(); }}>
                  <div>
                    <div className={`text-[12px] font-bold transition-colors ${isDark ? 'text-white group-hover:text-blue-400' : 'text-gray-800 group-hover:text-blue-600'}`}>{f.name}</div>
                    <div className={`text-[9px] mt-1 ${isDark ? 'text-[#666]' : 'text-gray-500'}`}>{f.updatedAt}</div>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={(e) => { e.stopPropagation(); onDelete(f.id); }} className={`border px-3 py-2 rounded-lg text-[10px] font-bold shadow-sm transition-all flex items-center justify-center w-8 ${isDark ? 'bg-[#333] text-[#aaa] border-[#555] hover:bg-blue-600 hover:text-white hover:border-blue-500' : 'bg-gray-100 text-gray-500 border-gray-300 hover:bg-blue-600 hover:text-white'}`} title="Delete Project">
                      <span>{Icons.Close}</span>
                    </button>
                    <button className={`border px-5 py-2 rounded-lg text-[10px] font-bold shadow-sm transition-all ${isDark ? 'bg-blue-600/20 text-blue-400 border-blue-500/30 group-hover:bg-blue-600 group-hover:text-white' : 'bg-blue-50 text-blue-600 border-blue-200 group-hover:bg-blue-600 group-hover:text-white'}`}>Load</button>
                  </div>
                </div>
              ))}
            </div>
          ) : (
            <div className="space-y-4 pt-8">
              <div className="space-y-2">
                <label className={`text-[10px] font-bold uppercase ${isDark ? 'text-white' : 'text-gray-800'}`}>Project Name</label>
                <input type="text" placeholder="e.g. Sales Report" value={sName} onChange={(e) => setSName(e.target.value)} className={`w-full p-3 rounded-xl outline-none focus:border-blue-500 text-sm transition-colors border ${isDark ? 'bg-[#252526] border-[#444] text-white' : 'bg-white border-gray-300 text-gray-800'}`} />
              </div>
              <button disabled={!sName} onClick={() => { onSave(sName); setSName(''); onClose(); }} className="w-full bg-blue-600 hover:bg-blue-500 disabled:opacity-30 text-white py-3 rounded-xl text-[11px] font-bold uppercase shadow-xl active:scale-95 transition-all">Save Now</button>
            </div>
          )}
        </div>
        <button onClick={onClose} className={`w-full p-3 text-[10px] font-bold uppercase border-t transition-colors ${isDark ? 'bg-[#252526] text-white hover:bg-[#333] border-[#444]' : 'bg-white text-gray-800 hover:bg-gray-100 border-gray-200'}`}>Close</button>
      </div>
    </div>
  );
};

const RangeSelectorModal = ({ isOpen, onClose, workbook, currentSheet, onRangesConfirm, initialRanges, initialUseHeader }: any) => {
  const { theme } = useContext(AppContext);
  const isDark = theme === 'dark';
  const [sRanges, setSRanges] = useState<string[]>(initialRanges || []);
  const [uHead, setUHead] = useState(initialUseHeader !== false);
  const [cWidth, setCWidth] = useState(120);
  const [dStart, setDStart] = useState<{ c: number, r: number } | null>(null);
  const [dEnd, setDEnd] = useState<{ c: number, r: number } | null>(null);
  const sData = useMemo(() => (!workbook || !currentSheet) ? [] : XLSX.utils.sheet_to_json(workbook.Sheets[currentSheet], { header: 1, defval: "", blankrows: true }).slice(0, 100) as any[][], [workbook, currentSheet]);
  const cols = useMemo(() => sData.length > 0 ? Array.from({ length: Math.max(...sData.map(r => r.length)) }, (_, i) => XLSX.utils.encode_col(i)) : [], [sData]);
  if (!isOpen) return null;
  return (
    <div className={`fixed inset-0 z-[100] flex items-center justify-center p-12 backdrop-blur-md no-print ${isDark ? 'bg-black/95' : 'bg-gray-900/50'}`}>
      <div className={`border rounded-3xl shadow-2xl w-full h-full flex flex-col overflow-hidden ring-1 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] ring-white/10' : 'bg-white border-gray-200 ring-black/5'}`}>
        <div className={`p-4 border-b flex flex-col font-sans transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444]' : 'bg-gray-50 border-gray-200'}`}>
          <div className="flex justify-between items-center mb-2">
            <h2 className={`text-[12px] font-bold uppercase tracking-[0.4em] flex items-center gap-2 ${isDark ? 'text-white' : 'text-gray-800'}`}>
              <span className="text-blue-500">{Icons.Select}</span> 抽出範囲の選択
            </h2>
            <div className="flex items-center gap-8">
              <div className="flex items-center gap-3">
                <span className={`text-[9px] font-bold uppercase tracking-widest ${isDark ? 'text-[#555]' : 'text-gray-500'}`}>Width</span>
                <input type="range" min="60" max="400" value={cWidth} onChange={(e) => setCWidth(Number(e.target.value))} className="w-32 accent-blue-500" />
              </div>
              <button onClick={onClose} className={`transition-colors text-xl flex items-center justify-center w-6 h-6 ${isDark ? 'text-[#666] hover:text-white' : 'text-gray-500 hover:text-gray-800'}`}>
                <span className="w-4 h-4 block flex items-center justify-center">{Icons.Close}</span>
              </button>
            </div>
          </div>
          <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-600'}`}>
            🖱️ 抽出したいデータの範囲を<strong>マウスでドラッグして選択</strong>してください。（複数の範囲を選択することも可能です）<br/>
            ✅ 選択後、右側のパネルで<strong>「1行目をヘッダーにする」</strong>オプションをオンにすると、最初の行が列名として認識されます。
          </p>
        </div>
        <div className="flex-1 flex overflow-hidden">
          <div className={`flex-1 overflow-auto custom-scrollbar transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
            <table className="border-collapse table-fixed">
              <thead className={`sticky top-0 z-10 transition-colors ${isDark ? 'bg-[#252526]' : 'bg-gray-100'}`}>
                <tr>
                  <th className={`w-12 border-b ${isDark ? 'border-[#444]' : 'border-gray-300'}`}></th>
                  {cols.map(c => <th key={c} style={{ width: cWidth }} className={`px-3 py-1.5 border-r border-b text-[10px] font-bold ${isDark ? 'border-[#444] text-[#888]' : 'border-gray-300 text-gray-600'}`}>{c}</th>)}
                </tr>
              </thead>
              <tbody>
                {sData.map((row, r) => (
                  <tr key={r}>
                    <td className={`w-12 border-r border-b text-center text-[10px] font-mono transition-colors ${isDark ? 'bg-[#252526] border-[#444] text-[#666]' : 'bg-gray-100 border-gray-300 text-gray-500'}`}>{r+1}</td>
                    {cols.map((_, c) => { 
                      const isSel = dStart && dEnd && c >= Math.min(dStart.c, dEnd.c) && c <= Math.max(dStart.c, dEnd.c) && r >= Math.min(dStart.r, dEnd.r) && r <= Math.max(dStart.r, dEnd.r); 
                      return (
                        <td 
                          key={c} 
                          onMouseDown={() => {setDStart({c, r}); setDEnd({c, r});}} 
                          onMouseOver={() => dStart && setDEnd({c, r})} 
                          onMouseUp={() => { if(dStart && dEnd){ const range = XLSX.utils.encode_range({s:{c:Math.min(dStart.c, dEnd.c),r:Math.min(dStart.r, dEnd.r)},e:{c:Math.max(dStart.c, dEnd.c),r:Math.max(dStart.r, dEnd.r)}}); setSRanges(p => [...p, range]); setDStart(null); setDEnd(null); }}} 
                          style={{ width: cWidth }} 
                          className={`px-2 py-1.5 border-r border-b truncate text-[11px] select-none cursor-crosshair transition-colors ${isSel ? (isDark ? 'bg-blue-500/50 ring-1 ring-blue-400 ring-inset text-white border-[#2a2a2a]' : 'bg-blue-100 ring-1 ring-blue-400 ring-inset text-gray-900 border-gray-200') : (isDark ? 'text-[#aaa] bg-transparent hover:bg-[#222] border-[#2a2a2a]' : 'text-gray-700 bg-white hover:bg-gray-100 border-gray-200')}`}
                        >
                          {row[c]}
                        </td> 
                      )
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className={`w-72 border-l p-6 flex flex-col gap-6 font-sans transition-colors ${isDark ? 'bg-[#252526] border-[#444]' : 'bg-gray-50 border-gray-300'}`}>
            <h3 className={`text-[10px] font-bold uppercase border-b pb-3 ${isDark ? 'text-[#888] border-[#444]' : 'text-gray-600 border-gray-200'}`}>Selection Settings</h3>
            <label className="flex items-center gap-3 cursor-pointer group">
              <input type="checkbox" checked={uHead} onChange={(e) => setUHead(e.target.checked)} className="w-4 h-4 accent-blue-500" />
              <span className={`text-[11px] uppercase font-bold transition-colors ${isDark ? 'text-[#ccc] group-hover:text-white' : 'text-gray-700 group-hover:text-gray-900'}`}>1行目をヘッダーにする</span>
            </label>
            <div className="flex-1 overflow-auto space-y-3 pr-2 custom-scrollbar">
              {sRanges.map((rs, i) => (
                <div key={i} className={`p-3 rounded-xl border flex justify-between items-center group transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] hover:border-blue-500/50' : 'bg-white border-gray-200 hover:border-blue-300'}`}>
                  <div className="flex flex-col">
                    <span className={`text-[8px] font-bold uppercase tracking-widest ${isDark ? 'text-[#555]' : 'text-gray-500'}`}>Range {i+1}</span>
                    <span className={`text-[11px] font-mono font-bold ${isDark ? 'text-blue-400' : 'text-blue-600'}`}>{rs}</span>
                  </div>
                  <button onClick={() => setSRanges(p => p.filter((_, idx) => idx !== i))} className={`transition-colors flex items-center justify-center w-5 h-5 ${isDark ? 'text-[#555] group-hover:text-white' : 'text-gray-400 group-hover:text-gray-800'}`}>
                    <span>{Icons.Close}</span>
                  </button>
                </div>
              ))}
            </div>
            <button onClick={() => {onRangesConfirm(sRanges, uHead); onClose();}} className="bg-blue-600 hover:bg-blue-500 py-3.5 rounded-xl text-[11px] font-bold text-white uppercase shadow-xl active:scale-95 transition-all">選択を適用</button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default function App() { return <ReactFlowProvider><FlowBuilder /></ReactFlowProvider>; }
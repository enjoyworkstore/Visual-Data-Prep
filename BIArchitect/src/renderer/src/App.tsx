import React, { useState, useCallback, useMemo, memo, useContext, useEffect } from 'react';
import { createPortal } from 'react-dom';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { ReactFlow, Controls, Background, addEdge, ReactFlowProvider, useReactFlow, useNodesState, useEdgesState, Panel } from '@xyflow/react';
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import type { Node, Edge } from '@xyflow/react';
import '@xyflow/react/dist/style.css';
import titleImage from '../Public/tytle.png';
import DataCheckNode from './components/nodes/DataCheckNode';
import NodeInput from './components/nodes/NodeInput';
import { AppContext, NodeWrap, useNodeLogic } from './components/nodes/shared';
import { getCheckOperatorLabel, matchesCondition } from './components/nodes/dataCheckUtils';
import type { SourceDataByNodeId, WorkerFlowResult } from './lib/flowCalcShared';

type CustomNode = Node<Record<string, any>>;
type CalcDataResult = { data: any[]; headers: string[] };
type CameraFocusReason = 'move' | 'delete' | 'resize' | 'connect' | 'create' | 'manual';
type CameraFocusConfig = Record<Exclude<CameraFocusReason, 'manual'>, boolean>;
const LAST_COLUMN_OPTION = '__last__';
const MAX_CHART_RENDER_POINTS = 1000;
const PREVIEW_AUTO_PAUSE_MAX_ROWS = 5000;
const PREVIEW_AUTO_PAUSE_MAX_CELLS = 120000;

const DEFAULT_CAMERA_FOCUS_CONFIG: CameraFocusConfig = {
  move: false,
  delete: true,
  resize: true,
  connect: false,
  create: true,
};

const createEmptyMatrix = (rows: number = 6, cols: number = 4): string[][] =>
  Array.from({ length: rows }, () => Array.from({ length: cols }, () => ''));

const sanitizeMatrix = (matrix: any): string[][] => {
  if (!Array.isArray(matrix) || matrix.length === 0) return createEmptyMatrix();
  const rows = matrix.map((row: any) => Array.isArray(row) ? row.map((cell: any) => cell == null ? '' : String(cell)) : []);
  const colCount = Math.max(1, ...rows.map((row: string[]) => row.length));
  return rows.map((row: string[]) => Array.from({ length: colCount }, (_, idx) => row[idx] ?? ''));
};

const parseDelimitedTextToMatrix = (text: string): string[][] => {
  if (!text.trim()) return createEmptyMatrix();
  const parsed = Papa.parse<string[]>(text.trim(), { skipEmptyLines: false });
  const rows = (parsed.data || []).map((row: any) => Array.isArray(row) ? row.map((cell: any) => cell == null ? '' : String(cell)) : []);
  return sanitizeMatrix(rows.filter((row: string[]) => row.length > 0));
};

const matrixToDelimitedText = (matrix: string[][]): string =>
  sanitizeMatrix(matrix).map((row) => row.join('\t')).join('\n');

const stringifyJsonCell = (value: unknown): string | number | boolean | null => {
  if (value == null) return null;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  return JSON.stringify(value);
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);

const isArrayOfArrays = (value: unknown[]): value is unknown[][] => value.every((item) => Array.isArray(item));

const isObjectArray = (value: unknown[]): value is Record<string, unknown>[] =>
  value.every((item) => isRecord(item));

const jsonToWorkbook = (text: string): XLSX.WorkBook => {
  const trimmed = text.trim();
  if (!trimmed) {
    throw new Error('JSONファイルが空です。内容を確認してください。');
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(trimmed) as unknown;
  } catch (err) {
    const detail = err instanceof Error && err.message ? ` (${err.message})` : '';
    throw new Error(`JSONファイルの解析に失敗しました。JSON形式として正しい内容か確認してください。${detail}`);
  }
  const wb = XLSX.utils.book_new();
  let ws: XLSX.WorkSheet;

  if (Array.isArray(parsed)) {
    const parsedArray: unknown[] = parsed;
    if (parsed.length === 0) {
      ws = XLSX.utils.json_to_sheet([]);
    } else if (isArrayOfArrays(parsedArray)) {
      ws = XLSX.utils.aoa_to_sheet(parsedArray.map((row) => row.map((cell) => stringifyJsonCell(cell))));
    } else if (isObjectArray(parsedArray)) {
      ws = XLSX.utils.json_to_sheet(
        parsedArray.map((row) =>
          Object.fromEntries(Object.entries(row).map(([key, value]) => [key, stringifyJsonCell(value)]))
        )
      );
    } else {
      ws = XLSX.utils.json_to_sheet(parsedArray.map((value) => ({ value: stringifyJsonCell(value) })));
    }
  } else if (isRecord(parsed)) {
    ws = XLSX.utils.json_to_sheet([
      Object.fromEntries(Object.entries(parsed).map(([key, value]) => [key, stringifyJsonCell(value)]))
    ]);
  } else {
    ws = XLSX.utils.json_to_sheet([{ value: stringifyJsonCell(parsed) }]);
  }

  XLSX.utils.book_append_sheet(wb, ws, 'JSON');
  return wb;
};

const parseJsonArrayCell = (value: unknown): unknown[] | null => {
  if (Array.isArray(value)) return value;
  if (typeof value !== 'string') return null;

  const trimmed = value.trim();
  if (!trimmed.startsWith('[') || !trimmed.endsWith(']')) return null;

  try {
    const parsed = JSON.parse(trimmed) as unknown;
    return Array.isArray(parsed) ? parsed : null;
  } catch {
    return null;
  }
};

const expandJsonArrayRows = (
  rows: any[],
  targetCol: string,
  valueKey: string = 'value',
  includeSourceColumns: boolean = false
): { data: any[]; headers: string[] } => {
  const expanded: any[] = [];
  const headers = new Set<string>();

  rows.forEach((row, rowIndex) => {
    const parsed = parseJsonArrayCell(row[targetCol]);
    if (!parsed) return;

    parsed.forEach((item, itemIndex) => {
      const nextRow: Record<string, any> = includeSourceColumns ? { ...row } : {};

      if (isRecord(item)) {
        Object.entries(item).forEach(([key, value]) => {
          nextRow[key] = stringifyJsonCell(value);
        });
      } else {
        nextRow[valueKey] = stringifyJsonCell(item);
      }

      nextRow._sourceRow = rowIndex + 1;
      nextRow._itemIndex = itemIndex;
      expanded.push(nextRow);
      Object.keys(nextRow).forEach((key) => headers.add(key));
    });
  });

  return { data: expanded, headers: Array.from(headers) };
};

const insertColumnAt = (headers: string[], columnName: string, insertAfterCol?: string): string[] => {
  const nextHeaders = headers.filter((header) => header !== columnName);
  if (!insertAfterCol || insertAfterCol === LAST_COLUMN_OPTION || !nextHeaders.includes(insertAfterCol)) {
    return [...nextHeaders, columnName];
  }
  const insertIndex = nextHeaders.indexOf(insertAfterCol) + 1;
  return [...nextHeaders.slice(0, insertIndex), columnName, ...nextHeaders.slice(insertIndex)];
};

const reorderRowsByHeaders = (rows: any[], headers: string[]): any[] =>
  rows.map((row) => {
    const ordered: Record<string, any> = {};
    headers.forEach((header) => {
      if (header in row) ordered[header] = row[header];
    });
    Object.keys(row).forEach((key) => {
      if (!(key in ordered)) ordered[key] = row[key];
    });
    return ordered;
  });

const sampleRowsForChart = (rows: any[], maxPoints: number = MAX_CHART_RENDER_POINTS): any[] => {
  if (rows.length <= maxPoints) return rows;
  const step = (rows.length - 1) / (maxPoints - 1);
  const sampled: any[] = [];
  for (let i = 0; i < maxPoints; i++) sampled.push(rows[Math.round(i * step)]);
  return sampled;
};

type CalcRuntime = {
  nodesById: Map<string, CustomNode>;
  primaryInputByTarget: Map<string, Edge>;
  inputAByTarget: Map<string, Edge>;
  inputBByTarget: Map<string, Edge>;
  resultCache: Map<string, CalcDataResult>;
};

const workbookExtractCache = new WeakMap<XLSX.WorkBook, Map<string, CalcDataResult>>();

const createCalcRuntime = (nodes: CustomNode[], edges: Edge[]): CalcRuntime => {
  const nodesById = new Map(nodes.map((node) => [node.id, node]));
  const primaryInputByTarget = new Map<string, Edge>();
  const inputAByTarget = new Map<string, Edge>();
  const inputBByTarget = new Map<string, Edge>();

  edges.forEach((edge) => {
    const targetHandle = (edge as any).targetHandle;
    if (targetHandle === 'input-a') inputAByTarget.set(edge.target, edge);
    else if (targetHandle === 'input-b') inputBByTarget.set(edge.target, edge);
    else if (!primaryInputByTarget.has(edge.target)) primaryInputByTarget.set(edge.target, edge);
  });

  return {
    nodesById,
    primaryInputByTarget,
    inputAByTarget,
    inputBByTarget,
    resultCache: new Map<string, CalcDataResult>(),
  };
};

const getCachedWorkbookExtract = (workbook: XLSX.WorkBook, cacheKey: string, compute: () => CalcDataResult): CalcDataResult => {
  let workbookCache = workbookExtractCache.get(workbook);
  if (!workbookCache) {
    workbookCache = new Map<string, CalcDataResult>();
    workbookExtractCache.set(workbook, workbookCache);
  }
  const cached = workbookCache.get(cacheKey);
  if (cached) return cached;
  const result = compute();
  workbookCache.set(cacheKey, result);
  return result;
};

const readWorkbookFromFile = (
  file: File | Blob,
  fileName: string,
  onSuccess: (workbook: XLSX.WorkBook) => void,
  onError: (error: Error) => void
) => {
  const lowerName = fileName.toLowerCase();
  const reader = new FileReader();

  reader.onload = (evt: ProgressEvent<FileReader>) => {
    try {
      const result = evt.target?.result;
      const workbook = lowerName.endsWith('.json')
        ? jsonToWorkbook(String(result || ''))
        : XLSX.read(result, { type: 'binary' });
      onSuccess(workbook);
    } catch (error) {
      onError(error instanceof Error ? error : new Error('Unknown file read error'));
    }
  };

  reader.onerror = () => {
    onError(new Error('ファイルの読み込み中にエラーが発生しました。'));
  };

  if (lowerName.endsWith('.json')) {
    reader.readAsText(file);
  } else {
    reader.readAsBinaryString(file);
  }
};

const extractDataFromMatrix = (
  matrix: any[][],
  ranges: string[] = [],
  useHdr: boolean = true
): { data: any[]; headers: string[] } => {
  const mat = sanitizeMatrix(matrix);
  if (mat.length === 0) return { data: [], headers: [] };

  const colCount = Math.max(...mat.map((row) => row.length), 0);
  if (colCount === 0) return { data: [], headers: [] };

  const defaultRange = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(0, mat.length - 1), c: Math.max(0, colCount - 1) }
  });
  const targetRanges = ranges.length > 0 ? ranges : [defaultRange];

  let headers: string[] = [];
  let data: any[] = [];

  targetRanges.forEach((rangeStr, idx) => {
    const range = XLSX.utils.decode_range(rangeStr);
    const extractedRows: string[][] = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = Array.from({ length: range.e.c - range.s.c + 1 }, (_, cIdx) => mat[r]?.[range.s.c + cIdx] ?? '');
      extractedRows.push(row);
    }
    if (extractedRows.length === 0) return;

    if (idx === 0) {
      headers = useHdr
        ? extractedRows[0].map((header, colIdx) => header ? String(header) : `Col_${colIdx + 1}`)
        : Array.from({ length: extractedRows[0].length }, (_, colIdx) => `Col_${colIdx + 1}`);

      for (let rowIdx = useHdr ? 1 : 0; rowIdx < extractedRows.length; rowIdx++) {
        const obj: Record<string, string> = {};
        headers.forEach((header, colIdx) => { obj[header] = extractedRows[rowIdx][colIdx] ?? ''; });
        data.push(obj);
      }
      return;
    }

    extractedRows.forEach((row) => {
      const obj: Record<string, string> = {};
      headers.forEach((header, colIdx) => { obj[header] = row[colIdx] ?? ''; });
      data.push(obj);
    });
  });

  return { data, headers };
};


const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:wght@700;800&family=Zen+Kaku+Gothic+New:wght@300;400;500;700&display=swap');
    body { font-family: 'Zen Kaku Gothic New', sans-serif; font-weight: 300; }
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    .dark ::-webkit-scrollbar-track { background: #1a1a1a; }
    .dark ::-webkit-scrollbar-thumb { background: #444; border-radius: 4px; }
    ::-webkit-scrollbar-track { background: #f3f4f6; }
    ::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 4px; }
    
    .react-flow__node { cursor: grab !important; }
    .react-flow__node:active { cursor: grabbing !important; }
    
    .react-flow__handle { 
      width: 18px !important; 
      height: 18px !important; 
      border: 3px solid #fff !important; 
      background-color: #3b82f6 !important; 
      transition: transform 0.1s ease; 
    }
    .dark .react-flow__handle { border: 3px solid #1e1e1e !important; background-color: #38bdf8 !important; }
    .react-flow__handle:hover { transform: scale(1.5); }
    .react-flow__handle::after {
      content: '';
      position: absolute;
      top: -20px; left: -20px; right: -20px; bottom: -20px;
      background: transparent;
    }

    .custom-scrollbar::-webkit-scrollbar { width: 4px; }
    @keyframes nodeIntroPop {
      0% { opacity: 0; transform: scale(0.96); }
      100% { opacity: 1; transform: scale(1); }
    }
    .node-intro-pop {
      animation: nodeIntroPop 220ms cubic-bezier(0.22, 1, 0.36, 1) both;
      will-change: opacity, transform;
    }
    
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

const IconSvg = ({ children, className = "w-[1em] h-[1em]" }: any) => (
  <svg viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" className={className}>{children}</svg>
);

const Icons = {
  Sun: <IconSvg><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></IconSvg>,
  Moon: <IconSvg><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></IconSvg>,
  Help: <IconSvg><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></IconSvg>,
  Source: <IconSvg><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/></IconSvg>,
  FolderAuto: <IconSvg><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/><path d="M12 11v6"/><polyline points="9 14 12 17 15 14"/></IconSvg>,
  Database: <IconSvg><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></IconSvg>,
  Zap: <IconSvg><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></IconSvg>,
  Layout: <IconSvg><rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="21" x2="9" y2="9"/></IconSvg>,
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
  Focus: <IconSvg><circle cx="12" cy="12" r="3"/><path d="M3 7V5a2 2 0 0 1 2-2h2"/><path d="M17 3h2a2 2 0 0 1 2 2v2"/><path d="M21 17v2a2 2 0 0 1-2 2h-2"/><path d="M7 21H5a2 2 0 0 1-2-2v-2"/></IconSvg>,
  Maximize: <IconSvg><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/></IconSvg>,
  Minimize: <IconSvg><path d="M8 3v3a2 2 0 0 1-2 2H3m18 0h-3a2 2 0 0 1-2-2V3m0 18v-3a2 2 0 0 1 2-2h3M3 16h3a2 2 0 0 1 2 2v3"/></IconSvg>
};

const LandingPage = () => {
  const landingSections = useMemo(
    () => [
      { id: 'hero', label: 'Home' },
      { id: 'features', label: 'Features' },
      { id: 'use-cases', label: 'Use Cases' },
      { id: 'how-it-works', label: 'Workflow' },
      { id: 'contact-request', label: 'Contact' },
    ],
    [],
  );
  const [activeSectionId, setActiveSectionId] = useState('hero');
  const activeSectionIndex = landingSections.findIndex((section) => section.id === activeSectionId);
  const goToSection = useCallback((sectionId: string) => {
    setActiveSectionId(sectionId);
  }, []);
  const goToSectionByIndex = useCallback(
    (index: number) => {
      if (index < 0 || index >= landingSections.length) return;
      setActiveSectionId(landingSections[index].id);
    },
    [landingSections],
  );
  const getSectionPanelClass = useCallback(
    (sectionId: string) =>
      `absolute inset-0 transition-all duration-500 ease-out ${
        activeSectionId === sectionId
          ? 'opacity-100 translate-x-0 pointer-events-auto'
          : 'opacity-0 translate-x-8 pointer-events-none'
      }`,
    [activeSectionId],
  );

  return (
    <div className="h-screen overflow-hidden bg-gray-50 text-gray-800 font-sans selection:bg-gray-200">
      <nav className="border-b border-gray-200 bg-white/80 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-5xl mx-auto px-6 py-4 flex justify-between items-center">
          <div className="flex items-center gap-3">

          </div>
          <div className="flex items-center gap-6">
            <button type="button" onClick={() => goToSection('features')} className={`text-xs font-bold transition-colors hidden md:block tracking-wider uppercase ${activeSectionId === 'features' ? 'text-gray-900' : 'text-gray-500 hover:text-gray-900'}`}>Features</button>
            <button type="button" onClick={() => goToSection('use-cases')} className={`text-xs font-bold transition-colors hidden md:block tracking-wider uppercase ${activeSectionId === 'use-cases' ? 'text-gray-900' : 'text-gray-500 hover:text-gray-900'}`}>Use Cases</button>
            <button type="button" onClick={() => goToSection('how-it-works')} className={`text-xs font-bold transition-colors hidden md:block tracking-wider uppercase ${activeSectionId === 'how-it-works' ? 'text-gray-900' : 'text-gray-500 hover:text-gray-900'}`}>Workflow</button>
            <button type="button" onClick={() => goToSection('contact-request')} className={`text-xs font-bold transition-colors hidden md:block tracking-wider uppercase ${activeSectionId === 'contact-request' ? 'text-gray-900' : 'text-gray-500 hover:text-gray-900'}`}>Contact</button>
            <a href="#/app" className="bg-gray-800 hover:bg-gray-700 text-white text-xs font-bold px-5 py-2.5 rounded-lg tracking-widest transition-all shadow-sm flex items-center gap-2">
              ツールを開く <span className="w-3 h-3 flex items-center justify-center">{Icons.ArrowRight}</span>
            </a>
          </div>
        </div>
      </nav>
      <main className="relative h-[calc(100vh-74px)] overflow-hidden">
        <section className={getSectionPanelClass('hero')}>
          <div className="h-full flex items-center justify-center px-6 py-10">
            <div className="max-w-5xl mx-auto relative z-10 text-center -translate-y-2 md:-translate-y-4">
              <div className="mb-5">
                <img
                  src={titleImage}
                  alt="Data shaping & Visual SQL Building"
                  className="w-full max-w-5xl mx-auto h-auto rounded-[2rem] shadow-[0_24px_60px_rgba(0,0,0,0.12)]"
                />
              </div>
              <p className="text-sm md:text-base text-gray-600 max-w-2xl mx-auto mb-8 leading-relaxed">
                ドラッグ＆ドロップの直感的な操作で、xlsx、CSVやJsonなどのデータを自由自在に結合・整形・計算。複雑なデータ処理パイプラインやSQLクエリを誰でも手軽に構築できるツールです。
              </p>
              <div className="flex justify-center items-center gap-4">
                <a href="#/app" className="bg-gray-800 hover:bg-gray-700 text-white text-sm font-bold px-8 py-3.5 rounded-xl tracking-widest transition-all shadow-md flex items-center gap-2 hover:scale-105">
                  使ってみる <span className="w-4 h-4 flex items-center justify-center">{Icons.ArrowRight}</span>
                </a>
              </div>
            </div>
          </div>
        </section>

        <section id="features" className={getSectionPanelClass('features')}>
          <div className="h-full flex items-center justify-center px-6 py-10">
            <div className="max-w-5xl mx-auto w-full">
              <div className="text-center mb-10">
                <h2 className="text-xl  tracking-[0.3em] text-gray-500 mb-2 uppercase">Features</h2>
              </div>

              <div className="grid md:grid-cols-2 lg:grid-cols-4 gap-6">
                {[
                  { i: Icons.Layout, t: "直感的なノーコードUI", d: "ノードをキャンバスに配置して線で繋ぐだけです。" },
                  { i: Icons.Database, t: "データソースの選択幅", d: "ローカルのCSV/Excelだけでなく、フォルダ自動監視やコピペ入力、複数ソースの結合・比較などにも対応。" },
                  { i: Icons.Zap, t: "簡単なクレンジング", d: "VLOOKUP的な結合、文字列抽出、ゼロ埋め、四則演算、条件によっての更新など…さまざまな種類のノードを用意しています。" },
                  { i: Icons.Code, t: "SQLの相互変換", d: "作成したフローからSELECT文を自動生成。逆にSQLからノードを自動配置することも可能。" }
                ].map((f, idx) => (
                  <div key={idx} className="bg-gray-50 border border-gray-200 p-5 rounded-xl hover:border-gray-400 transition-colors">
                    <div className="w-10 h-10 bg-white rounded-lg flex items-center justify-center text-gray-800 mb-3 border border-gray-200 shadow-sm">
                      <span className="w-5 h-5 flex items-center justify-center">{f.i}</span>
                    </div>
                    <h4 className="text-sm font-bold text-gray-900 mb-1.5">{f.t}</h4>
                    <p className="text-xs text-gray-600 leading-relaxed">{f.d}</p>
                  </div>
                ))}
              </div>

            </div>
          </div>
        </section>

        <section id="use-cases" className={getSectionPanelClass('use-cases')}>
          <div className="h-full flex items-center justify-center px-6 py-10 bg-gray-50">
            <div className="max-w-5xl mx-auto w-full">
              <div className="text-center mb-10">
                <h2 className="text-xl tracking-[0.3em] text-gray-500 mb-2 uppercase">Use Cases</h2>
              </div>

              <div className="grid md:grid-cols-3 gap-6">
                {[
                  { s: "Case 1", t: "複数システムのデータ統合", d: "販売管理システムと顧客管理システムなど、別々にエクスポートされたCSVデータを共通キー（顧客IDなど）で手軽にJOINし、分析用の統合データを作成できます。" },
                  { s: "Case 2", t: "定期レポート作成の自動化", d: "Auto Folderノードで指定フォルダの「最新の売上データ」を自動読み込み。フローを一度作れば、毎月同じ整形・集計処理を手作業で行う手間を省けます。" },
                  { s: "Case 3", t: "データクレンジングと名寄せ", d: "表記ゆれや不要な文字の削除、ゼロ埋め、条件分岐などを視覚的に設定。エンジニアに依頼することなく、現場の担当者だけでデータの正規化を完結させられます。" }
                ].map((step, idx) => (
                  <div key={idx} className="relative">
                    {idx !== 2 && <div className="hidden md:block absolute top-6 left-1/2 w-full h-[1px] border-t border-dashed border-gray-300 z-0"></div>}
                    <div className="bg-gray-50 border border-gray-200 p-5 rounded-xl relative z-10 h-full hover:border-gray-400 transition-colors">
                      <div className="text-gray-500 font-bold tracking-widest text-[10px] mb-1.5 uppercase">{step.s}</div>
                      <h4 className="text-sm font-bold text-gray-900 mb-1.5 uppercase">{step.t}</h4>
                      <p className="text-xs text-gray-600 leading-relaxed">{step.d}</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </section>

        <section id="how-it-works" className={getSectionPanelClass('how-it-works')}>
          <div className="h-full flex items-center justify-center px-6 py-10">
            <div className="max-w-5xl mx-auto w-full">
              <div className="text-center mb-10">
                <h2 className="text-xl tracking-[0.3em] text-gray-500 mb-2 uppercase">Workflow</h2>
              </div>

              <div className="grid md:grid-cols-3 gap-6">
                {[
                  { s: "Step 1", t: "Add Nodes", d: "左側のToolboxから、読み込み(Source)や結合(Join)などのノードをドラッグ＆ドロップで配置します。" },
                  { s: "Step 2", t: "Connect Flow", d: "ノード同士の端子をマウスで繋ぎます。データが左から右へと水のように流れて処理されます。" },
                  { s: "Step 3", t: "Preview & Export", d: "画面下部に結果がリアルタイム表示されます。グラフ化や、CSV・Excelへのエクスポート・保存が可能です。" }
                ].map((step, idx) => (
                  <div key={idx} className="relative">
                    {idx !== 2 && <div className="hidden md:block absolute top-6 left-1/2 w-full h-[1px] border-t border-dashed border-gray-300 z-0"></div>}
                    <div className="bg-gray-50 border border-gray-200 p-5 rounded-xl relative z-10 h-full hover:border-gray-400 transition-colors">
                      <div className="text-gray-500 font-bold tracking-widest text-[10px] mb-1.5 uppercase">{step.s}</div>
                      <h4 className="text-sm font-bold text-gray-900 mb-1.5 uppercase">{step.t}</h4>
                      <p className="text-xs text-gray-600 leading-relaxed">{step.d}</p>
                    </div>
                  </div>
                ))}
              </div>

            </div>
          </div>
        </section>

        <section id="contact-request" className={getSectionPanelClass('contact-request')}>
          <div className="h-full flex flex-col justify-between bg-gray-50">
            <div className="flex-1 flex items-center justify-center px-6 py-10">
              <div className="max-w-5xl mx-auto w-full">
                <div className="text-center">
                  <h2 className="text-xs font-bold tracking-[0.3em] text-gray-500 mb-2 uppercase">Contact / Request</h2>
                  <h3 className="text-2xl font-bold text-gray-900">お問い合わせ・ご要望</h3>
                  <a
                    href="mailto:support@enjoyworkstore.com"
                    className="inline-flex mt-4 text-base font-bold text-gray-700 hover:text-gray-900 underline underline-offset-4 transition-colors"
                  >
                    support@enjoyworkstore.com
                  </a>
                </div>
              </div>
            </div>
            <footer className="border-t border-gray-200 bg-white py-6">
              <div className="max-w-5xl mx-auto px-6 flex justify-end">
                <p className="text-[10px] font-bold tracking-widest text-gray-500 uppercase text-right">
                  &copy; {new Date().getFullYear()} enjoyworkstore. All rights reserved.
                </p>
              </div>
            </footer>
          </div>
        </section>

        <div className="absolute bottom-3 md:bottom-2 left-1/2 -translate-x-1/2 z-40 flex items-center gap-3 no-print">
          <button
            type="button"
            onClick={() => goToSectionByIndex(activeSectionIndex - 1)}
            disabled={activeSectionIndex <= 0}
            className={`w-10 h-10 rounded-full border flex items-center justify-center transition-colors ${
              activeSectionIndex <= 0
                ? 'bg-white/60 border-gray-200 text-gray-300'
                : 'bg-white border-gray-200 text-gray-700 hover:text-gray-900 hover:bg-gray-50 shadow-sm'
            }`}
          >
            {Icons.ArrowLeft}
          </button>
          <div className="flex items-center gap-2 rounded-full bg-white/90 border border-gray-200 px-4 py-2 shadow-sm">
            {landingSections.map((section, index) => (
              <button
                key={section.id}
                type="button"
                onClick={() => goToSection(section.id)}
                className={`w-2.5 h-2.5 rounded-full transition-all ${
                  activeSectionId === section.id ? 'bg-gray-800 scale-125' : 'bg-gray-300 hover:bg-gray-400'
                }`}
                aria-label={`${index + 1}ページ目へ移動`}
              />
            ))}
          </div>
          <button
            type="button"
            onClick={() => goToSectionByIndex(activeSectionIndex + 1)}
            disabled={activeSectionIndex >= landingSections.length - 1}
            className={`w-10 h-10 rounded-full border flex items-center justify-center transition-colors ${
              activeSectionIndex >= landingSections.length - 1
                ? 'bg-white/60 border-gray-200 text-gray-300'
                : 'bg-white border-gray-200 text-gray-700 hover:text-gray-900 hover:bg-gray-50 shadow-sm'
            }`}
          >
            {Icons.ArrowRight}
          </button>
        </div>
      </main>
    </div>
  );
}

const calcData = (nId: string, nodes: CustomNode[], edges: Edge[], wbs: any, runtime?: CalcRuntime): CalcDataResult => {
  const activeRuntime = runtime || createCalcRuntime(nodes, edges);
  const cached = activeRuntime.resultCache.get(nId);
  if (cached) return cached;

  const node = activeRuntime.nodesById.get(nId);
  if (!node) return { data: [], headers: [] };

  let result: CalcDataResult = { data: [], headers: [] };

  if (node.type === 'pasteNode') {
    try {
      const tableData = Array.isArray(node.data.tableData) && node.data.tableData.length > 0
        ? node.data.tableData
        : parseDelimitedTextToMatrix(node.data.rawData || '');
      result = extractDataFromMatrix(tableData, node.data.ranges || [], node.data.useFirstRowAsHeader !== false);
    } catch (e) {
      result = { data: [], headers: [] };
    }
    activeRuntime.resultCache.set(nId, result);
    return result;
  }

  if (node.type === 'dataNode' || node.type === 'folderSourceNode') {
    if (node.data.needsUpload) {
      activeRuntime.resultCache.set(nId, result);
      return result;
    }
    const wb = wbs[node.id];
    if (!wb) {
      activeRuntime.resultCache.set(nId, result);
      return result;
    }
    const ws = wb.Sheets[node.data.currentSheet || wb.SheetNames[0]];
    if (!ws) {
      activeRuntime.resultCache.set(nId, result);
      return result;
    }
    try {
      const ranges = (node.data.ranges || []).length === 0 && ws['!ref']
        ? [XLSX.utils.encode_range(XLSX.utils.decode_range(ws['!ref']))]
        : (node.data.ranges || []);
      const cacheKey = [
        node.data.currentSheet || wb.SheetNames[0],
        node.data.useFirstRowAsHeader !== false ? 'header' : 'no-header',
        ranges.join('|'),
      ].join('::');
      result = getCachedWorkbookExtract(wb, cacheKey, () => {
        const mat = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: true }) as any[][];
        if (!mat || mat.length === 0) return { data: [], headers: [] };
        return extractDataFromMatrix(mat, ranges, node.data.useFirstRowAsHeader !== false);
      });
    } catch (e) {
      result = { data: [], headers: [] };
    }
    activeRuntime.resultCache.set(nId, result);
    return result;
  }

  if (node.type === 'unionNode' || node.type === 'joinNode' || node.type === 'minusNode' || node.type === 'vlookupNode') {
    const eA = activeRuntime.inputAByTarget.get(nId);
    const eB = activeRuntime.inputBByTarget.get(nId);
    if (!eA || !eB) {
      activeRuntime.resultCache.set(nId, result);
      return result;
    }
    const rA = calcData(eA.source, nodes, edges, wbs, activeRuntime);
    const rB = calcData(eB.source, nodes, edges, wbs, activeRuntime);

    if (node.type === 'unionNode') {
      result = { data: [...rA.data, ...rB.data], headers: rA.headers };
      activeRuntime.resultCache.set(nId, result);
      return result;
    }

    if (node.type === 'minusNode') {
      const { keyA, keyB } = node.data;
      if (!keyA || !keyB) {
        activeRuntime.resultCache.set(nId, rA);
        return rA;
      }
      const bKeys = new Set(rB.data.map((b) => String(b[keyB as string])));
      result = { data: rA.data.filter((a) => !bKeys.has(String(a[keyA as string]))), headers: rA.headers };
      activeRuntime.resultCache.set(nId, result);
      return result;
    }

    if (node.type === 'vlookupNode') {
      const { keyA, keyB, fetchCol, targetCol } = node.data;
      if (!keyA || !keyB || !fetchCol || !targetCol) {
        activeRuntime.resultCache.set(nId, rA);
        return rA;
      }

      const bMap = new Map<string, any>();
      rB.data.forEach((b) => {
        bMap.set(String(b[keyB as string]), b[fetchCol as string]);
      });

      const vData = rA.data.map((a) => {
        const key = String(a[keyA as string]);
        const val = bMap.has(key) ? bMap.get(key) : null;
        return { ...a, [targetCol]: val };
      });
      const nextHeaders = insertColumnAt(rA.headers, targetCol, node.data.insertAfterCol);
      result = { data: reorderRowsByHeaders(vData, nextHeaders), headers: nextHeaders };
      activeRuntime.resultCache.set(nId, result);
      return result;
    }

    const { keyA, keyB, joinType = 'inner' } = node.data;
    if (!keyA || !keyB) {
      activeRuntime.resultCache.set(nId, rA);
      return rA;
    }

    const joined: any[] = [];
    if (joinType === 'inner' || joinType === 'left') {
      const bGrouped = new Map<string, any[]>();
      rB.data.forEach((b) => {
        const key = String(b[keyB as string]);
        const current = bGrouped.get(key);
        if (current) current.push(b);
        else bGrouped.set(key, [b]);
      });

      rA.data.forEach((a) => {
        const matches = bGrouped.get(String(a[keyA as string])) || [];
        if (matches.length > 0) {
          matches.forEach((b) => joined.push({ ...a, ...b }));
        } else if (joinType === 'left') {
          joined.push({ ...a });
        }
      });
    } else if (joinType === 'right') {
      const aGrouped = new Map<string, any[]>();
      rA.data.forEach((a) => {
        const key = String(a[keyA as string]);
        const current = aGrouped.get(key);
        if (current) current.push(a);
        else aGrouped.set(key, [a]);
      });

      rB.data.forEach((b) => {
        const matches = aGrouped.get(String(b[keyB as string])) || [];
        if (matches.length > 0) {
          matches.forEach((a) => joined.push({ ...a, ...b }));
        } else {
          joined.push({ ...b });
        }
      });
    }

    result = { data: joined, headers: [...new Set([...rA.headers, ...rB.headers])] };
    activeRuntime.resultCache.set(nId, result);
    return result;
  }

  const inEdge = activeRuntime.primaryInputByTarget.get(nId);
  if (!inEdge) {
    activeRuntime.resultCache.set(nId, result);
    return result;
  }
  const input = calcData(inEdge.source, nodes, edges, wbs, activeRuntime);
  let out = [...input.data], h = [...input.headers];

  if (node.type === 'sortNode') {
    const { sortCol, sortOrder } = node.data;
    if (sortCol) out.sort((a, b) => sortOrder === 'desc' ? String(b[sortCol as string]).localeCompare(String(a[sortCol as string]), undefined, { numeric: true }) : String(a[sortCol as string]).localeCompare(String(b[sortCol as string]), undefined, { numeric: true }));
  }

  if (node.type === 'filterNode') {
    const { filterCol, filterVal, matchType = 'includes' } = node.data;
    if (filterCol && filterVal !== undefined && filterVal !== '') {
      out = out.filter(r => matchesCondition(r, filterCol as string, filterVal, matchType));
    }
  }

  if (node.type === 'dataCheckNode') {
    return { data: out, headers: h };
  }

  if (node.type === 'selectNode') {
    const sel = node.data.selectedColumns || [];
    if (sel.length > 0) { h = sel; out = out.map(r => { const nr: any = {}; sel.forEach((c: string) => nr[c] = r[c]); return nr; }); }
  }

  if (node.type === 'jsonArrayNode') {
    const { targetCol, valueKey = 'value', includeSourceColumns = false } = node.data;
    if (!targetCol) {
      result = { data: out, headers: h };
      activeRuntime.resultCache.set(nId, result);
      return result;
    }
    result = expandJsonArrayRows(out, targetCol, valueKey, includeSourceColumns);
    activeRuntime.resultCache.set(nId, result);
    return result;
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
      out = Object.values(grps).map((g: any) => ({ [groupCol as string]: g[groupCol as string], [aggCol as string]: aggType === 'count' ? g._c : g._v }));
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
    } else if (command === 'auto_number') {
      const outCol = (node.data.createNewCol && node.data.newColName) ? node.data.newColName : targetCol;
      if (outCol) {
        const mode = node.data.autoNumberMode || 'number';
        const prefix = String(node.data.autoNumberPrefix || '');
        const digits = Math.max(0, Number(node.data.autoNumberDigits) || 0);
        out = out.map((r, idx) => {
          const base = idx + 1;
          const padded = digits > 0 ? String(base).padStart(digits, '0') : String(base);
          const nextVal = mode === 'prefix'
            ? `${prefix}${padded}`
            : (digits > 0 ? padded : base);
          return { ...r, [outCol as string]: nextVal };
        });
        if (node.data.createNewCol && node.data.newColName) {
          h = insertColumnAt(h, node.data.newColName, node.data.insertAfterCol);
          out = reorderRowsByHeaders(out, h);
        } else if (!h.includes(outCol)) {
          h = [...h, outCol];
        }
      }
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

        const outCol = (node.data.createNewCol && node.data.newColName) ? node.data.newColName : targetCol;
        return { ...r, [outCol as string]: v };
      });
      if (node.data.createNewCol && node.data.newColName && !h.includes(node.data.newColName)) {
        h = insertColumnAt(h, node.data.newColName, node.data.insertAfterCol);
        out = reorderRowsByHeaders(out, h);
      }
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
      h = insertColumnAt(h, newColName, node.data.insertAfterCol);
      out = reorderRowsByHeaders(out, h);
    }
  }

  result = { data: out, headers: h };
  activeRuntime.resultCache.set(nId, result);
  return result;
};

const DataNode = memo(({ id, data }: any) => {
  const { setWorkbooks, setRangeModalNode, setPasteEditorNode, focusNode, theme } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const isDark = theme === 'dark';
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const processFile = (f: File | Blob, fileName: string, pathStr: string) => {
    readWorkbookFromFile(
      f,
      fileName,
      (wb) => {
      setWorkbooks((p: any) => ({ ...p, [id]: wb }));
      updateNodeData(id, { fileName, filePath: pathStr, sheetNames: wb.SheetNames, currentSheet: wb.SheetNames[0], needsUpload: false });
      focusNode(id, false, false, 'resize');
      },
      (error) => {
        console.error('Failed to parse source file:', error);
        const isJson = fileName.toLowerCase().endsWith('.json');
        alert(
          isJson
            ? error.message || 'JSONファイルの解析に失敗しました。JSON形式として正しい内容か確認してください。'
            : 'ファイルの解析に失敗しました。対応していない形式か、ファイルが破損しています。'
        );
      }
    );
  };

  const onUp = (e: any) => {
    const f = e.target.files?.[0]; 
    if (!f) return;
    processFile(f, f.name, (f as any).path || f.name);
    e.target.value = '';
  };

  const tryAutoReload = async (e: React.MouseEvent) => {
    e.preventDefault();
    const targetPath = data.filePath || data.fileName;
    if (!targetPath) {
      fileInputRef.current?.click();
      return;
    }
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), 10000);
    try {
      const res = await fetch(targetPath, { signal: controller.signal });
      if (!res.ok) {
        throw new Error(`HTTP ${res.status}: ${res.statusText || 'Fetch failed'}`);
      }
      const blob = await res.blob();
      processFile(blob, data.fileName, targetPath);
    } catch (err) {
      console.error('Auto reload failed:', err);
      let feedback = 'ファイルの自動再読み込みに失敗しました。手動でファイルを再選択してください。';
      if (err instanceof DOMException && err.name === 'AbortError') {
        feedback = 'ファイルの自動再読み込みがタイムアウトしました。ネットワークやファイルパスを確認し、手動でファイルを再選択してください。';
      } else if (err instanceof TypeError) {
        feedback = 'ファイルの自動再読み込み中にネットワークエラーが発生しました。接続先またはファイルパスを確認し、手動でファイルを再選択してください。';
      } else if (err instanceof Error) {
        feedback = `ファイルの自動再読み込みに失敗しました (${err.message})。手動でファイルを再選択してください。`;
      }
      alert(feedback);
      fileInputRef.current?.click();
    } finally {
      window.clearTimeout(timeoutId);
    }
  };

  const summary = data.fileName ? data.fileName : '';
  return (
    <NodeWrap id={id} data={data} title="Source" col={isDark ? "text-blue-400" : "text-blue-600"} showTgt={false} summary={summary} helpText="ローカルのCSV・Excel・JSONファイルを選択して読み込みます。パネルから抽出範囲やヘッダーの設定が可能です。">
      <input type="file" accept=".csv,.xlsx,.json,application/json" className="hidden" ref={fileInputRef} onChange={onUp} />
      {data.needsUpload ? (
        <div className="space-y-3">
          <div className={`text-[10px] ${isDark ? 'text-white bg-blue-500/20 border-blue-500/50' : 'text-gray-800 bg-blue-50 border-blue-200'} flex items-center gap-2 p-2 rounded border`}>
            <span className={`${isDark ? 'text-blue-400' : 'text-blue-600'} flex items-center justify-center`}>{Icons.Warning}</span> Missing: {data.fileName}
          </div>
          <button onClick={tryAutoReload} className={`w-full ${isDark ? 'text-blue-400 border-blue-500/50 hover:bg-blue-500/20' : 'text-blue-600 border-blue-300 hover:bg-blue-50'} text-[10px] border border-dashed p-3 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors shadow-sm animate-pulse nodrag`}>
            <span className="flex items-center justify-center">{Icons.Folder}</span> 再設定
          </button>
        </div>
      ) : !data.fileName ? (
        <button onClick={() => fileInputRef.current?.click()} className={`w-full ${isDark ? 'text-blue-400 border-blue-500/50 hover:bg-blue-500/10' : 'text-blue-600 border-blue-300 hover:bg-blue-50'} text-[10px] border border-dashed p-4 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors nodrag`}>
          <span className="flex items-center justify-center w-4 h-4">{Icons.Folder}</span> Load File
        </button>
      ) : (
        <div className="space-y-3">
          <div className={`flex justify-between items-center ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'} p-2 rounded border transition-colors`}>
            <div className={`text-[10px] ${isDark ? 'text-white' : 'text-gray-800'} font-bold truncate flex items-center gap-2`}>
              <span className={`${isDark ? 'text-blue-400' : 'text-blue-600'} flex items-center justify-center`}>{Icons.File}</span> {data.fileName}
            </div>
            <button onClick={() => fileInputRef.current?.click()} className={`cursor-pointer ${isDark ? 'text-blue-400 hover:text-white' : 'text-blue-600 hover:text-gray-900'} text-[12px] font-bold uppercase transition-colors nodrag`} title="Change File">
              <span className="flex items-center justify-center">{Icons.Refresh}</span>
            </button>
          </div>
          <button onClick={() => setPasteEditorNode({ nodeId: id, selectionMode: false })} className={`w-full py-2 ${isDark ? 'bg-[#252526] text-blue-400 border-[#444] hover:bg-[#333]' : 'bg-white text-blue-600 border-gray-300 hover:bg-blue-50'} text-[10px] font-bold rounded border uppercase tracking-widest transition-colors nodrag`}>表を編集</button>
          <button onClick={() => setRangeModalNode(id)} className={`w-full py-2 ${isDark ? 'bg-blue-600/20 text-blue-400 border-blue-500/30 hover:bg-blue-600/40' : 'bg-blue-50 text-blue-600 border-blue-200 hover:bg-blue-100'} text-[10px] font-bold rounded border uppercase tracking-widest transition-colors nodrag`}>範囲選択</button>
          <label className="flex items-center gap-2 pt-2 cursor-pointer group"><input type="checkbox" checked={data.useFirstRowAsHeader !== false} onChange={(e) => updateNodeData(id, { useFirstRowAsHeader: e.target.checked })} className="accent-blue-500 w-4 h-4 cursor-pointer nodrag" /><span className={`text-[10px] ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-600 group-hover:text-gray-900'} font-bold uppercase transition-colors`}>1行目をヘッダーにする</span></label>
        </div>
      )}
    </NodeWrap>
  );
});

const FolderSourceNode = memo(({ id, data }: any) => {
  const { setWorkbooks, setRangeModalNode, setPasteEditorNode, focusNode, theme } = useContext(AppContext);
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
        if (lowerName.endsWith('.csv') || lowerName.endsWith('.xlsx') || lowerName.endsWith('.json')) {
          if (file.lastModified > latestTime) {
            latestTime = file.lastModified;
            latestFile = file;
          }
        }
      }

      if (latestFile) {
        readWorkbookFromFile(
          latestFile,
          latestFile.name,
          (wb) => {
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
            focusNode(id, false, false, 'resize');
            setIsLoading(false);
          },
          (error) => {
            console.error('Failed to parse folder source file:', error);
            const isJson = latestFile!.name.toLowerCase().endsWith('.json');
            alert(
              isJson
                ? error.message || 'JSONファイルの解析に失敗しました。JSON形式として正しい内容か確認してください。'
                : 'ファイルの解析に失敗しました。対応していない形式か、ファイルが破損しています。'
            );
            setIsLoading(false);
          }
        );
      } else {
        alert("選択されたフォルダ内に、CSV・Excel・JSONファイル（.csv, .xlsx, .json）が見つかりませんでした。");
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
    <NodeWrap id={id} data={data} title="Auto Folder" col={isDark ? "text-indigo-400" : "text-indigo-600"} showTgt={false} summary={summary} helpText="指定したフォルダを監視し、その中の『最も新しく更新されたCSV・Excel・JSONファイル』を自動で読み込みます。毎月の売上データ追加など、定期的な更新作業に便利です。">
      {data.needsUpload ? (
        <div className="space-y-3">
          <div className={`text-[10px] ${isDark ? 'text-white bg-indigo-500/20 border-indigo-500/50' : 'text-gray-800 bg-indigo-50 border-indigo-200'} flex items-center gap-2 p-2 rounded border`}>
            <span className={`${isDark ? 'text-indigo-400' : 'text-indigo-500'} flex items-center justify-center`}>{Icons.Warning}</span> Missing: {data.folderName}
          </div>
          <button onClick={triggerClick} disabled={isLoading} className={`w-full ${isDark ? 'text-indigo-400 border-indigo-500/50 hover:bg-indigo-500/20' : 'text-indigo-600 border-indigo-300 hover:bg-indigo-50'} text-[10px] border border-dashed p-3 rounded flex items-center justify-center gap-2 font-bold uppercase transition-colors shadow-sm animate-pulse disabled:opacity-50 nodrag`}>
            <span className="flex items-center justify-center">{Icons.FolderAuto}</span> フォルダを再選択
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
          <button onClick={() => setPasteEditorNode({ nodeId: id, selectionMode: false })} className={`w-full py-2 ${isDark ? 'bg-[#252526] text-indigo-400 border-[#444] hover:bg-[#333]' : 'bg-white text-indigo-600 border-gray-300 hover:bg-indigo-50'} text-[10px] font-bold rounded border uppercase tracking-widest transition-colors nodrag`}>表を編集</button>
          <button onClick={() => setRangeModalNode(id)} className={`w-full py-2 ${isDark ? 'bg-indigo-600/20 text-indigo-400 border-indigo-500/30 hover:bg-indigo-600/40' : 'bg-indigo-50 text-indigo-600 border-indigo-200 hover:bg-indigo-100'} text-[10px] font-bold rounded border uppercase tracking-widest transition-colors nodrag`}>範囲選択</button>
          <label className="flex items-center gap-2 pt-2 cursor-pointer group"><input type="checkbox" checked={data.useFirstRowAsHeader !== false} onChange={(e) => updateNodeData(id, { useFirstRowAsHeader: e.target.checked })} className="accent-indigo-500 w-4 h-4 cursor-pointer nodrag" /><span className={`text-[10px] ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-600 group-hover:text-gray-900'} font-bold uppercase transition-colors`}>1行目をヘッダーにする</span></label>
        </div>
      )}
      <input type="file" ref={inputRef} className="hidden" onChange={handleFolderSelect} {...{ webkitdirectory: "true", directory: "true" } as any} />
    </NodeWrap>
  );
});

const PasteNode = memo(({ id, data }: any) => {
  const { focusNode, theme, setPasteEditorNode } = useContext(AppContext);
  const { updateNodeData } = useReactFlow();
  const isDark = theme === 'dark';
  const matrix = sanitizeMatrix(data.tableData || parseDelimitedTextToMatrix(data.rawData || ''));
  const rowCount = matrix.length;
  const colCount = Math.max(...matrix.map((row) => row.length), 0);
  const summary = rowCount > 0 && colCount > 0 ? `${rowCount}x${colCount}` : '';

  return (
    <NodeWrap id={id} data={data} title="Paste Data" col={isDark ? "text-orange-400" : "text-orange-600"} showTgt={false} summary={summary} helpText="貼り付けたデータをそのまま表形式で編集できます。セル編集、行列追加、ヘッダー設定、範囲選択に対応します。">
      <div className="space-y-3">
        <div className={`rounded-lg border p-3 space-y-2 transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'}`}>
          <div className={`text-[10px] font-bold ${isDark ? 'text-white' : 'text-gray-800'}`}>
            {rowCount} rows / {colCount} cols
          </div>
          <div className={`text-[9px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>
            直接入力、Excelの貼り付け、セル編集、抽出範囲の設定に対応。
          </div>
          {data.ranges?.length > 0 && (
            <div className={`text-[8px] font-mono rounded border px-2 py-1 truncate ${isDark ? 'text-orange-300 bg-orange-500/10 border-orange-500/30' : 'text-orange-700 bg-orange-50 border-orange-200'}`}>
              Range: {data.ranges.join(', ')}
            </div>
          )}
        </div>
        <button
          onClick={() => setPasteEditorNode({ nodeId: id, selectionMode: false })}
          className="w-full bg-gray-800 hover:bg-gray-700 text-white text-[10px] font-bold py-2 rounded nodrag shadow-sm transition-all active:scale-95 flex items-center justify-center gap-1.5"
        >
          <span className="flex items-center justify-center">{Icons.Paste}</span> 表を編集
        </button>
        <button
          onClick={() => {
            setPasteEditorNode({ nodeId: id, selectionMode: true });
            focusNode(id, false, false, 'resize');
          }}
          className={`w-full py-2 border rounded text-[10px] font-bold uppercase tracking-widest transition-colors nodrag ${isDark ? 'bg-orange-600/20 text-orange-300 border-orange-500/30 hover:bg-orange-600/40' : 'bg-orange-50 text-orange-700 border-orange-200 hover:bg-orange-100'}`}
        >
          範囲選択
        </button>
        <label className="flex items-center gap-2 pt-1 cursor-pointer group">
          <input
            type="checkbox"
            checked={data.useFirstRowAsHeader !== false}
            onChange={(e) => updateNodeData(id, { useFirstRowAsHeader: e.target.checked })}
            className="accent-orange-500 w-4 h-4 cursor-pointer nodrag"
          />
          <span className={`text-[10px] font-bold uppercase transition-colors ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-600 group-hover:text-gray-900'}`}>1行目をヘッダーにする</span>
        </label>
      </div>
    </NodeWrap>
  )
});

const UnionNode = memo(({ id, data }: any) => {
  const { isDark } = useNodeLogic(id);
  return <NodeWrap id={id} data={data} title="Union" col={isDark ? "text-blue-400" : "text-blue-600"} multi={true} summary="Append" helpText="2つのデータを「縦」に繋ぎ合わせます。（例: 1月のデータと2月のデータを1つの表にする）"><div className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'} text-center italic tracking-widest uppercase py-2`}>Merge Vertically</div></NodeWrap>
});

const JoinNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.keyA && data.keyB ? `${data.joinType || 'INNER'} JOIN` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] hover:border-blue-400' : 'bg-white border-gray-300 hover:border-blue-500'}`;
  
  return (
    <NodeWrap id={id} data={data} title="Join" col={isDark ? "text-blue-400" : "text-blue-600"} multi={true} summary={summary} helpText="2つのデータを共通の「キー（列）」を使って「横」に繋ぎ合わせます。">
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
    <NodeWrap id={id} data={data} title="VLOOKUP" col={isDark ? "text-pink-400" : "text-pink-600"} multi={true} summary={summary} helpText="上(A)のデータの指定列をキーとして、下(B)のマスタデータを検索し、一致する行の特定列の値を新しい列として追加します。">
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
        <div className="space-y-1">
          <label className={`text-[8px] ${isDark ? 'text-[#888]' : 'text-gray-500'} uppercase tracking-widest font-bold`}>Insert Position</label>
          <select className={inputClass} value={data.insertAfterCol || LAST_COLUMN_OPTION} onChange={(e) => onChg('insertAfterCol', e.target.value)}>
            <option value={LAST_COLUMN_OPTION}>最後の列</option>
            {fData.headersA?.map((h: string) => <option key={h} value={h}>{h} の次</option>)}
          </select>
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
    <NodeWrap id={id} data={data} title="Minus" col={isDark ? "text-rose-500" : "text-rose-600"} multi={true} summary={summary} helpText="上(A)のデータから、下(B)のデータに存在するレコードを差し引いて残りのデータを抽出します。">
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

const CalculateNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.newColName ? `Add ${data.newColName}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-teal-400' : 'bg-white border-gray-300 text-gray-700 hover:border-teal-500'}`;

  return (
    <NodeWrap id={id} data={data} title="Calculate" col={isDark ? "text-teal-400" : "text-teal-600"} summary={summary} helpText="2つの列を使って計算（足し算や文字列結合など）を行い、新しい列として追加します。">
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
        <div className="space-y-1">
          <label className={`text-[8px] uppercase tracking-widest font-bold ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Insert Position</label>
          <select className={inputClass} value={data.insertAfterCol || LAST_COLUMN_OPTION} onChange={(e) => onChg('insertAfterCol', e.target.value)}>
            <option value={LAST_COLUMN_OPTION}>最後の列</option>
            {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h} の次</option>)}
          </select>
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
    <NodeWrap id={id} data={data} title="Sort" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="指定した列の値を使って、データを昇順（小さい順）または降順（大きい順）に並び替えます。">
      <div className="space-y-2">
        <select className={inputClass} value={data.sortCol || ''} onChange={(e) => onChg('sortCol', e.target.value)}><option value="">Target Column...</option>{fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}</select>
        <select className={inputClass} value={data.sortOrder || 'asc'} onChange={(e) => onChg('sortOrder', e.target.value)}><option value="asc">Ascending ↑</option><option value="desc">Descending ↓</option></select>
      </div>
    </NodeWrap>
  );
});

const TransformNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const transformOutCol = data.createNewCol && data.newColName ? data.newColName : data.targetCol;
  const summary = data.command === 'auto_number'
    ? (transformOutCol ? `AUTO NUMBER -> ${transformOutCol}` : 'AUTO NUMBER')
    : (data.targetCol || data.command === 'remove_duplicates' ? `${data.command === 'case_when' ? 'CASE WHEN' : (data.command || '...')} on ${data.targetCol || 'All'}` : '');
  
  let ph = "Parameter (ex: ',' or '100')";
  if (data.command === 'zero_padding') ph = "桁数を入力 (例: 3)";
  else if (data.command === 'substring') ph = "開始位置, 文字数 (例: 1, 3)";
  else if (data.command === 'round') ph = "小数点以下の桁数 (例: 0)";
  else if (data.command === 'mod') ph = "割る数 (例: 2)";

  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] focus:border-blue-400 hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 focus:border-blue-500 hover:border-blue-500'}`;
  const inputClassWhite = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400 hover:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500 hover:border-blue-500'}`;
  const isAutoNumber = data.command === 'auto_number';

  return (
    <NodeWrap id={id} data={data} title="Transform" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="データの内容を書き換えたり、型を変換したり、欠損値を補填したりする強力なクレンジングノードです。">
      <div className="space-y-2">
        <select className={inputClass} value={data.targetCol || ''} onChange={(e) => onChg('targetCol', e.target.value)}>
          <option value="">
            {data.command === 'remove_duplicates'
              ? '全体で重複判定 (All Columns)'
              : isAutoNumber
                ? '既存列に上書きする場合は列を選択'
                : 'Target Column...'}
          </option>
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
          <option value="auto_number">オートナンバー追加</option>
          <option value="remove_duplicates">重複行を削除 (Remove Duplicates)</option>
        </select>

        {data.command && data.command !== 'remove_duplicates' && data.command !== 'case_when' && !isAutoNumber && (
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

        {isAutoNumber && (
          <div className={`space-y-2 pt-2 border-t ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
            <select className={`${inputClassWhite} font-bold`} value={data.autoNumberMode || 'number'} onChange={(e) => onChg('autoNumberMode', e.target.value)}>
              <option value="number">通常の数字</option>
              <option value="prefix">指定文字列 + オートナンバー</option>
            </select>
            {data.autoNumberMode === 'prefix' && (
              <NodeInput
                className={inputClassWhite}
                placeholder="接頭辞 (例: NO-)"
                value={data.autoNumberPrefix || ''}
                onChange={(v: any) => onChg('autoNumberPrefix', v)}
              />
            )}
            <NodeInput
              className={inputClass}
              placeholder="桁数 (例: 3 -> 001, 002)"
              value={data.autoNumberDigits || ''}
              onChange={(v: any) => onChg('autoNumberDigits', v.replace(/[^\d]/g, ''))}
            />
            <div className={`text-[9px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>
              新しい列として追加する場合は下の設定で列名を指定、既存列へ入れる場合は一番上の列選択を使います。
            </div>
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

        {data.command && !['case_when', 'to_string', 'to_number', 'fill_zero', 'remove_duplicates', 'auto_number'].includes(data.command) && (
          <NodeInput 
            className={`${inputClassWhite} mt-2`} 
            placeholder={ph} 
            value={data.param0 || ''} 
            onChange={(v: any) => onChg('param0', v)} 
          />
        )}

        {data.command && data.command !== 'remove_duplicates' && (
          <div className={`pt-2 border-t ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
            <label className="flex items-center gap-2 cursor-pointer group mb-2">
              <input type="checkbox" checked={data.createNewCol || false} onChange={(e) => onChg('createNewCol', e.target.checked)} className="accent-blue-500 w-3 h-3 cursor-pointer nodrag" />
              <span className={`text-[9px] font-bold transition-colors ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-500 group-hover:text-gray-900'}`}>新しい列として追加する</span>
            </label>
            {data.createNewCol && (
              <div className="space-y-2">
                <NodeInput className={`w-full text-[10px] p-2 border rounded outline-none transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444] text-white focus:border-blue-400' : 'bg-white border-gray-300 text-gray-900 focus:border-blue-500'}`} placeholder="新しい列名 (例: New_Price)" value={data.newColName || ''} onChange={(v: any) => onChg('newColName', v)} />
                <select className={inputClass} value={data.insertAfterCol || LAST_COLUMN_OPTION} onChange={(e) => onChg('insertAfterCol', e.target.value)}>
                  <option value={LAST_COLUMN_OPTION}>最後の列</option>
                  {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h} の次</option>)}
                </select>
              </div>
            )}
          </div>
        )}
      </div>
    </NodeWrap>
  );
});

const FilterNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const chgVal = (v: string, t: string) => onChg('filterVal', (t === 'gt' || t === 'lt') ? v.replace(/[０-９．－]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[^0-9.-]/g, '') : v);
  const op = getCheckOperatorLabel(data.matchType);
  const summary = data.filterCol ? `${data.filterCol} ${op} ${data.filterVal || ''}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] hover:border-blue-400' : 'bg-white border-gray-300 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} data={data} title="Filter" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="指定した列の値が、入力した条件に一致する行だけを抽出して残します。（例: 売上が1000以上、など）">
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
  const selectedColumns = data.selectedColumns || [];
  const summary = data.selectedColumns?.length ? `${data.selectedColumns.length} cols selected` : '';
  const allHeaders = fData.incomingHeaders || [];

  const toggleColumn = (header: string, checked: boolean) => {
    const current = data.selectedColumns || [];
    onChg('selectedColumns', checked ? [...current, header] : current.filter((x: string) => x !== header));
  };

  const selectAllColumns = () => {
    onChg('selectedColumns', [...allHeaders]);
  };

  const clearSelectedColumns = () => {
    onChg('selectedColumns', []);
  };

  const handleScrollableWheel = (e: React.WheelEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.scrollTop += e.deltaY;
  };

  return (
    <NodeWrap id={id} data={data} title="Select" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="必要な列だけを選んで残します。列数が多い場合でも、全選択やホイールスクロールで素早く操作できます。">
      <div className="space-y-3">
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={selectAllColumns}
            disabled={allHeaders.length === 0}
            className={`px-2 py-1 rounded text-[9px] font-bold nodrag transition-colors ${isDark ? 'bg-[#1a1a1a] border border-[#444] text-[#ccc] hover:text-white hover:border-blue-400 disabled:text-[#555]' : 'bg-white border border-gray-300 text-gray-700 hover:text-gray-900 hover:border-blue-400 disabled:text-gray-400'}`}
          >
            全選択
          </button>
          <button
            type="button"
            onClick={clearSelectedColumns}
            disabled={selectedColumns.length === 0}
            className={`px-2 py-1 rounded text-[9px] font-bold nodrag transition-colors ${isDark ? 'bg-[#1a1a1a] border border-[#444] text-[#ccc] hover:text-white hover:border-rose-400 disabled:text-[#555]' : 'bg-white border border-gray-300 text-gray-700 hover:text-gray-900 hover:border-rose-400 disabled:text-gray-400'}`}
          >
            クリア
          </button>
          <div className={`ml-auto text-[8px] font-bold tracking-widest ${isDark ? 'text-[#666]' : 'text-gray-400'}`}>
            {selectedColumns.length}/{allHeaders.length}
          </div>
        </div>

        {selectedColumns.length > 0 && (
          <div>
            <div className={`text-[8px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Selected Columns</div>
            <div
              className={`max-h-28 overflow-y-auto mr-6 pr-2 space-y-1 rounded border custom-scrollbar overscroll-contain ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'}`}
              onWheel={handleScrollableWheel}
              onWheelCapture={(e) => e.stopPropagation()}
            >
              {selectedColumns.map((h: string) => (
                <div key={h} className={`flex items-center gap-1 p-1.5 rounded ${isDark ? 'text-[#ccc] hover:bg-[#333]' : 'text-gray-700 hover:bg-gray-200'}`}>
                  <span className="flex-1 truncate text-[10px] font-medium">{h}</span>
                  <button type="button" onClick={() => toggleColumn(h, false)} className={`px-1.5 h-5 rounded text-[9px] font-bold nodrag transition-colors ${isDark ? 'text-rose-300 hover:text-white hover:bg-rose-500/20' : 'text-rose-600 hover:text-rose-700 hover:bg-rose-100'}`}>x</button>
                </div>
              ))}
            </div>
          </div>
        )}

        <div>
          <div className={`text-[8px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Columns</div>
          <div
            className={`max-h-44 overflow-y-auto mr-6 pr-2 space-y-1 p-1 rounded border custom-scrollbar overscroll-contain ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'}`}
            onWheel={handleScrollableWheel}
            onWheelCapture={(e) => e.stopPropagation()}
          >
            {allHeaders.length > 0 ? allHeaders.map((h: string) => (
              <label key={h} className={`flex items-center gap-2 text-[10px] p-1.5 rounded cursor-pointer group ${isDark ? 'text-[#ccc] hover:bg-[#333]' : 'text-gray-700 hover:bg-gray-200'}`}>
                <input
                  type="checkbox"
                  checked={selectedColumns.includes(h)}
                  onChange={(e) => toggleColumn(h, e.target.checked)}
                  className="accent-blue-500 w-3 h-3 nodrag"
                />
                <span className={`truncate ${isDark ? 'group-hover:text-white' : 'group-hover:text-gray-900'}`}>{h}</span>
              </label>
            )) : <div className={`text-[9px] text-center py-4 ${isDark ? 'text-[#555]' : 'text-gray-500'}`}>Connect to input data</div>}
          </div>
        </div>
      </div>
    </NodeWrap>
  );
});

const JsonArrayNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.targetCol ? `${data.targetCol} -> ${data.valueKey || 'value'}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} data={data} title="JSON Array" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText='["json","sample","demo"] のようなJSON配列文字列が入った列を展開し、1要素1行のテーブルとして抽出します。配列要素がオブジェクトなら、そのキーを列として展開します。'>
      <div className="space-y-2">
        <select className={inputClass} value={data.targetCol || ''} onChange={(e) => onChg('targetCol', e.target.value)}>
          <option value="">JSON Array Column...</option>
          {fData.incomingHeaders?.map((h: string) => <option key={h} value={h}>{h}</option>)}
        </select>
        <NodeInput
          className={inputClass}
          placeholder="展開後の列名 (既定: value)"
          value={data.valueKey || ''}
          onChange={(v: any) => onChg('valueKey', v)}
        />
        <label className="flex items-center gap-2 cursor-pointer group">
          <input
            type="checkbox"
            checked={data.includeSourceColumns || false}
            onChange={(e) => onChg('includeSourceColumns', e.target.checked)}
            className="accent-blue-500 w-3 h-3 cursor-pointer nodrag"
          />
          <span className={`text-[9px] font-bold transition-colors ${isDark ? 'text-[#aaa] group-hover:text-white' : 'text-gray-500 group-hover:text-gray-900'}`}>
            元の行の列も残す
          </span>
        </label>
      </div>
    </NodeWrap>
  );
});

const GroupByNode = memo(({ id, data }: any) => {
  const { fData, onChg, isDark } = useNodeLogic(id);
  const summary = data.groupCol ? `By ${data.groupCol}` : '';
  const inputClass = `w-full text-[10px] p-2 border rounded outline-none transition-colors nodrag ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-300 text-gray-700 hover:border-blue-500'}`;

  return (
    <NodeWrap id={id} data={data} title="Group By" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="指定したキーでデータをグループ化し、数値を合計(SUM)したり件数をカウント(CNT)したりして集計します。">
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
    <NodeWrap id={id} data={data} title="Visualizer" col={isDark ? "text-blue-400" : "text-blue-600"} summary={summary} helpText="データをグラフとして描画します。配置したグラフは下部の「Dashboard」タブで一覧表示できます。">
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

const nodeTypesObj = { dataNode: DataNode, folderSourceNode: FolderSourceNode, pasteNode: PasteNode, unionNode: UnionNode, joinNode: JoinNode, vlookupNode: VlookupNode, minusNode: MinusNode, groupByNode: GroupByNode, sortNode: SortNode, transformNode: TransformNode, calculateNode: CalculateNode, selectNode: SelectNode, jsonArrayNode: JsonArrayNode, filterNode: FilterNode, dataCheckNode: DataCheckNode, chartNode: ChartNode };

const NodeNavigator = memo(({ tList, nodes }: { tList: any[], nodes: CustomNode[] }) => {
  const { focusNode, theme, nodeFlowData } = useContext(AppContext);
  const [isMinimized, setIsMinimized] = useState(true);
  const isDark = theme === 'dark';

  if (isMinimized) {
    return (
      <Panel position="top-left" className={`${isDark ? 'bg-[#252526]/90 border-[#444] hover:bg-[#333]' : 'bg-white/90 border-gray-200 hover:bg-gray-50'} backdrop-blur-md border rounded-xl shadow-xl z-50 m-4 ml-6 cursor-pointer transition-colors no-print`} onClick={() => setIsMinimized(false)}>
        <div className={`flex items-center gap-2 p-2 px-3 text-[10px] font-bold uppercase tracking-widest ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>
          <span className={`${isDark ? 'text-blue-400' : 'text-blue-500'} flex items-center justify-center`}>{Icons.Diamond}</span>
          <span className={isDark ? 'text-white' : 'text-gray-800'}>{nodes.length} Nodes</span>
        </div>
      </Panel>
    );
  }

  return (
    <Panel position="top-left" className={`${isDark ? 'bg-[#252526]/90 border-[#444]' : 'bg-white/90 border-gray-200'} backdrop-blur-md border p-3 rounded-xl shadow-xl max-h-[300px] overflow-y-auto custom-scrollbar flex flex-col gap-1.5 w-60 z-50 m-4 ml-6 no-print transition-colors`}>
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
        else if (n.type === 'pasteNode' && n.data.tableData?.length) subText = `${n.data.tableData.length} rows`;
        else if (n.type === 'filterNode' && n.data.filterCol) subText = `${n.data.filterCol} ${getCheckOperatorLabel(n.data.matchType)} ${n.data.filterVal || ''}`;
        else if (n.type === 'dataCheckNode') {
          const checkResult = nodeFlowData[n.id]?.checkResult;
          subText = checkResult?.isConfigured ? (checkResult.hasMatches ? `NG ${checkResult.count}件` : 'OK 0件') : 'Check Data';
        }
        else if (n.type === 'chartNode' && n.data.chartType) subText = `${n.data.chartType} chart`;
        else if (n.type === 'transformNode' && (n.data.targetCol || n.data.command === 'remove_duplicates')) subText = `${n.data.command === 'case_when' ? 'CASE WHEN' : (n.data.command || '...')} on ${n.data.targetCol || 'All'}`;
        else if (n.type === 'calculateNode' && n.data.newColName) subText = `Add ${n.data.newColName}`;
        else if (n.type === 'sortNode' && n.data.sortCol) subText = `${n.data.sortCol} ${n.data.sortOrder}`;
        else if (n.type === 'groupByNode' && n.data.groupCol) subText = `By ${n.data.groupCol}`;
        else if (n.type === 'selectNode' && n.data.selectedColumns) subText = `${n.data.selectedColumns.length} cols selected`;
        else if (n.type === 'jsonArrayNode' && n.data.targetCol) subText = `Expand ${n.data.targetCol}`;
        else if (n.type === 'joinNode' || n.type === 'unionNode' || n.type === 'minusNode') subText = "Merge Data";
        else if (n.type === 'vlookupNode' && n.data.targetCol) subText = `Add ${n.data.targetCol}`;
        else if ((n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.useFirstRowAsHeader) subText = "Setup Required";

        return (
          <button 
            key={n.id} 
            onClick={() => focusNode(n.id, true, true)} 
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
});

const FlowBuilder = () => {
  const [workbooks, setWorkbooks] = useState<Record<string, XLSX.WorkBook>>({});
  const [rangeModalNode, setRangeModalNode] = useState<string | null>(null);
  const [pasteEditorNode, setPasteEditorNode] = useState<{ nodeId: string; selectionMode: boolean } | null>(null);
  
  const { screenToFlowPosition, setCenter, getZoom, getNode } = useReactFlow();
  
  const [nodes, _setNodes, onNodesChange] = useNodesState<CustomNode>([{ id: 'n-1', type: 'dataNode', position: { x: 50, y: 150 }, data: { useFirstRowAsHeader: true } }]);
  const [edges, _setEdges, onEdgesChange] = useEdgesState<Edge>([]);
  
  const [previewTab, setPreviewTab] = useState<'table' | 'chart' | 'dashboard'>('table');
  const [isSidebarOpen, setIsSidebarOpen] = useState(() => localStorage.getItem('bi-architect-sidebar-open') !== 'false');
  const [isPreviewOpen, setIsPreviewOpen] = useState(true);
  const [isPreviewFullscreen, setIsPreviewFullscreen] = useState(false);
  const [isPreviewAutoPauseOverridden, setIsPreviewAutoPauseOverridden] = useState(false);
  const [tablePreviewRowLimit, setTablePreviewRowLimit] = useState(() => {
    const saved = Number(localStorage.getItem('bi-architect-table-preview-row-limit'));
    return Number.isFinite(saved) && saved >= 0 ? saved : 80;
  });
  const [isSaveLoadOpen, setIsSaveLoadOpen] = useState(false);
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [isSqlModalOpen, setIsSqlModalOpen] = useState(false);
  const [isAutoLayoutConfirmOpen, setIsAutoLayoutConfirmOpen] = useState(false);
  
  const [showTutorial, setShowTutorial] = useState(() => !localStorage.getItem('bi-architect-visited'));
  
  const [contextMenu, setContextMenu] = useState<{ id: string, top: number, left: number } | null>(null);
  const [paneAddMenu, setPaneAddMenu] = useState<{ top: number; left: number; position: { x: number; y: number } } | null>(null);
  const [toolboxTooltip, setToolboxTooltip] = useState<{ label: string; desc: string; left: number; top: number } | null>(null);

  const [savedFlows, setSavedFlows] = useState<any[]>([]);
  const [bottomHeight, setBottomHeight] = useState(() => {
    const saved = Number(localStorage.getItem('bi-architect-bottom-height'));
    return Number.isFinite(saved) && saved >= 100 ? saved : 300;
  });
  const [isDragging, setIsDragging] = useState(false);
  const [isNodeDragging, setIsNodeDragging] = useState(false);
  const [draggingNodeId, setDraggingNodeId] = useState<string | null>(null);
  const [isTrashHover, setIsTrashHover] = useState(false);
  const trashRef = React.useRef<HTMLDivElement | null>(null);
  const isTrashHoverRef = React.useRef(false);
  
  const [cameraFocusConfig, setCameraFocusConfig] = useState<CameraFocusConfig>(() => {
    const saved = localStorage.getItem('bi-architect-camera-focus-config');
    if (saved) {
      try {
        return { ...DEFAULT_CAMERA_FOCUS_CONFIG, ...JSON.parse(saved) };
      } catch {
        return DEFAULT_CAMERA_FOCUS_CONFIG;
      }
    }
    if (localStorage.getItem('bi-architect-auto-camera') === 'false') {
      return { move: false, delete: false, resize: false, connect: false, create: false };
    }
    return DEFAULT_CAMERA_FOCUS_CONFIG;
  });
  const [isAutoConnectNewNode, setIsAutoConnectNewNode] = useState(() => localStorage.getItem('bi-architect-auto-connect-new-node') !== 'false');
  const [showTooltips, setShowTooltips] = useState(() => localStorage.getItem('bi-architect-show-tooltips') !== 'false');
  const [previewNodeId, setPreviewNodeId] = useState<string | null>(null);
  const [isFlowReady, setIsFlowReady] = useState(false);
  const [introNodeId, setIntroNodeId] = useState<string | null>('n-1');

  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    const savedTheme = localStorage.getItem('bi-architect-theme');
    return savedTheme === 'dark' ? 'dark' : 'light';
  });
  const lastFocusedNodeIdRef = React.useRef<string | null>(null);
  const prevFocusedNodeIdRef = React.useRef<string | null>(null);
  const lastCreatedNodeIdRef = React.useRef<string | null>('n-1');
  const [isCameraFocusMenuOpen, setIsCameraFocusMenuOpen] = useState(false);
  
  useEffect(() => {
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

  useEffect(() => {
    localStorage.setItem('bi-architect-show-tooltips', String(showTooltips));
  }, [showTooltips]);

  useEffect(() => {
    localStorage.setItem('bi-architect-camera-focus-config', JSON.stringify(cameraFocusConfig));
    localStorage.setItem('bi-architect-auto-camera', String(Object.values(cameraFocusConfig).some(Boolean)));
  }, [cameraFocusConfig]);

  useEffect(() => {
    localStorage.setItem('bi-architect-auto-connect-new-node', String(isAutoConnectNewNode));
  }, [isAutoConnectNewNode]);

  useEffect(() => {
    localStorage.setItem('bi-architect-sidebar-open', String(isSidebarOpen));
  }, [isSidebarOpen]);

  useEffect(() => {
    if (!isDragging) {
      localStorage.setItem('bi-architect-bottom-height', String(bottomHeight));
    }
  }, [bottomHeight, isDragging]);

  useEffect(() => {
    localStorage.setItem('bi-architect-table-preview-row-limit', String(tablePreviewRowLimit));
  }, [tablePreviewRowLimit]);

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

  const isAnyCameraFocusEnabled = useMemo(() => Object.values(cameraFocusConfig).some(Boolean), [cameraFocusConfig]);

  const focusNode = useCallback((nodeId: string, force: boolean = false, instant: boolean = false, reason: CameraFocusReason = 'manual') => {
    if (!force) {
      if (reason === 'manual' && !isAnyCameraFocusEnabled) return;
      if (reason !== 'manual' && !cameraFocusConfig[reason]) return;
    }

    if (nodeId && nodeId !== lastFocusedNodeIdRef.current) {
      prevFocusedNodeIdRef.current = lastFocusedNodeIdRef.current;
      lastFocusedNodeIdRef.current = nodeId;
    }
    
    setTimeout(() => {
      const n = getNode(nodeId) as CustomNode | undefined;
      if (n) {
        const w = n.measured?.width || 260;
        const h = n.measured?.height || 150;
        setCenter(n.position.x + w / 2, n.position.y + h / 2, { zoom: getZoom(), duration: instant ? 0 : 1200 });
      }
    }, 50);
  }, [cameraFocusConfig, getNode, isAnyCameraFocusEnabled, setCenter, getZoom]);

  const onInit = useCallback(() => {
    setIsFlowReady(false);
    setIntroNodeId('n-1');
    setTimeout(() => {
      focusNode('n-1', true, true);
      setTimeout(() => {
        setIsFlowReady(true);
        setTimeout(() => {
          setIntroNodeId(null);
        }, 240);
      }, 140);
    }, 100);
  }, [focusNode]);

  const sanitizeFlow = useCallback((flowNodes: any[], flowEdges: any[]) => {
    const validNodes = (flowNodes || [])
      .filter((n: any) => n.type !== 'webSourceNode')
      .map((n: any) => {
        if ((n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.fileName) {
          return { ...n, data: { ...n.data, needsUpload: true } };
        }
        if (n.type === 'pasteNode') {
          const tableData = Array.isArray(n.data?.tableData) && n.data.tableData.length > 0
            ? sanitizeMatrix(n.data.tableData)
            : parseDelimitedTextToMatrix(n.data?.rawData || '');
          return {
            ...n,
            data: {
              ...n.data,
              tableData,
              rawData: matrixToDelimitedText(tableData),
              useFirstRowAsHeader: n.data?.useFirstRowAsHeader !== false
            }
          };
        }
        return n;
      });
    const validIds = new Set(validNodes.map((n: any) => n.id));
    const validEdges = (flowEdges || []).filter((e: any) => validIds.has(e.source) && validIds.has(e.target));
    return { nodes: validNodes, edges: validEdges };
  }, []);

  const handleSave = (name: string) => { const up = [...savedFlows, { id: Date.now().toString(), name, updatedAt: new Date().toLocaleString(), flow: { nodes, edges } }]; setSavedFlows(up); localStorage.setItem('bi-architect-flows', JSON.stringify(up)); };
  const handleLoad = (f: any) => { 
    const sanitized = sanitizeFlow(f.flow.nodes || [], f.flow.edges || []);
    _setNodes(sanitized.nodes); 
    _setEdges(sanitized.edges); 
    setWorkbooks({}); 
    lastCreatedNodeIdRef.current = sanitized.nodes[sanitized.nodes.length - 1]?.id || null;
  };
  const onEdgeContextMenu = useCallback((e: React.MouseEvent, edge: Edge) => { e.preventDefault(); _setEdges((eds: any) => eds.filter((e: any) => e.id !== edge.id)); }, [_setEdges]);

  const onNodeContextMenu = useCallback((e: React.MouseEvent, node: Node) => {
    e.preventDefault();
    setPaneAddMenu(null);
    setContextMenu({ id: node.id, top: e.clientY, left: e.clientX });
  }, []);

  const onPaneContextMenu = useCallback((e: MouseEvent | React.MouseEvent<Element, MouseEvent>) => {
    e.preventDefault();
    setContextMenu(null);
    const menuWidth = 320;
    const menuHeight = 250;
    const top = Math.min(e.clientY, window.innerHeight - menuHeight - 16);
    const left = Math.min(e.clientX, window.innerWidth - menuWidth - 16);
    setPaneAddMenu({
      top: Math.max(16, top),
      left: Math.max(16, left),
      position: screenToFlowPosition({ x: e.clientX, y: e.clientY }),
    });
  }, [screenToFlowPosition]);

  const handleContextDuplicate = () => {
    if (!contextMenu) return;
    const node = nodes.find(n => n.id === contextMenu.id);
    if (!node) return;
    const newNode = { ...node, id: `n-${Date.now()}`, position: { x: node.position.x + 50, y: node.position.y + 50 }, selected: true };
    _setNodes(nds => nds.map(n => ({...n, selected: false})).concat(newNode as any));
    setContextMenu(null);
  };
  const handleContextToggleCollapse = () => {
    if (!contextMenu) return;
    _setNodes(nds => nds.map(n => n.id === contextMenu.id ? {...n, data: {...n.data, isCollapsed: !(n.data as any).isCollapsed}} : n));
    setContextMenu(null);
  };
  const handleContextDelete = () => {
    if (!contextMenu) return;
    deleteNodeById(contextMenu.id);
  };

  const deleteNodeById = useCallback((nodeId: string) => {
    const prevFocused = prevFocusedNodeIdRef.current;
    const lastFocused = lastFocusedNodeIdRef.current;
    const incomingEdge = edges.find((e) => e.target === nodeId);
    const incomingSource = incomingEdge?.source;
    const fallbackId =
      (prevFocused && prevFocused !== nodeId && nodes.some((n) => n.id === prevFocused) ? prevFocused : null) ||
      (incomingSource && incomingSource !== nodeId && nodes.some((n) => n.id === incomingSource) ? incomingSource : null) ||
      (lastFocused && lastFocused !== nodeId && nodes.some((n) => n.id === lastFocused) ? lastFocused : null) ||
      (nodes.find((n) => n.id !== nodeId)?.id ?? null);

    _setNodes((nds) => nds.filter((n) => n.id !== nodeId));
    _setEdges((eds) => eds.filter((e) => e.source !== nodeId && e.target !== nodeId));
    setContextMenu((prev) => (prev?.id === nodeId ? null : prev));

    if (fallbackId) {
      if (lastFocusedNodeIdRef.current === nodeId) lastFocusedNodeIdRef.current = fallbackId;
      if (lastCreatedNodeIdRef.current === nodeId) lastCreatedNodeIdRef.current = fallbackId;
      setTimeout(() => {
        focusNode(fallbackId, false, false, 'delete');
      }, 0);
    }
  }, [_setNodes, _setEdges, edges, nodes, focusNode]);

  const isPointInTrash = useCallback((clientX: number, clientY: number) => {
    const el = trashRef.current;
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    return clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom;
  }, []);

  const handleAutoLayout = useCallback(() => {
    const levels: Record<string, number> = {};
    nodes.forEach(n => levels[n.id] = 0);
    for (let i = 0; i < nodes.length; i++) {
      edges.forEach(e => {
        if (levels[e.target] <= levels[e.source]) {
          levels[e.target] = levels[e.source] + 1;
        }
      });
    }
    const levelCounts: Record<number, number> = {};
    const newNodes = nodes.map(n => {
      const lvl = levels[n.id] || 0;
      levelCounts[lvl] = (levelCounts[lvl] || 0);
      const y = levelCounts[lvl] * 250 + 100;
      levelCounts[lvl]++;
      return { ...n, position: { x: lvl * 350 + 50, y } };
    });
    _setNodes(newNodes);
  }, [nodes, edges, _setNodes]);

  const handleAutoLayoutConfirm = useCallback(() => {
    setIsAutoLayoutConfirmOpen(true);
  }, []);

  const handleConnect = useCallback(
    (p: any) => {
      _setEdges((eds: any) =>
        addEdge({ ...p, animated: true, style: { stroke: '#38bdf8', strokeWidth: 4 } } as any, eds),
      );
      if (p?.target) focusNode(p.target, false, false, 'connect');
    },
    [_setEdges, focusNode],
  );


  const handleReset = () => {
    _setNodes([{ id: 'n-1', type: 'dataNode', position: { x: 50, y: 150 }, data: { useFirstRowAsHeader: true } }]);
    _setEdges([]);
    setWorkbooks({});
    setIsResetModalOpen(false);
    lastCreatedNodeIdRef.current = 'n-1';
    focusNode('n-1', true);
  };
  
  const handleDeleteFlow = (id: string) => {
    const updated = savedFlows.filter((f: any) => f.id !== id);
    setSavedFlows(updated);
    localStorage.setItem('bi-architect-flows', JSON.stringify(updated));
  };

  // ★ JSONエクスポート処理
  const handleExportFile = () => {
    const flow = { nodes, edges };
    const blob = new Blob([JSON.stringify(flow, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `VisualDataPrep_Flow_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ★ JSONインポート処理
  const handleImportFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const flow = JSON.parse(event.target?.result as string);
        if (flow.nodes && flow.edges) {
          const sanitized = sanitizeFlow(flow.nodes, flow.edges);
          _setNodes(sanitized.nodes);
          _setEdges(sanitized.edges);
          setWorkbooks({});
          setIsSaveLoadOpen(false);
          setTimeout(() => focusNode('n-1', true, true), 100);
        }
      } catch(err) {
        alert('ファイルの読み込みに失敗しました。');
      }
    };
    reader.readAsText(file);
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

  const calcGraphRef = React.useRef<{
    nodesSnapshot: CustomNode[];
    edgesSnapshot: Edge[];
    calcNodes: CustomNode[];
    calcEdges: Edge[];
  } | null>(null);

  const calcGraph = useMemo(() => {
    const prev = calcGraphRef.current;
    const sameNodes = !!prev &&
      nodes.length === prev.nodesSnapshot.length &&
      nodes.every((node, idx) => {
        const prevNode = prev.nodesSnapshot[idx];
        return !!prevNode && node.id === prevNode.id && node.type === prevNode.type && node.data === prevNode.data;
      });
    const sameEdges = !!prev &&
      edges.length === prev.edgesSnapshot.length &&
      edges.every((edge, idx) => {
        const prevEdge = prev.edgesSnapshot[idx];
        return !!prevEdge &&
          edge.id === prevEdge.id &&
          edge.source === prevEdge.source &&
          edge.target === prevEdge.target &&
          (edge as any).sourceHandle === (prevEdge as any).sourceHandle &&
          (edge as any).targetHandle === (prevEdge as any).targetHandle;
      });

    if (sameNodes && sameEdges) {
      return { calcNodes: prev.calcNodes, calcEdges: prev.calcEdges };
    }

    const next = { nodesSnapshot: nodes, edgesSnapshot: edges, calcNodes: nodes, calcEdges: edges };
    calcGraphRef.current = next;
    return { calcNodes: next.calcNodes, calcEdges: next.calcEdges };
  }, [nodes, edges]);

  const activePreviewId = useMemo(() => {
    if (previewNodeId && calcGraph.calcNodes.find(n => n.id === previewNodeId)) return previewNodeId;
    const term = calcGraph.calcNodes.find(n => !calcGraph.calcEdges.some(e => e.source === n.id));
    return term?.id || null;
  }, [previewNodeId, calcGraph.calcNodes, calcGraph.calcEdges]);

  const sourceDataByNodeId = useMemo<SourceDataByNodeId>(() => {
    const map: SourceDataByNodeId = {};

    calcGraph.calcNodes.forEach((node) => {
      if (node.type === 'pasteNode') {
        try {
          const tableData = Array.isArray(node.data.tableData) && node.data.tableData.length > 0
            ? node.data.tableData
            : parseDelimitedTextToMatrix(node.data.rawData || '');
          map[node.id] = extractDataFromMatrix(tableData, node.data.ranges || [], node.data.useFirstRowAsHeader !== false);
        } catch {
          map[node.id] = { data: [], headers: [] };
        }
        return;
      }

      if (node.type === 'dataNode' || node.type === 'folderSourceNode') {
        if (node.data.needsUpload) {
          map[node.id] = { data: [], headers: [] };
          return;
        }
        const wb = workbooks[node.id];
        if (!wb) {
          map[node.id] = { data: [], headers: [] };
          return;
        }
        const ws = wb.Sheets[node.data.currentSheet || wb.SheetNames[0]];
        if (!ws) {
          map[node.id] = { data: [], headers: [] };
          return;
        }
        try {
          const ranges = (node.data.ranges || []).length === 0 && ws['!ref']
            ? [XLSX.utils.encode_range(XLSX.utils.decode_range(ws['!ref']))]
            : (node.data.ranges || []);
          const cacheKey = [
            node.data.currentSheet || wb.SheetNames[0],
            node.data.useFirstRowAsHeader !== false ? 'header' : 'no-header',
            ranges.join('|'),
          ].join('::');
          map[node.id] = getCachedWorkbookExtract(wb, cacheKey, () => {
            const mat = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: true }) as any[][];
            if (!mat || mat.length === 0) return { data: [], headers: [] };
            return extractDataFromMatrix(mat, ranges, node.data.useFirstRowAsHeader !== false);
          });
        } catch {
          map[node.id] = { data: [], headers: [] };
        }
      }
    });

    return map;
  }, [calcGraph.calcNodes, workbooks]);

  const flowWorkerRef = React.useRef<Worker | null>(null);
  const flowWorkerRequestIdRef = React.useRef(0);
  const [workerFlowResult, setWorkerFlowResult] = useState<WorkerFlowResult>({
    nodeFlowData: {},
    final: { data: [], headers: [] },
    dashboardsData: [],
  });

  useEffect(() => {
    const worker = new Worker(new URL('./flowCalc.worker.ts', import.meta.url), { type: 'module' });
    flowWorkerRef.current = worker;
    worker.onmessage = (event: MessageEvent<{ requestId: number; result: WorkerFlowResult }>) => {
      const { requestId, result } = event.data;
      if (requestId !== flowWorkerRequestIdRef.current) return;
      setWorkerFlowResult(result);
    };
    return () => {
      worker.terminate();
      flowWorkerRef.current = null;
    };
  }, []);

  useEffect(() => {
    const worker = flowWorkerRef.current;
    if (!worker) return;
    const requestId = flowWorkerRequestIdRef.current + 1;
    flowWorkerRequestIdRef.current = requestId;
    worker.postMessage({
      requestId,
      payload: {
        nodes: calcGraph.calcNodes,
        edges: calcGraph.calcEdges,
        sourceDataByNodeId,
        activePreviewId,
        maxChartPoints: MAX_CHART_RENDER_POINTS,
      },
    });
  }, [calcGraph.calcNodes, calcGraph.calcEdges, sourceDataByNodeId, activePreviewId]);

  const nodeFlowData = workerFlowResult.nodeFlowData;
  const final = useMemo(() => {
    const targetNode = activePreviewId ? nodes.find((n) => n.id === activePreviewId) : null;
    return {
      ...workerFlowResult.final,
      chartConfig: targetNode?.type === 'chartNode' ? targetNode.data : null,
    };
  }, [workerFlowResult.final, nodes, activePreviewId]);
  const dashboardsData = workerFlowResult.dashboardsData;
  const chartPreviewData = useMemo(() => sampleRowsForChart(final.data), [final.data]);
  const displayedTableRows = useMemo(
    () => (tablePreviewRowLimit <= 0 ? [] : final.data.slice(0, tablePreviewRowLimit)),
    [final.data, tablePreviewRowLimit],
  );
  const previewAutoPauseReason = useMemo(() => {
    if (!isPreviewOpen || previewTab === 'dashboard') return null;
    if (previewTab === 'table' && tablePreviewRowLimit <= 0) return null;

    const totalRows = final.data.length;
    const totalCols = Math.max(1, final.headers.length);
    const totalCells = totalRows * totalCols;

    if (totalRows > PREVIEW_AUTO_PAUSE_MAX_ROWS) {
      return `結果件数が ${PREVIEW_AUTO_PAUSE_MAX_ROWS.toLocaleString()} 行を超えたため自動停止`;
    }
    if (totalCells > PREVIEW_AUTO_PAUSE_MAX_CELLS) {
      return `表示セル数が ${PREVIEW_AUTO_PAUSE_MAX_CELLS.toLocaleString()} を超えたため自動停止`;
    }
    return null;
  }, [isPreviewOpen, previewTab, tablePreviewRowLimit, final.data.length, final.headers.length]);
  const isPreviewAutoPaused = !!previewAutoPauseReason && !isPreviewAutoPauseOverridden;

  useEffect(() => {
    setIsPreviewAutoPauseOverridden(false);
  }, [activePreviewId, previewTab, final.data.length, final.headers.length, tablePreviewRowLimit]);

  const appContextValue = useMemo(() => ({
    workbooks,
    setWorkbooks,
    setRangeModalNode,
    setPasteEditorNode,
    nodeFlowData,
    showTooltips,
    focusNode,
    theme,
    activePreviewId: activePreviewId as string | null,
    introNodeId,
  }), [workbooks, nodeFlowData, showTooltips, focusNode, theme, activePreviewId, introNodeId]);
  const tablePreviewLimitOptions = useMemo(() => [
    { value: 0, label: '非表示' },
    { value: 5, label: '5件' },
    { value: 20, label: '20件' },
    { value: 50, label: '50件' },
    { value: 80, label: '80件' },
    { value: 150, label: '150件' },
    { value: 300, label: '300件' },
  ], []);

  const handleExport = (format: 'csv' | 'xlsx' | 'json') => {
    if (final.data.length === 0) return;
    if (format === 'csv') { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([[0xEF, 0xBB, 0xBF] as any, Papa.unparse(final.data)], { type: 'text/csv' })); a.download = 'export.csv'; a.click(); }
    else if (format === 'json') { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify(final.data, null, 2)], { type: 'application/json' })); a.download = 'export.json'; a.click(); }
    else if (format === 'xlsx') { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(final.data), "Result"); XLSX.writeFile(wb, "export.xlsx"); }
  };

  const tList = useMemo(() => [
    { t: 'dataNode', l: 'Source', i: Icons.Source, c: 'text-blue-500 dark:text-blue-400', desc: 'CSV・Excel・JSONファイルを選択して読み込みます。' },
    { t: 'folderSourceNode', l: 'Auto Folder', i: Icons.FolderAuto, c: 'text-indigo-500 dark:text-indigo-400', desc: '指定フォルダを監視し、中の最新CSV・Excel・JSONファイルを自動で読み込みます。' },
    { t: 'pasteNode', l: 'Paste Data', i: Icons.Paste, c: 'text-orange-500 dark:text-orange-400', desc: 'Excel等のデータを表形式で直接貼り付け・編集し、範囲選択して読み込みます。' },
    { t: 'unionNode', l: 'Union', i: Icons.Union, c: 'text-blue-500 dark:text-blue-400', desc: '2つのデータを「縦」に繋ぎ合わせます。(データの追加)' },
    { t: 'joinNode', l: 'Join', i: Icons.Join, c: 'text-blue-500 dark:text-blue-400', desc: '2つのデータを共通のキーで「横」に繋ぎます。' },
    { t: 'vlookupNode', l: 'VLOOKUP', i: Icons.Vlookup, c: 'text-pink-500 dark:text-pink-400', desc: '別データから一致するキーを検索し、特定の列の値を新しい列として追加します。' },
    { t: 'minusNode', l: 'Minus', i: Icons.Minus, c: 'text-rose-600 dark:text-rose-500', desc: '上のデータから下のデータに存在するレコードを差し引きます。' },
    { t: 'groupByNode', l: 'Group By', i: Icons.GroupBy, c: 'text-blue-500 dark:text-blue-400', desc: '指定したキーでデータをグループ化し、合計や件数を集計します。' },
    { t: 'sortNode', l: 'Sort', i: Icons.Sort, c: 'text-blue-500 dark:text-blue-400', desc: '指定した列を基準に、データを昇順・降順に並び替えます。' },
    { t: 'transformNode', l: 'Transform', i: Icons.Transform, c: 'text-blue-500 dark:text-blue-400', desc: 'データの内容を書き換えたり、型変換や0埋めなどを行うクレンジングノードです。' },
    { t: 'calculateNode', l: 'Calculate', i: Icons.Calculate, c: 'text-teal-500 dark:text-teal-400', desc: '2つの列の値を計算（足し算、文字結合など）し、新しい列として追加します。' },
    { t: 'selectNode', l: 'Select', i: Icons.Select, c: 'text-blue-500 dark:text-blue-400', desc: '必要な列(カラム)だけを選んで残し、不要な列を削除します。' },
    { t: 'jsonArrayNode', l: 'JSON Array', i: Icons.Database, c: 'text-sky-500 dark:text-sky-400', desc: '["json","sample","demo"] のようなJSON配列文字列が入った列を展開し、1要素1行のテーブルとして抽出します。' },
    { t: 'filterNode', l: 'Filter', i: Icons.Filter, c: 'text-blue-500 dark:text-blue-400', desc: '条件に一致する行だけを抽出します。(例: 売上1000以上)' },
    { t: 'dataCheckNode', l: 'Data Check', i: Icons.Warning, c: 'text-sky-500 dark:text-sky-400', desc: '指定条件に一致するデータの有無をチェックし、件数と対象行を表示します。ヒット時は赤、該当なしは青で表示します。' },
    { t: 'chartNode', l: 'Visualizer', i: Icons.Chart, c: 'text-blue-500 dark:text-blue-400', desc: 'データをグラフ化します。Dashboardタブで一覧表示できます。' }
  ], []);

  const previewNodeOptions = useMemo(() => {
    return calcGraph.calcNodes.map((n) => {
      const typeInfo = tList.find((t) => t.t === n.type);
      const label = typeInfo ? typeInfo.l : n.type;
      let sub = n.id;
      if ((n.type === 'dataNode' || n.type === 'folderSourceNode') && n.data.fileName) sub = n.data.fileName;
      else if (n.type === 'pasteNode' && n.data.tableData?.length) sub = `${n.data.tableData.length} rows`;
      return { id: n.id, label: `${label} - ${sub}` };
    });
  }, [calcGraph.calcNodes, tList]);

  const isDark = theme === 'dark';
  const btnClasses = isSidebarOpen ? 'p-3 gap-4 w-[calc(100%-0.5rem)] mr-2' : 'p-2 justify-center w-9 h-10 mr-1';
  const NEW_NODE_DROP_OFFSET = { x: 130, y: 32 };
  const TOOLBOX_TOOLTIP_WIDTH = 224;
  const sourceOnlyNodeTypes = new Set(['dataNode', 'folderSourceNode', 'pasteNode']);
  const multiInputNodeTypes = new Set(['unionNode', 'joinNode', 'minusNode', 'vlookupNode']);
  const showToolboxTooltip = useCallback((event: React.MouseEvent<HTMLDivElement>, item: { l: string; desc: string }) => {
    const rect = event.currentTarget.getBoundingClientRect();
    const left = Math.max(12, Math.min(rect.left - TOOLBOX_TOOLTIP_WIDTH - 12, window.innerWidth - TOOLBOX_TOOLTIP_WIDTH - 12));
    const top = Math.max(12, Math.min(rect.top + rect.height / 2 - 48, window.innerHeight - 120));
    setToolboxTooltip({ label: item.l, desc: item.desc, left, top });
  }, []);
  const createNodeAtPosition = useCallback((type: string, position: { x: number; y: number }) => {
    const id = `n-${Date.now()}`;
    const prevId = lastCreatedNodeIdRef.current;
    _setNodes((nds: any) => nds.concat({ id, type, position, data: { useFirstRowAsHeader: true } }));
    if (
      isAutoConnectNewNode &&
      prevId &&
      prevId !== id &&
      !sourceOnlyNodeTypes.has(type) &&
      nodes.some((n) => n.id === prevId)
    ) {
      _setEdges((eds: any) =>
        addEdge(
          {
            source: prevId,
            target: id,
            targetHandle: multiInputNodeTypes.has(type) ? 'input-a' : undefined,
            animated: true,
            style: { stroke: '#38bdf8', strokeWidth: 4 }
          } as any,
          eds,
        ),
      );
    }
    lastCreatedNodeIdRef.current = id;
    setPaneAddMenu(null);
    setContextMenu(null);
    focusNode(id, false, false, 'create');
  }, [_setNodes, _setEdges, focusNode, isAutoConnectNewNode, nodes]);

  return (
    <AppContext.Provider value={appContextValue}>
      <div className={`h-screen w-screen flex flex-col font-sans overflow-hidden transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`} onClick={() => { setContextMenu(null); setPaneAddMenu(null); setToolboxTooltip(null); setIsCameraFocusMenuOpen(false); }}>
        <GlobalStyle />
        <div className={`border-b px-6 py-3 flex justify-between items-center z-40 gap-4 no-print transition-colors ${isDark ? 'bg-[#181818] border-[#333] shadow-md' : 'bg-white border-gray-200 shadow-sm'}`}>
          <a href="#/" className={`text-[13px] font-bold tracking-[0.5em] uppercase flex items-center gap-3 shrink-0 hover:opacity-80 transition-opacity ${isDark ? 'text-white' : 'text-gray-800'}`} title="Back to Home">
            <span className="text-blue-500 w-4 h-4 flex items-center justify-center">{Icons.Diamond}</span>
            Visual Data Prep
          </a>
          
          <div className="flex items-center gap-3 shrink-0">
            <button onClick={handleAutoLayoutConfirm} className={`text-[10px] p-2 rounded-lg font-bold uppercase tracking-widest flex items-center justify-center shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-[#666]' : 'bg-white hover:bg-gray-50 border-gray-200 text-gray-600'}`} title="整列">
              <span className="flex items-center justify-center text-lg">{Icons.Layout}</span>
            </button>

            <button onClick={() => setShowTutorial(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-2 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-[#666]' : 'bg-white hover:bg-gray-50 border-gray-200 text-gray-600'}`} title="Tutorial">
              <span className="flex items-center justify-center text-lg">{Icons.Help}</span>
            </button>

            <button onClick={() => setShowTooltips(!showTooltips)} className={`text-[10px] px-3 py-2 rounded-lg font-bold tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444]' : 'bg-white hover:bg-gray-50 border-gray-200'} ${showTooltips ? (isDark ? 'text-blue-400' : 'text-blue-500') : (isDark ? 'text-[#666]' : 'text-gray-500')}`} title="Toggle Tooltips">
              <span className="flex items-center justify-center text-lg">{Icons.Help}</span> Tooltips: {showTooltips ? 'ON' : 'OFF'}
            </button>

            <button onClick={toggleTheme} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-2 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-[#666]' : 'bg-white hover:bg-gray-50 border-gray-200 text-gray-600'}`} title="Toggle Theme">
              <span className="flex items-center justify-center text-lg">{isDark ? Icons.Sun : Icons.Moon}</span>
            </button>

            <div className="relative" onClick={(e) => e.stopPropagation()}>
              <button onClick={() => setIsCameraFocusMenuOpen((v) => !v)} className={`text-[10px] px-3 py-2 rounded-lg font-bold tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444]' : 'bg-white hover:bg-gray-50 border-gray-200'} ${isAnyCameraFocusEnabled ? (isDark ? 'text-blue-400' : 'text-blue-500') : (isDark ? 'text-[#666]' : 'text-gray-500')}`} title="Camera Focus Settings">
                <span className="flex items-center justify-center gap-1">{Icons.Focus}</span> CameraFocus
              </button>
              {isCameraFocusMenuOpen && (
                <div className={`absolute right-0 top-full mt-2 w-56 rounded-xl border shadow-2xl p-3 z-[120] ${isDark ? 'bg-[#252526] border-[#444]' : 'bg-white border-gray-200'}`}>
                  <div className={`text-[9px] font-bold uppercase tracking-widest mb-2 ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Auto Focus Cases</div>
                  {[
                    { key: 'move', label: 'ノード移動後' },
                    { key: 'delete', label: 'ノード削除後' },
                    { key: 'resize', label: 'ノード高さ変更時' },
                    { key: 'connect', label: 'ノード接続時' },
                    { key: 'create', label: 'ノード追加時' },
                  ].map((item) => (
                    <label key={item.key} className="flex items-center gap-2 py-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={cameraFocusConfig[item.key as keyof CameraFocusConfig]}
                        onChange={(e) => setCameraFocusConfig((prev) => ({ ...prev, [item.key]: e.target.checked }))}
                        className="accent-blue-500 w-3.5 h-3.5"
                      />
                      <span className={`text-[10px] font-bold ${isDark ? 'text-[#ccc]' : 'text-gray-700'}`}>{item.label}</span>
                    </label>
                  ))}
                </div>
              )}
            </div>

            <button onClick={() => setIsAutoConnectNewNode(!isAutoConnectNewNode)} className={`text-[10px] px-3 py-2 rounded-lg font-bold tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444]' : 'bg-white hover:bg-gray-50 border-gray-200'} ${isAutoConnectNewNode ? (isDark ? 'text-blue-400' : 'text-blue-500') : (isDark ? 'text-[#666]' : 'text-gray-500')}`} title="Auto Connect New Node">
              <span className="flex items-center justify-center gap-1">{Icons.Join}</span> AutoConnect: {isAutoConnectNewNode ? 'ON' : 'OFF'}
            </button>
            
            <button onClick={() => setIsSqlModalOpen(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border hidden md:flex ${isDark ? 'bg-[#252526] hover:bg-blue-900/30 border-[#444] hover:border-blue-500/50 text-[#aaa] hover:text-blue-400' : 'bg-white hover:bg-gray-50 border-gray-200 text-gray-600'}`}>
              <span className="flex items-center justify-center gap-1">{Icons.Code}</span> SQL
            </button>
            <button onClick={() => setIsResetModalOpen(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border hidden md:flex ${isDark ? 'bg-[#252526] hover:bg-red-900/30 border-[#444] hover:border-red-500/50 text-[#aaa] hover:text-red-400' : 'bg-white hover:bg-gray-50 border-gray-200 text-gray-600'}`}>
              <span className="flex items-center justify-center gap-1">{Icons.Trash}</span> RESET
            </button>
            <button onClick={() => setIsSaveLoadOpen(true)} className={`text-[10px] px-3 py-2 rounded-lg font-bold uppercase tracking-widest flex items-center gap-1.5 shadow-sm active:scale-95 transition-colors border ${isDark ? 'bg-[#252526] hover:bg-[#333] border-[#444] text-white' : 'bg-gray-800 hover:bg-gray-700 border-gray-800 text-white'}`}>
              <span className="flex items-center justify-center gap-1">{Icons.Save}</span> / <span className="flex items-center justify-center gap-1">{Icons.Folder}</span> PROJECTS
            </button>
          </div>
        </div>
        
        <div className="flex-1 flex overflow-hidden relative no-print">
          <div className={`flex-1 relative transition-opacity duration-150 ${isFlowReady ? 'opacity-100' : 'opacity-0 pointer-events-none'}`}>
            <ReactFlow 
              nodes={nodes} 
              edges={edges} 
              onInit={onInit}
              onNodesChange={onNodesChange} 
              onEdgesChange={onEdgesChange} 
              onConnect={handleConnect} 
              onEdgeContextMenu={onEdgeContextMenu} 
              onNodeContextMenu={onNodeContextMenu}
              onPaneContextMenu={onPaneContextMenu}
              onPaneClick={() => { setContextMenu(null); setPaneAddMenu(null); }}
              nodeTypes={nodeTypesObj} 
              connectionRadius={50}
              onNodeDragStart={(_, node: any) => {
                setIsNodeDragging(true);
                setDraggingNodeId(node.id);
                isTrashHoverRef.current = false;
                setIsTrashHover(false);
              }}
              onNodeDrag={(event, node: any) => {
                if (draggingNodeId !== node.id) return;
                if (!(event instanceof MouseEvent)) return;
                const nextHover = isPointInTrash(event.clientX, event.clientY);
                if (nextHover !== isTrashHoverRef.current) {
                  isTrashHoverRef.current = nextHover;
                  setIsTrashHover(nextHover);
                }
              }}
              onNodeDragStop={(event, node: any) => {
                if (event instanceof MouseEvent && isPointInTrash(event.clientX, event.clientY)) {
                  deleteNodeById(node.id);
                } else {
                  focusNode(node.id, false, false, 'move');
                }
                setIsNodeDragging(false);
                setDraggingNodeId(null);
                isTrashHoverRef.current = false;
                setIsTrashHover(false);
              }}
              onDrop={(e) => { 
                const t = e.dataTransfer.getData('application/reactflow'); 
                if (t) {
                  // New nodes feel more natural when they appear around the cursor,
                  // instead of placing their top-left corner exactly at the drop point.
                  const position = screenToFlowPosition({
                    x: e.clientX - NEW_NODE_DROP_OFFSET.x,
                    y: e.clientY - NEW_NODE_DROP_OFFSET.y
                  });
                  createNodeAtPosition(t, position);
                }
              }} 
              onDragOver={(e) => e.preventDefault()} 
              defaultViewport={{ x: 250, y: 100, zoom: 0.9 }}
            >
              <Background color={isDark ? "#333" : "#d1d5db"} gap={24} size={1} />
              <Controls className={`border fill-gray-600 ${isDark ? 'bg-[#252526] border-[#444]' : 'bg-white border-gray-200'}`} />
              <NodeNavigator tList={tList} nodes={calcGraph.calcNodes} />
            </ReactFlow>

            {isNodeDragging && (
              <div
                ref={trashRef}
                className={`pointer-events-none absolute bottom-6 left-1/2 -translate-x-1/2 z-[60] w-16 h-16 rounded-2xl border flex items-center justify-center transition-colors shadow-2xl ${
                  isTrashHover
                    ? (isDark ? 'bg-rose-600/30 border-rose-500 text-rose-300' : 'bg-rose-100 border-rose-300 text-rose-600')
                    : (isDark ? 'bg-[#252526]/90 border-[#444] text-[#aaa]' : 'bg-white/90 border-gray-200 text-gray-600')
                }`}
              >
                <span className="flex items-center justify-center text-2xl">{Icons.Trash}</span>
              </div>
            )}

            {contextMenu && (
              <div 
                style={{ top: contextMenu.top, left: contextMenu.left }} 
                className={`fixed z-[1000] border rounded-xl shadow-2xl p-1.5 w-48 font-bold text-[11px] flex flex-col gap-0.5 transition-colors ${isDark ? 'bg-[#252526] border-[#444]' : 'bg-white border-gray-200'}`}
                onClick={(e) => e.stopPropagation()}
              >
                <button onClick={handleContextDuplicate} className={`w-full text-left px-3 py-2.5 rounded-lg flex items-center gap-2 transition-colors ${isDark ? 'text-[#ccc] hover:bg-[#333] hover:text-white' : 'text-gray-700 hover:bg-gray-100 hover:text-gray-900'}`}><span className="w-4 h-4 flex items-center justify-center">{Icons.Copy}</span> 複製 (Duplicate)</button>
                <button onClick={handleContextToggleCollapse} className={`w-full text-left px-3 py-2.5 rounded-lg flex items-center gap-2 transition-colors ${isDark ? 'text-[#ccc] hover:bg-[#333] hover:text-white' : 'text-gray-700 hover:bg-gray-100 hover:text-gray-900'}`}><span className="w-4 h-4 flex items-center justify-center">{Icons.ChevronDown}</span> 最小化 / 展開</button>
                <button onClick={handleContextDelete} className={`w-full text-left px-3 py-2.5 rounded-lg flex items-center gap-2 transition-colors ${isDark ? 'text-rose-400 hover:bg-rose-500/20' : 'text-rose-600 hover:bg-rose-50'}`}><span className="w-4 h-4 flex items-center justify-center">{Icons.Trash}</span> 削除 (Delete)</button>
              </div>
            )}

            {paneAddMenu && (
              <div
                style={{ top: paneAddMenu.top, left: paneAddMenu.left }}
                className={`fixed z-[1001] border rounded-2xl shadow-2xl p-3 w-[320px] transition-colors ${
                  isDark ? 'bg-[#252526]/98 border-[#444]' : 'bg-white/98 border-gray-200'
                }`}
                onClick={(e) => e.stopPropagation()}
              >
                <div className={`flex items-center justify-between mb-3 px-1 ${isDark ? 'text-[#aaa]' : 'text-gray-600'}`}>
                  <span className="text-[10px] font-bold tracking-widest uppercase">Add Node</span>
                  <span className="text-[9px]">空白を右クリック</span>
                </div>
                <div className="grid grid-cols-4 gap-2">
                  {tList.map((item) => (
                    <button
                      key={item.t}
                      type="button"
                      title={item.l}
                      onClick={() => createNodeAtPosition(item.t, paneAddMenu.position)}
                      className={`rounded-xl border p-2 flex flex-col items-center justify-center gap-1.5 transition-all active:scale-95 ${
                        isDark
                          ? 'bg-[#1a1a1a] border-[#333] hover:border-blue-500 hover:bg-[#222]'
                          : 'bg-gray-50 border-gray-200 hover:border-blue-400 hover:bg-white'
                      }`}
                    >
                      <span className={`text-lg flex items-center justify-center ${item.c}`}>{item.i}</span>
                      <span className={`text-[8px] font-bold leading-tight text-center ${isDark ? 'text-[#bbb]' : 'text-gray-700'}`}>{item.l}</span>
                    </button>
                  ))}
                </div>
              </div>
            )}
          </div>
          
          <aside className={`border-l z-20 flex flex-col transition-all duration-300 ease-in-out ${isDark ? 'bg-[#181818] border-[#333]' : 'bg-white border-gray-200'} ${isSidebarOpen ? 'w-64 py-4 pr-4 pl-2' : 'w-16 py-4 px-2 items-center'}`}>
            <div className={`flex items-center ${isSidebarOpen ? 'justify-between mb-4 pl-2' : 'justify-center mb-6'} border-b pb-2 ${isDark ? 'border-[#333]' : 'border-gray-200'}`}>
              <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className={`p-1 rounded transition-colors flex items-center justify-center w-6 h-6 ${isDark ? 'text-[#888] hover:text-white hover:bg-[#333]' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-100'}`}>
                {isSidebarOpen ? Icons.ArrowRight : Icons.ArrowLeft}
              </button>
              {isSidebarOpen && <div className={`text-[10px] font-bold tracking-[0.3em] uppercase ${isDark ? 'text-white' : 'text-gray-800'}`}>Toolbox</div>}
            </div>
            <div className={`flex flex-col ${isSidebarOpen ? 'gap-3 pl-2 pr-1' : 'gap-4 pl-1 pr-1 w-full items-center'} overflow-y-auto overflow-x-visible pb-10 custom-scrollbar relative`}>
              {tList.map(item => {
                const hoverBorderClass = isDark ? (item.t === 'calculateNode' ? 'hover:border-teal-500' : item.t === 'vlookupNode' ? 'hover:border-pink-500' : item.t === 'minusNode' ? 'hover:border-rose-500' : item.t === 'folderSourceNode' ? 'hover:border-indigo-500' : item.t === 'pasteNode' ? 'hover:border-orange-500' : 'hover:border-blue-500') 
                                              : (item.t === 'calculateNode' ? 'hover:border-teal-400' : item.t === 'vlookupNode' ? 'hover:border-pink-400' : item.t === 'minusNode' ? 'hover:border-rose-400' : item.t === 'folderSourceNode' ? 'hover:border-indigo-400' : item.t === 'pasteNode' ? 'hover:border-orange-400' : 'hover:border-blue-400');
                return (
                <div
                  key={item.t}
                  className={`relative group/btn rounded-xl cursor-grab flex items-center transition-all shadow-sm active:scale-95 border ${isDark ? 'bg-[#252526] border-[#333]' : 'bg-white border-gray-200'} ${hoverBorderClass} ${btnClasses}`}
                  onDragStart={(e) => e.dataTransfer.setData('application/reactflow', item.t)}
                  onMouseEnter={(e) => showToolboxTooltip(e, item)}
                  onMouseLeave={() => setToolboxTooltip((prev) => (prev?.label === item.l ? null : prev))}
                  draggable
                >
                  <div className={`${item.c} text-lg group-hover/btn:scale-125 transition-transform flex items-center justify-center ${isSidebarOpen ? '' : 'text-xl'}`}>{item.i}</div>
                  {isSidebarOpen && <span className={`text-[10px] font-bold uppercase tracking-wider truncate ${isDark ? 'text-[#888] group-hover/btn:text-white' : 'text-gray-600 group-hover/btn:text-gray-900'}`}>{item.l}</span>}
                </div>
              )})}
            </div>
          </aside>
        </div>

        {showTooltips && toolboxTooltip && createPortal(
          <div
            style={{ left: toolboxTooltip.left, top: toolboxTooltip.top }}
            className={`fixed w-56 text-[11px] p-3 rounded-lg border pointer-events-none z-[99999] shadow-2xl normal-case leading-relaxed ${isDark ? 'bg-[#111] text-[#ccc] border-[#555]' : 'bg-white text-gray-700 border-gray-300'}`}
          >
            <div className={`font-bold mb-1 tracking-widest ${isDark ? 'text-white' : 'text-gray-900'}`}>{toolboxTooltip.label}</div>
            {toolboxTooltip.desc}
          </div>,
          document.body
        )}
        
        <div 
          style={isPreviewFullscreen ? {} : { height: isPreviewOpen ? bottomHeight : 48, transition: isDragging ? 'none' : 'height 0.3s cubic-bezier(0.4, 0, 0.2, 1)' }} 
          className={isPreviewFullscreen 
            ? `fixed inset-0 z-[500] flex flex-col transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-white'}` 
            : `flex flex-col border-t z-30 relative transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#333] shadow-[0_-10px_40px_rgba(0,0,0,0.5)]' : 'bg-white border-gray-200 shadow-[0_-10px_40px_rgba(0,0,0,0.1)]'}`}
        >
          {isPreviewOpen && !isPreviewFullscreen && (
            <div 
              onMouseDown={startResize} 
              className={`absolute top-[-8px] left-0 w-full h-4 cursor-row-resize z-50 transition-colors no-print ${isDark ? 'hover:bg-blue-500/30' : 'hover:bg-blue-400/30'}`} 
              title="Drag to resize"
            />
          )}
          <div className={`px-6 py-2 flex justify-between items-center border-b h-[48px] shrink-0 no-print transition-colors ${isDark ? 'bg-[#252526] border-[#333]' : 'bg-gray-50 border-gray-200'}`}>
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
                <div className={`flex items-center gap-2 mr-2 border-r pr-4 ${isDark ? 'border-[#444]' : 'border-gray-300'}`}>
                  <span className={`text-[10px] font-bold uppercase tracking-widest ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Preview:</span>
                  <select 
                    className={`text-[10px] p-1.5 border rounded outline-none transition-colors cursor-pointer w-40 truncate ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-200 text-gray-700 hover:border-blue-500'}`}
                    value={previewNodeId || ''}
                    onChange={(e) => setPreviewNodeId(e.target.value || null)}
                  >
                    <option value="">Auto (末端のノード)</option>
                    {previewNodeOptions.map((option) => (
                      <option key={option.id} value={option.id}>{option.label}</option>
                    ))}
                  </select>
                </div>
              )}

              {isPreviewOpen && previewTab === 'table' && (
                <div className={`flex items-center gap-2 mr-2 border-r pr-4 ${isDark ? 'border-[#444]' : 'border-gray-300'}`}>
                  <span className={`text-[10px] font-bold uppercase tracking-widest ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Rows:</span>
                  <select
                    className={`text-[10px] p-1.5 border rounded outline-none transition-colors cursor-pointer w-24 ${isDark ? 'bg-[#1a1a1a] border-[#444] text-[#ccc] hover:border-blue-400' : 'bg-white border-gray-200 text-gray-700 hover:border-blue-500'}`}
                    value={tablePreviewRowLimit}
                    onChange={(e) => setTablePreviewRowLimit(Number(e.target.value))}
                  >
                    {tablePreviewLimitOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
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
              {previewTab !== 'dashboard' && (
                <span className={`text-[10px] font-bold ml-4 ${isDark ? 'text-blue-400' : 'text-blue-600'}`}>
                  {previewTab === 'table' && tablePreviewRowLimit > 0
                    ? `${displayedTableRows.length}/${final.data.length} rows`
                    : `${final.data.length} rows`}
                </span>
              )}
              {isPreviewOpen && previewTab !== 'dashboard' && isPreviewAutoPaused && (
                <span className={`text-[10px] font-bold uppercase tracking-widest ${isDark ? 'text-amber-400' : 'text-amber-600'}`}>
                  Auto Paused
                </span>
              )}
              {isPreviewOpen && (
                <button onClick={() => setIsPreviewFullscreen(!isPreviewFullscreen)} className={`p-1.5 ml-2 rounded transition-colors flex items-center justify-center w-7 h-7 ${isDark ? 'text-[#888] hover:text-white hover:bg-[#333]' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-200'}`} title={isPreviewFullscreen ? "元のサイズに戻す" : "全画面表示"}>
                  <span className="w-4 h-4 flex items-center justify-center">{isPreviewFullscreen ? Icons.Minimize : Icons.Maximize}</span>
                </button>
              )}
            </div>
          </div>
          {isPreviewOpen && (
            <div className={`flex-1 overflow-auto print-preview-area transition-colors ${isDark ? 'bg-[#1e1e1e]' : 'bg-white'}`}>
              {isPreviewAutoPaused ? (
                <div className="h-full flex items-center justify-center p-8">
                  <div className={`max-w-md w-full rounded-2xl border p-6 text-center shadow-xl ${isDark ? 'bg-[#181818] border-[#333]' : 'bg-gray-50 border-gray-200'}`}>
                    <div className={`text-[11px] font-bold uppercase tracking-[0.3em] mb-3 ${isDark ? 'text-amber-400' : 'text-amber-600'}`}>Preview Auto Paused</div>
                    <div className={`text-sm font-bold mb-2 ${isDark ? 'text-white' : 'text-gray-900'}`}>{previewAutoPauseReason}</div>
                    <div className={`text-[11px] leading-relaxed mb-4 ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>
                      Rows を減らすか、必要な時だけ一時表示してください。
                    </div>
                    <div className="flex items-center justify-center gap-3">
                      <button
                        onClick={() => setIsPreviewAutoPauseOverridden(true)}
                        className={`text-[10px] px-4 py-2 rounded-lg font-bold uppercase tracking-widest transition-colors shadow-sm ${isDark ? 'bg-amber-500/20 text-amber-300 hover:bg-amber-500/30' : 'bg-amber-100 text-amber-700 hover:bg-amber-200'}`}
                      >
                        一時表示
                      </button>
                      {previewTab === 'table' && (
                        <button
                          onClick={() => setTablePreviewRowLimit(5)}
                          className={`text-[10px] px-4 py-2 rounded-lg font-bold uppercase tracking-widest transition-colors shadow-sm ${isDark ? 'bg-[#252526] text-[#ccc] hover:bg-[#333]' : 'bg-white text-gray-700 hover:bg-gray-100 border border-gray-200'}`}
                        >
                          5件表示へ変更
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              ) : previewTab === 'table' ? (
                tablePreviewRowLimit <= 0 ? (
                  <div className={`h-full flex items-center justify-center text-[11px] italic tracking-widest uppercase ${isDark ? 'text-[#555]' : 'text-gray-400'}`}>
                    Data Table Hidden
                  </div>
                ) : (
                  <table className="w-full text-left text-[11px] whitespace-nowrap border-collapse">
                    <thead className={`sticky top-0 border-b z-10 shadow-sm transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#333]' : 'bg-gray-50 border-gray-200'}`}>
                      <tr>{final.headers.map((h, i) => <th key={i} className={`px-5 py-3 font-bold border-r uppercase tracking-wider ${isDark ? 'text-[#888] border-[#333]' : 'text-gray-600 border-gray-200'}`}>{h}</th>)}</tr>
                    </thead>
                    <tbody>
                      {displayedTableRows.map((row, i) => (
                        <tr key={i} className={`transition-colors border-b ${isDark ? 'hover:bg-[#252526] border-[#222]' : 'hover:bg-gray-50 border-gray-200'}`}>
                          {final.headers.map((h, j) => <td key={j} className={`px-5 py-2 border-r font-mono ${isDark ? 'text-[#ccc] border-[#222]' : 'text-gray-800 border-gray-200'}`}>{row[h]}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )
              ) : previewTab === 'chart' ? (
                <div style={{ width: '100%', height: '100%', minHeight: '300px', padding: '24px' }}>
                  {final.chartConfig?.xAxis && final.chartConfig?.yAxis ? (
                    <ResponsiveContainer width="100%" height="100%" minWidth={10} minHeight={10}>
                      {final.chartConfig.chartType === 'line' ? (
                        <LineChart data={chartPreviewData}>
                          <CartesianGrid strokeDasharray="3 3" stroke={isDark ? "#333" : "#e5e7eb"} vertical={false} />
                          <XAxis dataKey={final.chartConfig.xAxis} stroke={isDark ? "#555" : "#9ca3af"} tick={{ fill: isDark ? '#888' : '#6b7280', fontSize: 11 }} tickLine={false} axisLine={false} dy={10} />
                          <YAxis stroke={isDark ? "#555" : "#9ca3af"} tick={{ fill: isDark ? '#888' : '#6b7280', fontSize: 11 }} tickLine={false} axisLine={false} dx={-10} />
                          <Tooltip contentStyle={{ backgroundColor: isDark ? '#1a1a1a' : '#fff', border: isDark ? '1px solid #444' : '1px solid #e5e7eb', fontSize: '11px', borderRadius: '8px' }} labelStyle={{ color: isDark ? '#fff' : '#1f2937', fontWeight: 'bold', paddingBottom: '4px' }} itemStyle={{ color: '#60a5fa' }} />
                          <Line type="monotone" dataKey={final.chartConfig.yAxis} stroke="#3b82f6" strokeWidth={3} dot={{ r: 4, fill: '#3b82f6', strokeWidth: 0 }} activeDot={{ r: 6 }} />
                        </LineChart>
                      ) : (
                        <BarChart data={chartPreviewData}>
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
              ) : previewTab === 'dashboard' ? (
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
              ) : null}
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

              <div className={`p-8 flex flex-col items-center space-y-6 transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-white'}`}>
                <div className={`text-[11px] leading-relaxed text-center ${isDark ? 'text-[#aaa]' : 'text-gray-600'}`}>
                  ”ノード”を繋いで視覚的にデータを作る、データ整形ツールです。<br/><br/>基本的な使い方は以下の3ステップです。
                </div>
                
                <div className="w-full space-y-4 text-left">
                  <div className={`p-4 rounded-xl border flex items-center gap-4 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] hover:border-blue-500/50' : 'bg-gray-50 border-gray-200 hover:border-gray-400'}`}>
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg shrink-0 border transition-colors ${isDark ? 'bg-[#252526] text-blue-400 border-[#444]' : 'bg-white text-gray-800 border-gray-200 shadow-sm'}`}>
                      <span className="w-5 h-5 flex items-center justify-center">{Icons.Source}</span>
                    </div>
                    <div className="flex-1">
                      <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-white' : 'text-gray-800'}`}>Step 1: Add Nodes</div>
                      <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>右の<strong>Toolbox</strong>からドラッグ＆ドロップするか、空白キャンバスを<strong>右クリック</strong>してノード一覧から追加します。</p>
                    </div>
                  </div>

                  <div className={`p-4 rounded-xl border flex items-center gap-4 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] hover:border-pink-500/50' : 'bg-gray-50 border-gray-200 hover:border-gray-400'}`}>
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg shrink-0 border transition-colors ${isDark ? 'bg-[#252526] text-pink-400 border-[#444]' : 'bg-white text-gray-800 border-gray-200 shadow-sm'}`}>
                      <span className="w-5 h-5 flex items-center justify-center">{Icons.Join}</span>
                    </div>
                    <div className="flex-1">
                      <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-white' : 'text-gray-800'}`}>Step 2: Connect Flow</div>
                      <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>ノード同士の<strong>青い○（ハンドル）</strong>をマウスで繋ぐと、データが左から右へと流れて処理されます。<strong>AutoConnect</strong> がONなら新規ノード追加時に直前ノードへ自動接続されます。</p>
                    </div>
                  </div>

                  <div className={`p-4 rounded-xl border flex items-center gap-4 transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444] hover:border-emerald-500/50' : 'bg-gray-50 border-gray-200 hover:border-gray-400'}`}>
                    <div className={`w-10 h-10 rounded-full flex items-center justify-center text-lg shrink-0 border transition-colors ${isDark ? 'bg-[#252526] text-emerald-400 border-[#444]' : 'bg-white text-gray-800 border-gray-200 shadow-sm'}`}>
                      <span className="w-5 h-5 flex items-center justify-center">{Icons.Dashboard}</span>
                    </div>
                    <div className="flex-1">
                      <div className={`text-[10px] font-bold uppercase tracking-widest mb-1 ${isDark ? 'text-white' : 'text-gray-800'}`}>Step 3: Preview & Export</div>
                      <p className={`text-[10px] leading-relaxed ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>画面下部の<strong>Preview</strong>に処理結果がリアルタイムで表示されます。Excel出力やSQL変換も可能です。</p>
                    </div>
                  </div>
                </div>
              </div>
              <button onClick={closeTutorial} className={`w-full p-4 text-[11px] font-bold uppercase tracking-widest flex items-center justify-center gap-2 transition-colors border-t ${isDark ? 'bg-[#252526] text-blue-400 border-[#444] hover:bg-[#333]' : 'bg-gray-800 text-white border-gray-800 hover:bg-gray-700'}`}>
                <span className="w-4 h-4 flex items-center justify-center">{Icons.Diamond}</span> 使ってみる
              </button>
            </div>
          </div>
        )}

        <SqlModal 
          isOpen={isSqlModalOpen} 
          onClose={() => setIsSqlModalOpen(false)} 
          nodes={nodes} 
          edges={edges} 
          onImport={(n: CustomNode[], e: Edge[]) => { _setNodes(n); _setEdges(e); setWorkbooks({}); lastCreatedNodeIdRef.current = n[n.length - 1]?.id || null; }} 
        />

        <SaveLoadModal 
          isOpen={isSaveLoadOpen} 
          onClose={() => setIsSaveLoadOpen(false)} 
          onSave={handleSave} 
          onLoad={handleLoad} 
          onDelete={handleDeleteFlow} 
          flows={savedFlows} 
          onExportFile={handleExportFile}
          onImportFile={handleImportFile}
        />

        {isAutoLayoutConfirmOpen && (
          <div className={`fixed inset-0 z-[300] flex items-center justify-center p-8 backdrop-blur-sm no-print ${isDark ? 'bg-black/80' : 'bg-gray-900/50'}`}>
            <div className={`border rounded-2xl shadow-2xl w-[360px] p-6 text-center space-y-6 ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
              <div className="space-y-2">
                <h3 className={`text-sm font-bold tracking-widest ${isDark ? 'text-white' : 'text-gray-900'}`}>Visual Data Prep</h3>
                <p className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>ノードを整列しますか？</p>
              </div>
              <div className="flex gap-3">
                <button onClick={() => setIsAutoLayoutConfirmOpen(false)} className={`flex-1 py-3 rounded-xl text-[10px] font-bold uppercase tracking-widest transition-colors ${isDark ? 'bg-[#333] hover:bg-[#444] text-white' : 'bg-gray-200 hover:bg-gray-300 text-gray-800'}`}>いいえ</button>
                <button onClick={() => { setIsAutoLayoutConfirmOpen(false); handleAutoLayout(); }} className="flex-1 py-3 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-[10px] font-bold uppercase tracking-widest shadow-xl transition-all">はい</button>
              </div>
            </div>
          </div>
        )}

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

        <PasteTableEditorModal
          isOpen={!!pasteEditorNode}
          onClose={() => setPasteEditorNode(null)}
          node={nodes.find((n) => n.id === pasteEditorNode?.nodeId)}
          initialSelectionMode={pasteEditorNode?.selectionMode}
          onApply={(nextData: any) => {
            if (!pasteEditorNode) return;
            const { workbook, ...nodeData } = nextData;
            if (workbook) {
              setWorkbooks((prev) => ({ ...prev, [pasteEditorNode.nodeId]: workbook }));
            }
            _setNodes((nds: any[]) => nds.map((n: any) => n.id === pasteEditorNode.nodeId ? { ...n, data: { ...n.data, ...nodeData } } : n));
          }}
        />
        {rangeModalNode && <RangeSelectorModal isOpen={true} onClose={() => setRangeModalNode(null)} workbook={workbooks[rangeModalNode]} currentSheet={nodes.find(n => n.id === rangeModalNode)?.data.currentSheet} initialRanges={nodes.find(n => n.id === rangeModalNode)?.data.ranges} initialUseHeader={nodes.find(n => n.id === rangeModalNode)?.data.useFirstRowAsHeader} onRangesConfirm={(r: string[], h: boolean) => { _setNodes((nds: any[]) => nds.map((n: any) => n.id === rangeModalNode ? { ...n, data: { ...n.data, ranges: r, useFirstRowAsHeader: h } } : n)); setRangeModalNode(null); }} />}
      </div>
    </AppContext.Provider>
  );
};

const PasteTableEditorModal = ({ isOpen, onClose, node, onApply, initialSelectionMode = false }: any) => {
  const { theme, workbooks } = useContext(AppContext);
  const isDark = theme === 'dark';
  const [tableData, setTableData] = useState<string[][]>(createEmptyMatrix());
  const [useHeader, setUseHeader] = useState(true);
  const [ranges, setRanges] = useState<string[]>([]);
  const [importText, setImportText] = useState('');
  const [selectionMode, setSelectionMode] = useState(false);
  const [colWidth, setColWidth] = useState(120);
  const [dragStart, setDragStart] = useState<{ r: number; c: number } | null>(null);
  const [dragEnd, setDragEnd] = useState<{ r: number; c: number } | null>(null);

  useEffect(() => {
    if (!isOpen) return;
    const isWorkbookNode = node?.type === 'dataNode' || node?.type === 'folderSourceNode';
    const workbook = node?.id ? workbooks[node.id] : null;
    const currentSheet = node?.data?.currentSheet || workbook?.SheetNames?.[0];
    const baseMatrix = isWorkbookNode
      ? sanitizeMatrix(
          workbook && currentSheet && workbook.Sheets[currentSheet]
            ? (XLSX.utils.sheet_to_json(workbook.Sheets[currentSheet], { header: 1, defval: "", blankrows: true }) as any[][])
            : createEmptyMatrix()
        )
      : (Array.isArray(node?.data?.tableData) && node.data.tableData.length > 0
          ? sanitizeMatrix(node.data.tableData)
          : parseDelimitedTextToMatrix(node?.data?.rawData || ''));
    setTableData(baseMatrix);
    setUseHeader(node?.data?.useFirstRowAsHeader !== false);
    setRanges(node?.data?.ranges || []);
    setImportText('');
    setSelectionMode(!!initialSelectionMode);
    setDragStart(null);
    setDragEnd(null);
  }, [isOpen, node, initialSelectionMode, workbooks]);

  if (!isOpen || !node) return null;

  const isWorkbookNode = node.type === 'dataNode' || node.type === 'folderSourceNode';
  const workbook = isWorkbookNode ? workbooks[node.id] : null;
  const currentSheet = node.data?.currentSheet || workbook?.SheetNames?.[0];
  const normalizedMatrix = sanitizeMatrix(tableData);
  const colCount = Math.max(...normalizedMatrix.map((row) => row.length), 0);
  const cols = Array.from({ length: colCount }, (_, idx) => XLSX.utils.encode_col(idx));

  const updateCell = (r: number, c: number, value: string) => {
    setTableData((prev) => {
      const next = sanitizeMatrix(prev);
      next[r][c] = value;
      return next;
    });
  };

  const addRow = () => {
    setTableData((prev) => [...sanitizeMatrix(prev), Array.from({ length: colCount || 4 }, () => '')]);
  };

  const addColumn = () => {
    setTableData((prev) => sanitizeMatrix(prev).map((row) => [...row, '']));
  };

  const clearTable = () => {
    setTableData(createEmptyMatrix());
    setRanges([]);
    setDragStart(null);
    setDragEnd(null);
  };

  const applyImportText = () => {
    const parsed = parseDelimitedTextToMatrix(importText);
    setTableData(parsed);
    setRanges([]);
    setDragStart(null);
    setDragEnd(null);
  };

  const isSelected = (r: number, c: number) => (
    dragStart &&
    dragEnd &&
    c >= Math.min(dragStart.c, dragEnd.c) &&
    c <= Math.max(dragStart.c, dragEnd.c) &&
    r >= Math.min(dragStart.r, dragEnd.r) &&
    r <= Math.max(dragStart.r, dragEnd.r)
  );

  const pushRange = () => {
    if (!dragStart || !dragEnd) return;
    const range = XLSX.utils.encode_range({
      s: { c: Math.min(dragStart.c, dragEnd.c), r: Math.min(dragStart.r, dragEnd.r) },
      e: { c: Math.max(dragStart.c, dragEnd.c), r: Math.max(dragStart.r, dragEnd.r) }
    });
    setRanges((prev) => prev.includes(range) ? prev : [...prev, range]);
    setDragStart(null);
    setDragEnd(null);
  };

  const saveTable = () => {
    const nextMatrix = sanitizeMatrix(tableData);
    if (isWorkbookNode && workbook && currentSheet) {
      const nextWorkbook = {
        ...workbook,
        Sheets: {
          ...workbook.Sheets,
          [currentSheet]: XLSX.utils.aoa_to_sheet(nextMatrix)
        }
      } as XLSX.WorkBook;
      onApply({
        workbook: nextWorkbook,
        ranges,
        useFirstRowAsHeader: useHeader
      });
    } else {
      onApply({
        tableData: nextMatrix,
        rawData: matrixToDelimitedText(nextMatrix),
        ranges,
        useFirstRowAsHeader: useHeader
      });
    }
    onClose();
  };

  return (
    <div className={`fixed inset-0 z-[350] flex items-center justify-center p-8 backdrop-blur-md no-print ${isDark ? 'bg-black/90' : 'bg-gray-900/50'}`}>
      <div className={`border rounded-3xl shadow-2xl w-full h-full flex flex-col overflow-hidden transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
        <div className={`p-4 border-b flex justify-between items-center transition-colors ${isDark ? 'bg-[#1a1a1a] border-[#444]' : 'bg-gray-50 border-gray-200'}`}>
          <div className="space-y-1">
            <h2 className={`text-[12px] font-bold uppercase tracking-[0.35em] flex items-center gap-2 ${isDark ? 'text-white' : 'text-gray-800'}`}>
              <span className="text-orange-500 w-4 h-4 flex items-center justify-center">{Icons.Paste}</span>
              {isWorkbookNode ? `${node.type === 'dataNode' ? 'Source' : 'Auto Folder'} Editor` : 'Paste Data Editor'}
            </h2>
            <p className={`text-[10px] ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>
              {isWorkbookNode ? '現在のシート内容を直接編集できます。保存するとこのノードの読み込み結果に反映されます。' : 'セルを直接編集し、必要に応じて範囲選択で抽出対象を絞り込みます。'}
            </p>
          </div>
          <button onClick={onClose} className={`transition-colors text-xl flex items-center justify-center w-6 h-6 ${isDark ? 'text-[#666] hover:text-white' : 'text-gray-500 hover:text-gray-800'}`}>
            <span className="w-4 h-4 block flex items-center justify-center">{Icons.Close}</span>
          </button>
        </div>

        <div className="flex-1 flex overflow-hidden">
          <div className={`flex-1 flex flex-col transition-colors ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
            <div className={`p-4 border-b space-y-3 transition-colors ${isDark ? 'border-[#333]' : 'border-gray-200'}`}>
              <div className="flex flex-wrap gap-2">
                <button onClick={addRow} className={`px-3 py-2 rounded-lg text-[10px] font-bold border transition-colors ${isDark ? 'bg-[#252526] border-[#444] text-white hover:bg-[#333]' : 'bg-white border-gray-300 text-gray-800 hover:bg-gray-100'}`}>行追加</button>
                <button onClick={addColumn} className={`px-3 py-2 rounded-lg text-[10px] font-bold border transition-colors ${isDark ? 'bg-[#252526] border-[#444] text-white hover:bg-[#333]' : 'bg-white border-gray-300 text-gray-800 hover:bg-gray-100'}`}>列追加</button>
                <button onClick={() => setSelectionMode((prev) => !prev)} className={`px-3 py-2 rounded-lg text-[10px] font-bold border transition-colors ${selectionMode ? 'bg-orange-600 text-white border-orange-500' : (isDark ? 'bg-[#252526] border-[#444] text-[#ccc] hover:bg-[#333]' : 'bg-white border-gray-300 text-gray-800 hover:bg-gray-100')}`}>{selectionMode ? '選択モード中' : '範囲選択モード'}</button>
                <button onClick={clearTable} className={`px-3 py-2 rounded-lg text-[10px] font-bold border transition-colors ${isDark ? 'bg-[#252526] border-[#444] text-rose-300 hover:bg-rose-500/20' : 'bg-white border-gray-300 text-rose-600 hover:bg-rose-50'}`}>表をクリア</button>
              </div>
              <div className="flex flex-col gap-2">
                <textarea
                  value={importText}
                  onChange={(e) => setImportText(e.target.value)}
                  placeholder="ExcelやCSVの内容をここに貼り付けてから、下のボタンで表へ反映できます。"
                  className={`w-full h-20 text-[11px] p-3 border rounded-xl outline-none resize-none custom-scrollbar transition-colors ${isDark ? 'bg-[#111] border-[#333] text-[#ddd] focus:border-orange-400' : 'bg-white border-gray-300 text-gray-800 focus:border-orange-500'}`}
                />
                <button onClick={applyImportText} disabled={!importText.trim()} className="self-start bg-orange-600 hover:bg-orange-500 disabled:opacity-30 text-white px-4 py-2 rounded-lg text-[10px] font-bold uppercase shadow-sm transition-all">
                  貼り付け内容を表に反映
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-auto custom-scrollbar">
              <table className="border-collapse table-fixed min-w-full">
                <thead className={`sticky top-0 z-10 transition-colors ${isDark ? 'bg-[#252526]' : 'bg-gray-100'}`}>
                  <tr>
                    <th className={`w-12 border-b ${isDark ? 'border-[#444]' : 'border-gray-300'}`}></th>
                    {cols.map((col) => (
                      <th key={col} style={{ width: colWidth }} className={`px-3 py-1.5 border-r border-b text-[10px] font-bold ${isDark ? 'border-[#444] text-[#888]' : 'border-gray-300 text-gray-600'}`}>{col}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {normalizedMatrix.map((row, r) => (
                    <tr key={r}>
                      <td className={`w-12 border-r border-b text-center text-[10px] font-mono ${isDark ? 'bg-[#252526] border-[#444] text-[#666]' : 'bg-gray-100 border-gray-300 text-gray-500'}`}>{r + 1}</td>
                      {row.map((value, c) => (
                        <td
                          key={`${r}-${c}`}
                          style={{ width: colWidth }}
                          onMouseDown={selectionMode ? () => { setDragStart({ r, c }); setDragEnd({ r, c }); } : undefined}
                          onMouseOver={selectionMode ? () => dragStart && setDragEnd({ r, c }) : undefined}
                          onMouseUp={selectionMode ? pushRange : undefined}
                          className={`border-r border-b align-top ${selectionMode ? 'cursor-crosshair select-none' : ''} ${isSelected(r, c) ? (isDark ? 'bg-orange-500/40 border-orange-400' : 'bg-orange-100 border-orange-300') : (isDark ? 'border-[#2a2a2a]' : 'border-gray-200')}`}
                        >
                          {selectionMode ? (
                            <div className={`px-2 py-1.5 h-[34px] truncate text-[11px] ${isDark ? 'text-[#ddd] hover:bg-[#222]' : 'text-gray-800 hover:bg-gray-100'}`}>
                              {value}
                            </div>
                          ) : (
                            <input
                              type="text"
                              value={value}
                              onChange={(e) => updateCell(r, c, e.target.value)}
                              className={`w-full px-2 py-1.5 text-[11px] outline-none nodrag transition-colors ${isDark ? 'bg-transparent text-white focus:bg-[#222]' : 'bg-white text-gray-800 focus:bg-orange-50'}`}
                            />
                          )}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className={`w-80 border-l p-6 flex flex-col gap-5 transition-colors ${isDark ? 'bg-[#252526] border-[#444]' : 'bg-white border-gray-200'}`}>
            <h3 className={`text-[10px] font-bold uppercase border-b pb-3 ${isDark ? 'text-[#888] border-[#444]' : 'text-gray-600 border-gray-200'}`}>Editor Settings</h3>
            <div className="space-y-3">
              <label className="flex items-center gap-3 cursor-pointer group">
                <input type="checkbox" checked={useHeader} onChange={(e) => setUseHeader(e.target.checked)} className="w-4 h-4 accent-orange-500" />
                <span className={`text-[11px] uppercase font-bold transition-colors ${isDark ? 'text-[#ccc] group-hover:text-white' : 'text-gray-700 group-hover:text-gray-900'}`}>1行目をヘッダーにする</span>
              </label>
              <div className="space-y-2">
                <div className={`text-[9px] font-bold uppercase ${isDark ? 'text-[#888]' : 'text-gray-500'}`}>Column Width</div>
                <input type="range" min="80" max="240" value={colWidth} onChange={(e) => setColWidth(Number(e.target.value))} className="w-full accent-orange-500" />
              </div>
              <div className={`text-[10px] leading-relaxed rounded-xl border p-3 ${isDark ? 'bg-[#1a1a1a] border-[#333] text-[#aaa]' : 'bg-gray-50 border-gray-200 text-gray-600'}`}>
                {selectionMode ? 'セルをドラッグすると範囲を追加できます。' : 'セルを直接編集できます。Excelの貼り付けは上部の入力欄を使います。'}
              </div>
            </div>

            <div className="flex-1 overflow-auto space-y-3 pr-2 custom-scrollbar">
              <div className="flex items-center justify-between">
                <div className={`text-[10px] font-bold uppercase ${isDark ? 'text-white' : 'text-gray-800'}`}>Selected Ranges</div>
                {ranges.length > 0 && (
                  <button onClick={() => setRanges([])} className={`text-[9px] font-bold transition-colors ${isDark ? 'text-rose-300 hover:text-rose-200' : 'text-rose-600 hover:text-rose-700'}`}>全削除</button>
                )}
              </div>
              {ranges.length === 0 && <div className={`text-[10px] ${isDark ? 'text-[#666]' : 'text-gray-400'}`}>未設定なら表全体を対象にします。</div>}
              {ranges.map((range, idx) => (
                <div key={range} className={`p-3 rounded-xl border flex justify-between items-center ${isDark ? 'bg-[#1a1a1a] border-[#444]' : 'bg-gray-50 border-gray-200'}`}>
                  <div className="flex flex-col">
                    <span className={`text-[8px] font-bold uppercase tracking-widest ${isDark ? 'text-[#555]' : 'text-gray-500'}`}>Range {idx + 1}</span>
                    <span className={`text-[11px] font-mono font-bold ${isDark ? 'text-orange-300' : 'text-orange-700'}`}>{range}</span>
                  </div>
                  <button onClick={() => setRanges((prev) => prev.filter((_, rangeIdx) => rangeIdx !== idx))} className={`transition-colors flex items-center justify-center w-5 h-5 ${isDark ? 'text-[#555] hover:text-white' : 'text-gray-400 hover:text-gray-800'}`}>
                    <span className="w-4 h-4 flex items-center justify-center">{Icons.Close}</span>
                  </button>
                </div>
              ))}
            </div>

            <div className="flex gap-3">
              <button onClick={onClose} className={`flex-1 py-3 rounded-xl text-[10px] font-bold uppercase tracking-widest transition-colors ${isDark ? 'bg-[#333] hover:bg-[#444] text-white' : 'bg-gray-200 hover:bg-gray-300 text-gray-800'}`}>Close</button>
              <button onClick={saveTable} className="flex-1 py-3 bg-orange-600 hover:bg-orange-500 text-white rounded-xl text-[10px] font-bold uppercase tracking-widest shadow-xl transition-all">Apply</button>
            </div>
          </div>
        </div>
      </div>
    </div>
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
        } else if (n.data.command === 'auto_number') {
          const outCol = (n.data.createNewCol && n.data.newColName) ? n.data.newColName : n.data.targetCol;
          if (outCol) {
            const digits = Math.max(0, Number(n.data.autoNumberDigits) || 0);
            const prefix = String(n.data.autoNumberPrefix || '');
            const rowNumExpr = digits > 0
              ? `LPAD(CAST(ROW_NUMBER() OVER () AS CHAR), ${digits}, '0')`
              : `CAST(ROW_NUMBER() OVER () AS CHAR)`;
            const func = n.data.autoNumberMode === 'prefix'
              ? `CONCAT('${prefix}', ${rowNumExpr})`
              : (digits > 0 ? rowNumExpr : `ROW_NUMBER() OVER ()`);
            select = select === "*" ? `*, ${func} AS \`${outCol}\`` : `${select}, ${func} AS \`${outCol}\``;
          }
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
             const outCol = (n.data.createNewCol && n.data.newColName) ? n.data.newColName : n.data.targetCol;
             select = select === "*" ? `*, ${func} AS \`${outCol}\`` : `${select}, ${func} AS \`${outCol}\``;
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

const SaveLoadModal = ({ isOpen, onClose, onSave, onLoad, onDelete, flows, onExportFile, onImportFile }: any) => {
  const { theme } = useContext(AppContext);
  const isDark = theme === 'dark';
  const [tab, setTab] = useState<'load' | 'save'>('load'); 
  const [sName, setSName] = useState('');
  const sortedFlows = useMemo(() => [...flows].sort((a, b) => Number(b.id) - Number(a.id)), [flows]);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onImportFile(e.target.files[0]);
    }
  };

  if (!isOpen) return null;
  return (
    <div className={`fixed inset-0 z-[200] flex items-center justify-center p-8 backdrop-blur-md no-print ${isDark ? 'bg-black/90' : 'bg-gray-900/50'}`}>
      <div className={`border rounded-2xl shadow-2xl w-[500px] overflow-hidden transition-colors ${isDark ? 'bg-[#1e1e1e] border-[#444]' : 'bg-white border-gray-200'}`}>
        <div className={`flex border-b ${isDark ? 'border-[#444]' : 'border-gray-200'}`}>
          <button onClick={() => setTab('load')} className={`flex-1 p-3 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'load' ? (isDark ? 'bg-[#252526] text-blue-400 border-b-2 border-blue-400' : 'bg-gray-50 text-blue-600 border-b-2 border-blue-600') : (isDark ? 'text-[#666] hover:bg-[#252526]' : 'text-gray-500 hover:bg-gray-50')}`}><span>{Icons.Folder}</span> Load</button>
          <button onClick={() => setTab('save')} className={`flex-1 p-3 text-[11px] font-bold uppercase tracking-widest transition-colors flex items-center justify-center gap-2 ${tab === 'save' ? (isDark ? 'bg-[#252526] text-white border-b-2 border-white' : 'bg-gray-50 text-gray-900 border-b-2 border-gray-900') : (isDark ? 'text-[#666] hover:bg-[#252526]' : 'text-gray-500 hover:bg-gray-50')}`}><span>{Icons.Save}</span> Save</button>
        </div>
        <div className={`p-6 h-[300px] overflow-y-auto custom-scrollbar transition-colors flex flex-col ${isDark ? 'bg-[#1a1a1a]' : 'bg-gray-50'}`}>
          {tab === 'load' ? (
            <div className="flex-1 flex flex-col">
              <div className="flex gap-2 mb-4 shrink-0">
                 <button onClick={onExportFile} className={`flex-1 py-2 text-[10px] font-bold rounded-lg border transition-colors ${isDark ? 'bg-[#333] border-[#444] text-[#ccc] hover:bg-[#444] hover:text-white' : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>JSONを出力</button>
                 <button onClick={() => fileInputRef.current?.click()} className={`flex-1 py-2 text-[10px] font-bold rounded-lg border transition-colors ${isDark ? 'bg-[#333] border-[#444] text-[#ccc] hover:bg-[#444] hover:text-white' : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'}`}>JSONを読込</button>
                 <input type="file" accept=".json" className="hidden" ref={fileInputRef} onChange={handleFileChange} />
              </div>
              <div className="space-y-3 overflow-y-auto flex-1 custom-scrollbar pr-2">
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
              <span className="text-blue-500 w-4 h-4 flex items-center justify-center">{Icons.Select}</span> 抽出範囲の選択
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
            抽出したいデータの範囲を<strong>マウスでドラッグして選択</strong>してください。（複数の範囲を選択することも可能です）<br/>
            選択後、右側のパネルで<strong>「1行目をヘッダーにする」</strong>オプションをオンにすると、最初の行が列名として認識されます。
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
                    <span className="w-4 h-4 flex items-center justify-center">{Icons.Close}</span>
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

export default function App() {
  const [hash, setHash] = useState(window.location.hash);

  useEffect(() => {
    const handleHashChange = () => setHash(window.location.hash);
    window.addEventListener('hashchange', handleHashChange);
    return () => window.removeEventListener('hashchange', handleHashChange);
  }, []);

  if (hash === '#/app') {
    return <ReactFlowProvider><FlowBuilder /></ReactFlowProvider>;
  }

  return <LandingPage />;
}

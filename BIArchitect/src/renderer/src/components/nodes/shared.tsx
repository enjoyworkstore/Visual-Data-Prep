import React, { createContext, memo, useContext, useState } from 'react';
import { Handle, Position, useReactFlow } from '@xyflow/react';
import type * as XLSX from 'xlsx';

export type AppContextType = {
  workbooks: Record<string, XLSX.WorkBook>;
  setWorkbooks: React.Dispatch<React.SetStateAction<Record<string, XLSX.WorkBook>>>;
  setRangeModalNode: React.Dispatch<React.SetStateAction<string | null>>;
  setPasteEditorNode: React.Dispatch<React.SetStateAction<{ nodeId: string; selectionMode: boolean } | null>>;
  nodeFlowData: Record<string, any>;
  isAutoCameraMove: boolean;
  showTooltips: boolean;
  focusNode: (id: string, force?: boolean, instant?: boolean) => void;
  theme: 'light' | 'dark';
  activePreviewId: string | null;
};

export const AppContext = createContext<AppContextType>({} as AppContextType);

export const useNodeLogic = (id: string) => {
  const { nodeFlowData, isAutoCameraMove, focusNode, theme, activePreviewId, showTooltips } = useContext(AppContext);
  const { updateNodeData, setNodes, setEdges, getEdges } = useReactFlow();
  const isDark = theme === 'dark';

  return {
    fData: nodeFlowData[id] || { incomingHeaders: [], headersA: [], headersB: [] },
    isDark,
    activePreviewId,
    showTooltips,
    onChg: (k: string, v: any) => {
      updateNodeData(id, { [k]: v });
      if (['command', 'joinType', 'chartType', 'matchType', 'aggType', 'groupCol', 'sortCol', 'targetCol', 'filterCol', 'xAxis', 'yAxis', 'applyCond', 'fetchCol', 'colA', 'colB', 'operator', 'newColName', 'createNewCol', 'checkCol', 'checkType'].includes(k)) {
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

const TgtHandle = ({ id, style }: { id?: string; style?: React.CSSProperties }) => (
  <Handle type="target" position={Position.Left} id={id} style={style} className="react-flow__handle z-10 -ml-2 nodrag" />
);

const SrcHandle = () => (
  <Handle type="source" position={Position.Right} className="react-flow__handle z-10 -mr-2 nodrag" />
);

const ChevronDownIcon = (
  <svg viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" className="w-[1em] h-[1em]">
    <polyline points="6 9 12 15 18 9" />
  </svg>
);

const ChevronUpIcon = (
  <svg viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" className="w-[1em] h-[1em]">
    <polyline points="18 15 12 9 6 15" />
  </svg>
);

export type NodeStatusTone = '' | 'red' | 'blue';

type NodeWrapProps = {
  id: string;
  data: any;
  title: string;
  col: string;
  children: React.ReactNode;
  showTgt?: boolean;
  multi?: boolean;
  summary?: string;
  helpText?: string;
  statusTone?: NodeStatusTone;
};

const statusBorderClassMap: Record<Exclude<NodeStatusTone, ''>, { dark: string; light: string }> = {
  red: {
    dark: 'border-rose-500 shadow-[0_0_18px_rgba(244,63,94,0.25)]',
    light: 'border-rose-400 shadow-[0_0_18px_rgba(244,63,94,0.18)]',
  },
  blue: {
    dark: 'border-sky-500 shadow-[0_0_18px_rgba(56,189,248,0.22)]',
    light: 'border-sky-400 shadow-[0_0_18px_rgba(56,189,248,0.16)]',
  },
};

const headerStatusClassMap: Record<Exclude<NodeStatusTone, ''>, { dark: string; light: string }> = {
  red: {
    dark: 'bg-rose-950/40 border-rose-500/40',
    light: 'bg-rose-50 border-rose-200',
  },
  blue: {
    dark: 'bg-sky-950/40 border-sky-500/40',
    light: 'bg-sky-50 border-sky-200',
  },
};

export const NodeWrap = memo(({
  id,
  data,
  title,
  col,
  children,
  showTgt = true,
  multi = false,
  summary = '',
  helpText = '',
  statusTone = '',
}: NodeWrapProps) => {
  const { isDark, activePreviewId, showTooltips } = useNodeLogic(id);
  const [showHelp, setShowHelp] = useState(false);
  const { updateNodeData } = useReactFlow();

  const isCollapsed = data?.isCollapsed || false;
  const isPreview = activePreviewId === id;
  const statusBorderClass = statusTone
    ? statusBorderClassMap[statusTone][isDark ? 'dark' : 'light']
    : (isDark ? 'border-[#444]' : 'border-gray-300');
  const headerStatusClass = statusTone
    ? headerStatusClassMap[statusTone][isDark ? 'dark' : 'light']
    : (isDark ? 'bg-[#1a1a1a] border-[#444]' : 'bg-gray-50 border-gray-300');
  const borderClass = isPreview
    ? `border-blue-500 ring-2 ring-blue-500 shadow-[0_0_15px_rgba(59,130,246,0.5)] ${isDark ? 'ring-offset-[#1a1a1a]' : 'ring-offset-gray-50'} ring-offset-2`
    : statusBorderClass;

  return (
    <div className={`${isDark ? 'bg-[#252526]' : 'bg-white'} border ${borderClass} rounded-xl shadow-2xl min-w-[260px] pb-1 relative group/node transition-colors`}>
      <div
        className={`${headerStatusClass} p-2 border-b flex justify-between items-center rounded-t-xl select-none group/header cursor-pointer transition-colors`}
        onDoubleClick={() => updateNodeData(id, { isCollapsed: !isCollapsed })}
      >
        <div className="flex items-center gap-2">
          <button
            onClick={(e) => { e.stopPropagation(); updateNodeData(id, { isCollapsed: !isCollapsed }); }}
            className={`${isDark ? 'text-[#888] hover:text-white' : 'text-gray-500 hover:text-gray-800'} transition-colors flex items-center justify-center w-4 h-4 nodrag`}
            title={isCollapsed ? '展開する' : '最小化する'}
          >
            {isCollapsed ? ChevronDownIcon : ChevronUpIcon}
          </button>
          <span className={`text-[10px] font-bold tracking-widest uppercase ${col}`}>{title}</span>
          {helpText && showTooltips && (
            <div className="relative flex items-center">
              <button
                onClick={(e) => { e.preventDefault(); e.stopPropagation(); setShowHelp(!showHelp); }}
                className={`text-[10px] flex items-center justify-center w-4 h-4 rounded-full border nodrag transition-colors ${showHelp ? 'bg-blue-500 text-white border-blue-500' : (isDark ? 'text-[#888] hover:text-white border-[#555] bg-[#222]' : 'text-gray-500 hover:text-gray-800 border-gray-300 bg-gray-100')}`}
              >
                ?
              </button>
              {showHelp && (
                <div className={`absolute left-0 top-full mt-2 w-56 text-[11px] p-3 rounded-lg border opacity-0 group-hover/node:opacity-100 pointer-events-none transition-opacity z-[999] shadow-2xl normal-case tracking-normal leading-relaxed ${isDark ? 'bg-[#111] text-[#ccc] border-[#555]' : 'bg-white text-gray-700 border-gray-300'}`}>
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

NodeWrap.displayName = 'NodeWrap';

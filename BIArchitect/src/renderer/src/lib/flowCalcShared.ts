import type { Edge, Node } from '@xyflow/react';
import { matchesCondition } from '../components/nodes/dataCheckUtils';

export type CustomNode = Node<Record<string, any>>;
export type CalcDataResult = { data: any[]; headers: string[] };
export type SourceDataByNodeId = Record<string, CalcDataResult>;
export type WorkerFlowPayload = {
  nodes: CustomNode[];
  edges: Edge[];
  sourceDataByNodeId: SourceDataByNodeId;
  activePreviewId: string | null;
  maxChartPoints?: number;
};
export type WorkerFlowResult = {
  nodeFlowData: Record<string, any>;
  final: CalcDataResult;
  dashboardsData: Array<{ id: string; config: any; data: any[] }>;
};

const LAST_COLUMN_OPTION = '__last__';
const DEFAULT_MAX_CHART_POINTS = 1000;

const stringifyJsonCell = (value: unknown): string | number | boolean | null => {
  if (value == null) return null;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  return JSON.stringify(value);
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);

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
  includeSourceColumns: boolean = false,
): CalcDataResult => {
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

const sampleRowsForChart = (rows: any[], maxPoints: number): any[] => {
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
  sourceDataByNodeId: SourceDataByNodeId;
};

const createCalcRuntime = (nodes: CustomNode[], edges: Edge[], sourceDataByNodeId: SourceDataByNodeId): CalcRuntime => {
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
    sourceDataByNodeId,
  };
};

const calcData = (nId: string, runtime: CalcRuntime): CalcDataResult => {
  const cached = runtime.resultCache.get(nId);
  if (cached) return cached;

  const node = runtime.nodesById.get(nId);
  if (!node) return { data: [], headers: [] };

  const sourceData = runtime.sourceDataByNodeId[nId];
  if (sourceData && (node.type === 'pasteNode' || node.type === 'dataNode' || node.type === 'folderSourceNode')) {
    runtime.resultCache.set(nId, sourceData);
    return sourceData;
  }

  let result: CalcDataResult = { data: [], headers: [] };

  if (node.type === 'unionNode' || node.type === 'joinNode' || node.type === 'minusNode' || node.type === 'vlookupNode') {
    const eA = runtime.inputAByTarget.get(nId);
    const eB = runtime.inputBByTarget.get(nId);
    if (!eA || !eB) {
      runtime.resultCache.set(nId, result);
      return result;
    }
    const rA = calcData(eA.source, runtime);
    const rB = calcData(eB.source, runtime);

    if (node.type === 'unionNode') {
      result = { data: [...rA.data, ...rB.data], headers: rA.headers };
      runtime.resultCache.set(nId, result);
      return result;
    }

    if (node.type === 'minusNode') {
      const { keyA, keyB } = node.data;
      if (!keyA || !keyB) {
        runtime.resultCache.set(nId, rA);
        return rA;
      }
      const bKeys = new Set(rB.data.map((b) => String(b[keyB as string])));
      result = { data: rA.data.filter((a) => !bKeys.has(String(a[keyA as string]))), headers: rA.headers };
      runtime.resultCache.set(nId, result);
      return result;
    }

    if (node.type === 'vlookupNode') {
      const { keyA, keyB, fetchCol, targetCol } = node.data;
      if (!keyA || !keyB || !fetchCol || !targetCol) {
        runtime.resultCache.set(nId, rA);
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
      runtime.resultCache.set(nId, result);
      return result;
    }

    const { keyA, keyB, joinType = 'inner' } = node.data;
    if (!keyA || !keyB) {
      runtime.resultCache.set(nId, rA);
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
    runtime.resultCache.set(nId, result);
    return result;
  }

  const inEdge = runtime.primaryInputByTarget.get(nId);
  if (!inEdge) {
    runtime.resultCache.set(nId, result);
    return result;
  }
  const input = calcData(inEdge.source, runtime);
  let out = [...input.data];
  let h = [...input.headers];

  if (node.type === 'sortNode') {
    const { sortCol, sortOrder } = node.data;
    if (sortCol) {
      out.sort((a, b) =>
        sortOrder === 'desc'
          ? String(b[sortCol as string]).localeCompare(String(a[sortCol as string]), undefined, { numeric: true })
          : String(a[sortCol as string]).localeCompare(String(b[sortCol as string]), undefined, { numeric: true }),
      );
    }
  }

  if (node.type === 'filterNode') {
    const { filterCol, filterVal, matchType = 'includes' } = node.data;
    if (filterCol && filterVal !== undefined && filterVal !== '') {
      out = out.filter((r) => matchesCondition(r, filterCol as string, filterVal, matchType));
    }
  }

  if (node.type === 'dataCheckNode') {
    result = { data: out, headers: h };
    runtime.resultCache.set(nId, result);
    return result;
  }

  if (node.type === 'selectNode') {
    const sel = node.data.selectedColumns || [];
    if (sel.length > 0) {
      h = sel;
      out = out.map((r) => {
        const nr: any = {};
        sel.forEach((c: string) => {
          nr[c] = r[c];
        });
        return nr;
      });
    }
  }

  if (node.type === 'jsonArrayNode') {
    const { targetCol, valueKey = 'value', includeSourceColumns = false } = node.data;
    result = !targetCol ? { data: out, headers: h } : expandJsonArrayRows(out, targetCol, valueKey, includeSourceColumns);
    runtime.resultCache.set(nId, result);
    return result;
  }

  if (node.type === 'groupByNode') {
    const { groupCol, aggCol, aggType } = node.data;
    if (groupCol && aggCol && out.length > 0) {
      const grps: Record<string, any> = {};
      out.forEach((r) => {
        const k = r[groupCol as string];
        if (!grps[k]) grps[k] = { [groupCol as string]: k, _v: 0, _c: 0 };
        grps[k]._v += Number(r[aggCol as string]) || 0;
        grps[k]._c++;
      });
      out = Object.values(grps).map((g: any) => ({
        [groupCol as string]: g[groupCol as string],
        [aggCol as string]: aggType === 'count' ? g._c : g._v,
      }));
      h = [groupCol, aggCol];
    }
  }

  if (node.type === 'transformNode') {
    const { targetCol, command, param0, applyCond, condCol, condOp, condVal } = node.data;
    if (command === 'remove_duplicates') {
      const seen = new Set();
      out = out.filter((r) => {
        const key = targetCol ? String(r[targetCol as string]) : JSON.stringify(r);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    } else if (command === 'auto_number') {
      const outCol = node.data.createNewCol && node.data.newColName ? node.data.newColName : targetCol;
      if (outCol) {
        const mode = node.data.autoNumberMode || 'number';
        const prefix = String(node.data.autoNumberPrefix || '');
        const digits = Math.max(0, Number(node.data.autoNumberDigits) || 0);
        out = out.map((r, idx) => {
          const base = idx + 1;
          const padded = digits > 0 ? String(base).padStart(digits, '0') : String(base);
          const nextVal = mode === 'prefix' ? `${prefix}${padded}` : digits > 0 ? padded : base;
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
      out = out.map((r) => {
        if (applyCond && condCol && condOp) {
          const cValCheck = String(r[condCol as string] || '').toLowerCase();
          const tValCheck = String(condVal || '').toLowerCase();
          const cNumCheck = Number(r[condCol as string]);
          const tNumCheck = Number(condVal);
          let isMatch = false;
          switch (condOp) {
            case 'exact':
              isMatch = cValCheck === tValCheck;
              break;
            case 'not':
              isMatch = cValCheck !== tValCheck;
              break;
            case 'gt':
              isMatch = !isNaN(cNumCheck) && !isNaN(tNumCheck) ? cNumCheck > tNumCheck : cValCheck > tValCheck;
              break;
            case 'lt':
              isMatch = !isNaN(cNumCheck) && !isNaN(tNumCheck) ? cNumCheck < tNumCheck : cValCheck < tValCheck;
              break;
            default:
              isMatch = cValCheck.includes(tValCheck);
          }
          if (!isMatch) return r;
        }

        let v = r[targetCol as string];
        const vStr = v === null || v === undefined ? '' : String(v);

        if (command === 'replace') v = vStr.replace(param0 || '', '');
        else if (command === 'math_mul') v = Number(vStr) * Number(param0 || 1);
        else if (command === 'add_suffix') v = vStr + (param0 || '');
        else if (command === 'concat') v = vStr + String(param0 || '');
        else if (command === 'to_string') v = vStr;
        else if (command === 'to_number') {
          const num = Number(vStr.replace(/[^0-9.-]/g, ''));
          v = isNaN(num) ? null : num;
        } else if (command === 'fill_zero') {
          if (vStr.trim() === '') v = 0;
        } else if (command === 'zero_padding') {
          v = vStr.padStart(Number(param0) || 1, '0');
        } else if (command === 'round') {
          const d = Number(param0) || 0;
          const m = Math.pow(10, d);
          v = Math.round(Number(vStr) * m) / m;
        } else if (command === 'mod') {
          const denom = Number(param0) || 1;
          v = Number(vStr) % denom;
          if (isNaN(v as number)) v = null;
        } else if (command === 'substring') {
          const params = String(param0 || '1').split(',').map((s) => Number(s.trim()));
          const start = params[0] || 1;
          const len = params.length > 1 ? params[1] : vStr.length;
          const sIdx = Math.max(0, start - 1);
          v = vStr.slice(sIdx, sIdx + len);
        } else if (command === 'case_when') {
          let match = false;
          const cwValCheck = vStr.toLowerCase();
          const cwTVal = String(node.data.cwCondVal || '').toLowerCase();
          const cwNum = Number(v);
          const cwTNum = Number(node.data.cwCondVal);
          switch (node.data.cwCondOp) {
            case 'exact':
              match = cwValCheck === cwTVal;
              break;
            case 'not':
              match = cwValCheck !== cwTVal;
              break;
            case 'gt':
              match = !isNaN(cwNum) && !isNaN(cwTNum) ? cwNum > cwTNum : cwValCheck > cwTVal;
              break;
            case 'lt':
              match = !isNaN(cwNum) && !isNaN(cwTNum) ? cwNum < cwTNum : cwValCheck < cwTVal;
              break;
            default:
              match = cwValCheck.includes(cwTVal);
          }
          v = match ? node.data.trueVal || '' : node.data.falseVal || '';
        }

        const outCol = node.data.createNewCol && node.data.newColName ? node.data.newColName : targetCol;
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
      out = out.map((r) => {
        const valA = r[colA];
        const valB = r[colB];
        let resultValue: any = null;

        if (operator === 'concat') {
          resultValue = String(valA || '') + String(valB || '');
        } else {
          const numA = Number(valA) || 0;
          const numB = Number(valB) || 0;
          if (operator === 'add') resultValue = numA + numB;
          else if (operator === 'sub') resultValue = numA - numB;
          else if (operator === 'mul') resultValue = numA * numB;
          else if (operator === 'div') resultValue = numB !== 0 ? numA / numB : null;
        }

        return { ...r, [newColName]: resultValue };
      });
      h = insertColumnAt(h, newColName, node.data.insertAfterCol);
      out = reorderRowsByHeaders(out, h);
    }
  }

  result = { data: out, headers: h };
  runtime.resultCache.set(nId, result);
  return result;
};

export const computeFlowArtifacts = (payload: WorkerFlowPayload): WorkerFlowResult => {
  const runtime = createCalcRuntime(payload.nodes, payload.edges, payload.sourceDataByNodeId);
  const nodeFlowData: Record<string, any> = {};

  payload.nodes.forEach((node) => {
    if (node.type === 'joinNode' || node.type === 'unionNode' || node.type === 'minusNode' || node.type === 'vlookupNode') {
      const eA = runtime.inputAByTarget.get(node.id);
      const eB = runtime.inputBByTarget.get(node.id);
      const hA = eA ? calcData(eA.source, runtime).headers : [];
      const hB = eB ? calcData(eB.source, runtime).headers : [];
      nodeFlowData[node.id] = { headersA: hA, headersB: hB, incomingHeaders: [...new Set([...hA, ...hB])] };
    } else {
      const inEdge = runtime.primaryInputByTarget.get(node.id);
      const incoming = inEdge ? calcData(inEdge.source, runtime) : { data: [], headers: [] };
      if (node.type === 'dataCheckNode') {
        const isConfigured = !!(node.data.checkCol && node.data.checkVal !== undefined && node.data.checkVal !== '');
        const matchedRows = isConfigured
          ? incoming.data.filter((row: any) => matchesCondition(row, node.data.checkCol, node.data.checkVal, node.data.checkType || 'includes'))
          : [];
        nodeFlowData[node.id] = {
          incomingHeaders: incoming.headers,
          checkResult: {
            count: matchedRows.length,
            rows: matchedRows.slice(0, 3),
            hasMatches: matchedRows.length > 0,
            isConfigured,
          },
        };
      } else {
        nodeFlowData[node.id] = { incomingHeaders: incoming.headers };
      }
    }
  });

  const final = payload.activePreviewId ? calcData(payload.activePreviewId, runtime) : { data: [], headers: [] };
  const maxChartPoints = payload.maxChartPoints || DEFAULT_MAX_CHART_POINTS;
  const dashboardsData = payload.nodes
    .filter((node) => node.type === 'chartNode')
    .map((node) => {
      const res = calcData(node.id, runtime);
      return { id: node.id, config: node.data, data: sampleRowsForChart(res.data, maxChartPoints) };
    });

  return { nodeFlowData, final, dashboardsData };
};

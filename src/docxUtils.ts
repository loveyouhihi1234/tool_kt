import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import {
  DOMParser,
  XMLSerializer,
  Document as XmlDocument,
  Element as XmlElement,
  Node as XmlNode,
} from '@xmldom/xmldom';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function xuLyCmtCccd(val: unknown): string {
  if (val === null || val === undefined) return '';
  const raw = String(val).trim();
  if (!raw || raw.toLowerCase() === 'nan') return '';
  let s = raw.split('.')[0].trim();
  if (s.length === 8 || s.length === 11) {
    s = '0' + s;
  }
  return s;
}

export function xuLyNgayExcel(val: unknown): string {
  if (val === null || val === undefined) return '';
  const raw = String(val).trim();
  if (!raw || raw.toLowerCase() === 'nan') return '';

  if (val instanceof Date && !isNaN(val.getTime())) {
    const d = String(val.getDate()).padStart(2, '0');
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const y = val.getFullYear();
    return `${d}/${m}/${y}`;
  }

  if (typeof val === 'number' && !isNaN(val)) {
    const baseMs = Date.UTC(1899, 11, 30);
    const dateObj = new Date(baseMs + Math.round(val * 86400000));
    const d = String(dateObj.getUTCDate()).padStart(2, '0');
    const m = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
    const y = dateObj.getUTCFullYear();
    return `${d}/${m}/${y}`;
  }

  if (typeof val === 'string') {
    const trimmed = val.trim();
    // Already DD/MM/YYYY
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(trimmed)) {
      return trimmed;
    }
    const parsed = new Date(trimmed);
    if (!isNaN(parsed.getTime()) && /^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
      const d = String(parsed.getUTCDate()).padStart(2, '0');
      const m = String(parsed.getUTCMonth() + 1).padStart(2, '0');
      const y = parsed.getUTCFullYear();
      return `${d}/${m}/${y}`;
    }
    return trimmed;
  }

  return raw;
}

export function xuLyMst(val: unknown): string {
  if (val === null || val === undefined) return '';
  let s = String(val).trim();
  if (!s || s.toLowerCase() === 'nan') return '';
  if (s.endsWith('.0')) s = s.slice(0, -2);
  if (/^\d+$/.test(s)) {
    if (s.length <= 10) return s.padStart(10, '0');
    if (s.length <= 13) {
      const s13 = s.padStart(13, '0');
      return `${s13.slice(0, 10)}-${s13.slice(10)}`;
    }
  }
  return s;
}

export function layGiaTri(row: Record<string, unknown>, keys: string[]): unknown {
  const cols = Object.keys(row);
  for (const k of keys) {
    for (const col of cols) {
      if (k === col.trim()) return row[col];
    }
  }
  for (const k of keys) {
    for (const col of cols) {
      if (col.includes(k)) return row[col];
    }
  }
  return '';
}

export function formatMoneyDot(val: unknown): string {
  if (val === null || val === undefined) return '';
  const s = String(val).trim();
  if (!s || s.toLowerCase() === 'nan') return '';
  const num = Number(s);
  if (isNaN(num)) return s;
  const rounded = Math.round(num);
  return rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

export function formatIntStr(val: unknown): string {
  if (val === null || val === undefined) return '';
  const s = String(val).trim();
  if (!s || s.toLowerCase() === 'nan') return '';
  const num = Number(s);
  if (isNaN(num)) return s;
  return Math.trunc(num).toString();
}

export function formatCurrencyApi3(val: unknown): string {
  if (val === null || val === undefined || val === '') return '0';
  const s = String(val).trim();
  if (!s || s.toLowerCase() === 'nan') return '0';
  const num = Number(s);
  if (isNaN(num)) return s;
  return Math.trunc(num).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

export function safeStr(val: unknown): string {
  if (val === null || val === undefined) return '';
  const s = String(val);
  if (s.toLowerCase() === 'nan') return '';
  return s;
}

export function cleanCodeStr(val: unknown): string {
  const s = safeStr(val).trim();
  return s.endsWith('.0') ? s.slice(0, -2) : s;
}

export interface TableCellSpec {
  text: string;
  bold?: boolean;
  align?: 'center' | 'right' | 'left';
  gridSpan?: number;
}

export interface TableSpec {
  borderMode: 'insideH' | 'full';
  useTableGridStyle?: boolean;
  lineSpacing13?: boolean;
  colCount: number;
  rows: TableCellSpec[][];
}

function buildTableElement(doc: XmlDocument, spec: TableSpec): XmlElement {
  const tbl = doc.createElementNS(W_NS, 'w:tbl');

  const tblPr = doc.createElementNS(W_NS, 'w:tblPr');
  if (spec.useTableGridStyle) {
    const tblStyle = doc.createElementNS(W_NS, 'w:tblStyle');
    tblStyle.setAttribute('w:val', 'TableGrid');
    tblPr.appendChild(tblStyle);
  }

  const tblW = doc.createElementNS(W_NS, 'w:tblW');
  tblW.setAttribute('w:w', '5000');
  tblW.setAttribute('w:type', 'pct');
  tblPr.appendChild(tblW);

  const tblBorders = doc.createElementNS(W_NS, 'w:tblBorders');
  const borderNames =
    spec.borderMode === 'insideH'
      ? ['insideH']
      : ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];

  for (const bName of borderNames) {
    const b = doc.createElementNS(W_NS, `w:${bName}`);
    b.setAttribute('w:val', 'single');
    b.setAttribute('w:sz', '4');
    b.setAttribute('w:space', '0');
    b.setAttribute('w:color', 'auto');
    tblBorders.appendChild(b);
  }
  tblPr.appendChild(tblBorders);
  tbl.appendChild(tblPr);

  const tblGrid = doc.createElementNS(W_NS, 'w:tblGrid');
  const colWidth = Math.floor(9000 / spec.colCount);
  for (let i = 0; i < spec.colCount; i++) {
    const gc = doc.createElementNS(W_NS, 'w:gridCol');
    gc.setAttribute('w:w', String(colWidth));
    tblGrid.appendChild(gc);
  }
  tbl.appendChild(tblGrid);

  for (const rowSpec of spec.rows) {
    const tr = doc.createElementNS(W_NS, 'w:tr');
    for (const cellSpec of rowSpec) {
      const tc = doc.createElementNS(W_NS, 'w:tc');
      const tcPr = doc.createElementNS(W_NS, 'w:tcPr');

      if (cellSpec.gridSpan && cellSpec.gridSpan > 1) {
        const gs = doc.createElementNS(W_NS, 'w:gridSpan');
        gs.setAttribute('w:val', String(cellSpec.gridSpan));
        tcPr.appendChild(gs);
      }

      const vAlign = doc.createElementNS(W_NS, 'w:vAlign');
      vAlign.setAttribute('w:val', 'center');
      tcPr.appendChild(vAlign);
      tc.appendChild(tcPr);

      const p = doc.createElementNS(W_NS, 'w:p');
      const pPr = doc.createElementNS(W_NS, 'w:pPr');

      const jc = doc.createElementNS(W_NS, 'w:jc');
      jc.setAttribute('w:val', cellSpec.align || 'center');
      pPr.appendChild(jc);

      const spacing = doc.createElementNS(W_NS, 'w:spacing');
      if (spec.lineSpacing13) {
        spacing.setAttribute('w:line', '312');
        spacing.setAttribute('w:lineRule', 'auto');
      } else {
        spacing.setAttribute('w:before', '120');
        spacing.setAttribute('w:after', '0');
      }
      pPr.appendChild(spacing);
      p.appendChild(pPr);

      const lines = (cellSpec.text ?? '').split('\n');
      const r = doc.createElementNS(W_NS, 'w:r');
      const rPr = doc.createElementNS(W_NS, 'w:rPr');

      const rFonts = doc.createElementNS(W_NS, 'w:rFonts');
      rFonts.setAttribute('w:ascii', 'Times New Roman');
      rFonts.setAttribute('w:hAnsi', 'Times New Roman');
      rFonts.setAttribute('w:eastAsia', 'Times New Roman');
      rFonts.setAttribute('w:cs', 'Times New Roman');
      rPr.appendChild(rFonts);

      if (cellSpec.bold) {
        const bEl = doc.createElementNS(W_NS, 'w:b');
        rPr.appendChild(bEl);
      }

      const sz = doc.createElementNS(W_NS, 'w:sz');
      sz.setAttribute('w:val', '24');
      rPr.appendChild(sz);

      const szCs = doc.createElementNS(W_NS, 'w:szCs');
      szCs.setAttribute('w:val', '24');
      rPr.appendChild(szCs);

      r.appendChild(rPr);

      lines.forEach((line, idx) => {
        if (idx > 0) {
          r.appendChild(doc.createElementNS(W_NS, 'w:br'));
        }
        const t = doc.createElementNS(W_NS, 'w:t');
        t.setAttribute('xml:space', 'preserve');
        t.textContent = line;
        r.appendChild(t);
      });

      p.appendChild(r);
      tc.appendChild(p);
      tr.appendChild(tc);
    }
    tbl.appendChild(tr);
  }

  return tbl;
}

export function renderDocxTemplate(
  templateBuffer: Buffer,
  context: Record<string, string>,
  tableReplacements: Record<string, TableSpec> = {}
): Buffer {
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    delimiters: { start: '{{', end: '}}' },
    paragraphLoop: true,
    linebreaks: true,
    parser: (tag: string) => {
      const cleanTag = tag.trim();
      return {
        get: (scope: Record<string, string>) => {
          return cleanTag in scope ? scope[cleanTag] : '';
        },
      };
    },
    nullGetter: () => '',
  });

  doc.render(context);

  const renderedZip = doc.getZip();

  if (Object.keys(tableReplacements).length > 0) {
    const docXmlStr = renderedZip.file('word/document.xml')?.asText();
    if (docXmlStr) {
      const dom = new DOMParser().parseFromString(docXmlStr, 'text/xml');
      const paragraphs = Array.from(dom.getElementsByTagNameNS(W_NS, 'p'));

      for (const [marker, tableSpec] of Object.entries(tableReplacements)) {
        for (const p of paragraphs) {
          const text = p.textContent || '';
          if (text.includes(marker)) {
            const tbl = buildTableElement(dom, tableSpec);
            p.parentNode?.insertBefore(tbl, p);
            p.parentNode?.removeChild(p);
            break;
          }
        }
      }

      const serialized = new XMLSerializer().serializeToString(dom);
      renderedZip.file('word/document.xml', serialized);
    }
  }

  return renderedZip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  }) as Buffer;
}

export function mergeDocxBuffers(buffers: Buffer[]): Buffer {
  if (buffers.length === 0) {
    throw new Error('No documents to merge');
  }
  if (buffers.length === 1) {
    return buffers[0];
  }

  const masterZip = new PizZip(buffers[0]);
  const masterXmlStr = masterZip.file('word/document.xml')?.asText();
  if (!masterXmlStr) return buffers[0];

  const parser = new DOMParser();
  const masterDom = parser.parseFromString(masterXmlStr, 'text/xml');
  const masterBody = masterDom.getElementsByTagNameNS(W_NS, 'body')[0];
  if (!masterBody) return buffers[0];

  let masterSectPr: XmlElement | null = null;
  const masterChildren = Array.from(masterBody.childNodes) as XmlNode[];
  for (let i = masterChildren.length - 1; i >= 0; i--) {
    const node = masterChildren[i] as XmlElement;
    if (node.nodeType === 1 && node.localName === 'sectPr') {
      masterSectPr = node;
      break;
    }
  }

  for (let idx = 1; idx < buffers.length; idx++) {
    const childZip = new PizZip(buffers[idx]);
    const childXmlStr = childZip.file('word/document.xml')?.asText();
    if (!childXmlStr) continue;

    const childDom = parser.parseFromString(childXmlStr, 'text/xml');
    const childBody = childDom.getElementsByTagNameNS(W_NS, 'body')[0];
    if (!childBody) continue;

    // Add page break paragraph before appending next document
    const pbPara = masterDom.createElementNS(W_NS, 'w:p');
    const pbRun = masterDom.createElementNS(W_NS, 'w:r');
    const pbBr = masterDom.createElementNS(W_NS, 'w:br');
    pbBr.setAttribute('w:type', 'page');
    pbRun.appendChild(pbBr);
    pbPara.appendChild(pbRun);

    if (masterSectPr) {
      masterBody.insertBefore(pbPara, masterSectPr);
    } else {
      masterBody.appendChild(pbPara);
    }

    const childNodes = Array.from(childBody.childNodes) as XmlNode[];
    for (const node of childNodes) {
      if (node.nodeType === 1 && (node as XmlElement).localName === 'sectPr') {
        continue;
      }
      const imported = masterDom.importNode(node, true);
      if (masterSectPr) {
        masterBody.insertBefore(imported, masterSectPr);
      } else {
        masterBody.appendChild(imported);
      }
    }
  }

  const serialized = new XMLSerializer().serializeToString(masterDom);
  masterZip.file('word/document.xml', serialized);

  return masterZip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  }) as Buffer;
}

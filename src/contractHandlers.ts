import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import {
  cleanCodeStr,
  formatCurrencyApi3,
  formatIntStr,
  formatMoneyDot,
  layGiaTri,
  mergeDocxBuffers,
  renderDocxTemplate,
  safeStr,
  TableSpec,
  xuLyCmtCccd,
  xuLyMst,
  xuLyNgayExcel,
} from './docxUtils.js';

function sanitizeFileName(name: string): string {
  return name.replace(/[/\\:]/g, '_');
}

function sumColumn(rows: Record<string, unknown>[], key: string): number {
  let total = 0;
  for (const r of rows) {
    const val = Number(r[key]);
    if (!isNaN(val)) total += val;
  }
  return total;
}

export async function handleHopDongGao(
  excelBuffer: Buffer,
  wordTemplateBuffer: Buffer
): Promise<Buffer> {
  const wb = XLSX.read(excelBuffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

  const rows = rawRows.filter((r) => {
    const v = r['sohd'];
    return v !== '' && v !== null && v !== undefined && String(v).toLowerCase() !== 'nan';
  });

  const sohdNumbers = rows
    .map((r) => parseInt(String(r['sohd']), 10))
    .filter((n) => !isNaN(n))
    .sort((a, b) => a - b);

  const sohdDau = sohdNumbers.length > 0 ? String(sohdNumbers[0]).padStart(2, '0') : '00';
  const sohdCuoi =
    sohdNumbers.length > 0 ? String(sohdNumbers[sohdNumbers.length - 1]).padStart(2, '0') : '00';

  const groups = new Map<string, Record<string, unknown>[]>();
  for (const r of rows) {
    const key = cleanCodeStr(r['sohd']);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(r);
  }

  const zip = new JSZip();
  const folder = zip.folder('ChiTiet_HopDongGao')!;
  const generatedDocs: Buffer[] = [];

  for (const [sohd, group] of groups.entries()) {
    const firstRow = group[0];
    const soHdStr = sohd.padStart(2, '0');
    const ngay = cleanCodeStr(firstRow['ngay']).padStart(2, '0');
    const thang = cleanCodeStr(firstRow['thang']).padStart(2, '0');
    const nam = cleanCodeStr(firstRow['nam']);

    const tongThue = sumColumn(group, 'thue');
    const tongTien = sumColumn(group, 'tongtien');

    let giaohang = '...';
    let thanggiaohang = '...';
    let namgiaohang = '...';
    const nInt = parseInt(nam, 10);
    const tInt = parseInt(thang, 10);
    const dInt = parseInt(ngay, 10);
    if (!isNaN(nInt) && !isNaN(tInt) && !isNaN(dInt)) {
      const dt = new Date(Date.UTC(nInt, tInt - 1, dInt));
      dt.setUTCDate(dt.getUTCDate() + 2);
      giaohang = String(dt.getUTCDate()).padStart(2, '0');
      thanggiaohang = String(dt.getUTCMonth() + 1).padStart(2, '0');
      namgiaohang = String(dt.getUTCFullYear());
    }

    const bang1Spec: TableSpec = {
      borderMode: 'insideH',
      colCount: 4,
      rows: [
        [
          { text: 'Tên hàng', bold: true, align: 'center' },
          { text: 'Số lượng', bold: true, align: 'center' },
          { text: 'Đơn giá', bold: true, align: 'center' },
          { text: 'Mặt hàng', bold: true, align: 'center' },
        ],
        ...group.map((row) => [
          { text: safeStr(row['tenhang']), align: 'center' as const },
          { text: `${formatMoneyDot(row['soluong'])} kg`, align: 'center' as const },
          { text: `${formatMoneyDot(row['dongia'])} đ/kg`, align: 'right' as const },
          { text: safeStr(row['mathang']), align: 'center' as const },
        ]),
      ],
    };

    const thueSuatTxt = `Thuế Suất GTGT: ${safeStr(firstRow['thuesuat'])}`;
    const bang2Spec: TableSpec = {
      borderMode: 'full',
      useTableGridStyle: true,
      colCount: 5,
      rows: [
        [
          { text: 'STT', bold: true, align: 'center' },
          { text: 'Tên hàng hóa', bold: true, align: 'center' },
          { text: 'Số lượng (Kg)', bold: true, align: 'center' },
          { text: 'Đơn Giá', bold: true, align: 'center' },
          { text: 'Thành Tiền', bold: true, align: 'center' },
        ],
        ...group.map((row, idx) => [
          { text: String(idx + 1), align: 'center' as const },
          { text: safeStr(row['tenhang']), align: 'center' as const },
          { text: formatMoneyDot(row['soluong']), align: 'center' as const },
          { text: formatMoneyDot(row['dongia']), align: 'right' as const },
          { text: formatMoneyDot(row['thanhtien']), align: 'right' as const },
        ]),
        [
          { text: thueSuatTxt, bold: true, align: 'center', gridSpan: 4 },
          { text: formatMoneyDot(tongThue), bold: true, align: 'right' },
        ],
        [
          { text: 'Tổng tiền thanh toán', bold: true, align: 'center', gridSpan: 4 },
          { text: formatMoneyDot(tongTien), bold: true, align: 'right' },
        ],
      ],
    };

    const context: Record<string, string> = {
      sohd: soHdStr,
      ngay,
      thang,
      nam,
      giaohang,
      thanggiaohang,
      namgiaohang,
      benmua: safeStr(firstRow['benmua']),
      diachi: safeStr(firstRow['diachi']),
      mst: xuLyMst(firstRow['mst']),
      gioitinh: safeStr(firstRow['gioitinh']),
      daidien: safeStr(firstRow['daidien']),
      bang_dieu_1: '__SUBDOC_BANG_DIEU_1__',
      bang_hang_hoa: '__SUBDOC_BANG_HANG_HOA__',
    };

    const docBuffer = renderDocxTemplate(wordTemplateBuffer, context, {
      __SUBDOC_BANG_DIEU_1__: bang1Spec,
      __SUBDOC_BANG_HANG_HOA__: bang2Spec,
    });

    const fileName = `1_${soHdStr}. Hop_Dong_Gao_${ngay}${thang}.docx`;
    folder.file(fileName, docBuffer);
    generatedDocs.push(docBuffer);
  }

  if (generatedDocs.length > 0) {
    const masterName = `1_MASTER_HOP_DONG_TONG_${sohdDau}_Den_${sohdCuoi}.docx`;
    zip.file(masterName, mergeDocxBuffers(generatedDocs));
  }

  return (await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })) as Buffer;
}

export async function handleHopDongXeTai(
  excelBuffer: Buffer,
  wordTemplateBuffer: Buffer
): Promise<Buffer | { error: string }> {
  const wb = XLSX.read(excelBuffer, { type: 'buffer', cellDates: false });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawMatrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null });

  const keywords = [
    'so',
    'stt',
    'số',
    'bks',
    'biển',
    'daidien',
    'đại diện',
    'xe',
    'oto',
    'ô tô',
    'trongtai',
    'tải',
    'cmt',
    'cccd',
    'ngay',
    'tien',
    'tiền',
    'thang',
    'tháng',
    'tong',
    'tổng',
  ];

  let headerIdx = -1;
  let maxScore = 0;

  for (let i = 0; i < rawMatrix.length; i++) {
    const row = rawMatrix[i] || [];
    const rowText = row
      .filter((v) => v !== null && v !== undefined && String(v).trim() !== '')
      .map((v) => String(v).toLowerCase())
      .join(' ');
    const score = keywords.reduce((acc, k) => (rowText.includes(k) ? acc + 1 : acc), 0);
    if (score > maxScore) {
      maxScore = score;
      headerIdx = i;
    }
  }

  if (maxScore < 3 || headerIdx === -1) {
    return { error: 'Không tìm thấy dòng tiêu đề trong file excel.' };
  }

  const headers = (rawMatrix[headerIdx] || []).map((v) =>
    v === null || v === undefined ? 'nan' : String(v).trim().toLowerCase()
  );

  const dataRows: Record<string, unknown>[] = [];
  for (let r = headerIdx + 1; r < rawMatrix.length; r++) {
    const rowArr = rawMatrix[r] || [];
    const rowObj: Record<string, unknown> = {};
    headers.forEach((col, cIdx) => {
      if (!(col in rowObj)) {
        rowObj[col] = rowArr[cIdx] ?? '';
      }
    });
    dataRows.push(rowObj);
  }

  const zip = new JSZip();
  const folder = zip.folder('ChiTiet_HopDongXeTai')!;
  const generatedDocs: Buffer[] = [];

  for (const row of dataRows) {
    const bksRaw = safeStr(layGiaTri(row, ['bks', 'biển'])).trim();
    const daiDienRaw = safeStr(layGiaTri(row, ['daidien', 'đại diện', 'tên']));

    if (!bksRaw || bksRaw.toLowerCase() === 'nan') continue;

    const soRaw = layGiaTri(row, ['so', 'stt', 'số']);
    const soNum = parseFloat(String(soRaw));
    const soStr = !isNaN(soNum) ? String(Math.trunc(soNum)).padStart(2, '0') : '00';

    const cmt = xuLyCmtCccd(layGiaTri(row, ['cmt', 'cccd', 'cmnd']));
    const ngaycap = xuLyNgayExcel(layGiaTri(row, ['ngaycap', 'ngày cấp', 'capngay']));
    const trongtai = formatIntStr(layGiaTri(row, ['trongtai', 'trọng tải', 'tải']));

    const tienThang = formatMoneyDot(layGiaTri(row, ['tien', 'tiền', 'giá']));
    const soThang = formatIntStr(layGiaTri(row, ['thang', 'tháng']));
    const tongTien = formatMoneyDot(layGiaTri(row, ['tong', 'tổng', 'thành']));

    const bangThanhToanSpec: TableSpec = {
      borderMode: 'full',
      colCount: 5,
      rows: [
        [
          { text: 'Từ tháng', bold: true, align: 'center' },
          { text: 'Đến tháng', bold: true, align: 'center' },
          { text: 'Số tháng', bold: true, align: 'center' },
          { text: 'Giá trị thuê / tháng', bold: true, align: 'center' },
          { text: 'Thành tiền', bold: true, align: 'center' },
        ],
        [
          { text: '01', align: 'center' },
          { text: '12', align: 'center' },
          { text: soThang, align: 'center' },
          { text: tienThang, align: 'right' },
          { text: tongTien, align: 'right' },
        ],
        [
          { text: 'Tổng cộng:', bold: true, align: 'center', gridSpan: 4 },
          { text: tongTien, bold: true, align: 'right' },
        ],
      ],
    };

    const diachiVal = safeStr(layGiaTri(row, ['diachi', 'địa chỉ']));
    const conganVal = safeStr(layGiaTri(row, ['congan', 'công an', 'nơi']));
    const otoVal = safeStr(layGiaTri(row, ['oto', 'ô tô', 'loại']));

    const context: Record<string, string> = {
      so: soStr,
      daidien: daiDienRaw,
      diachi: diachiVal,
      cmt,
      ngaycap,
      capngay: ngaycap,
      congan: conganVal,
      bks: bksRaw,
      oto: otoVal,
      trongtai,
      bang_thanh_toan: '__SUBDOC_BANG_THANH_TOAN__',
    };

    const docBuffer = renderDocxTemplate(wordTemplateBuffer, context, {
      __SUBDOC_BANG_THANH_TOAN__: bangThanhToanSpec,
    });

    const fileName = `2_${soStr}_${sanitizeFileName(bksRaw)}.docx`;
    folder.file(fileName, docBuffer);
    generatedDocs.push(docBuffer);
  }

  if (generatedDocs.length > 0) {
    zip.file('2_00_Ban_In_Tat_Ca_Hop_Dong.docx', mergeDocxBuffers(generatedDocs));
  }

  return (await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })) as Buffer;
}

export async function handleHopDongNguyenTacGao(
  excelBuffer: Buffer,
  hdntTemplateBuffer: Buffer,
  bbgnTemplateBuffer: Buffer
): Promise<Buffer> {
  const wb = XLSX.read(excelBuffer, { type: 'buffer' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

  const rows: Record<string, unknown>[] = rawRows.map((r) => ({
    ...r,
    mst: cleanCodeStr(r['mst']),
    sohd: cleanCodeStr(r['sohd']),
  }));

  const zip = new JSZip();
  const hdntFolder = zip.folder('ChiTiet_HopDongNguyenTac')!;
  const bbgnFolder = zip.folder('ChiTiet_BienBanGiaoNhan')!;

  const danhSachHdnt: Buffer[] = [];
  const danhSachBbgn: Buffer[] = [];

  // Part 1: Hợp Đồng Nguyên Tắc (buyers with >= 2 unique invoices)
  const buyerInvoices = new Map<string, Set<string>>();
  const buyerFirstRow = new Map<string, Record<string, unknown>>();
  for (const r of rows) {
    const buyer = safeStr(r['benmua']);
    if (!buyer) continue;
    if (!buyerInvoices.has(buyer)) {
      buyerInvoices.set(buyer, new Set());
      buyerFirstRow.set(buyer, r);
    }
    buyerInvoices.get(buyer)!.add(safeStr(r['sohd']));
  }

  let hdntCounter = 1;
  for (const [buyer, invSet] of buyerInvoices.entries()) {
    if (invSet.size < 2) continue;
    const buyerData = buyerFirstRow.get(buyer)!;
    const ngayNum = parseInt(String(buyerData['ngay']), 10) || 0;
    const thangNum = parseInt(String(buyerData['thang']), 10) || 0;
    const ngayStr = String(ngayNum).padStart(2, '0');
    const thangStr = String(thangNum).padStart(2, '0');
    const counterStr = String(hdntCounter).padStart(2, '0');

    const contextHdnt: Record<string, string> = {
      so_hdnt: counterStr,
      ky_hieu: safeStr(buyerData['kyhieu']),
      ngay: ngayStr,
      thang: thangStr,
      nam: cleanCodeStr(buyerData['nam']),
      ben_mua: safeStr(buyerData['benmua']),
      dia_chi: safeStr(buyerData['diachi']),
      mst: safeStr(buyerData['mst']),
      gioi_tinh: safeStr(buyerData['gioitinh']),
      dai_dien: safeStr(buyerData['daidien']),
    };

    const docBuffer = renderDocxTemplate(hdntTemplateBuffer, contextHdnt);
    const safeName = sanitizeFileName(safeStr(buyerData['benmua']));
    const outName = `3_HDNT_${counterStr}_${safeName}_${ngayStr}${thangStr}.docx`;
    hdntFolder.file(outName, docBuffer);
    danhSachHdnt.push(docBuffer);
    hdntCounter++;
  }

  // Part 2: Biên Bản Giao Nhận (grouped by sohd)
  const groupedByInvoice = new Map<string, Record<string, unknown>[]>();
  for (const r of rows) {
    const sohd = safeStr(r['sohd']);
    if (!sohd) continue;
    if (!groupedByInvoice.has(sohd)) groupedByInvoice.set(sohd, []);
    groupedByInvoice.get(sohd)!.push(r);
  }

  for (const [sohd, group] of groupedByInvoice.entries()) {
    const info = group[0];
    const ngayNum = parseInt(String(info['ngay']), 10) || 0;
    const thangNum = parseInt(String(info['thang']), 10) || 0;
    const ngayStr = String(ngayNum).padStart(2, '0');
    const thangStr = String(thangNum).padStart(2, '0');

    const tongThue = group.length === 1 ? info['thue'] : sumColumn(group, 'thue');
    const tongTien = group.length === 1 ? info['tongtien'] : sumColumn(group, 'tongtien');
    const thueSuatStr = safeStr(info['thuesuat']);

    const bbgnTableSpec: TableSpec = {
      borderMode: 'full',
      useTableGridStyle: true,
      lineSpacing13: true,
      colCount: 5,
      rows: [
        [
          { text: 'STT', bold: true, align: 'center' },
          { text: 'Tên hàng hóa', bold: true, align: 'center' },
          { text: 'Số lượng (Kg)\n(±10%)', bold: true, align: 'center' },
          { text: 'Đơn Giá\n(Đồng/Kg)', bold: true, align: 'center' },
          { text: 'Thành Tiền\n(VNĐ)', bold: true, align: 'center' },
        ],
        ...group.map((row, idx) => [
          { text: String(idx + 1), align: 'center' as const },
          { text: safeStr(row['tenhang']), align: 'center' as const },
          { text: formatCurrencyApi3(row['soluong']), align: 'center' as const },
          { text: formatCurrencyApi3(row['dongia']), align: 'right' as const },
          { text: formatCurrencyApi3(row['thanhtien']), align: 'right' as const },
        ]),
        [
          { text: `Thuế Suất GTGT: ${thueSuatStr}`, align: 'center', gridSpan: 4 },
          { text: formatCurrencyApi3(tongThue), align: 'right' },
        ],
        [
          { text: 'Tổng tiền thanh toán', align: 'center', gridSpan: 4 },
          { text: formatCurrencyApi3(tongTien), align: 'right' },
        ],
      ],
    };

    const contextBbgn: Record<string, string> = {
      ngay: ngayStr,
      thang: thangStr,
      nam: cleanCodeStr(info['nam']),
      ben_mua: safeStr(info['benmua']),
      dia_chi: safeStr(info['diachi']),
      mst: safeStr(info['mst']),
      gioi_tinh: safeStr(info['gioitinh']),
      dai_dien: safeStr(info['daidien']),
    };

    const docBuffer = renderDocxTemplate(bbgnTemplateBuffer, contextBbgn, {
      '[CHEN_BANG_VAO_DAY]': bbgnTableSpec,
    });

    const safeName = sanitizeFileName(safeStr(info['benmua']));
    const outName = `3_BBGN_HD${sohd}_${safeName}_${ngayStr}${thangStr}.docx`;
    bbgnFolder.file(outName, docBuffer);
    danhSachBbgn.push(docBuffer);
  }

  // Part 3: Merge Master files
  if (danhSachHdnt.length > 0) {
    zip.file('3_1_Tong_Hop_Hop_Dong_Nguyen_Tac.docx', mergeDocxBuffers(danhSachHdnt));
  }
  if (danhSachBbgn.length > 0) {
    zip.file('3_2_Tong_Hop_Bien_Ban_Giao_Nhan.docx', mergeDocxBuffers(danhSachBbgn));
  }
  const allDocs = [...danhSachHdnt, ...danhSachBbgn];
  if (allDocs.length > 0) {
    zip.file('3_3_Tong_Hop_Tat_Ca_HDNT_va_BBGN.docx', mergeDocxBuffers(allDocs));
  }

  return (await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })) as Buffer;
}

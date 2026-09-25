import * as XLSX from 'xlsx';

function normalize(value) {
  return String(value ?? '').trim();
}

export function readProjectAssetsFromWorkbook(workbook) {
  const parsedAssets = [];
  let fallbackId = 0;
  const normalizedSheetNames = new Set();

  workbook.SheetNames.forEach((sheetName) => {
    const normalizedName = normalize(sheetName).normalize('NFKC').replace(/\s+/g, ' ').toLocaleLowerCase();
    if (normalizedSheetNames.has(normalizedName)) throw new Error('DUPLICATE_SHEET_NAMES');
    normalizedSheetNames.add(normalizedName);
  });

  workbook.SheetNames.forEach((sheetName) => {
    const type = normalize(sheetName);
    const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, defval: '' });

    rawRows.forEach((row) => {
      fallbackId += 1;
      const asset = {
        sourceId: normalize(row.ID ?? row.id ?? fallbackId),
        pallet: normalize(row.pallet ?? row.Pallet ?? row.PALLET),
        sn: normalize(row.SN ?? row.sn ?? row['Serial Number']),
        type,
      };
      if (asset.sn) parsedAssets.push(asset);
    });
  });

  if (!parsedAssets.length) throw new Error('NO_ASSETS');

  const idCounts = new Map();
  parsedAssets.forEach((asset) => idCounts.set(asset.sourceId, (idCounts.get(asset.sourceId) || 0) + 1));
  const assets = parsedAssets.map(({ sourceId, ...asset }) => ({
    id: idCounts.get(sourceId) > 1 ? `${asset.type}::${sourceId}` : sourceId,
    ...asset,
  }));
  const ids = new Set();
  const serials = new Set();
  for (const asset of assets) {
    if (ids.has(asset.id) || serials.has(asset.sn)) throw new Error('DUPLICATE_ASSETS');
    ids.add(asset.id);
    serials.add(asset.sn);
  }

  const reviewRows = parsedAssets.map(({ sourceId, ...asset }) => ({ id: sourceId, ...asset }));
  return { assets, sheetNames: [...workbook.SheetNames], reviewRows };
}

export async function readProjectAssets(file) {
  let workbook;
  try {
    workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  } catch (error) {
    if (String(error?.message).includes('Invalid HTML')) throw new Error('INVALID_EXCEL_HTML');
    throw new Error('INVALID_EXCEL_FILE');
  }
  return readProjectAssetsFromWorkbook(workbook);
}

export function projectFileErrorMessage(error) {
  if (error.message === 'INVALID_EXCEL_HTML') return 'ไฟล์นี้เป็น HTML และไม่มีตารางข้อมูล กรุณาบันทึกใหม่เป็นไฟล์ Excel .xlsx หรือ .xls';
  if (error.message === 'INVALID_EXCEL_FILE') return 'ไฟล์เสียหรือไม่ใช่ไฟล์ Excel กรุณาบันทึกใหม่เป็น .xlsx หรือ .xls';
  if (error.message === 'DUPLICATE_SHEET_NAMES') return 'พบชื่อชีตซ้ำกัน กรุณาตั้งชื่อประเภทอุปกรณ์ของแต่ละชีตไม่ให้ซ้ำกัน';
  if (error.message === 'NO_ASSETS') return 'ไม่พบ Serial Number ในไฟล์ กรุณาตรวจหัวคอลัมน์ SN';
  if (error.message === 'DUPLICATE_ASSETS') return 'พบ ID ซ้ำภายในชีตเดียวกัน หรือ Serial Number ซ้ำในไฟล์';
  return 'อ่านไฟล์ไม่สำเร็จ กรุณาตรวจสอบรูปแบบไฟล์ Excel';
}

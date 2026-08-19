import { applicationDefault, cert, initializeApp } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';
import XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';

const serviceAccountPath = process.env.FIREBASE_SERVICE_ACCOUNT;
const projectId = process.env.FIREBASE_PROJECT_ID;

if (!serviceAccountPath && !process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  console.error('กรุณากำหนด FIREBASE_SERVICE_ACCOUNT เป็น path ของไฟล์ Service Account JSON');
  process.exit(1);
}

const credential = serviceAccountPath
  ? cert(JSON.parse(fs.readFileSync(path.resolve(serviceAccountPath), 'utf8')))
  : applicationDefault();

initializeApp({ credential, projectId });
const db = getFirestore();
const workbook = XLSX.readFile(path.resolve('serial.xlsx'), { raw: false });
const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { raw: false, defval: '' });
const assets = rows.map((row, index) => ({
  id: String(row.ID || index + 1).trim(),
  pallet: String(row.pallet || '').trim(),
  sn: String(row.SN || '').trim(),
})).filter((row) => /^\d+$/.test(row.sn));

const skippedRows = rows.length - assets.length;
if (skippedRows) console.log(`ข้าม ${skippedRows} แถวที่ว่างหรือไม่ใช่ Serial Number ตัวเลข`);

const seenIds = new Set();
const seenSerials = new Set();
for (const asset of assets) {
  if (seenIds.has(asset.id)) throw new Error(`พบ ID ซ้ำ: ${asset.id}`);
  const key = asset.sn.toLocaleLowerCase();
  if (seenSerials.has(key)) throw new Error(`พบ Serial Number ซ้ำ: ${asset.sn}`);
  seenIds.add(asset.id);
  seenSerials.add(key);
}

for (let start = 0; start < assets.length; start += 450) {
  const batch = db.batch();
  for (const asset of assets.slice(start, start + 450)) {
    batch.set(db.collection('assets').doc(asset.id), {
      sn: asset.sn,
      snSearch: asset.sn.toLocaleLowerCase(),
      pallet: asset.pallet,
    });
  }
  await batch.commit();
  console.log(`อัปโหลดแล้ว ${Math.min(start + 450, assets.length)}/${assets.length}`);
}

await db.collection('system').doc('stats').set({
  totalAssets: assets.length,
  importedAt: new Date().toISOString(),
});
console.log(`นำเข้าข้อมูลสำเร็จทั้งหมด ${assets.length} รายการ`);

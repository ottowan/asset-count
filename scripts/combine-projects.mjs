import { cert, initializeApp } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';
import fs from 'node:fs';
import path from 'node:path';

const sourceProjects = [
  { id: 'project-1788502940157', type: 'Power Supply', prefix: 'ps' },
  { id: 'project-1788496274089', type: 'Monitor แบบที่ 2', prefix: 'monitor2' },
  { id: 'project-1788496260345', type: 'Monitor แบบที่ 1', prefix: 'monitor1' },
  { id: 'project-1788496213784', type: 'PC แบบที่ 2', prefix: 'pc2' },
  { id: 'project-1788495695860', type: 'PC แบบที่ 1', prefix: 'pc1' },
];

const accountPath = path.resolve(process.env.FIREBASE_SERVICE_ACCOUNT || 'firebase-service-account.json');
const account = JSON.parse(fs.readFileSync(accountPath, 'utf8'));
initializeApp({ credential: cert(account), projectId: process.env.FIREBASE_PROJECT_ID || account.project_id });
const db = getFirestore();

const requestedId = process.argv.find((argument) => argument.startsWith('--id='))?.slice(5);
const projectId = requestedId || 'project-1789009347137';
const projectName = 'ตรวจนับรวม PC, Monitor และ Power Supply (69)';
const assets = [];
const selections = [];

for (const source of sourceProjects) {
  const [dataSnapshot, auditSnapshot] = await Promise.all([
    db.collection('project_data').doc(source.id).get(),
    db.collection('random_audits').doc(source.id).get(),
  ]);
  if (!dataSnapshot.exists) throw new Error(`ไม่พบข้อมูลโครงการ ${source.id}`);
  if (!auditSnapshot.exists) throw new Error(`ไม่พบรายการสุ่มของโครงการ ${source.id}`);
  const sourceAssets = dataSnapshot.data().assets || [];
  const sourceAssetIds = new Set(sourceAssets.map((asset) => String(asset.id).trim()));
  for (const asset of sourceAssets) {
    assets.push({
      id: `${source.prefix}-${String(asset.id).trim()}`,
      originalId: String(asset.id).trim(),
      sn: String(asset.sn || '').trim(),
      pallet: String(asset.pallet || '').trim(),
      type: source.type,
      sourceProjectId: source.id,
    });
  }
  for (const selection of auditSnapshot.data().selections || []) {
    const originalAssetId = String(selection.assetId).trim();
    if (!sourceAssetIds.has(originalAssetId)) throw new Error(`รายการสุ่ม ${originalAssetId} ไม่มีในโครงการ ${source.id}`);
    selections.push({
      assetId: `${source.prefix}-${originalAssetId}`,
      round: Number(selection.round) || 1,
      sourceProjectId: source.id,
    });
  }
}

const ensureUnique = (field) => {
  const seen = new Set();
  for (const asset of assets) {
    const value = String(asset[field]).toLocaleLowerCase();
    if (!value || seen.has(value)) throw new Error(`ข้อมูล ${field} ว่างหรือซ้ำ: ${asset[field]}`);
    seen.add(value);
  }
};
ensureUnique('id');
ensureUnique('sn');
if (new Set(selections.map(({ assetId }) => assetId)).size !== selections.length) {
  throw new Error('พบ Asset ID ซ้ำในรายการสุ่มรวม');
}
const combinedAssetIds = new Set(assets.map(({ id }) => id));
if (selections.some(({ assetId }) => !combinedAssetIds.has(assetId))) {
  throw new Error('รายการสุ่มรวมอ้างถึงครุภัณฑ์ที่ไม่มีในโครงการรวม');
}

const projectRef = db.collection('count_projects').doc(projectId);
const dataRef = db.collection('project_data').doc(projectId);
const auditRef = db.collection('random_audits').doc(projectId);
const [existingProject, existingData] = await Promise.all([projectRef.get(), dataRef.get()]);
if (existingProject.exists !== existingData.exists) throw new Error(`ข้อมูลโครงการ ${projectId} ไม่สมบูรณ์`);
if (existingProject.exists) {
  const currentSources = existingProject.data().sourceProjectIds || [];
  if (JSON.stringify(currentSources) !== JSON.stringify(sourceProjects.map(({ id }) => id))) {
    throw new Error(`โครงการ ${projectId} ไม่ได้สร้างจากชุดโครงการต้นทางนี้`);
  }
}

const createdAt = new Date().toISOString();
const batch = db.batch();
batch.set(projectRef, {
  name: projectName,
  status: existingProject.data()?.status || 'closed',
  createdAt: existingProject.data()?.createdAt || createdAt,
  totalAssets: assets.length,
  targetCount: selections.length,
  targetPercent: 30,
  fileName: 'รวมข้อมูลจาก 5 โครงการ',
  assetTypes: sourceProjects.map(({ type }) => type),
  sourceProjectIds: sourceProjects.map(({ id }) => id),
  updatedAt: createdAt,
}, { merge: true });
if (!existingData.exists) {
  batch.set(dataRef, {
    assets,
    totalAssets: assets.length,
    importedAt: createdAt,
    sourceProjectIds: sourceProjects.map(({ id }) => id),
  });
}
batch.set(auditRef, {
  projectId,
  mode: 'all',
  roundSizes: [selections.length],
  selections,
  generatedAt: createdAt,
  sourceProjectIds: sourceProjects.map(({ id }) => id),
});
await batch.commit();

console.log(JSON.stringify({ projectId, projectName, totalAssets: assets.length, randomSelections: selections.length, targetPercent: 30, status: existingProject.data()?.status || 'closed' }, null, 2));

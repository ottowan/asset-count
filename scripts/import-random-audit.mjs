import { applicationDefault, cert, initializeApp } from 'firebase-admin/app';
import { getFirestore } from 'firebase-admin/firestore';
import XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const normalize = (value) => String(value ?? '').trim();

export function validateSelections(rows, assets) {
  if (!rows.length) throw new Error('No random selections found');
  const assetMap = new Map(assets.map((asset) => [normalize(asset.id), asset]));
  const ids = new Set();
  const serials = new Set();
  return rows.map((row) => {
    const id = normalize(row.ID);
    const sn = normalize(row.SN);
    if (!id || !sn) throw new Error('Every selection must have ID and SN');
    if (ids.has(id) || serials.has(sn.toLowerCase())) throw new Error(`Duplicate selection: ${id}`);
    const asset = assetMap.get(id);
    if (!asset || normalize(asset.sn) !== sn) throw new Error(`ID/SN does not match project data: ${id}`);
    ids.add(id);
    serials.add(sn.toLowerCase());
    return { assetId: id, round: 1 };
  });
}

export function readRandomAudit(filePath) {
  const projectId = path.basename(filePath).match(/^random_(project-\d+)\.xlsx$/)?.[1];
  if (!projectId) throw new Error('Expected filename: random_project-<id>.xlsx');
  const workbook = XLSX.readFile(filePath);
  if (!workbook.Sheets['สุ่ม'] || !workbook.Sheets.Sheet1) throw new Error('Expected sheets: สุ่ม and Sheet1');
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets['สุ่ม'], { raw: false, defval: '' });
  const assets = XLSX.utils.sheet_to_json(workbook.Sheets.Sheet1, { raw: false, defval: '' })
    .map((row) => ({ id: normalize(row.ID), sn: normalize(row.SN) }));
  validateSelections(rows, assets);
  return { projectId, rows };
}

async function main() {
  const args = process.argv.slice(2);
  const filePath = path.resolve(args.find((arg) => !arg.startsWith('--')) || 'source/random_project-1788495695860.xlsx');
  const write = args.includes('--write');
  const validateOnly = args.includes('--validate-only');
  if (args.some((arg) => arg.startsWith('--') && !['--write', '--validate-only'].includes(arg)) || (write && validateOnly)) {
    throw new Error('Use [file.xlsx] with --validate-only, --write, or no flag (Firestore dry run)');
  }
  const { projectId, rows } = readRandomAudit(filePath);
  console.log(`Workbook validated: ${projectId}, ${rows.length} random selections`);
  if (validateOnly) return;
  const serviceAccountPath = process.env.FIREBASE_SERVICE_ACCOUNT;
  if (!serviceAccountPath && !process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    throw new Error('Set FIREBASE_SERVICE_ACCOUNT or GOOGLE_APPLICATION_CREDENTIALS to update/validate against Firestore');
  }
  const credential = serviceAccountPath
    ? cert(JSON.parse(fs.readFileSync(path.resolve(serviceAccountPath), 'utf8')))
    : applicationDefault();
  initializeApp({ credential, ...(process.env.FIREBASE_PROJECT_ID ? { projectId: process.env.FIREBASE_PROJECT_ID } : {}) });
  const db = getFirestore();
  const projectRef = db.doc(`count_projects/${projectId}`);
  const dataRef = db.doc(`project_data/${projectId}`);
  const auditRef = db.doc(`random_audits/${projectId}`);
  const backupPath = path.join(path.dirname(filePath), `${projectId}-random-audit-backup-${Date.now()}.json`);
  await db.runTransaction(async (transaction) => {
    const [project, data, previous] = await transaction.getAll(projectRef, dataRef, auditRef);
    if (!project.exists || !data.exists) throw new Error(`Project does not exist: ${projectId}`);
    const selections = validateSelections(rows, data.data().assets || []);
    console.log(`Matched ${selections.length} selections against project: ${project.data().name}`);
    if (!write) return;
    fs.writeFileSync(backupPath, JSON.stringify({
      projectId, firebaseProjectId: db.projectId,
      backedUpAt: new Date().toISOString(),
      previous: previous.exists ? previous.data() : null,
    }, null, 2));
    transaction.set(auditRef, {
      projectId,
      mode: 'rounds',
      roundSizes: [selections.length],
      selections,
      generatedAt: new Date().toISOString(),
    });
  });
  if (!write) {
    console.log('Dry run complete. Add --write to save the random selections.');
    return;
  }
  const saved = await auditRef.get();
  const expected = validateSelections(rows, (await dataRef.get()).data().assets);
  if (JSON.stringify(saved.data()?.selections) !== JSON.stringify(expected)) throw new Error('Post-write verification failed');
  console.log(`Updated and verified ${projectId}: ${expected.length} selections. Backup: ${backupPath}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  main().catch((error) => { console.error(error.message); process.exitCode = 1; });
}

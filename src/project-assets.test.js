import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { exportAssetId, projectFileErrorMessage, readProjectAssets, readProjectAssetsFromWorkbook } from './project-assets.js';

function workbookWithSheets(sheets) {
  const workbook = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name);
  });
  return workbook;
}

test('combines every sheet and uses each sheet name as the asset type', () => {
  const workbook = workbookWithSheets({
    Monitor: [{ ID: 1, PALLET: 'P-01', SN: 'MON-001', Type: 'ignored' }],
    Printer: [{ ID: 2, Pallet: 'P-02', 'Serial Number': 'PRN-001' }],
  });

  assert.deepEqual(readProjectAssetsFromWorkbook(workbook), {
    assets: [
      { id: '1', sourceId: '1', pallet: 'P-01', sn: 'MON-001', type: 'Monitor' },
      { id: '2', sourceId: '2', pallet: 'P-02', sn: 'PRN-001', type: 'Printer' },
    ],
    sheetNames: ['Monitor', 'Printer'],
    reviewRows: [
      { id: '1', pallet: 'P-01', sn: 'MON-001', type: 'Monitor' },
      { id: '2', pallet: 'P-02', sn: 'PRN-001', type: 'Printer' },
    ],
  });
});

test('ignores empty sheets and still reports all sheet names', () => {
  const workbook = workbookWithSheets({
    Monitor: [{ ID: 1, SN: 'MON-001' }],
    Empty: [],
  });

  const result = readProjectAssetsFromWorkbook(workbook);
  assert.equal(result.assets.length, 1);
  assert.deepEqual(result.sheetNames, ['Monitor', 'Empty']);
});

test('namespaces IDs that restart in each sheet', () => {
  const result = readProjectAssetsFromWorkbook(workbookWithSheets({
    Monitor: [{ ID: 1, SN: 'DUP-001' }],
    Printer: [{ ID: 1, SN: 'PRN-001' }],
  }));

  assert.deepEqual(result.assets.map((asset) => asset.id), ['Monitor::1', 'Printer::1']);
});

test('exports the original ID as a number without the sheet namespace', () => {
  assert.equal(exportAssetId({ id: 'Monitor::42', type: 'Monitor' }), 42);
  assert.equal(exportAssetId({ id: 'internal', sourceId: '7', type: 'Monitor' }), 7);
  assert.equal(exportAssetId({ id: 'Monitor::001', type: 'Monitor' }), '001');
  assert.equal(exportAssetId({ id: 'Monitor::ABC', type: 'Monitor' }), 'ABC');
});

test('rejects sheet names that differ only by spacing or letter case', () => {
  const workbook = {
    SheetNames: ['PC Monitor', ' pc  monitor '],
    Sheets: {
      'PC Monitor': XLSX.utils.json_to_sheet([{ ID: 1, SN: 'MON-001' }]),
      ' pc  monitor ': XLSX.utils.json_to_sheet([{ ID: 2, SN: 'MON-002' }]),
    },
  };

  assert.throws(() => readProjectAssetsFromWorkbook(workbook), /DUPLICATE_SHEET_NAMES/);
});

test('rejects duplicate IDs in one sheet or duplicate serial numbers across sheets', () => {
  assert.throws(() => readProjectAssetsFromWorkbook(workbookWithSheets({
    Monitor: [{ ID: 1, SN: 'MON-001' }, { ID: 1, SN: 'MON-002' }],
  })), /DUPLICATE_ASSETS/);

  assert.throws(() => readProjectAssetsFromWorkbook(workbookWithSheets({
    Monitor: [{ ID: 1, SN: 'DUP-001' }],
    Printer: [{ ID: 2, SN: 'DUP-001' }],
  })), /DUPLICATE_ASSETS/);
});

test('reports HTML files without a table as invalid Excel instead of leaking the parser error', async () => {
  const bytes = new TextEncoder().encode('<html><body>not an Excel workbook</body></html>');
  const file = { arrayBuffer: async () => bytes.buffer };

  await assert.rejects(() => readProjectAssets(file), /INVALID_EXCEL_HTML/);
  assert.match(projectFileErrorMessage(new Error('INVALID_EXCEL_HTML')), /HTML/);
});

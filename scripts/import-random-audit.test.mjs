import test from 'node:test';
import assert from 'node:assert/strict';
import { validateSelections } from './import-random-audit.mjs';

const assets = [{ id: '1', sn: 'ABC001' }, { id: '2', sn: 'ABC002' }];

test('preserves selected IDs and puts the imported selections in round 1', () => {
  assert.deepEqual(validateSelections([{ ID: '2', SN: 'ABC002' }], assets), [{ assetId: '2', round: 1 }]);
});

test('rejects a serial from another project even when its ID matches', () => {
  assert.throws(() => validateSelections([{ ID: '1', SN: 'OTHER001' }], assets), /does not match/);
});

test('rejects missing IDs, blank serials, and empty selection sets', () => {
  assert.throws(() => validateSelections([{ ID: '3', SN: 'ABC003' }], assets), /does not match/);
  assert.throws(() => validateSelections([{ ID: '1', SN: '' }], assets), /must have/);
  assert.throws(() => validateSelections([], assets), /No random/);
});

test('rejects duplicate IDs and duplicate serials', () => {
  assert.throws(() => validateSelections([{ ID: '1', SN: 'ABC001' }, { ID: '1', SN: 'ABC001' }], assets), /Duplicate/);
  assert.throws(() => validateSelections([{ ID: '1', SN: 'ABC001' }, { ID: '2', SN: 'ABC001' }], assets), /Duplicate/);
});

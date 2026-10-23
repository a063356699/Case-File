import test from 'node:test';
import assert from 'node:assert/strict';
import { uniqueMeetingRecords, isMeetingChange } from '../app/meetingMerge.ts';

test('one property across weeks uses its latest record without mutating either source', () => {
  const old = Object.freeze({ id: 'old-id', propertyNo: 'eg0540881 ', lastModifiedAt: '2026-10-15', price: '1000' });
  const current = Object.freeze({ id: 'new-id', propertyNo: 'EG0540881', lastModifiedAt: '2026-10-22', price: '1500', status: '售出下架' });
  assert.deepEqual(uniqueMeetingRecords([old, current]), [current]);
  assert.deepEqual(uniqueMeetingRecords([current, old]), [current]);
  assert.equal(old.price, '1000');
  assert.equal(isMeetingChange(current), true);
});

test('separate contracts at the same address and unnumbered drafts remain separate', () => {
  const records = [{ id: 'a', propertyNo: 'A', address: '同址' }, { id: 'b', propertyNo: 'B', address: '同址' }, { id: 'draft-a' }, { id: 'draft-b' }];
  assert.deepEqual(uniqueMeetingRecords([...records, records[2]]), records);
});

test('archived and sold properties go to change reports, active listings stay in introductions', () => {
  assert.equal(isMeetingChange({ id: 'a', status: '委託中' }), false);
  assert.equal(isMeetingChange({ id: 'b', status: '租出下架' }), true);
  assert.equal(isMeetingChange({ id: 'c', archived: '2026-10-22' }), true);
});

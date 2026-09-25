import test from 'node:test';
import assert from 'node:assert/strict';
import { authErrorMessage, firestoreErrorMessage } from './error-messages.js';

test('maps Google login errors to their actual causes', () => {
  assert.match(authErrorMessage({ code: 'auth/popup-blocked' }), /Popup/);
  assert.match(authErrorMessage({ code: 'auth/unauthorized-domain' }), /โดเมน/);
  assert.match(authErrorMessage({ code: 'auth/network-request-failed' }), /เครือข่าย/);
});

test('maps Firestore errors to their actual causes and keeps unknown codes visible', () => {
  assert.match(firestoreErrorMessage({ code: 'permission-denied' }, 'โหลดข้อมูล'), /ไม่มีสิทธิ์/);
  assert.match(firestoreErrorMessage({ code: 'resource-exhausted' }, 'บันทึกข้อมูล'), /โควตา/);
  assert.match(firestoreErrorMessage({ code: 'custom-error' }, 'โหลดข้อมูล'), /custom-error/);
});

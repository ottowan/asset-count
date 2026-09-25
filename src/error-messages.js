function errorCode(error) {
  return String(error?.code || '').trim();
}

export function authErrorMessage(error) {
  const code = errorCode(error);
  const messages = {
    'auth/popup-closed-by-user': 'ยกเลิกการเข้าสู่ระบบแล้ว เนื่องจากหน้าต่าง Google ถูกปิด',
    'auth/popup-blocked': 'Browser บล็อกหน้าต่าง Google กรุณาอนุญาต Popup แล้วลองใหม่',
    'auth/unauthorized-domain': 'Google Login ไม่อนุญาตโดเมนนี้ กรุณาเพิ่มโดเมนใน Firebase Authentication',
    'auth/network-request-failed': 'เข้าสู่ระบบไม่ได้ เนื่องจากเครือข่ายขัดข้อง กรุณาตรวจสอบอินเทอร์เน็ต',
    'auth/cancelled-popup-request': 'คำขอเข้าสู่ระบบก่อนหน้าถูกยกเลิก กรุณากดเข้าสู่ระบบอีกครั้ง',
    'auth/account-exists-with-different-credential': 'Email นี้มีบัญชีด้วยวิธีเข้าสู่ระบบอื่นอยู่แล้ว',
    'auth/operation-not-allowed': 'ยังไม่ได้เปิดใช้งาน Google Login ใน Firebase Authentication',
  };
  return messages[code] || `เข้าสู่ระบบด้วย Google ไม่สำเร็จ${code ? ` (${code})` : ''}`;
}

export function firestoreErrorMessage(error, action) {
  const code = errorCode(error);
  const messages = {
    'permission-denied': `${action}ไม่ได้ เนื่องจากบัญชีนี้ไม่มีสิทธิ์หรือ Firestore Rules ปฏิเสธ`,
    unauthenticated: `${action}ไม่ได้ กรุณาเข้าสู่ระบบใหม่`,
    'resource-exhausted': `${action}ไม่ได้ เนื่องจากโควตา Firestore เต็ม`,
    unavailable: `${action}ไม่ได้ เนื่องจาก Firestore หรือเครือข่ายไม่พร้อมใช้งาน`,
    'deadline-exceeded': `${action}ไม่สำเร็จภายในเวลาที่กำหนด กรุณาลองใหม่`,
    'failed-precondition': `${action}ไม่ได้ เนื่องจากการตั้งค่า Firestore หรือ Index ยังไม่พร้อม`,
    'invalid-argument': `${action}ไม่ได้ เนื่องจากรูปแบบข้อมูลไม่ถูกต้อง`,
    'not-found': `${action}ไม่ได้ เนื่องจากไม่พบข้อมูลที่ต้องการ`,
    'already-exists': `${action}ไม่ได้ เนื่องจากมีข้อมูลนี้อยู่แล้ว`,
  };
  return messages[code] || `${action}ไม่สำเร็จ${code ? ` (${code})` : ''}`;
}

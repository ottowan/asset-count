import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { initializeApp } from 'firebase/app';
import { collection, deleteDoc, doc, getDoc, getDocs, getFirestore, limit, onSnapshot, query as firestoreQuery, setDoc, where } from 'firebase/firestore';
import * as XLSX from 'xlsx';
import './styles.css';

const STORAGE_KEY = 'asset-count-confirmed-v1';
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};
const firebaseEnabled = Boolean(firebaseConfig.apiKey && firebaseConfig.projectId && firebaseConfig.appId);
const db = firebaseEnabled ? getFirestore(initializeApp(firebaseConfig)) : null;

function normalize(value) {
  return String(value ?? '').trim();
}

function extractSerialFromScan(value) {
  const text = normalize(value);
  const afterSeparator = text.match(/--\s*(\d+)/);
  if (afterSeparator) return afterSeparator[1];
  if (/^\d+$/.test(text)) return text;
  const numberGroups = text.match(/\d+/g);
  return numberGroups?.at(-1) || '';
}

function App() {
  const [assets, setAssets] = useState([]);
  const [counted, setCounted] = useState(() => {
    try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {}; } catch { return {}; }
  });
  const [countDetails, setCountDetails] = useState({});
  const [sharedTotal, setSharedTotal] = useState(0);
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState(null);
  const [searchMatches, setSearchMatches] = useState([]);
  const [assetCondition, setAssetCondition] = useState('good');
  const [scannerOpen, setScannerOpen] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [showSummary, setShowSummary] = useState(false);
  const [summaryFilter, setSummaryFilter] = useState('all');
  const [summaryQuery, setSummaryQuery] = useState('');
  const [summaryDate, setSummaryDate] = useState('');
  const [summaryLimit, setSummaryLimit] = useState(200);
  const [status, setStatus] = useState({ type: 'loading', text: 'กำลังโหลดข้อมูลครุภัณฑ์…' });
  const inputRef = useRef(null);
  const scannerVideoRef = useRef(null);

  useEffect(() => {
    if (db) {
      getDoc(doc(db, 'system', 'assets_index'))
        .then((snapshot) => {
          if (!snapshot.exists()) throw new Error('ASSET_INDEX_NOT_FOUND');
          const clean = snapshot.data().assets || [];
          setAssets(clean);
          setStatus({ type: 'ready', text: `พร้อมตรวจนับ ${clean.length.toLocaleString('th-TH')} รายการ` });
        })
        .catch(() => setStatus({ type: 'error', text: 'โหลดรายการไม่สำเร็จ อาจเกินโควตา Firestore กรุณาลองอีกครั้งภายหลัง' }));
      return;
    }
    fetch('/serial.xlsx')
      .then((response) => {
        if (!response.ok) throw new Error('ไม่พบไฟล์ serial.xlsx');
        return response.arrayBuffer();
      })
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { raw: false, defval: '' });
        const clean = rows.map((row, index) => ({
          id: normalize(row.ID || index + 1),
          pallet: normalize(row.pallet),
          sn: normalize(row.SN),
        })).filter((row) => /^\d+$/.test(row.sn));
        setAssets(clean);
        setStatus({ type: 'ready', text: `พร้อมตรวจนับ ${clean.length.toLocaleString('th-TH')} รายการ` });
      })
      .catch((error) => setStatus({ type: 'error', text: error.message || 'โหลดข้อมูลไม่สำเร็จ' }));
  }, []);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(counted));
  }, [counted]);

  useEffect(() => {
    if (!db) return;
    const unsubscribeCounts = onSnapshot(collection(db, 'asset_counts'), (snapshot) => {
      const sharedCounts = {};
      const sharedDetails = {};
      snapshot.forEach((item) => {
        const data = item.data();
        sharedCounts[item.id] = data.countedAt;
        sharedDetails[item.id] = { id: data.assetId, sn: data.sn, pallet: data.pallet, condition: data.condition || '', countedAt: data.countedAt };
      });
      setCounted(sharedCounts);
      setCountDetails(sharedDetails);
    }, () => {
      setStatus({ type: 'error', text: 'เชื่อมต่อยอดส่วนกลางไม่สำเร็จ กรุณาตรวจสอบ Firebase และ Firestore Rules' });
    });
    const unsubscribeStats = onSnapshot(doc(db, 'system', 'stats'), (snapshot) => {
      if (snapshot.exists()) setSharedTotal(Number(snapshot.data().totalAssets) || 0);
    });
    return () => { unsubscribeCounts(); unsubscribeStats(); };
  }, []);

  const countedAssets = useMemo(() => db
    ? Object.values(countDetails).map((item) => ({ id: item.id, sn: item.sn, pallet: item.pallet }))
    : assets.filter((asset) => counted[asset.id]), [assets, counted, countDetails]);
  const total = db ? sharedTotal : assets.length;
  const done = countedAssets.length;
  const remaining = Math.max(total - done, 0);
  const percent = total ? Math.round((done / total) * 100) : 0;
  const todayDone = useMemo(() => {
    const now = new Date();
    const todayKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    return Object.values(counted).filter((value) => {
      const date = new Date(value);
      const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
      return key === todayKey;
    }).length;
  }, [counted]);
  const summaryRows = useMemo(() => {
    const term = summaryQuery.trim();
    return assets.filter((asset) => {
      const isCounted = Boolean(counted[asset.id]);
      const countedDate = isCounted ? (() => {
        const date = new Date(counted[asset.id]);
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      })() : '';
      const matchesFilter = summaryFilter === 'all'
        || (summaryFilter === 'counted' && isCounted)
        || (summaryFilter === 'pending' && !isCounted);
      const matchesDate = !summaryDate || countedDate === summaryDate;
      return matchesFilter && matchesDate && (!term || asset.sn.includes(term));
    });
  }, [assets, counted, summaryFilter, summaryQuery, summaryDate]);

  useEffect(() => { setSummaryLimit(200); }, [summaryFilter, summaryQuery, summaryDate]);

  useEffect(() => {
    if (!scannerOpen) return undefined;
    let stream;
    let animationFrame;
    let active = true;
    const startScanner = async () => {
      if (!('BarcodeDetector' in window)) {
        setStatus({ type: 'error', text: 'Browser นี้ไม่รองรับการสแกน กรุณาเปิดด้วย Chrome หรือ Edge เวอร์ชันล่าสุด' });
        setScannerOpen(false);
        return;
      }
      try {
        stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' } }, audio: false });
        const video = scannerVideoRef.current;
        if (!video || !active) return;
        video.srcObject = stream;
        await video.play();
        const detector = new window.BarcodeDetector({ formats: ['qr_code'] });
        const scanFrame = async () => {
          if (!active) return;
          try {
            const results = await detector.detect(video);
            if (results.length) {
              const decodedText = results[0].rawValue;
        const value = extractSerialFromScan(decodedText);
        if (!value) {
          setStatus({ type: 'error', text: 'QR Code ไม่มี Serial Number ตัวเลข' });
                animationFrame = requestAnimationFrame(scanFrame);
                return;
        }
        const exact = assets.find((asset) => asset.sn === value);
        const matches = exact ? [] : assets.filter((asset) => asset.sn.includes(value));
        setQuery(value);
        setSearchMatches(matches);
        if (exact) {
          setSelected(exact);
          setAssetCondition(countDetails[exact.id]?.condition || 'good');
          setStatus(counted[exact.id]
            ? { type: 'warning', text: 'สแกนพบรายการที่นับแล้ว สามารถยกเลิกการนับได้' }
            : { type: 'found', text: 'สแกนสำเร็จ กรุณาตรวจสอบและกดยืนยัน' });
        } else {
          setSelected(null);
          setStatus(matches.length
            ? { type: 'found', text: `สแกนแล้วพบ ${matches.length.toLocaleString('th-TH')} รายการ กรุณาเลือก` }
            : { type: 'error', text: 'ไม่พบ Serial Number จาก QR Code ในระบบ' });
        }
        setScannerOpen(false);
              return;
            }
          } catch { /* รอภาพจากกล้องเฟรมถัดไป */ }
          animationFrame = requestAnimationFrame(scanFrame);
        };
        scanFrame();
      } catch {
        setStatus({ type: 'error', text: 'เปิดกล้องไม่สำเร็จ กรุณาอนุญาตสิทธิ์กล้องใน Browser' });
        setScannerOpen(false);
      }
    };
    startScanner();
    return () => {
      active = false;
      if (animationFrame) cancelAnimationFrame(animationFrame);
      stream?.getTracks().forEach((track) => track.stop());
    };
  }, [scannerOpen, assets, counted, countDetails]);

  const handleQueryChange = (event) => {
    const value = event.target.value.replace(/\D/g, '');
    setQuery(value);
    setSelected(null);
    if (!value) {
      setSearchMatches([]);
      setStatus({ type: 'ready', text: `พร้อมตรวจนับ ${assets.length.toLocaleString('th-TH')} รายการ` });
      return;
    }
    const matches = assets.filter((asset) => asset.sn.includes(value));
    setSearchMatches(matches);
    setStatus(matches.length
      ? { type: 'found', text: `แนะนำ ${matches.length.toLocaleString('th-TH')} รายการ กดเลือกรายการที่ต้องการ` }
      : { type: 'warning', text: 'ยังไม่พบรายการ ลองกรอกตัวเลขเพิ่มหรือตรวจสอบ Serial Number' });
  };

  const handleSearch = async (event) => {
    event.preventDefault();
    const term = normalize(query).toLocaleLowerCase();
    if (!term) {
      setSelected(null);
      setSearchMatches([]);
      setStatus({ type: 'warning', text: 'กรุณากรอก Serial Number' });
      return;
    }
    setStatus({ type: 'loading', text: 'กำลังค้นหา Serial Number…' });
    const partialMatches = assets.filter((asset) => asset.sn.includes(term));
    const exactLocal = partialMatches.find((asset) => asset.sn === term);
    if (exactLocal) {
      setSelected(exactLocal);
      setAssetCondition(countDetails[exactLocal.id]?.condition || 'good');
      setSearchMatches([]);
      setStatus(counted[exactLocal.id]
        ? { type: 'warning', text: 'Serial Number นี้ถูกนับแล้ว สามารถยกเลิกการนับได้' }
        : { type: 'found', text: 'พบรายการ กรุณาตรวจสอบและกดยืนยัน' });
      inputRef.current?.blur();
      return;
    }
    if (partialMatches.length) {
      setSelected(null);
      setSearchMatches(partialMatches);
      setStatus({ type: 'found', text: `พบ ${partialMatches.length.toLocaleString('th-TH')} รายการ กรุณาเลือก Pallet และ Serial Number` });
      inputRef.current?.blur();
      return;
    }
    let exact;
    if (db) {
      try {
        const result = await getDocs(firestoreQuery(collection(db, 'assets'), where('snSearch', '==', term), limit(1)));
        if (!result.empty) {
          const record = result.docs[0];
          exact = { id: record.id, ...record.data() };
        }
      } catch {
        exact = assets.find((asset) => asset.sn === term);
        if (!exact) {
          setSelected(null);
          setStatus({ type: 'error', text: 'ค้นหาจาก Firestore ไม่สำเร็จ กรุณาตรวจสอบอินเทอร์เน็ต' });
          return;
        }
      }
      if (!exact) exact = assets.find((asset) => asset.sn === term);
    } else {
      exact = assets.find((asset) => asset.sn.toLocaleLowerCase() === term);
    }
    if (exact) {
      setSelected(exact);
      setAssetCondition(countDetails[exact.id]?.condition || 'good');
      setSearchMatches([]);
      setStatus(counted[exact.id]
        ? { type: 'warning', text: 'Serial Number นี้ถูกนับแล้ว' }
        : { type: 'found', text: 'พบรายการ กรุณาตรวจสอบและกดยืนยัน' });
    } else {
      setSelected(null);
      setSearchMatches([]);
      setStatus({ type: 'error', text: 'ไม่พบ Serial Number นี้ในระบบ' });
    }
    inputRef.current?.blur();
  };

  const confirmCount = async () => {
    if (!selected || counted[selected.id] || isSaving) return;
    const now = new Date().toISOString();
    setIsSaving(true);
    if (db) {
      try {
        const itemRef = doc(db, 'asset_counts', String(selected.id));
        await setDoc(itemRef, {
          assetId: String(selected.id),
          sn: selected.sn,
          pallet: selected.pallet,
          condition: assetCondition,
          countedAt: now,
        });
      } catch (error) {
        const quotaExceeded = error.code === 'resource-exhausted';
        setStatus({ type: quotaExceeded ? 'warning' : 'error', text: quotaExceeded ? 'โควตา Firestore วันนี้เต็ม กรุณาลองใหม่หลังโควตารีเซ็ต' : 'บันทึกไม่สำเร็จ รายการอาจถูกนับแล้วหรือ Rules ยังไม่อัปเดต' });
        setIsSaving(false);
        return;
      }
    }
    setCounted((current) => ({ ...current, [selected.id]: now }));
    if (db) setCountDetails((current) => ({ ...current, [selected.id]: { ...selected, condition: assetCondition, countedAt: now } }));
    setStatus({ type: 'success', text: `บันทึก SN ${selected.sn} สำเร็จ ยอดรวมเพิ่มขึ้น 1` });
    setSelected(null);
    setSearchMatches([]);
    setQuery('');
    setIsSaving(false);
    if (!window.matchMedia('(max-width: 720px)').matches) setTimeout(() => inputRef.current?.focus(), 0);
  };

  const cancelCount = async () => {
    if (!selected || !counted[selected.id] || isSaving) return;
    setIsSaving(true);
    try {
      if (db) await deleteDoc(doc(db, 'asset_counts', String(selected.id)));
      setCounted((current) => {
        const next = { ...current };
        delete next[selected.id];
        return next;
      });
      if (db) setCountDetails((current) => {
        const next = { ...current };
        delete next[selected.id];
        return next;
      });
      setStatus({ type: 'success', text: `ยกเลิกการนับ SN ${selected.sn} สำเร็จ ยอดรวมลดลง 1` });
      setSelected(null);
      setSearchMatches([]);
      setQuery('');
      if (!window.matchMedia('(max-width: 720px)').matches) setTimeout(() => inputRef.current?.focus(), 0);
    } catch {
      setStatus({ type: 'error', text: 'ยกเลิกการนับไม่สำเร็จ กรุณาตรวจสอบ Firestore Rules' });
    } finally {
      setIsSaving(false);
    }
  };

  const exportExcel = () => {
    const rows = countedAssets
      .sort((a, b) => new Date(counted[b.id]) - new Date(counted[a.id]))
      .map((asset, index) => ({
        ลำดับ: index + 1,
        ID: asset.id,
        Pallet: asset.pallet,
        'Serial Number': asset.sn,
        สถานะ: 'นับแล้ว',
        สภาพ: countDetails[asset.id]?.condition === 'damaged' ? 'เสีย' : countDetails[asset.id]?.condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ',
        'วันเวลาที่นับ': new Date(counted[asset.id]).toLocaleString('th-TH'),
      }));
    const summary = [
      { รายการ: 'จำนวนทั้งหมด', จำนวน: total },
      { รายการ: 'นับแล้ว', จำนวน: done },
      { รายการ: 'คงเหลือ', จำนวน: remaining },
      { รายการ: 'ความคืบหน้า', จำนวน: `${percent}%` },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'ผลการตรวจนับ');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'สรุปยอด');
    XLSX.writeFile(workbook, `asset-count-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportSummaryExcel = () => {
    const rows = summaryRows.map((asset, index) => {
      const isCounted = Boolean(counted[asset.id]);
      return {
        ลำดับ: index + 1,
        ID: asset.id,
        'Serial Number': asset.sn,
        Pallet: asset.pallet || '-',
        สถานะ: isCounted ? 'นับแล้ว' : 'ยังไม่นับ',
        สภาพ: isCounted ? (countDetails[asset.id]?.condition === 'damaged' ? 'เสีย' : countDetails[asset.id]?.condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ') : '-',
        'วันเวลาที่นับ': isCounted ? new Date(counted[asset.id]).toLocaleString('th-TH') : '-',
      };
    });
    const filterName = summaryFilter === 'counted' ? 'นับแล้ว' : summaryFilter === 'pending' ? 'ยังไม่นับ' : 'ทั้งหมด';
    const summary = [
      { รายการ: 'ตัวกรองที่ส่งออก', จำนวน: filterName },
      { รายการ: 'วันที่นับ', จำนวน: summaryDate || 'ทุกวัน' },
      { รายการ: 'จำนวนในไฟล์', จำนวน: rows.length },
      { รายการ: 'จำนวนทั้งหมด', จำนวน: total },
      { รายการ: 'นับแล้ว', จำนวน: done },
      { รายการ: 'ยังไม่นับ', จำนวน: remaining },
    ];
    const workbook = XLSX.utils.book_new();
    const dataSheet = XLSX.utils.json_to_sheet(rows);
    dataSheet['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(workbook, dataSheet, 'รายการครุภัณฑ์');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'สรุปยอด');
    const suffix = summaryFilter === 'counted' ? 'counted' : summaryFilter === 'pending' ? 'pending' : 'all';
    XLSX.writeFile(workbook, `asset-summary-${suffix}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <main>
      <header className="topbar">
        <div className="brand-mark">AC</div>
        <div>
          <p className="eyebrow">ASSET CONTROL</p>
          <h1>ระบบนับครุภัณฑ์</h1>
        </div>
        <button className="summary-button" onClick={() => setShowSummary(true)} aria-label="ดูสรุปรายการทั้งหมด">▤ <span>สรุปรายการ</span></button>
        <button className="export-button" onClick={exportExcel} disabled={!done} aria-label="ส่งออกผลเป็น Excel">
          <span>⇩</span> Export Excel
        </button>
      </header>

      <section className="hero">
        <div className="hero-copy">
          <span className="live-dot"></span> {db ? 'เชื่อมต่อ Firebase แบบเรียลไทม์' : 'ตรวจนับแบบเรียลไทม์บนอุปกรณ์นี้'}
          <h2>ค้นหา ตรวจสอบ<br />แล้วกดยืนยัน</h2>
          <p>กรอก Serial Number เพื่อค้นหาครุภัณฑ์ ยอดจะเพิ่มขึ้นทันทีหลังยืนยันรายการ</p>
        </div>
        <div className="hero-stats">
        <div className="today-card"><span>ยอดที่นับวันนี้</span><strong>{todayDone.toLocaleString('th-TH')}</strong><small>รายการ</small></div>
        <div className="total-card">
          <div className="mobile-stats">
            <div className="primary"><span>วันนี้</span><strong>{todayDone.toLocaleString('th-TH')}</strong></div>
            <div><span>นับแล้ว</span><strong>{done.toLocaleString('th-TH')}</strong></div>
            <div><span>คงเหลือ</span><strong>{remaining.toLocaleString('th-TH')}</strong></div>
            <div><span>ทั้งหมด</span><strong>{total.toLocaleString('th-TH')}</strong></div>
          </div>
          <span>ยอดที่นับแล้ว</span>
          <strong>{done.toLocaleString('th-TH')}</strong>
          <small>จากทั้งหมด {total.toLocaleString('th-TH')} รายการ</small>
          <em className="today-count">วันนี้นับเพิ่ม {todayDone.toLocaleString('th-TH')} รายการ</em>
          <div className="progress"><i style={{ width: `${percent}%` }} /></div>
          <b>{percent}% สำเร็จ</b>
        </div>
        </div>
      </section>

      <section className="dashboard">
        <div className="work-grid">
          <section className="search-card">
            <div className="section-heading"><span>01</span><div><h3>ค้นหา Serial Number</h3><p>กรอกหมายเลขให้ตรงกับข้อมูลในระบบ</p></div></div>
            <form onSubmit={handleSearch}>
              <label htmlFor="sn">SERIAL NUMBER</label>
              <div className="search-row">
                <div className="input-wrap"><span>⌕</span><input ref={inputRef} id="sn" type="text" value={query} onChange={handleQueryChange} placeholder="พิมพ์ SN เพื่อดูคำแนะนำ" autoComplete="off" inputMode="numeric" pattern="[0-9]*" aria-label="กรอก Serial Number เป็นตัวเลข" /></div>
                <button className="scan-button" type="button" onClick={() => setScannerOpen(true)} aria-label="สแกน QR Code">▣ <span>สแกน</span></button>
                <button type="submit" disabled={!db && !assets.length}>ค้นหา</button>
              </div>
            </form>
            <div className={`notice ${status.type}`}><span>{status.type === 'success' ? '✓' : status.type === 'error' ? '!' : 'i'}</span>{status.text}</div>
            {searchMatches.length > 0 && (
              <div className="search-results">
                {searchMatches.slice(0, 50).map((asset) => (
                  <button key={asset.id} type="button" onClick={() => {
                    setSelected(asset);
                    setAssetCondition(countDetails[asset.id]?.condition || 'good');
                    setQuery(asset.sn);
                    setSearchMatches([]);
                    setStatus(counted[asset.id]
                      ? { type: 'warning', text: 'รายการนี้ถูกนับแล้ว สามารถยกเลิกการนับได้' }
                      : { type: 'found', text: 'เลือกรายการแล้ว กรุณาตรวจสอบและกดยืนยัน' });
                  }}>
                    <span><small>PALLET</small><strong>{asset.pallet || '-'}</strong></span>
                    <span><small>SERIAL NUMBER</small><strong>{asset.sn}</strong></span>
                    <i>{counted[asset.id] ? 'นับแล้ว' : 'เลือก'}</i>
                  </button>
                ))}
                {searchMatches.length > 50 && <p>แสดง 50 จาก {searchMatches.length.toLocaleString('th-TH')} รายการ กรุณากรอกตัวเลขเพิ่มเพื่อจำกัดผลลัพธ์</p>}
              </div>
            )}
          </section>

          <section className={`confirm-card ${selected ? 'active' : ''}`}>
            <div className="section-heading"><span>02</span><div><h3>ยืนยันรายการ</h3><p>ตรวจสอบข้อมูลก่อนบันทึกยอด</p></div>{selected && <b className="confirm-id">ID: {selected.id}</b>}{selected && <button className="confirm-close" onClick={() => setSelected(null)} aria-label="ปิดรายการ">×</button>}</div>
            {selected ? (
              <div className="asset-result">
                <dl><div><dt>Pallet</dt><dd className="pallet-value">{selected.pallet || '-'}</dd></div><div><dt>Serial Number</dt><dd>{selected.sn}</dd></div></dl>
                {counted[selected.id] ? (
                  <div className={`condition-readonly ${countDetails[selected.id]?.condition === 'damaged' ? 'damaged' : 'good'}`}><span>สภาพครุภัณฑ์</span><strong>{countDetails[selected.id]?.condition === 'damaged' ? 'เสีย' : countDetails[selected.id]?.condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</strong></div>
                ) : (
                  <fieldset className="condition-picker"><legend>สภาพครุภัณฑ์</legend><label className={assetCondition === 'good' ? 'selected' : ''}><input type="radio" name="condition" value="good" checked={assetCondition === 'good'} onChange={() => setAssetCondition('good')} /><span>✓</span><strong>ไม่เสีย</strong></label><label className={assetCondition === 'damaged' ? 'selected damaged' : ''}><input type="radio" name="condition" value="damaged" checked={assetCondition === 'damaged'} onChange={() => setAssetCondition('damaged')} /><span>!</span><strong>เสีย</strong></label></fieldset>
                )}
                {counted[selected.id] ? (
                  <button className="confirm-button cancel-button" onClick={cancelCount} disabled={isSaving}>× {isSaving ? 'กำลังยกเลิก…' : 'ยกเลิกการนับรายการนี้'}</button>
                ) : (
                  <button className="confirm-button" onClick={confirmCount} disabled={isSaving}>✓ {isSaving ? 'กำลังบันทึก…' : 'ยืนยันนับรายการ'}</button>
                )}
              </div>
            ) : (
              <div className="empty-state"><span>✓</span><p>ข้อมูลรายการจะแสดงที่นี่<br />หลังจากค้นหา Serial Number</p></div>
            )}
          </section>
        </div>

        <section className="recent-card">
          <div className="recent-header"><div><h3>รายการที่นับล่าสุด</h3><p>แสดง 5 รายการล่าสุดบนอุปกรณ์นี้</p></div><button onClick={exportExcel} disabled={!done}>Export Excel</button></div>
          {done ? <div className="recent-list">{countedAssets.sort((a,b) => new Date(counted[b.id]) - new Date(counted[a.id])).slice(0,5).map((asset) => <div className="recent-row" key={asset.id}><span className="check">✓</span><div><strong>{asset.sn}</strong><small>Pallet {asset.pallet || '-'}</small></div><time>{new Date(counted[asset.id]).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })}</time></div>)}</div> : <p className="no-records">ยังไม่มีรายการที่ยืนยันการนับ</p>}
        </section>
      </section>
      <footer>{db ? 'ยอดรวมเชื่อมต่อ Firebase แบบ Real-time และอัปเดตทุกอุปกรณ์โดยไม่ต้อง Refresh' : 'โหมด Local: ข้อมูลการนับบันทึกในเบราว์เซอร์ของอุปกรณ์นี้'}</footer>

      {showSummary && (
        <div className="summary-modal" role="dialog" aria-modal="true" aria-labelledby="summary-title">
          <div className="summary-panel">
            <header className="summary-panel-header">
              <div><p>ASSET OVERVIEW</p><h2 id="summary-title">สรุปรายการครุภัณฑ์</h2></div>
              <button onClick={() => setShowSummary(false)} aria-label="ปิดหน้าสรุป">×</button>
            </header>
            <div className="summary-totals">
              <button className={summaryFilter === 'all' ? 'active' : ''} onClick={() => setSummaryFilter('all')}><span>ทั้งหมด</span><strong>{total.toLocaleString('th-TH')}</strong></button>
              <button className={summaryFilter === 'counted' ? 'active counted' : 'counted'} onClick={() => setSummaryFilter('counted')}><span>นับแล้ว</span><strong>{done.toLocaleString('th-TH')}</strong></button>
              <button className={summaryFilter === 'pending' ? 'active pending' : 'pending'} onClick={() => setSummaryFilter('pending')}><span>ยังไม่นับ</span><strong>{remaining.toLocaleString('th-TH')}</strong></button>
            </div>
            <div className="summary-tools">
              <div className="summary-search">⌕<input value={summaryQuery} onChange={(event) => setSummaryQuery(event.target.value.replace(/\D/g, ''))} inputMode="numeric" placeholder="ค้นหา Serial Number" /></div>
              <div className="date-filter"><label htmlFor="count-date">วันที่นับ</label><input id="count-date" type="date" value={summaryDate} onChange={(event) => setSummaryDate(event.target.value)} />{summaryDate && <button onClick={() => setSummaryDate('')} aria-label="ล้างวันที่">×</button>}</div>
              <span>พบ {summaryRows.length.toLocaleString('th-TH')} รายการ</span>
              <button className="summary-export" onClick={exportSummaryExcel} disabled={!summaryRows.length}>⇩ Export Excel</button>
            </div>
            <div className="asset-table-wrap">
              <table className="asset-table">
                <thead><tr><th>ลำดับ</th><th>Serial Number</th><th>Pallet</th><th>สถานะ</th><th>สภาพ</th><th>เวลาที่นับ</th></tr></thead>
                <tbody>{summaryRows.slice(0, summaryLimit).map((asset, index) => {
                  const isCounted = Boolean(counted[asset.id]);
                  const condition = countDetails[asset.id]?.condition;
                  return <tr key={asset.id}><td>{index + 1}</td><td><strong>{asset.sn}</strong><small>ID {asset.id}</small></td><td>{asset.pallet || '-'}</td><td><span className={`status-pill ${isCounted ? 'is-counted' : 'is-pending'}`}>{isCounted ? '✓ นับแล้ว' : '– ยังไม่นับ'}</span></td><td><span className={`condition-pill ${condition === 'damaged' ? 'is-damaged' : condition === 'good' ? 'is-good' : ''}`}>{!isCounted ? '-' : condition === 'damaged' ? 'เสีย' : condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</span></td><td>{isCounted ? new Date(counted[asset.id]).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' }) : '-'}</td></tr>;
                })}</tbody>
              </table>
              {!summaryRows.length && <div className="summary-empty">ไม่พบรายการ</div>}
            </div>
            {summaryRows.length > summaryLimit && <button className="load-more" onClick={() => setSummaryLimit((value) => value + 200)}>แสดงเพิ่มอีก {Math.min(200, summaryRows.length - summaryLimit).toLocaleString('th-TH')} รายการ</button>}
          </div>
        </div>
      )}
      {scannerOpen && (
        <div className="scanner-modal" role="dialog" aria-modal="true" aria-label="สแกน QR Code">
          <div className="scanner-panel"><div className="scanner-header"><div><strong>สแกน QR Code</strong><small>วาง QR Code ให้อยู่ในกรอบ</small></div><button onClick={() => setScannerOpen(false)} aria-label="ปิดกล้อง">×</button></div><div className="camera-view"><video ref={scannerVideoRef} playsInline muted /><i /></div><p>กล้องจะอ่าน Serial Number และค้นหาให้อัตโนมัติ</p></div>
        </div>
      )}
    </main>
  );
}

createRoot(document.getElementById('root')).render(<React.StrictMode><App /></React.StrictMode>);

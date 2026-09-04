import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { BrowserQRCodeReader } from '@zxing/browser';
import { initializeApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider, onAuthStateChanged, signInWithPopup, signOut } from 'firebase/auth';
import { collection, deleteDoc, doc, getDoc, getDocs, getFirestore, limit, onSnapshot, query as firestoreQuery, setDoc, updateDoc, where } from 'firebase/firestore';
import * as XLSX from 'xlsx';
import './styles.css';

const STORAGE_KEY = 'asset-count-confirmed-v1';
const ACTIVE_PROJECT_KEY = 'asset-count-active-project-v1';
const ADMIN_EMAIL = 'parinya.coj@gmail.com';
const LEGACY_PROJECT = { id: 'legacy', name: 'โครงการเดิม', status: 'open', isLegacy: true, managed: false };
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
};
const firebaseEnabled = Boolean(firebaseConfig.apiKey && firebaseConfig.projectId && firebaseConfig.appId);
const firebaseApp = firebaseEnabled ? initializeApp(firebaseConfig) : null;
const db = firebaseApp ? getFirestore(firebaseApp) : null;
const auth = firebaseApp ? getAuth(firebaseApp) : null;

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
  const [projects, setProjects] = useState([LEGACY_PROJECT]);
  const [activeProjectId, setActiveProjectId] = useState(() => localStorage.getItem(ACTIVE_PROJECT_KEY) || 'legacy');
  const [viewingProjectId, setViewingProjectId] = useState(null);
  const [currentPage, setCurrentPage] = useState('count');
  const [newProjectName, setNewProjectName] = useState('');
  const [newProjectFile, setNewProjectFile] = useState(null);
  const [newProjectTarget, setNewProjectTarget] = useState('');
  const [editingProject, setEditingProject] = useState(null);
  const [isSavingProject, setIsSavingProject] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [authReady, setAuthReady] = useState(!auth);
  const [authorizedEmails, setAuthorizedEmails] = useState([]);
  const [accessReady, setAccessReady] = useState(!auth);
  const [newAuthorizedEmail, setNewAuthorizedEmail] = useState('');
  const [isSavingAccess, setIsSavingAccess] = useState(false);
  const [query, setQuery] = useState('');
  const [selected, setSelected] = useState(null);
  const [searchMatches, setSearchMatches] = useState([]);
  const [assetCondition, setAssetCondition] = useState('good');
  const [scannerOpen, setScannerOpen] = useState(false);
  const [showRandomAudit, setShowRandomAudit] = useState(false);
  const [showRandomReport, setShowRandomReport] = useState(false);
  const [randomAuditRows, setRandomAuditRows] = useState([]);
  const [isSavingRandomAudit, setIsSavingRandomAudit] = useState(false);
  const [selectedRandomPallet, setSelectedRandomPallet] = useState(null);
  const [randomAuditMode, setRandomAuditMode] = useState('rounds');
  const [randomRoundCount, setRandomRoundCount] = useState(3);
  const [randomRoundSizes, setRandomRoundSizes] = useState(['', '', '']);
  const [isSaving, setIsSaving] = useState(false);
  const [showSummary, setShowSummary] = useState(false);
  const [summaryView, setSummaryView] = useState('pallets');
  const [summaryFilter, setSummaryFilter] = useState('all');
  const [summaryQuery, setSummaryQuery] = useState('');
  const [summaryDate, setSummaryDate] = useState('');
  const [randomReportFilter, setRandomReportFilter] = useState('all');
  const [randomReportView, setRandomReportView] = useState('pallets');
  const [summaryLimit, setSummaryLimit] = useState(200);
  const [selectedPalletSummary, setSelectedPalletSummary] = useState(null);
  const [editingCountDate, setEditingCountDate] = useState(null);
  const [outsideAuditAsset, setOutsideAuditAsset] = useState(null);
  const [isUpdatingCountDate, setIsUpdatingCountDate] = useState(false);
  const [status, setStatus] = useState({ type: 'loading', text: 'กำลังโหลดข้อมูลครุภัณฑ์…' });
  const inputRef = useRef(null);
  const scannerVideoRef = useRef(null);
  const currentProjectId = viewingProjectId || activeProjectId;
  const activeProject = projects.find((project) => project.id === currentProjectId) || LEGACY_PROJECT;
  const projectFileName = (activeProject.name || 'project').trim().replace(/[<>:"/\\|?*\x00-\x1F]+/g, '-').replace(/\s+/g, '-').replace(/[.-]+$/g, '') || 'project';
  const isViewingClosedProject = Boolean(viewingProjectId && activeProject.status === 'closed');
  const projectIsOpen = activeProject.status === 'open';
  const countDocumentId = (assetId) => currentProjectId === 'legacy' ? String(assetId) : `${currentProjectId}__${assetId}`;
  const normalizedUserEmail = currentUser?.email?.trim().toLocaleLowerCase() || '';
  const isAdmin = normalizedUserEmail === ADMIN_EMAIL;
  const hasAccess = !auth || isAdmin || authorizedEmails.includes(normalizedUserEmail);

  useEffect(() => {
    if (!auth) return undefined;
    return onAuthStateChanged(auth, (user) => {
      if (user) setAccessReady(false);
      setCurrentUser(user);
      setAuthReady(true);
    });
  }, []);

  useEffect(() => {
    if (!db || !currentUser) { setAuthorizedEmails([]); setAccessReady(!currentUser); return undefined; }
    setAccessReady(false);
    return onSnapshot(collection(db, 'authorized_users'), (snapshot) => {
      setAuthorizedEmails(snapshot.docs.map((item) => String(item.data().email || item.id).trim().toLocaleLowerCase()));
      setAccessReady(true);
    }, () => {
      setAccessReady(false);
      setStatus({ type: 'error', code: 'access-check-failed', text: 'ตรวจสอบสิทธิ์ผู้ใช้งานไม่สำเร็จ กรุณาตรวจสอบอินเทอร์เน็ตแล้วลองใหม่' });
    });
  }, [currentUser]);

  useEffect(() => {
    if (auth && authReady && (!currentUser || (!hasAccess && accessReady)) && ['projects', 'access'].includes(currentPage)) setCurrentPage('count');
  }, [authReady, accessReady, currentUser, currentPage, hasAccess]);

  useEffect(() => {
    if (!currentUser || !accessReady) return;
    if (!hasAccess) {
      setStatus({ type: 'error', code: 'access-denied', text: `บัญชี ${currentUser.email} ยังไม่ได้รับสิทธิ์ใช้งาน กรุณาติดต่อผู้ดูแลระบบ` });
      return;
    }
    setStatus((current) => current.code === 'access-denied' || current.code === 'access-check-failed'
      ? { type: 'ready', text: `พร้อมตรวจนับ ${assets.length.toLocaleString('th-TH')} รายการ` }
      : current);
  }, [currentUser, accessReady, hasAccess, assets.length]);

  const runAuthenticated = async (action) => {
    if (!auth || currentUser) { action(); return; }
    try {
      const result = await signInWithPopup(auth, new GoogleAuthProvider());
      setAccessReady(false);
      setCurrentUser(result.user);
      action();
    } catch {
      setStatus({ type: 'error', text: 'เข้าสู่ระบบด้วย Google ไม่สำเร็จ กรุณาลองใหม่' });
    }
  };

  const runAuthorized = async (action) => {
    if (!currentUser) { await runAuthenticated(() => {}); return; }
    if (!hasAccess) { setStatus({ type: 'error', text: `บัญชี ${currentUser.email} ยังไม่ได้รับสิทธิ์ใช้งาน` }); return; }
    action();
  };

  useEffect(() => {
    if (db) {
      const sourceRef = currentProjectId === 'legacy' ? doc(db, 'system', 'assets_index') : doc(db, 'project_data', currentProjectId);
      setAssets([]);
      setSharedTotal(0);
      getDoc(sourceRef)
        .then((snapshot) => {
          if (!snapshot.exists()) throw new Error('PROJECT_DATA_NOT_FOUND');
          const clean = snapshot.data().assets || [];
          setAssets(clean);
          setSharedTotal(clean.length);
          setStatus({ type: 'ready', text: `พร้อมตรวจนับ ${clean.length.toLocaleString('th-TH')} รายการ` });
        })
        .catch(() => setStatus({ type: 'error', text: 'โหลดรายการไม่สำเร็จ อาจเกินโควตา Firestore กรุณาลองอีกครั้งภายหลัง' }));
      return;
    }
    if (currentProjectId !== 'legacy') return;
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
  }, [currentProjectId]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(counted));
  }, [counted]);

  useEffect(() => { localStorage.setItem(ACTIVE_PROJECT_KEY, activeProjectId); }, [activeProjectId]);

  useEffect(() => {
    const isProjectsPage = currentPage === 'projects';
    document.documentElement.classList.toggle('projects-page-open', isProjectsPage);
    document.body.classList.toggle('projects-page-open', isProjectsPage);
    return () => {
      document.documentElement.classList.remove('projects-page-open');
      document.body.classList.remove('projects-page-open');
    };
  }, [currentPage]);

  useEffect(() => {
    if (!db) return;
    return onSnapshot(collection(db, 'count_projects'), (snapshot) => {
      const remote = snapshot.docs.map((item) => ({ id: item.id, ...item.data(), managed: true }));
      const remoteLegacy = remote.find((project) => project.id === 'legacy');
      const regularProjects = remote.filter((project) => project.id !== 'legacy');
      const nextProjects = [{ ...LEGACY_PROJECT, ...remoteLegacy, isLegacy: true }, ...regularProjects.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))];
      setProjects(nextProjects);
      const openProject = nextProjects.find((project) => project.status === 'open');
      if (openProject) setActiveProjectId(openProject.id);
    });
  }, []);

  useEffect(() => {
    if (!db) return;
    setCounted({});
    setCountDetails({});
    const unsubscribeCounts = onSnapshot(collection(db, 'asset_counts'), (snapshot) => {
      const sharedCounts = {};
      const sharedDetails = {};
      snapshot.forEach((item) => {
        const data = item.data();
        const belongsToProject = currentProjectId === 'legacy' ? !data.projectId : data.projectId === currentProjectId;
        if (!belongsToProject) return;
        sharedCounts[data.assetId] = data.countedAt;
        sharedDetails[data.assetId] = { id: data.assetId, sn: data.sn, pallet: data.pallet, condition: data.condition || '', countedAt: data.countedAt };
      });
      setCounted(sharedCounts);
      setCountDetails(sharedDetails);
    }, () => {
      setStatus({ type: 'error', text: 'เชื่อมต่อยอดส่วนกลางไม่สำเร็จ กรุณาตรวจสอบ Firebase และ Firestore Rules' });
    });
    return () => { unsubscribeCounts(); };
  }, [currentProjectId]);

  useEffect(() => {
    if (!db || !assets.length) return undefined;
    const auditId = currentProjectId;
    return onSnapshot(doc(db, 'random_audits', auditId), (snapshot) => {
      if (!snapshot.exists()) { setRandomAuditRows([]); return; }
      const assetMap = new Map(assets.map((asset) => [String(asset.id), asset]));
      const rows = (snapshot.data().selections || []).map((selection) => {
        const asset = assetMap.get(String(selection.assetId));
        return asset ? { ...asset, round: Number(selection.round) || 0 } : null;
      }).filter(Boolean);
      rows.sort((a, b) => a.round - b.round || (a.pallet || '').localeCompare(b.pallet || '', 'th', { numeric: true }) || a.sn.localeCompare(b.sn, 'th', { numeric: true }));
      setRandomAuditRows(rows);
    });
  }, [currentProjectId, assets]);

  const randomAuditIdSet = useMemo(() => new Set(randomAuditRows.map((asset) => String(asset.id))), [randomAuditRows]);
  const hasRandomAudit = randomAuditRows.length > 0;
  const isRandomEligible = (assetId) => !hasRandomAudit || randomAuditIdSet.has(String(assetId));
  const notifyOutsideRandomAudit = (asset) => {
    setOutsideAuditAsset(asset);
  };
  const dismissOutsideAuditAlert = () => {
    setOutsideAuditAsset(null);
    setSelected(null);
    setSearchMatches([]);
    setQuery('');
    setStatus({ type: 'ready', text: `พร้อมตรวจนับ ${assets.length.toLocaleString('th-TH')} รายการ` });
    if (!window.matchMedia('(max-width: 720px)').matches) setTimeout(() => inputRef.current?.focus(), 0);
  };

  const countedAssets = useMemo(() => db
    ? Object.values(countDetails).map((item) => ({ id: item.id, sn: item.sn, pallet: item.pallet }))
    : assets.filter((asset) => counted[asset.id]), [assets, counted, countDetails]);
  const total = db ? sharedTotal : assets.length;
  const done = countedAssets.length;
  const remaining = Math.max(total - done, 0);
  const targetTotal = Math.min(Math.max(Number(activeProject.targetCount) || total, 1), total || 1);
  const targetPercent = Number(activeProject.targetPercent) || (total ? (targetTotal / total) * 100 : 0);
  const targetRemaining = Math.max(targetTotal - done, 0);
  const randomAvailableTarget = Math.min(targetRemaining, remaining);
  const randomRoundTotal = randomRoundSizes.reduce((sum, value) => sum + (Number(value) || 0), 0);

  const openRandomAudit = () => {
    const availableTarget = Math.min(targetRemaining, remaining);
    setRandomAuditMode('rounds');
    setRandomRoundCount(3);
    const base = Math.floor(availableTarget / 3);
    setRandomRoundSizes([String(base), String(base), String(availableTarget - (base * 2))]);
    setShowRandomAudit(true);
  };

  const changeRandomRoundCount = (value) => {
    const count = Math.min(Math.max(Number(value) || 1, 1), 20);
    const availableTarget = Math.min(targetRemaining, remaining);
    const base = Math.floor(availableTarget / count);
    setRandomRoundCount(count);
    setRandomRoundSizes(Array.from({ length: count }, (_, index) => String(index === count - 1 ? availableTarget - (base * (count - 1)) : base)));
  };

  const generateRandomAudit = async () => {
    const availableTarget = Math.min(targetRemaining, remaining);
    const roundSizes = randomAuditMode === 'all' ? [availableTarget] : randomRoundSizes.map((value) => Number(value) || 0);
    const requested = roundSizes.reduce((sum, value) => sum + value, 0);
    if (requested > availableTarget) return;
    if (!requested && !done) return;
    const shuffle = (items) => {
      const result = [...items];
      for (let index = result.length - 1; index > 0; index -= 1) {
        const swapIndex = Math.floor(Math.random() * (index + 1));
        [result[index], result[swapIndex]] = [result[swapIndex], result[index]];
      }
      return result;
    };
    const countedPerPallet = new Map();
    assets.forEach((asset) => {
      const pallet = asset.pallet || 'ไม่ระบุ Pallet';
      if (!countedPerPallet.has(pallet)) countedPerPallet.set(pallet, 0);
      if (counted[asset.id]) countedPerPallet.set(pallet, countedPerPallet.get(pallet) + 1);
    });
    const grouped = new Map();
    assets.filter((asset) => !counted[asset.id]).forEach((asset) => {
      const pallet = asset.pallet || 'ไม่ระบุ Pallet';
      if (!grouped.has(pallet)) grouped.set(pallet, []);
      grouped.get(pallet).push(asset);
    });
    let activePallets = [...grouped.entries()].map(([pallet, items]) => ({
      pallet,
      items: shuffle(items),
      cursor: 0,
      effectiveCount: countedPerPallet.get(pallet) || 0,
    }));
    const selectedRows = [];
    while (selectedRows.length < requested && activePallets.length) {
      const lowestCount = Math.min(...activePallets.map((group) => group.effectiveCount));
      const lowestPallets = activePallets.filter((group) => group.effectiveCount === lowestCount);
      const group = lowestPallets[Math.floor(Math.random() * lowestPallets.length)];
      selectedRows.push({ ...group.items[group.cursor] });
      group.cursor += 1;
      group.effectiveCount += 1;
      activePallets = activePallets.filter((group) => group.cursor < group.items.length);
    }
    let rowCursor = 0;
    const rowsWithRounds = assets.filter((asset) => counted[asset.id]).map((asset) => ({ ...asset, round: 0 }));
    roundSizes.forEach((size, roundIndex) => {
      selectedRows.slice(rowCursor, rowCursor + size).forEach((asset) => rowsWithRounds.push({ ...asset, round: roundIndex + 1 }));
      rowCursor += size;
    });
    rowsWithRounds.sort((a, b) => a.round - b.round
      || (a.pallet || '').localeCompare(b.pallet || '', 'th', { numeric: true, sensitivity: 'base' })
      || a.sn.localeCompare(b.sn, 'th', { numeric: true, sensitivity: 'base' }));
    setIsSavingRandomAudit(true);
    try {
      if (db) await setDoc(doc(db, 'random_audits', currentProjectId), {
        projectId: currentProjectId,
        mode: randomAuditMode,
        roundSizes,
        selections: rowsWithRounds.map((asset) => ({ assetId: String(asset.id), round: asset.round })),
        generatedAt: new Date().toISOString(),
      });
      setRandomAuditRows(rowsWithRounds);
    } catch {
      setStatus({ type: 'error', text: 'บันทึกรายการสุ่มไม่สำเร็จ กรุณาอัปเดต Firestore Rules' });
    } finally { setIsSavingRandomAudit(false); }
  };

  const clearRandomAudit = async () => {
    if (!window.confirm('ล้างเฉพาะรายการสุ่มของโครงการนี้ใช่หรือไม่? ผลการนับเดิมจะไม่ถูกลบ')) return;
    setIsSavingRandomAudit(true);
    try {
      if (db) await deleteDoc(doc(db, 'random_audits', currentProjectId));
      setRandomAuditRows([]);
      setSelectedRandomPallet(null);
    } catch {
      setStatus({ type: 'error', text: 'ล้างรายการสุ่มไม่สำเร็จ กรุณาตรวจสอบ Firestore Rules' });
    } finally { setIsSavingRandomAudit(false); }
  };

  const exportRandomAudit = (includePrior = false) => {
    if (!randomAuditRows.length) return;
    const sourceRows = includePrior === true ? randomAuditRows : randomAuditRows.filter((asset) => asset.round > 0);
    const rows = sourceRows.map((asset, index) => ({ ลำดับ: index + 1, รอบ: asset.round === 0 ? 'นับก่อนสุ่ม' : asset.round, Pallet: asset.pallet || '-', 'Serial Number': asset.sn, ID: asset.id, สถานะ: counted[asset.id] ? 'นับแล้ว' : 'ยังไม่นับ', สภาพ: counted[asset.id] ? (countDetails[asset.id]?.condition === 'damaged' ? 'เสีย' : countDetails[asset.id]?.condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ') : '-' }));
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet['!cols'] = [{ wch: 8 }, { wch: 8 }, { wch: 18 }, { wch: 22 }, { wch: 12 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(workbook, sheet, 'รายการสุ่มตรวจ');
    XLSX.writeFile(workbook, `random-audit-${new Date().toISOString().slice(0, 10)}-${projectFileName}.xlsx`);
  };
  const exportRandomReport = () => {
    const reportAssets = filteredRandomReportAssets;
    if (!reportAssets.length) return;
    const rows = reportAssets.map((asset, index) => ({ ลำดับ: index + 1, รอบ: asset.round === 0 ? 'นับก่อนสุ่ม' : asset.round, Pallet: asset.pallet || '-', 'Serial Number': asset.sn, ID: asset.id, สถานะ: counted[asset.id] ? 'นับแล้ว' : 'ยังไม่นับ', สภาพ: counted[asset.id] ? (countDetails[asset.id]?.condition === 'damaged' ? 'เสีย' : countDetails[asset.id]?.condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ') : '-', 'วันเวลาที่นับ': counted[asset.id] ? new Date(counted[asset.id]).toLocaleString('th-TH') : '-' }));
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows);
    sheet['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 18 }, { wch: 22 }, { wch: 12 }, { wch: 14 }, { wch: 12 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(workbook, sheet, 'รายงานสุ่มและตรวจนับ');
    XLSX.writeFile(workbook, `random-audit-report-${new Date().toISOString().slice(0, 10)}-${projectFileName}.xlsx`);
  };
  const percent = targetTotal ? Math.min(Math.ceil((done / targetTotal) * 100), 100) : 0;
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

  const palletRows = useMemo(() => {
    const groups = new Map();
    assets.forEach((asset) => {
      const pallet = asset.pallet || 'ไม่ระบุ Pallet';
      if (!groups.has(pallet)) groups.set(pallet, { pallet, assets: [], countedCount: 0, goodCount: 0, damagedCount: 0, latestCountedAt: '' });
      const group = groups.get(pallet);
      group.assets.push(asset);
      if (counted[asset.id]) {
        group.countedCount += 1;
        if (countDetails[asset.id]?.condition === 'good') group.goodCount += 1;
        if (countDetails[asset.id]?.condition === 'damaged') group.damagedCount += 1;
        if (!group.latestCountedAt || new Date(counted[asset.id]) > new Date(group.latestCountedAt)) group.latestCountedAt = counted[asset.id];
      }
    });
    const term = summaryQuery.trim().toLocaleLowerCase();
    return [...groups.values()].map((group) => {
      const totalCount = group.assets.length;
      const status = group.countedCount === totalCount ? 'counted' : group.countedCount ? 'partial' : 'pending';
      return { ...group, totalCount, status, percent: Math.ceil((group.countedCount / totalCount) * 100) };
    }).filter((group) => {
      const matchesFilter = summaryFilter === 'all' || summaryFilter === group.status || (summaryFilter === 'damaged' && group.damagedCount > 0);
      const matchesQuery = !term || group.pallet.toLocaleLowerCase().includes(term) || group.assets.some((asset) => asset.sn.includes(term));
      const matchesDate = !summaryDate || group.assets.some((asset) => counted[asset.id]?.slice(0, 10) === summaryDate);
      return matchesFilter && matchesQuery && matchesDate;
    }).sort((a, b) => a.pallet.localeCompare(b.pallet, 'th', { numeric: true }));
  }, [assets, counted, countDetails, summaryFilter, summaryQuery, summaryDate]);

  const palletTotals = useMemo(() => {
    const all = new Map();
    assets.forEach((asset) => {
      const key = asset.pallet || 'ไม่ระบุ Pallet';
      if (!all.has(key)) all.set(key, { total: 0, done: 0 });
      const item = all.get(key);
      item.total += 1;
      if (counted[asset.id]) item.done += 1;
    });
    const values = [...all.values()];
    return {
      all: values.length,
      counted: values.filter((item) => item.done === item.total).length,
      partial: values.filter((item) => item.done > 0 && item.done < item.total).length,
      pending: values.filter((item) => item.done === 0).length,
    };
  }, [assets, counted]);
  const damagedPalletTotal = useMemo(() => new Set(assets.filter((asset) => counted[asset.id] && countDetails[asset.id]?.condition === 'damaged').map((asset) => asset.pallet || 'ไม่ระบุ Pallet')).size, [assets, counted, countDetails]);
  const reportAuditRows = useMemo(() => {
    if (randomAuditRows.length) return randomAuditRows;
    return assets.filter((asset) => counted[asset.id]).map((asset) => ({ ...asset, round: 0 }));
  }, [assets, counted, randomAuditRows]);

  const randomPalletRows = useMemo(() => {
    const sampledIds = new Set(reportAuditRows.map((asset) => String(asset.id)));
    const groups = new Map();
    assets.forEach((asset) => {
      const pallet = asset.pallet || 'ไม่ระบุ Pallet';
      if (!groups.has(pallet)) groups.set(pallet, { pallet, assets: [], sampled: 0, sampledCounted: 0, good: 0, damaged: 0, outsideCounted: 0 });
    });
    reportAuditRows.forEach((asset) => {
      const pallet = asset.pallet || 'ไม่ระบุ Pallet';
      if (!groups.has(pallet)) groups.set(pallet, { pallet, assets: [], sampled: 0, sampledCounted: 0, good: 0, damaged: 0, outsideCounted: 0 });
      const group = groups.get(pallet);
      group.assets.push(asset); group.sampled += 1;
      if (counted[asset.id]) {
        group.sampledCounted += 1;
        if (countDetails[asset.id]?.condition === 'damaged') group.damaged += 1;
        if (countDetails[asset.id]?.condition === 'good') group.good += 1;
      }
    });
    assets.forEach((asset) => {
      if (!counted[asset.id] || sampledIds.has(String(asset.id))) return;
      const pallet = asset.pallet || 'ไม่ระบุ Pallet';
      if (!groups.has(pallet)) groups.set(pallet, { pallet, assets: [], sampled: 0, sampledCounted: 0, good: 0, damaged: 0, outsideCounted: 0 });
      const group = groups.get(pallet);
      group.assets.push({ ...asset, round: 0, outsidePlan: true });
      group.outsideCounted += 1;
      if (countDetails[asset.id]?.condition === 'damaged') group.damaged += 1;
      if (countDetails[asset.id]?.condition === 'good') group.good += 1;
    });
    return [...groups.values()].sort((a, b) => a.pallet.localeCompare(b.pallet, 'th', { numeric: true }));
  }, [assets, counted, countDetails, reportAuditRows]);
  const filteredRandomPalletRows = useMemo(() => {
    const term = summaryQuery.trim().toLocaleLowerCase();
    return randomPalletRows.filter((group) => {
      const matchesQuery = !term || group.pallet.toLocaleLowerCase().includes(term) || group.assets.some((asset) => asset.sn.toLocaleLowerCase().includes(term));
      const matchesDate = !summaryDate || group.assets.some((asset) => counted[asset.id]?.slice(0, 10) === summaryDate);
      const matchesStatus = randomReportView === 'assets' || randomReportFilter === 'all'
        || (randomReportFilter === 'counted' && group.sampled > 0 && group.sampledCounted === group.sampled)
        || (randomReportFilter === 'pending' && (group.sampled === 0 || group.sampledCounted < group.sampled))
        || (randomReportFilter === 'outside' && group.outsideCounted > 0);
      return matchesQuery && matchesDate && matchesStatus;
    });
  }, [randomPalletRows, summaryQuery, summaryDate, counted, randomReportFilter, randomReportView]);
  const filteredRandomReportAssets = useMemo(() => {
    const term = summaryQuery.trim().toLocaleLowerCase();
    return filteredRandomPalletRows.flatMap((group) => {
      const palletMatches = !term || group.pallet.toLocaleLowerCase().includes(term);
      return group.assets.filter((asset) => {
        const matchesQuery = palletMatches || asset.sn.toLocaleLowerCase().includes(term);
        const matchesDate = !summaryDate || counted[asset.id]?.slice(0, 10) === summaryDate;
        const matchesStatus = randomReportView === 'pallets' || randomReportFilter === 'all'
          || (randomReportFilter === 'counted' && counted[asset.id] && !asset.outsidePlan)
          || (randomReportFilter === 'pending' && !counted[asset.id])
          || (randomReportFilter === 'outside' && asset.outsidePlan);
        return matchesQuery && matchesDate && matchesStatus;
      });
    });
  }, [filteredRandomPalletRows, summaryQuery, summaryDate, counted, randomReportFilter, randomReportView]);
  const randomPalletTotals = useMemo(() => ({
    all: randomPalletRows.length,
    counted: randomPalletRows.filter((group) => group.sampled > 0 && group.sampledCounted === group.sampled).length,
    pending: randomPalletRows.filter((group) => group.sampled === 0 || group.sampledCounted < group.sampled).length,
    outside: randomPalletRows.filter((group) => group.outsideCounted > 0).length,
  }), [randomPalletRows]);
  const randomSerialTotals = useMemo(() => {
    const rows = randomPalletRows.flatMap((group) => group.assets);
    return {
      all: rows.length,
      counted: rows.filter((asset) => counted[asset.id] && !asset.outsidePlan).length,
      pending: rows.filter((asset) => !counted[asset.id]).length,
      outside: rows.filter((asset) => asset.outsidePlan).length,
    };
  }, [randomPalletRows, counted]);
  const randomReportTotals = randomReportView === 'pallets' ? randomPalletTotals : randomSerialTotals;
  const randomReportUnit = randomReportView === 'pallets' ? 'Pallet' : 'SN';

  useEffect(() => { setSummaryLimit(200); }, [summaryView, summaryFilter, summaryQuery, summaryDate]);

  useEffect(() => {
    if (!scannerOpen) return undefined;
    const codeReader = new BrowserQRCodeReader();
    let controls;
    let active = true;
    const startScanner = async () => {
      try {
        const video = scannerVideoRef.current;
        if (!video || !active) return;
        controls = await codeReader.decodeFromConstraints(
          { video: { facingMode: { ideal: 'environment' } }, audio: false },
          video,
          (result) => {
            if (!active || !result) return;
            const decodedText = result.getText();
        const value = extractSerialFromScan(decodedText);
        if (!value) {
          setStatus({ type: 'error', text: 'QR Code ไม่มี Serial Number ตัวเลข' });
              return;
        }
        const exact = assets.find((asset) => asset.sn === value);
        const matches = exact ? [] : assets.filter((asset) => asset.sn.includes(value));
        setQuery(value);
        setSearchMatches(matches);
        if (exact && !isRandomEligible(exact.id)) {
          setScannerOpen(false);
          notifyOutsideRandomAudit(exact);
          return;
        }
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
          },
        );
      } catch {
        setStatus({ type: 'error', text: 'เปิดกล้องไม่สำเร็จ กรุณาอนุญาตสิทธิ์กล้องใน Browser' });
        setScannerOpen(false);
      }
    };
    startScanner();
    return () => {
      active = false;
      controls?.stop();
    };
  }, [scannerOpen, assets, counted, countDetails, randomAuditRows]);

  const handleQueryChange = (event) => {
    const value = event.target.value;
    const searchValue = value.trim().toLocaleLowerCase();
    setQuery(value);
    setSelected(null);
    if (!value) {
      setSearchMatches([]);
      setStatus({ type: 'ready', text: `พร้อมตรวจนับ ${assets.length.toLocaleString('th-TH')} รายการ` });
      return;
    }
    const matches = assets.filter((asset) => asset.sn.toLocaleLowerCase().includes(searchValue) || asset.pallet.toLocaleLowerCase().includes(searchValue));
    setSearchMatches(matches);
    setStatus(matches.length
      ? { type: 'found', text: `พบ ${matches.length.toLocaleString('th-TH')} รายการจาก Serial Number หรือ Pallet` }
      : { type: 'warning', text: 'ยังไม่พบรายการ กรุณาตรวจสอบ Serial Number หรือ Pallet' });
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
    const partialMatches = assets.filter((asset) => asset.sn.toLocaleLowerCase().includes(term) || asset.pallet.toLocaleLowerCase().includes(term));
    const exactLocal = partialMatches.find((asset) => asset.sn === term);
    if (exactLocal) {
      if (!isRandomEligible(exactLocal.id)) {
        notifyOutsideRandomAudit(exactLocal);
        return;
      }
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
    if (exact && !isRandomEligible(exact.id)) {
      notifyOutsideRandomAudit(exact);
      return;
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
    if (!selected || counted[selected.id] || isSaving || !projectIsOpen) return;
    if (!isRandomEligible(selected.id)) {
      setStatus({ type: 'warning', text: `SN ${selected.sn} ไม่อยู่ในรายการที่สุ่มไว้ ไม่สามารถบันทึกการนับได้` });
      return;
    }
    const now = new Date().toISOString();
    setIsSaving(true);
    if (db) {
      try {
        const itemRef = doc(db, 'asset_counts', countDocumentId(selected.id));
        await setDoc(itemRef, {
          assetId: String(selected.id),
          ...(currentProjectId !== 'legacy' ? { projectId: currentProjectId } : {}),
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
    if (!selected || !counted[selected.id] || isSaving || !projectIsOpen) return;
    setIsSaving(true);
    try {
      if (db) await deleteDoc(doc(db, 'asset_counts', countDocumentId(selected.id)));
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

  const openCountDateEditor = (asset) => {
    if (!projectIsOpen) return;
    const date = new Date(counted[asset.id]);
    if (Number.isNaN(date.getTime())) return;
    const localDate = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
    setEditingCountDate({ asset, value: localDate.toISOString().slice(0, 16) });
  };

  const saveCountDate = async () => {
    if (!editingCountDate || isUpdatingCountDate || !projectIsOpen) return;
    const nextDate = new Date(editingCountDate.value);
    if (Number.isNaN(nextDate.getTime())) return;
    const assetId = String(editingCountDate.asset.id);
    const countedAt = nextDate.toISOString();
    setIsUpdatingCountDate(true);
    try {
      if (db) await updateDoc(doc(db, 'asset_counts', countDocumentId(assetId)), { countedAt });
      setCounted((current) => ({ ...current, [assetId]: countedAt }));
      if (db) setCountDetails((current) => ({
        ...current,
        [assetId]: { ...current[assetId], countedAt },
      }));
      setEditingCountDate(null);
      setStatus({ type: 'success', text: `แก้ไขวันเวลาที่นับ SN ${editingCountDate.asset.sn} สำเร็จ` });
    } catch {
      setStatus({ type: 'error', text: 'แก้ไขวันเวลาที่นับไม่สำเร็จ กรุณาตรวจสอบ Firestore Rules' });
    } finally {
      setIsUpdatingCountDate(false);
    }
  };

  const createProject = async (event) => {
    event.preventDefault();
    const name = normalize(newProjectName);
    if (!name || !newProjectFile || isSavingProject) return;
    const projectId = `project-${Date.now()}`;
    setIsSavingProject(true);
    setStatus({ type: 'loading', text: 'กำลังอ่านไฟล์และสร้างโครงการ…' });
    try {
      const buffer = await newProjectFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { raw: false, defval: '' });
      const projectAssets = rawRows.map((row, index) => ({
        id: normalize(row.ID ?? row.id ?? index + 1),
        pallet: normalize(row.pallet ?? row.Pallet),
        sn: normalize(row.SN ?? row.sn ?? row['Serial Number']),
      })).filter((row) => row.sn.length > 0);
      if (!projectAssets.length) throw new Error('NO_ASSETS');
      const ids = new Set();
      const serials = new Set();
      for (const asset of projectAssets) {
        if (ids.has(asset.id) || serials.has(asset.sn)) throw new Error('DUPLICATE_ASSETS');
        ids.add(asset.id); serials.add(asset.sn);
      }
      const targetPercentValue = Math.min(Math.max(Number(newProjectTarget) || 100, 1), 100);
      const targetCount = Math.ceil((projectAssets.length * targetPercentValue) / 100);
      const createdAt = new Date().toISOString();
      const project = { id: projectId, name, status: 'closed', createdAt, totalAssets: projectAssets.length, targetCount, targetPercent: targetPercentValue, fileName: newProjectFile.name };
      if (db) {
        await setDoc(doc(db, 'count_projects', projectId), { name, status: 'closed', createdAt, totalAssets: projectAssets.length, targetCount, targetPercent: targetPercentValue, fileName: newProjectFile.name });
        await setDoc(doc(db, 'project_data', projectId), { assets: projectAssets, totalAssets: projectAssets.length, importedAt: createdAt });
      }
      else setProjects((current) => [...current, project]);
      setNewProjectName('');
      setNewProjectFile(null);
      setNewProjectTarget('');
      setStatus({ type: 'success', text: `สร้างโครงการ “${name}” สถานะปิด พร้อมข้อมูล ${projectAssets.length.toLocaleString('th-TH')} รายการ` });
    } catch (error) {
      const message = error.message === 'NO_ASSETS' ? 'ไม่พบ Serial Number ในไฟล์ กรุณาตรวจหัวคอลัมน์ SN' : error.message === 'DUPLICATE_ASSETS' ? 'พบ ID หรือ Serial Number ซ้ำในไฟล์' : 'สร้างโครงการไม่สำเร็จ กรุณาตรวจไฟล์และ Firestore Rules';
      setStatus({ type: 'error', text: message });
    } finally { setIsSavingProject(false); }
  };

  const openProjectEditor = (project) => {
    const projectTotal = project.isLegacy ? assets.length : Number(project.totalAssets) || 0;
    const percentValue = Number(project.targetPercent) || (projectTotal ? Math.ceil(((Number(project.targetCount) || projectTotal) / projectTotal) * 100) : 100);
    setEditingProject({ ...project, nameValue: project.name, targetValue: String(percentValue) });
  };

  const saveProject = async (event) => {
    event.preventDefault();
    if (!editingProject || isSavingProject) return;
    const name = normalize(editingProject.nameValue);
    const projectTotal = editingProject.isLegacy ? assets.length : Number(editingProject.totalAssets) || 0;
    const targetPercentValue = Math.min(Math.max(Number(editingProject.targetValue) || 100, 1), 100);
    const targetCount = Math.ceil((projectTotal * targetPercentValue) / 100);
    if (!name) return;
    setIsSavingProject(true);
    try {
      const updatedAt = new Date().toISOString();
      if (db) {
        if (editingProject.isLegacy && !editingProject.managed) await setDoc(doc(db, 'count_projects', 'legacy'), { name, status: editingProject.status, createdAt: updatedAt, updatedAt, totalAssets: projectTotal, targetCount, targetPercent: targetPercentValue });
        else await updateDoc(doc(db, 'count_projects', editingProject.id), { name, targetCount, targetPercent: targetPercentValue, updatedAt });
      } else setProjects((current) => current.map((item) => item.id === editingProject.id ? { ...item, name, targetCount, targetPercent: targetPercentValue } : item));
      setEditingProject(null);
      setStatus({ type: 'success', text: `แก้ไขโครงการ “${name}” สำเร็จ` });
    } catch {
      setStatus({ type: 'error', text: 'แก้ไขโครงการไม่สำเร็จ กรุณาอัปเดต Firestore Rules' });
    } finally { setIsSavingProject(false); }
  };

  const toggleProjectStatus = async (project) => {
    if (isSavingProject) return;
    const nextStatus = project.status === 'open' ? 'closed' : 'open';
    if (nextStatus === 'open') {
      const openProject = projects.find((item) => item.id !== project.id && item.status === 'open');
      if (openProject) {
        setStatus({ type: 'error', text: `เปิดโครงการไม่ได้ กรุณาปิดโครงการ “${openProject.name}” ก่อน` });
        return;
      }
    }
    setIsSavingProject(true);
    try {
      if (db) {
        if (project.isLegacy && !project.managed) await setDoc(doc(db, 'count_projects', 'legacy'), { name: project.name, status: nextStatus, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() });
        else await updateDoc(doc(db, 'count_projects', project.id), { status: nextStatus, updatedAt: new Date().toISOString() });
      }
      else setProjects((current) => current.map((item) => item.id === project.id ? { ...item, status: nextStatus } : item));
      setStatus({ type: 'success', text: `${nextStatus === 'open' ? 'เปิด' : 'ปิด'}โครงการ “${project.name}” แล้ว` });
      if (nextStatus === 'open') setActiveProjectId(project.id);
      if (nextStatus === 'closed') setSelected(null);
    } catch {
      setStatus({ type: 'error', text: 'เปลี่ยนสถานะโครงการไม่สำเร็จ กรุณาตรวจสอบ Firestore Rules' });
    } finally { setIsSavingProject(false); }
  };

  const addAuthorizedEmail = async (event) => {
    event.preventDefault();
    const email = newAuthorizedEmail.trim().toLocaleLowerCase();
    if (!isAdmin || !/^\S+@\S+\.\S+$/.test(email) || isSavingAccess) return;
    setIsSavingAccess(true);
    try {
      await setDoc(doc(db, 'authorized_users', email), { email, addedAt: new Date().toISOString(), addedBy: normalizedUserEmail });
      setNewAuthorizedEmail('');
    } catch { setStatus({ type: 'error', text: 'เพิ่ม Email ไม่สำเร็จ กรุณาอัปเดต Firestore Rules' }); }
    finally { setIsSavingAccess(false); }
  };

  const removeAuthorizedEmail = async (email) => {
    if (!isAdmin || email === ADMIN_EMAIL || isSavingAccess) return;
    setIsSavingAccess(true);
    try { await deleteDoc(doc(db, 'authorized_users', email)); }
    catch { setStatus({ type: 'error', text: 'ลบ Email ไม่สำเร็จ กรุณาตรวจสอบ Firestore Rules' }); }
    finally { setIsSavingAccess(false); }
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
      { รายการ: 'เป้าหมายการตรวจนับ', จำนวน: targetTotal },
      { รายการ: 'นับแล้ว', จำนวน: done },
      { รายการ: 'คงเหลือ', จำนวน: remaining },
      { รายการ: 'ความคืบหน้า', จำนวน: `${percent}%` },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'ผลการตรวจนับ');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'สรุปยอด');
    XLSX.writeFile(workbook, `asset-count-${new Date().toISOString().slice(0, 10)}-${projectFileName}.xlsx`);
  };

  const exportSummaryExcel = () => {
    if (summaryView === 'pallets') {
      const rows = palletRows.map((group, index) => ({
        ลำดับ: index + 1,
        Pallet: group.pallet,
        'นับแล้ว (รายการ)': group.countedCount,
        'ทั้งหมด (รายการ)': group.totalCount,
        ความคืบหน้า: `${group.percent}%`,
        สถานะ: group.status === 'counted' ? 'ครบแล้ว' : group.status === 'partial' ? 'กำลังนับ' : 'ยังไม่เริ่ม',
        'นับล่าสุด': group.latestCountedAt ? new Date(group.latestCountedAt).toLocaleString('th-TH') : '-',
      }));
      const workbook = XLSX.utils.book_new();
      const sheet = XLSX.utils.json_to_sheet(rows);
      sheet['!cols'] = [{ wch: 8 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 15 }, { wch: 14 }, { wch: 24 }];
      XLSX.utils.book_append_sheet(workbook, sheet, 'สรุปตาม Pallet');
      XLSX.writeFile(workbook, `pallet-summary-${new Date().toISOString().slice(0, 10)}-${projectFileName}.xlsx`);
      return;
    }
    const rows = summaryRows.map((asset, index) => {
      const isCounted = Boolean(counted[asset.id]);
      return {
        ลำดับ: index + 1,
        ID: asset.id,
        Pallet: asset.pallet || '-',
        'Serial Number': asset.sn,
        สถานะ: isCounted ? 'นับแล้ว' : 'ยังไม่นับ',
        สภาพ: isCounted ? (countDetails[asset.id]?.condition === 'damaged' ? 'เสีย' : countDetails[asset.id]?.condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ') : '-',
        'วันเวลาที่นับ': isCounted ? new Date(counted[asset.id]).toLocaleString('th-TH') : '-',
      };
    });
    const filterName = summaryFilter === 'counted' ? 'นับแล้ว' : summaryFilter === 'pending' ? 'ยังไม่นับ' : summaryFilter === 'damaged' ? 'พบของเสีย' : 'ทั้งหมด';
    const summary = [
      { รายการ: 'ตัวกรองที่ส่งออก', จำนวน: filterName },
      { รายการ: 'วันที่นับ', จำนวน: summaryDate || 'ทุกวัน' },
      { รายการ: 'จำนวนในไฟล์', จำนวน: rows.length },
      { รายการ: 'จำนวนทั้งหมด', จำนวน: total },
      { รายการ: 'เป้าหมายการตรวจนับ', จำนวน: targetTotal },
      { รายการ: 'นับแล้ว', จำนวน: done },
      { รายการ: 'ยังไม่นับ', จำนวน: remaining },
    ];
    const workbook = XLSX.utils.book_new();
    const dataSheet = XLSX.utils.json_to_sheet(rows);
    dataSheet['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 14 }, { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(workbook, dataSheet, 'รายการครุภัณฑ์');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'สรุปยอด');
    const suffix = summaryFilter === 'counted' ? 'counted' : summaryFilter === 'pending' ? 'pending' : summaryFilter === 'damaged' ? 'damaged' : 'all';
    XLSX.writeFile(workbook, `asset-summary-${suffix}-${new Date().toISOString().slice(0, 10)}-${projectFileName}.xlsx`);
  };

  if (currentPage === 'access' && isAdmin) return (
    <main className="projects-page access-page"><header className="topbar projects-topbar"><button className="back-button" onClick={() => setCurrentPage('count')}>← กลับหน้าตรวจนับ</button><div><p className="eyebrow">ACCESS CONTROL</p><h1>จัดการสิทธิ์ผู้ใช้งาน</h1></div></header><section className="projects-page-content"><div className="projects-page-heading"><div><span>AUTHORIZED EMAILS</span><h2>Email ที่เข้าใช้งานได้</h2><p>ผู้ใช้ต้อง Login Google ด้วย Email ที่อยู่ในรายการนี้</p></div><strong>{(authorizedEmails.length + (authorizedEmails.includes(ADMIN_EMAIL) ? 0 : 1)).toLocaleString('th-TH')} บัญชี</strong></div><section className="access-create-card"><form onSubmit={addAuthorizedEmail}><label htmlFor="authorized-email">เพิ่ม Gmail หรือ Email ผู้ใช้งาน</label><div><input id="authorized-email" type="email" value={newAuthorizedEmail} onChange={(event) => setNewAuthorizedEmail(event.target.value)} placeholder="name@gmail.com" autoComplete="email" /><button type="submit" disabled={!/^\S+@\S+\.\S+$/.test(newAuthorizedEmail.trim()) || isSavingAccess}>＋ เพิ่มสิทธิ์</button></div></form></section><section className="access-list">{[...new Set([ADMIN_EMAIL, ...authorizedEmails])].sort().map((email) => <article key={email}><span>{email.slice(0, 1).toUpperCase()}</span><div><strong>{email}</strong><small>{email === ADMIN_EMAIL ? 'ผู้ดูแลระบบหลัก' : 'ผู้ใช้งานที่ได้รับอนุญาต'}</small></div>{email === ADMIN_EMAIL ? <b>ADMIN</b> : <button onClick={() => removeAuthorizedEmail(email)} disabled={isSavingAccess}>ลบสิทธิ์</button>}</article>)}</section></section></main>
  );

  if (currentPage === 'projects') return (
    <main className="projects-page">
      <header className="topbar projects-topbar">
        <button className="back-button" onClick={() => setCurrentPage('count')}>← กลับหน้าตรวจนับ</button>
        <div><p className="eyebrow">COUNT PROJECTS</p><h1>โครงการตรวจนับ</h1></div>
      </header>
      <section className="projects-page-content">
        <div className="projects-page-heading"><div><span>PROJECT LIST</span><h2>รายการโครงการ</h2><p>แต่ละโครงการมีชุดข้อมูล Pallet, Serial Number และผลการนับแยกจากกัน</p></div><strong>{projects.length.toLocaleString('th-TH')} โครงการ</strong></div>
        {status.type !== 'ready' && <div className={`project-page-notice ${status.type}`}><span>{status.type === 'success' ? '✓' : status.type === 'error' ? '!' : 'i'}</span><p>{status.text}</p></div>}
        <section className="project-create-card">
          <div><small>NEW PROJECT</small><h3>สร้างโครงการใหม่</h3><p>กรอกชื่อและอัปโหลดไฟล์ Excel ที่มีคอลัมน์ ID, pallet และ SN</p></div>
          <form onSubmit={createProject}>
            <label htmlFor="project-name-page">ชื่อโครงการ</label>
            <input id="project-name-page" value={newProjectName} onChange={(event) => setNewProjectName(event.target.value)} placeholder="เช่น ตรวจนับประจำปี 2569" maxLength="80" />
            <label htmlFor="project-file">ไฟล์ข้อมูล Excel</label>
            <label className={`project-file-input ${newProjectFile ? 'has-file' : ''}`} htmlFor="project-file"><span>{newProjectFile ? '✓' : '⇧'}</span><div><strong>{newProjectFile?.name || 'เลือกไฟล์ .xlsx หรือ .xls'}</strong><small>{newProjectFile ? `${(newProjectFile.size / 1024).toFixed(1)} KB` : 'หัวคอลัมน์: ID, pallet, SN'}</small></div></label>
            <input id="project-file" className="visually-hidden" type="file" accept=".xlsx,.xls" onChange={(event) => setNewProjectFile(event.target.files?.[0] || null)} />
            <label htmlFor="project-target">เปอร์เซ็นต์ที่จะนับ</label>
            <input id="project-target" type="number" min="1" max="100" value={newProjectTarget} onChange={(event) => setNewProjectTarget(event.target.value.replace(/\D/g, '').slice(0, 3))} placeholder="เช่น 30 (เว้นว่าง = 100)" />
            <button type="submit" disabled={!newProjectName.trim() || !newProjectFile || Number(newProjectTarget) > 100 || isSavingProject}>{isSavingProject ? 'กำลังอ่านไฟล์และบันทึก…' : !newProjectName.trim() ? 'กรุณากรอกชื่อโครงการ' : !newProjectFile ? 'กรุณาเลือกไฟล์ Excel' : Number(newProjectTarget) > 100 ? 'เปอร์เซ็นต์ต้องไม่เกิน 100' : '＋ สร้างและเปิดโครงการ'}</button>
          </form>
        </section>
        <section className="project-page-list">
          {projects.map((project) => <article className={`${project.id === currentProjectId ? 'active' : ''} is-${project.status}`} key={project.id}>
            <div className="project-page-icon">{project.status === 'open' ? '●' : '○'}</div>
            <div className="project-page-info"><small>{project.isLegacy ? 'LEGACY PROJECT' : 'COUNT PROJECT'}</small><h3>{project.name}</h3><p className="project-code">รหัสโครงการ: <code>{project.id}</code></p><p>{project.isLegacy ? 'ข้อมูลรายการและผลการนับเดิม' : `${Number(project.totalAssets || 0).toLocaleString('th-TH')} รายการ · ${project.fileName || 'ไฟล์ Excel'}`}</p><b>เป้าหมาย {Number(project.targetPercent) || 100}% = {(Number(project.targetCount) || Number(project.totalAssets) || (project.isLegacy ? assets.length : 0)).toLocaleString('th-TH')} รายการ (ทั้งหมด {(Number(project.totalAssets) || (project.isLegacy ? assets.length : 0)).toLocaleString('th-TH')} รายการ)</b></div>
            <span className={`project-page-status is-${project.status}`}>{project.status === 'open' ? 'เปิดอยู่' : 'ปิดแล้ว'}</span>
            <div className="project-page-actions">
              <button className="open-project-button" onClick={() => { if (project.status === 'open') { setActiveProjectId(project.id); setViewingProjectId(null); } else setViewingProjectId(project.id); setSelected(null); setQuery(''); setSearchMatches([]); setCurrentPage('count'); }}>{project.status === 'open' ? 'เปิดดูโครงการ' : 'ดูแบบอ่านอย่างเดียว'}</button>
              <button className="edit-project-button" onClick={() => openProjectEditor(project)}>แก้ไข</button>
              <button className="toggle-project-button" onClick={() => toggleProjectStatus(project)} disabled={isSavingProject}>{project.status === 'open' ? 'ปิดโครงการ' : 'เปิดโครงการ'}</button>
            </div>
          </article>)}
        </section>
      </section>
      {editingProject && <div className="date-editor-modal" role="dialog" aria-modal="true" onMouseDown={(event) => { if (event.target === event.currentTarget) setEditingProject(null); }}><form className="date-editor-panel" onSubmit={saveProject}><h3>แก้ไขโครงการ</h3><p>กำหนดชื่อและเปอร์เซ็นต์เป้าหมายที่จะตรวจนับ</p><label htmlFor="edit-project-name">ชื่อโครงการ</label><input id="edit-project-name" value={editingProject.nameValue} onChange={(event) => setEditingProject((current) => ({ ...current, nameValue: event.target.value }))} maxLength="80" /><label htmlFor="edit-project-target">เปอร์เซ็นต์ที่จะนับ</label><input id="edit-project-target" type="number" min="1" max="100" value={editingProject.targetValue} onChange={(event) => setEditingProject((current) => ({ ...current, targetValue: event.target.value.replace(/\D/g, '').slice(0, 3) }))} /><small className="project-target-hint">ข้อมูลทั้งหมด {(editingProject.isLegacy ? assets.length : Number(editingProject.totalAssets) || 0).toLocaleString('th-TH')} × {Math.min(Number(editingProject.targetValue) || 0, 100)}% = {Math.ceil(((editingProject.isLegacy ? assets.length : Number(editingProject.totalAssets) || 0) * Math.min(Number(editingProject.targetValue) || 0, 100)) / 100).toLocaleString('th-TH')} รายการ</small><div className="date-editor-actions"><button type="button" className="secondary" onClick={() => setEditingProject(null)}>ยกเลิก</button><button type="submit" className="primary" disabled={!editingProject.nameValue.trim() || !editingProject.targetValue || Number(editingProject.targetValue) > 100 || isSavingProject}>{isSavingProject ? 'กำลังบันทึก…' : 'บันทึก'}</button></div></form></div>}
    </main>
  );

  return (
    <main>
      <header className="topbar">
        <div className="brand-mark">AC</div>
        <div>
          <p className="eyebrow">ASSET CONTROL</p>
          <h1>ระบบนับครุภัณฑ์</h1>
        </div>
        <button className={`project-button ${projectIsOpen ? 'is-open' : 'is-closed'}`} onClick={() => runAuthorized(() => setCurrentPage('projects'))} disabled={!authReady || !accessReady}>
          <span className="project-dot" /> <span><small>{isViewingClosedProject ? 'กำลังดูโครงการที่ปิด' : 'โครงการปัจจุบัน'}</small><strong>{activeProject.name}</strong></span><b>{isViewingClosedProject ? 'อ่านอย่างเดียว' : projectIsOpen ? 'เปิด' : 'ปิด'}</b>
        </button>
        {currentUser && hasAccess && projectIsOpen && <button className="random-audit-button" onClick={openRandomAudit} disabled={!remaining}>⌘ <span>สุ่มตรวจ</span></button>}
        {currentUser && hasAccess && <button className="summary-button" onClick={() => setShowSummary(true)} aria-label="ดูสรุปรายการทั้งหมด">▤ <span>สรุปรายการ</span></button>}
        {isAdmin && <button className="access-button" onClick={() => setCurrentPage('access')}>♙ <span>สิทธิ์ผู้ใช้</span></button>}
        <button className={`auth-button ${currentUser ? 'signed-in' : ''}`} onClick={() => currentUser ? signOut(auth) : runAuthenticated(() => {})} disabled={!authReady} title={currentUser ? `ออกจากระบบ ${currentUser.email}` : 'เข้าสู่ระบบด้วย Google'}>{currentUser?.photoURL ? <img src={currentUser.photoURL} alt="" /> : 'G'}<span>{currentUser ? 'ออกจากระบบ' : 'เข้าสู่ระบบ'}</span></button>
      </header>

      <section className="hero">
        <div className="hero-copy">
          <span className="live-dot"></span> {db ? 'เชื่อมต่อ Firebase แบบเรียลไทม์' : 'ตรวจนับแบบเรียลไทม์บนอุปกรณ์นี้'}
          <h2>ค้นหา ตรวจสอบ<br />แล้วกดยืนยัน</h2>
          <p>กรอก Serial Number เพื่อค้นหาครุภัณฑ์ ยอดจะเพิ่มขึ้นทันทีหลังยืนยันรายการ</p>
        </div>
        <div className="hero-stats">
        <div className="today-card"><span>ยอดที่นับวันนี้</span><strong>{todayDone.toLocaleString('th-TH')}</strong><small>รายการ</small><em>คงเหลือ {targetRemaining.toLocaleString('th-TH')} เครื่อง</em></div>
        <div className="total-card">
          <div className="mobile-stats">
            <div className="primary"><span>วันนี้</span><strong>{todayDone.toLocaleString('th-TH')}</strong></div>
            <div><span>นับแล้ว</span><strong>{done.toLocaleString('th-TH')}</strong></div>
            <div><span>เหลือถึงเป้า</span><strong>{targetRemaining.toLocaleString('th-TH')}</strong></div>
            <div><span>เป้าหมาย</span><strong>{targetTotal.toLocaleString('th-TH')}</strong></div>
          </div>
          <span>ยอดที่นับแล้ว</span>
          <strong>{done.toLocaleString('th-TH')}</strong>
          <small>จากเป้าหมาย {targetTotal.toLocaleString('th-TH')} รายการ ({targetPercent}% ของข้อมูลทั้งหมด)</small>
          <em className="today-count">ยอดทั้งหมด {total.toLocaleString('th-TH')} รายการ</em>
          <div className="progress"><i style={{ width: `${percent}%` }} /></div>
          <b>{percent}% สำเร็จ</b>
        </div>
        </div>
      </section>

      <section className="dashboard">
        <div className="work-grid">
          <section className="search-card">
            <div className="section-heading"><span>01</span><div><h3>ค้นหา Serial Number หรือ Pallet</h3><p>ค้นหาได้ทั้งรหัสเต็มและบางส่วน</p></div></div>
            <form onSubmit={handleSearch}>
              <label htmlFor="sn">SERIAL NUMBER / PALLET</label>
              <div className="search-row">
                <div className="input-wrap"><span>⌕</span><input ref={inputRef} id="sn" type="text" value={query} onChange={handleQueryChange} placeholder={projectIsOpen ? 'พิมพ์ Serial Number หรือ Pallet' : 'โครงการนี้ปิดแล้ว'} autoComplete="off" inputMode="search" aria-label="ค้นหา Serial Number หรือ Pallet" disabled={!projectIsOpen} /></div>
                <button className="scan-button" type="button" onClick={() => setScannerOpen(true)} aria-label="สแกน QR Code" disabled={!projectIsOpen}>▣ <span>สแกน</span></button>
                <button type="submit" disabled={!projectIsOpen || (!db && !assets.length)}>ค้นหา</button>
              </div>
            </form>
            <div className={`notice ${status.type}`}><span>{status.type === 'success' ? '✓' : status.type === 'error' ? '!' : 'i'}</span>{status.text}</div>
            {searchMatches.length > 0 && (
              <div className="search-results">
                {searchMatches.slice(0, 50).map((asset) => (
                  <button key={asset.id} type="button" onClick={() => {
                    if (!isRandomEligible(asset.id)) {
                      notifyOutsideRandomAudit(asset);
                      return;
                    }
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
                    <i>{counted[asset.id] ? 'นับแล้ว' : !isRandomEligible(asset.id) ? 'นอกแผน' : 'เลือก'}</i>
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
                  <button className="confirm-button cancel-button" onClick={cancelCount} disabled={isSaving || !projectIsOpen}>× {!projectIsOpen ? 'โครงการปิดแล้ว' : isSaving ? 'กำลังยกเลิก…' : 'ยกเลิกการนับรายการนี้'}</button>
                ) : (
                  <button className="confirm-button" onClick={confirmCount} disabled={isSaving || !projectIsOpen || !isRandomEligible(selected.id)}>✓ {!projectIsOpen ? 'โครงการปิดแล้ว' : !isRandomEligible(selected.id) ? 'ไม่อยู่ในรายการสุ่ม' : isSaving ? 'กำลังบันทึก…' : 'ยืนยันนับรายการ'}</button>
                )}
              </div>
            ) : (
              <div className="empty-state"><span>✓</span><p>ข้อมูลรายการจะแสดงที่นี่<br />หลังจากค้นหา Serial Number</p></div>
            )}
          </section>
        </div>

        <section className="recent-card">
          <div className="recent-header"><div><h3>รายการที่นับล่าสุด</h3><p>แสดง 5 รายการล่าสุดบนอุปกรณ์นี้</p></div></div>
          {done ? <div className="recent-list">{countedAssets.sort((a,b) => new Date(counted[b.id]) - new Date(counted[a.id])).slice(0,5).map((asset) => <div className="recent-row" key={asset.id}><span className="check">✓</span><div><strong>{asset.sn}</strong><small>Pallet {asset.pallet || '-'}</small></div><time>{new Date(counted[asset.id]).toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })}</time></div>)}</div> : <p className="no-records">ยังไม่มีรายการที่ยืนยันการนับ</p>}
        </section>
      </section>
      <footer>{db ? 'ยอดรวมเชื่อมต่อ Firebase แบบ Real-time และอัปเดตทุกอุปกรณ์โดยไม่ต้อง Refresh' : 'โหมด Local: ข้อมูลการนับบันทึกในเบราว์เซอร์ของอุปกรณ์นี้'}</footer>

      {showRandomAudit && <div className="random-audit-modal" role="dialog" aria-modal="true" aria-labelledby="random-audit-title" onMouseDown={(event) => { if (event.target === event.currentTarget) setShowRandomAudit(false); }}><section className="random-audit-panel">
        <header><div><small>RANDOM AUDIT</small><h2 id="random-audit-title">สุ่มรายการตรวจนับ</h2><p>สุ่มจากรายการที่ยังไม่นับ และกระจายจำนวนให้แต่ละ Pallet ใกล้เคียงกัน</p></div><button onClick={() => setShowRandomAudit(false)} aria-label="ปิด">×</button></header>
        <div className="random-audit-config"><div><span>เป้าหมายโครงการ</span><strong>{targetTotal.toLocaleString('th-TH')}</strong><small>รายการ</small></div><div><span>นับแล้ว</span><strong>{done.toLocaleString('th-TH')}</strong><small>รายการ</small></div><div><span>สุ่มได้ไม่เกิน</span><strong>{randomAvailableTarget.toLocaleString('th-TH')}</strong><small>รายการ</small></div></div>
        <div className="random-mode-tabs"><button className={randomAuditMode === 'rounds' ? 'active' : ''} onClick={() => setRandomAuditMode('rounds')}>สุ่มแบบรอบ</button><button className={randomAuditMode === 'all' ? 'active' : ''} onClick={() => setRandomAuditMode('all')}>สุ่มทั้งหมด</button></div>
        <form id="random-audit-form" className="random-round-form" onSubmit={(event) => { event.preventDefault(); generateRandomAudit(); }}>
          {randomAuditMode === 'rounds' ? <><div className="random-round-heading"><label htmlFor="round-count">จำนวนรอบ</label><input id="round-count" type="number" min="1" max="20" value={randomRoundCount} onChange={(event) => changeRandomRoundCount(event.target.value)} /><span>ผลรวม <strong className={randomRoundTotal > randomAvailableTarget ? 'over' : ''}>{randomRoundTotal.toLocaleString('th-TH')}</strong> / {randomAvailableTarget.toLocaleString('th-TH')} รายการ</span></div><div className="random-round-inputs">{randomRoundSizes.map((value, index) => <label key={index}><span>รอบ {index + 1}</span><input type="number" min="0" value={value} onChange={(event) => { const next = [...randomRoundSizes]; next[index] = event.target.value.replace(/\D/g, ''); setRandomRoundSizes(next); }} /><small>เครื่อง</small></label>)}</div></> : <div className="random-all-message"><strong>สุ่มทั้งหมด {randomAvailableTarget.toLocaleString('th-TH')} รายการ</strong><span>จากรายการที่ยังไม่นับตามยอดเป้าหมายโครงการ</span></div>}
        </form>
        <div className="random-audit-actions"><div><button className="generate-random-button" type="submit" form="random-audit-form" disabled={isSavingRandomAudit || (!randomAvailableTarget && !done) || (randomAuditMode === 'rounds' && randomAvailableTarget > 0 && (!randomRoundTotal || randomRoundTotal > randomAvailableTarget))}>⌘ {isSavingRandomAudit ? 'กำลังบันทึก…' : randomAuditMode === 'all' ? 'สุ่มทั้งหมด' : 'สุ่มตามรอบ'}</button>{randomAuditRows.some((asset) => asset.round > 0) && <button className="clear-random-button" type="button" onClick={clearRandomAudit} disabled={isSavingRandomAudit}>ล้างเลขสุ่ม</button>}</div>{randomAuditRows.some((asset) => asset.round > 0) && <button className="export-random-button" type="button" onClick={exportRandomAudit}>⇩ Excel</button>}</div>
        {randomAuditRows.some((asset) => asset.round > 0) ? <><div className="random-audit-result-header"><p>รายการสุ่มใหม่ <strong>{randomAuditRows.filter((asset) => asset.round > 0).length.toLocaleString('th-TH')}</strong> รายการ</p></div><div className="random-audit-table-wrap"><table className="asset-table"><thead><tr><th>ลำดับ</th><th>รอบ</th><th>Pallet</th><th>Serial Number</th><th>สถานะ</th><th>สภาพ</th></tr></thead><tbody>{randomAuditRows.filter((asset) => asset.round > 0).map((asset, index) => { const isCounted = Boolean(counted[asset.id]); const condition = countDetails[asset.id]?.condition; return <tr key={asset.id}><td>{index + 1}</td><td>{asset.round}</td><td><strong>{asset.pallet || '-'}</strong></td><td><strong>{asset.sn}</strong><small>ID: {asset.id}</small></td><td><span className={`status-pill ${isCounted ? 'is-counted' : 'is-pending'}`}>{isCounted ? '✓ นับแล้ว' : '– ยังไม่นับ'}</span></td><td>{!isCounted ? '-' : condition === 'damaged' ? 'เสีย' : condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</td></tr>; })}</tbody></table></div></> : <div className="random-audit-empty"><span>⌘</span><p>กำหนดรอบแล้วกดสุ่ม<br />หน้านี้จะแสดงเฉพาะรายการสุ่มใหม่</p></div>}
      </section></div>}
      {showRandomReport && <div className="random-audit-modal" role="dialog" aria-modal="true" onMouseDown={(event) => { if (event.target === event.currentTarget) setShowRandomReport(false); }}><section className="random-audit-panel random-report-panel"><header><div><small>RANDOM AUDIT REPORT</small><h2>รายงานการสุ่มและตรวจนับ</h2><p>ติดตามผลตาม Pallet รวมรายการในแผนและรายการที่นับนอกแผน</p></div><button onClick={() => setShowRandomReport(false)}>×</button></header><div className="random-report-totals"><div><span>รายการในแผน</span><strong>{reportAuditRows.length.toLocaleString('th-TH')}</strong></div><div><span>นับแล้วในแผน</span><strong>{reportAuditRows.filter((asset) => counted[asset.id]).length.toLocaleString('th-TH')}</strong></div><div><span>นับนอกแผน</span><strong>{randomPalletRows.reduce((sum, group) => sum + group.outsideCounted, 0).toLocaleString('th-TH')}</strong></div><button onClick={exportRandomReport}>⇩ Export Excel</button></div><div className="random-pallet-card-wrap">{randomPalletRows.map((group) => <article className="random-pallet-card" key={group.pallet} onClick={() => setSelectedRandomPallet(group)}><header><div><small>PALLET</small><h3>{group.pallet}</h3></div><span>{group.sampledCounted}/{group.sampled}</span></header><div className="random-pallet-progress"><i style={{ width: `${group.sampled ? Math.ceil((group.sampledCounted / group.sampled) * 100) : 0}%` }} /></div><div className="random-pallet-stats"><span><b>{group.sampled}</b>ในแผน</span><span><b>{group.sampledCounted}</b>นับแล้ว</span><span><b>{group.outsideCounted}</b>นอกแผน</span><span className="good"><b>{group.good}</b>ไม่เสีย</span><span className="damaged"><b>{group.damaged}</b>เสีย</span></div><button>ดูรายการใน Pallet →</button></article>)}</div></section></div>}
      {selectedRandomPallet && <div className="pallet-detail-modal" role="dialog" aria-modal="true" onMouseDown={(event) => { if (event.target === event.currentTarget) setSelectedRandomPallet(null); }}><section className="pallet-detail-panel"><header><div><small>RANDOM AUDIT DETAIL</small><h2>Pallet {selectedRandomPallet.pallet}</h2><p>เลขที่สุ่ม {selectedRandomPallet.sampled.toLocaleString('th-TH')} · นับแล้ว {selectedRandomPallet.sampledCounted.toLocaleString('th-TH')} · นอกแผน {selectedRandomPallet.outsideCounted.toLocaleString('th-TH')}</p></div><button onClick={() => setSelectedRandomPallet(null)}>×</button></header><div className="pallet-detail-table-wrap"><table className="asset-table"><thead><tr><th>รอบ</th><th>Serial Number</th><th>สถานะ</th><th>สภาพ</th><th>เวลาที่นับ</th></tr></thead><tbody>{selectedRandomPallet.assets.sort((a, b) => a.round - b.round || a.sn.localeCompare(b.sn, 'th', { numeric: true })).map((asset) => { const isCounted = Boolean(counted[asset.id]); const condition = countDetails[asset.id]?.condition; return <tr key={asset.id}><td>{asset.round === 0 ? 'ก่อนสุ่ม' : asset.round}</td><td><strong>{asset.sn}</strong><small>ID: {asset.id}</small></td><td><span className={`status-pill ${isCounted ? 'is-counted' : 'is-pending'}`}>{isCounted ? '✓ นับแล้ว' : '– ยังไม่นับ'}</span></td><td>{!isCounted ? '-' : condition === 'damaged' ? 'เสีย' : condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</td><td>{isCounted ? new Date(counted[asset.id]).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' }) : '-'}</td></tr>; })}</tbody></table></div></section></div>}

      {showSummary && (
        <div className="summary-modal" role="dialog" aria-modal="true" aria-labelledby="summary-title">
          <div className={`summary-panel view-${summaryView}`}>
            <header className="summary-panel-header">
              <div><p>ASSET OVERVIEW</p><h2 id="summary-title">สรุปผลการตรวจนับ</h2></div>
              <button onClick={() => setShowSummary(false)} aria-label="ปิดหน้าสรุป">×</button>
            </header>
            <div className="summary-view-tabs">
              <button className={summaryView === 'pallets' ? 'active' : ''} onClick={() => { setSummaryView('pallets'); setSummaryFilter('all'); }}>▦ ตาม Pallet</button>
              <button className={summaryView === 'assets' ? 'active' : ''} onClick={() => { setSummaryView('assets'); setSummaryFilter('all'); }}>☷ ราย Serial Number</button>
              <button className={summaryView === 'randomReport' ? 'active random-report-tab' : 'random-report-tab'} onClick={() => setSummaryView('randomReport')} disabled={!randomAuditRows.length && !done}>⌘ รายงานสุ่มและตรวจนับ</button>
            </div>
            {summaryView === 'randomReport' && <div className="summary-random-report">
              <div className="random-report-totals pallet-filter-totals">
                <button className={randomReportFilter === 'all' ? 'active' : ''} onClick={() => setRandomReportFilter('all')}><span>ทั้งหมด ({randomReportUnit})</span><strong>{randomReportTotals.all.toLocaleString('th-TH')}</strong></button>
                <button className={randomReportFilter === 'counted' ? 'active counted' : 'counted'} onClick={() => setRandomReportFilter('counted')}><span>นับแล้ว ({randomReportUnit})</span><strong>{randomReportTotals.counted.toLocaleString('th-TH')}</strong></button>
                <button className={randomReportFilter === 'pending' ? 'active pending' : 'pending'} onClick={() => setRandomReportFilter('pending')}><span>นับไม่เสร็จ ({randomReportUnit})</span><strong>{randomReportTotals.pending.toLocaleString('th-TH')}</strong></button>
                <button className={randomReportFilter === 'outside' ? 'active outside' : 'outside'} onClick={() => setRandomReportFilter('outside')}><span>นับนอกแผน ({randomReportUnit})</span><strong>{randomReportTotals.outside.toLocaleString('th-TH')}</strong></button>
              </div>
              <div className="summary-tools random-report-tools"><div className="summary-search">⌕<input value={summaryQuery} onChange={(event) => setSummaryQuery(event.target.value)} placeholder="ค้นหา Pallet หรือ Serial Number" /></div><div className="date-filter"><label htmlFor="random-report-date">วันที่นับ</label><input id="random-report-date" type="date" value={summaryDate} onChange={(event) => setSummaryDate(event.target.value)} />{summaryDate && <button onClick={() => setSummaryDate('')} aria-label="ล้างวันที่">×</button>}</div><span>พบ {(randomReportView === 'pallets' ? filteredRandomPalletRows.length : filteredRandomReportAssets.length).toLocaleString('th-TH')} {randomReportView === 'pallets' ? 'Pallet' : 'รายการ'}</span><div className="random-report-view-toggle"><button className={randomReportView === 'pallets' ? 'active' : ''} onClick={() => setRandomReportView('pallets')}>▦ Pallet</button><button className={randomReportView === 'assets' ? 'active' : ''} onClick={() => setRandomReportView('assets')}>☷ Serial Number</button></div><button className="summary-export" onClick={exportRandomReport} disabled={!filteredRandomReportAssets.length}>⇩ Export Excel</button></div>
              {randomReportView === 'pallets' ? <div className="random-pallet-card-wrap">{filteredRandomPalletRows.map((group) => <article className="random-pallet-card" key={group.pallet} onClick={() => setSelectedRandomPallet(group)}><header><div><small>PALLET</small><h3>{group.pallet}</h3></div><span>{group.sampledCounted}/{group.sampled}</span></header><div className="random-pallet-progress"><i style={{ width: `${group.sampled ? Math.ceil((group.sampledCounted / group.sampled) * 100) : 0}%` }} /></div><div className="random-pallet-stats"><span><b>{group.sampled}</b>ในแผน</span><span><b>{group.sampledCounted}</b>นับแล้ว</span><span><b>{group.outsideCounted}</b>นอกแผน</span><span className="good"><b>{group.good}</b>ไม่เสีย</span><span className="damaged"><b>{group.damaged}</b>เสีย</span></div><button>ดูรายการใน Pallet →</button></article>)}</div> : <div className="asset-table-wrap random-report-asset-table"><table className="asset-table"><thead><tr><th>ลำดับ</th><th>รอบ</th><th>Pallet</th><th>Serial Number</th><th>สถานะ</th><th>สภาพ</th><th>เวลาที่นับ</th></tr></thead><tbody>{filteredRandomReportAssets.map((asset, index) => { const isCounted = Boolean(counted[asset.id]); const condition = countDetails[asset.id]?.condition; return <tr key={asset.id}><td>{index + 1}</td><td>{asset.round === 0 ? 'ก่อนหน้า' : asset.round}</td><td><strong>{asset.pallet || '-'}</strong></td><td><strong>{asset.sn}</strong><small>ID: {asset.id}</small></td><td><span className={`status-pill ${isCounted ? 'is-counted' : 'is-pending'}`}>{isCounted ? '✓ นับแล้ว' : '– ยังไม่นับ'}</span></td><td>{!isCounted ? '-' : condition === 'damaged' ? 'เสีย' : condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</td><td>{isCounted ? new Date(counted[asset.id]).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' }) : '-'}</td></tr>; })}</tbody></table>{!filteredRandomReportAssets.length && <div className="summary-empty">ไม่พบรายการ</div>}</div>}
            </div>}
            <div className="summary-totals">
              <button className={summaryFilter === 'all' ? 'active' : ''} onClick={() => setSummaryFilter('all')}><span>ทั้งหมด</span><strong>{(summaryView === 'pallets' ? palletTotals.all : total).toLocaleString('th-TH')}</strong></button>
              <button className={summaryFilter === 'counted' ? 'active counted' : 'counted'} onClick={() => setSummaryFilter('counted')}><span>{summaryView === 'pallets' ? 'ครบแล้ว' : 'นับแล้ว'}</span><strong>{(summaryView === 'pallets' ? palletTotals.counted : done).toLocaleString('th-TH')}</strong></button>
              {summaryView === 'pallets' && <button className={summaryFilter === 'partial' ? 'active partial' : 'partial'} onClick={() => setSummaryFilter('partial')}><span>กำลังนับ</span><strong>{palletTotals.partial.toLocaleString('th-TH')}</strong></button>}
              <button className={summaryFilter === 'pending' ? 'active pending' : 'pending'} onClick={() => setSummaryFilter('pending')}><span>{summaryView === 'pallets' ? 'ยังไม่เริ่ม' : 'ยังไม่นับ'}</span><strong>{(summaryView === 'pallets' ? palletTotals.pending : remaining).toLocaleString('th-TH')}</strong></button>
              {summaryView === 'pallets' && <button className={summaryFilter === 'damaged' ? 'active damaged' : 'damaged'} onClick={() => setSummaryFilter('damaged')}><span>พบของเสีย</span><strong>{damagedPalletTotal.toLocaleString('th-TH')}</strong></button>}
            </div>
            <div className="summary-tools">
              <div className="summary-search">⌕<input value={summaryQuery} onChange={(event) => setSummaryQuery(summaryView === 'assets' ? event.target.value.replace(/\D/g, '') : event.target.value)} inputMode={summaryView === 'assets' ? 'numeric' : 'search'} placeholder={summaryView === 'pallets' ? 'ค้นหา Pallet หรือ Serial Number' : 'ค้นหา Serial Number'} /></div>
              <div className="date-filter"><label htmlFor="count-date">วันที่นับ</label><input id="count-date" type="date" value={summaryDate} onChange={(event) => setSummaryDate(event.target.value)} />{summaryDate && <button onClick={() => setSummaryDate('')} aria-label="ล้างวันที่">×</button>}</div>
              <span>พบ {(summaryView === 'pallets' ? palletRows.length : summaryRows.length).toLocaleString('th-TH')} รายการ</span>
              <button className="summary-export" onClick={exportSummaryExcel} disabled={summaryView === 'pallets' ? !palletRows.length : !summaryRows.length}>⇩ Export Excel</button>
            </div>
            <div className={summaryView === 'pallets' ? 'pallet-card-wrap' : 'asset-table-wrap'}>
              {summaryView === 'pallets' ? <div className="pallet-card-grid">
                {palletRows.map((group) => <article className={`pallet-card is-${group.status}`} key={group.pallet} role="button" tabIndex="0" onClick={() => setSelectedPalletSummary(group)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') setSelectedPalletSummary(group); }}>
                  <div className="pallet-card-heading">
                    <div><small>PALLET</small><h3>{group.pallet}</h3></div>
                    <span className={`status-pill is-${group.status}`}>{group.status === 'counted' ? '✓ ครบแล้ว' : group.status === 'partial' ? '◐ กำลังนับ' : '– ยังไม่เริ่ม'}</span>
                  </div>
                  <div className="pallet-card-count"><strong>{group.countedCount.toLocaleString('th-TH')}</strong><span>/ {group.totalCount.toLocaleString('th-TH')} รายการ</span><b>{group.percent}%</b></div>
                  <div className="pallet-card-progress"><i style={{ width: `${group.percent}%` }} /></div>
                  <div className="pallet-condition-summary"><span className="good">✓ ไม่เสีย <b>{group.goodCount.toLocaleString('th-TH')}</b></span><span className="damaged">! เสีย <b>{group.damagedCount.toLocaleString('th-TH')}</b></span></div>
                  <footer>{group.latestCountedAt ? <>นับล่าสุด <time>{new Date(group.latestCountedAt).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' })}</time></> : 'ยังไม่มีรายการที่นับ'}</footer>
                  <button className="pallet-detail-button" type="button">ดูรายการใน Pallet →</button>
                </article>)}
              </div> : <>
              <table className="asset-table">
                <thead><tr><th>ลำดับ</th><th>Pallet</th><th>Serial Number</th><th>สถานะ</th><th>สภาพ</th><th>เวลาที่นับ</th></tr></thead>
                <tbody>{summaryRows.slice(0, summaryLimit).map((asset, index) => {
                  const isCounted = Boolean(counted[asset.id]);
                  const condition = countDetails[asset.id]?.condition;
                  return <tr key={asset.id}><td>{index + 1}</td><td><strong>{asset.pallet || '-'}</strong></td><td>{isCounted ? <button className="serial-edit-button" onClick={() => openCountDateEditor(asset)} title="คลิกเพื่อแก้ไขวันเวลาที่นับ"><strong>{asset.sn}</strong><small>แตะเพื่อแก้วันนับ</small></button> : <strong>{asset.sn}</strong>}</td><td><span className={`status-pill ${isCounted ? 'is-counted' : 'is-pending'}`}>{isCounted ? '✓ นับแล้ว' : '– ยังไม่นับ'}</span></td><td><span className={`condition-pill ${condition === 'damaged' ? 'is-damaged' : condition === 'good' ? 'is-good' : ''}`}>{!isCounted ? '-' : condition === 'damaged' ? 'เสีย' : condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</span></td><td>{isCounted ? <button className="count-date-button" onClick={() => openCountDateEditor(asset)} title="คลิกเพื่อแก้ไขวันเวลาที่นับ">{new Date(counted[asset.id]).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' })}</button> : '-'}</td></tr>;
                })}</tbody>
              </table>
              </>}
              {!(summaryView === 'pallets' ? palletRows.length : summaryRows.length) && <div className="summary-empty">ไม่พบรายการ</div>}
            </div>
            {summaryView === 'assets' && summaryRows.length > summaryLimit && <button className="load-more" onClick={() => setSummaryLimit((value) => value + 200)}>แสดงเพิ่มอีก {Math.min(200, summaryRows.length - summaryLimit).toLocaleString('th-TH')} รายการ</button>}
          </div>
        </div>
      )}
      {selectedPalletSummary && (
        <div className="pallet-detail-modal" role="dialog" aria-modal="true" aria-labelledby="pallet-detail-title" onMouseDown={(event) => { if (event.target === event.currentTarget) setSelectedPalletSummary(null); }}>
          <section className="pallet-detail-panel">
            <header><div><small>PALLET DETAIL</small><h2 id="pallet-detail-title">Pallet {selectedPalletSummary.pallet}</h2><p>นับแล้ว {selectedPalletSummary.countedCount.toLocaleString('th-TH')} จาก {selectedPalletSummary.totalCount.toLocaleString('th-TH')} รายการ · {selectedPalletSummary.percent}%</p></div><button onClick={() => setSelectedPalletSummary(null)} aria-label="ปิด">×</button></header>
            <div className="pallet-detail-progress"><i style={{ width: `${selectedPalletSummary.percent}%` }} /></div>
            <div className="pallet-detail-table-wrap"><table className="asset-table"><thead><tr><th>ลำดับ</th><th>Serial Number</th><th>สถานะ</th><th>สภาพ</th><th>เวลาที่นับ</th></tr></thead><tbody>{selectedPalletSummary.assets.map((asset, index) => { const isCounted = Boolean(counted[asset.id]); const condition = countDetails[asset.id]?.condition; return <tr key={asset.id}><td>{index + 1}</td><td><strong>{asset.sn}</strong><small>ID: {asset.id}</small></td><td><span className={`status-pill ${isCounted ? 'is-counted' : 'is-pending'}`}>{isCounted ? '✓ นับแล้ว' : '– ยังไม่นับ'}</span></td><td><span className={`condition-pill ${condition === 'damaged' ? 'is-damaged' : condition === 'good' ? 'is-good' : ''}`}>{!isCounted ? '-' : condition === 'damaged' ? 'เสีย' : condition === 'good' ? 'ไม่เสีย' : 'ไม่ระบุ'}</span></td><td>{isCounted ? new Date(counted[asset.id]).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' }) : '-'}</td></tr>; })}</tbody></table></div>
          </section>
        </div>
      )}
      {editingCountDate && (
        <div className="date-editor-modal" role="dialog" aria-modal="true" aria-labelledby="date-editor-title" onMouseDown={(event) => { if (event.target === event.currentTarget && !isUpdatingCountDate) setEditingCountDate(null); }}>
          <div className="date-editor-panel">
            <h3 id="date-editor-title">แก้ไขวันเวลาที่นับ</h3>
            <p>Serial Number <strong>{editingCountDate.asset.sn}</strong></p>
            <label htmlFor="edit-count-date">วันและเวลาที่นับ</label>
            <input id="edit-count-date" type="datetime-local" value={editingCountDate.value} onChange={(event) => setEditingCountDate((current) => ({ ...current, value: event.target.value }))} />
            <div className="date-editor-actions">
              <button type="button" className="secondary" onClick={() => setEditingCountDate(null)} disabled={isUpdatingCountDate}>ยกเลิก</button>
              <button type="button" className="primary" onClick={saveCountDate} disabled={!editingCountDate.value || isUpdatingCountDate}>{isUpdatingCountDate ? 'กำลังบันทึก…' : 'บันทึก'}</button>
            </div>
          </div>
        </div>
      )}
      {outsideAuditAsset && (
        <div className="date-editor-modal" role="dialog" aria-modal="true" aria-labelledby="outside-audit-title" onMouseDown={(event) => { if (event.target === event.currentTarget) dismissOutsideAuditAlert(); }}>
          <div className="date-editor-panel outside-audit-panel">
            <h3 id="outside-audit-title">ไม่อยู่ในรายการสุ่ม</h3>
            <p>SN <strong>{outsideAuditAsset.sn}</strong> ไม่ได้อยู่ในรายการที่สุ่มไว้ ไม่สามารถตรวจนับได้</p>
            <div className="date-editor-actions outside-audit-actions">
              <button type="button" className="danger" onClick={dismissOutsideAuditAlert}>ตกลง</button>
            </div>
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

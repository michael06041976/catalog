import React, { useState, useEffect, useRef } from 'react';
import {
  LayoutGrid, Plus, Building2, Save, Download, FileSpreadsheet,
  FileText, Search, FileEdit, ImagePlus, X, Upload, Pencil, Trash2,
  PackageSearch, Check, DatabaseBackup, UploadCloud, FileUp, Shield,
  PlayCircle, Loader2, CloudCog, CloudLightning, Menu, List, ArrowDownWideNarrow, ArrowUpNarrowWide, Sparkles
} from 'lucide-react';
import { Routes, Route, Navigate, NavLink, useLocation, useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { collection, doc, onSnapshot, setDoc, deleteDoc, serverTimestamp } from 'firebase/firestore';
import { db, auth, uploadProductImage } from './firebase';
import Admin from './Admin';
import { useGeminiProduct } from './useGeminiProduct';
import { useSemanticSearch, cosineSimilarity } from './useSemanticSearch';

interface Product {
  id: string; // Changed to string for Firestore
  internalId: string;
  sku: string;
  desc: string;
  supplier: string;
  category?: string;
  price?: number | string;
  stock?: number | string;
  prices?: Record<string, number | string>;
  img: string | null;
  created: string;
  updatedAt?: any;
  embedding?: number[];
}

interface Company {
  name: string;
  hp?: string;
  addr: string;
  phone: string;
  email: string;
  web: string;
  logo: string | null;
  updatedAt?: any;
}

interface PriceListDef {
  id: string;
  name: string;
}

export default function MainApp({ userRole, userMenus }: { userRole: string, userMenus?: string[] | null }) {
  // State
  const [products, setProducts] = useState<Product[]>([]);
  const [company, setCompany] = useState<Company>({ name: '', hp: '', addr: '', phone: '', email: '', web: '', logo: null });
  const [priceLists, setPriceLists] = useState<PriceListDef[]>([]);
  
  const navigate = useNavigate();
  const location = useLocation();

  // ... rest of the state
  const [search, setSearch] = useState('');
  const [searchMode, setSearchMode] = useState<'classic' | 'ai'>('classic');
  const [searchQueryEmbedding, setSearchQueryEmbedding] = useState<number[] | null>(null);
  const { getEmbedding, isSearchingSemantic } = useSemanticSearch();

  useEffect(() => {
    if (searchMode !== 'ai') {
      setSearchQueryEmbedding(null);
      return;
    }
    const timer = setTimeout(async () => {
      if (search.trim()) {
        const emb = await getEmbedding(search);
        setSearchQueryEmbedding(emb);
      } else {
        setSearchQueryEmbedding(null);
      }
    }, 800);
    return () => clearTimeout(timer);
  }, [search, searchMode, getEmbedding]);

  const [categoryFilter, setCategoryFilter] = useState('');
  const [supplierFilter, setSupplierFilter] = useState('');
  const [dateStartFilter, setDateStartFilter] = useState('');
  const [dateEndFilter, setDateEndFilter] = useState('');
  const [editId, setEditId] = useState<string | null>(null);
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [fImgFile, setFImgFile] = useState<File | null>(null);
  const [syncStatus, setSyncStatus] = useState<'synced' | 'syncing' | 'error'>('synced');
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  
  const [newPriceListName, setNewPriceListName] = useState('');
  const [editingPriceListId, setEditingPriceListId] = useState<string | null>(null);
  const [editingPriceListName, setEditingPriceListName] = useState('');
  const [priceListToDelete, setPriceListToDelete] = useState<string | null>(null);
  
  const [catalogView, setCatalogView] = useState<'grid' | 'list'>('grid');
  const [displayPriceId, setDisplayPriceId] = useState<string>('none');
  const [sortConfig, setSortConfig] = useState<{ key: keyof Product; direction: 'asc' | 'desc' } | null>(null);

  // Excel Import State
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [importData, setImportData] = useState<any[]>([]);
  const [importHeaders, setImportHeaders] = useState<string[]>([]);
  const [fieldMapping, setFieldMapping] = useState<Record<string, string>>({});

  const [importPlModalOpen, setImportPlModalOpen] = useState(false);
  const [plImportData, setPlImportData] = useState<any[]>([]);
  const [plImportHeaders, setPlImportHeaders] = useState<string[]>([]);
  const [plFieldMapping, setPlFieldMapping] = useState<Record<string, string>>({});
  const [importingPl, setImportingPl] = useState<PriceListDef | null>(null);

  // Form State
  const [fId, setFId] = useState('');
  const [fSku, setFSku] = useState('');
  const [fDesc, setFDesc] = useState('');
  const [fSup, setFSup] = useState('');
  const [fCategory, setFCategory] = useState('');
  const [fImg, setFImg] = useState<string | null>(null);
  const [fPrice, setFPrice] = useState<number | string>('');
  const [fStock, setFStock] = useState<number | string>('');
  const [fPrices, setFPrices] = useState<Record<string, number | string>>({});

  const [cName, setCName] = useState('');
  const [cHp, setCHp] = useState('');
  const [cAddr, setCAddr] = useState('');
  const [cPhone, setCPhone] = useState('');
  const [cEmail, setCEmail] = useState('');
  const [cWeb, setCWeb] = useState('');
  const [cLogo, setCLogo] = useState<string | null>(null);

  const [aiGeneratedFields, setAiGeneratedFields] = useState<Record<string, boolean>>({});
  const { analyzeImage, isAnalyzing, error: aiError } = useGeminiProduct(fImg);

  const handleGenerateAI = async () => {
    if (!fImg || !fImg.startsWith('data:')) {
      showToast('אנא העלה תמונה חדשה תחילה');
      return;
    }
    const result = await analyzeImage();
    if (result) {
      setAiGeneratedFields({ sku: !!result.sku, desc: !!result.desc, category: !!result.category, supplier: !!result.supplier });
      if (result.sku) setFSku(result.sku);
      if (result.desc) setFDesc(result.desc);
      if (result.category) setFCategory(result.category);
      if (result.supplier) setFSup(result.supplier);
      showToast('✓ הנתונים מולאו בהצלחה בעזרת בינה מלאכותית');
    } else {
      showToast('⚠️ הפעולה נכשלה, נסה שוב או נסה תמונה אחרת');
    }
  };

  const isAdmin = userRole === 'admin';
  const isEditor = isAdmin || userRole === 'editor';

  const canSeeTab = (tabId: string) => {
    // If specific menus are set, follow them strictly
    if (userMenus && userMenus.length > 0) {
      // admin user allows seeing everything as fallback in case list is corrupted, but if menus is explicit, respect it.
      if (isAdmin && tabId === 'admin') return true; // always allow admin to see admin tab to prevent lockout
      return userMenus.includes(tabId);
    }
    // Fallback logic
    if (tabId === 'catalog') return true;
    if (tabId === 'add') return isEditor;
    if (tabId === 'company' || tabId === 'admin' || tabId === 'pricelists') return isAdmin;
    return false;
  };

  // Init Data from Firestore
  useEffect(() => {
    const unsubProducts = onSnapshot(collection(db, 'products'), { includeMetadataChanges: true }, (snapshot) => {
      setSyncStatus(snapshot.metadata.hasPendingWrites ? 'syncing' : 'synced');
      setProducts(snapshot.docs.map(d => ({ id: d.id, ...d.data() } as Product)));
    }, err => {
      console.error("Products err", err);
      setSyncStatus('error');
    });

    const unsubCompany = onSnapshot(doc(db, 'settings', 'company'), { includeMetadataChanges: true }, (d) => {
      if (d.metadata.hasPendingWrites) setSyncStatus('syncing');
      else setSyncStatus('synced');
      
      if (d.exists()) {
        const co = d.data() as Company;
        setCompany(co);
        setCName(co.name || '');
        setCHp(co.hp || '');
        setCAddr(co.addr || '');
        setCPhone(co.phone || '');
        setCEmail(co.email || '');
        setCWeb(co.web || '');
        setCLogo(co.logo || null);
      }
    }, err => console.error("Company err", err));

    const unsubPriceLists = onSnapshot(doc(db, 'settings', 'pricelists'), (d) => {
      if (d.exists()) {
        setPriceLists(d.data().lists || []);
      }
    });

    return () => {
      unsubProducts();
      unsubCompany();
      unsubPriceLists();
    };
  }, []);

  const handleFirestoreError = (err: any) => {
    console.error(err);
    if (err.code === 'permission-denied') {
       showToast('⚠️ אין לך הרשאות מתאימות לביצוע פעולה זו (נדרשת הרשאת עריכה/ניהול)');
    } else {
       showToast('⚠️ עקב התקלה, עדכון לא התבצע');
    }
  };

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  const onImgPick = (e: React.ChangeEvent<HTMLInputElement>, setter: (val: string) => void, setFile?: (f: File | null) => void) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 5000000) { showToast('⚠️ הקובץ גדול מדי (עד 5MB)'); return; }
    
    if (setFile) {
      setFile(file);
    }

    const r = new FileReader();
    r.onload = (ev) => {
      const result = ev.target?.result as string;
      if (!result) return;
      
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        let width = img.width;
        let height = img.height;
        const maxSize = 800; // max dimension
        
        if (width > height && width > maxSize) {
          height *= maxSize / width;
          width = maxSize;
        } else if (height > maxSize) {
          width *= maxSize / height;
          height = maxSize;
        }
        
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        if (ctx) {
           ctx.drawImage(img, 0, 0, width, height);
           // save quality to 0.7 for jpeg
           // If image is png, we can just use jpeg to enforce small size
           const dataUrl = canvas.toDataURL('image/jpeg', 0.7);
           setter(dataUrl);
        } else {
           setter(result); // fallback
        }
      };
      img.src = result;
    };
    r.readAsDataURL(file);
  };

  const getResizedDataUrl = (img: HTMLImageElement): string => {
    const canvas = document.createElement('canvas');
    const MAX_WIDTH = 800;
    const MAX_HEIGHT = 800;
    let width = img.width;
    let height = img.height;

    if (width > height) {
      if (width > MAX_WIDTH) {
        height *= MAX_WIDTH / width;
        width = MAX_WIDTH;
      }
    } else {
      if (height > MAX_HEIGHT) {
        width *= MAX_HEIGHT / height;
        height = MAX_HEIGHT;
      }
    }
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx?.drawImage(img, 0, 0, width, height);
    return canvas.toDataURL('image/jpeg', 0.8);
  };

  const handleInlineImageUpload = async (id: string, file: File) => {
    if (!isEditor) return;
    if (file.size > 5 * 1024 * 1024) { showToast('⚠️ התמונה גדולה מ-5MB'); return; }
    try {
      const url = await uploadProductImage(file, id);
      await setDoc(doc(db, 'products', id), { img: url, updatedAt: serverTimestamp() }, { merge: true });
      showToast('✓ התמונה עודכנה בהצלחה');
    } catch (err) { handleFirestoreError(err); }
  };

  const saveProduct = async () => {
    if (!fId || !fSku || !fDesc) {
      showToast('⚠️ מלא: מק"ט פנימי, מק"ט כללי ותיאור');
      return;
    }
    
    if (!isEditor) {
      showToast('⚠️ אין לך הרשאת עריכה');
      return;
    }

    setIsUploading(true);
    let finalImgUrl = fImg || '';
    const docId = editId || Date.now().toString() + Math.floor(Math.random()*1000);

    try {
      if (fImgFile) {
        finalImgUrl = await uploadProductImage(fImgFile, docId);
      }
    } catch (err) {
      handleFirestoreError(err);
      setIsUploading(false);
      return;
    }

    const data = {
      internalId: fId.substring(0, 100),
      sku: fSku.substring(0, 100),
      desc: fDesc.substring(0, 5000),
      supplier: fSup.substring(0, 200),
      category: fCategory.substring(0, 200),
      price: fPrice !== '' ? Number(fPrice) : '',
      stock: fStock !== '' ? Number(fStock) : '',
      prices: fPrices,
      img: finalImgUrl,
      created: editId ? products.find(p => p.id === editId)?.created || new Date().toLocaleDateString('he-IL') : new Date().toLocaleDateString('he-IL'),
      updatedAt: serverTimestamp()
    };

    try {
      await setDoc(doc(db, 'products', docId), data);
      showToast(editId ? '✓ הפריט עודכן בהצלחה' : '✓ הפריט נוסף לקטלוג');
      cancelEdit();
      navigate('/catalog');
    } catch (err) {
      handleFirestoreError(err);
    } finally {
      setIsUploading(false);
    }
  };

  const cancelEdit = () => {
    setEditId(null);
    setFId('');
    setFSku('');
    setFDesc('');
    setFSup('');
    setFCategory('');
    setFImg(null);
    setFPrice('');
    setFStock('');
    setFPrices({});
    setFImgFile(null);
    setAiGeneratedFields({});
  };

  const startEdit = (id: string) => {
    if (!isEditor) { showToast('⚠️ אין לך הרשאת עריכה'); return; }
    const p = products.find(x => x.id === id);
    if (!p) return;
    setEditId(id);
    setFId(p.internalId || '');
    setFSku(p.sku || '');
    setFDesc(p.desc || '');
    setFSup(p.supplier || '');
    setFCategory(p.category || '');
    setFImg(p.img || null);
    setFPrice(p.price || '');
    setFStock(p.stock || '');
    setFPrices(p.prices || {});
    setFImgFile(null);
    setAiGeneratedFields({});
    navigate('/add');
  };

  const delProduct = async (id: string) => {
    if (!isEditor) { showToast('⚠️ אין לך הרשאת מחיקה'); return; }
    if (window.confirm("האם אתה בטוח שברצונך למחוק פריט זה?")) {
      try {
        await deleteDoc(doc(db, 'products', id));
        showToast('🗑 הפריט נמחק');
      } catch (err) {
         handleFirestoreError(err);
      }
    }
  };

  const uniqueCategories = Array.from(new Set(products.map(p => p.category).filter(Boolean))) as string[];
  const uniqueSuppliers = Array.from(new Set(products.map(p => p.supplier).filter(Boolean))) as string[];

  const parseDate = (dString: string) => {
    if (!dString) return new Date(0);
    const parts = dString.split(/[\.\/\-]/);
    if (parts.length === 3) {
      const year = parts[2].length === 2 ? `20${parts[2]}` : parts[2];
      return new Date(Number(year), Number(parts[1]) - 1, Number(parts[0]));
    }
    return new Date(0);
  };

  const filteredProducts = products.filter(p => {
    let matchesSearch = true;
    (p as any)._sim = 0; // reset similarity

    if (searchMode === 'classic') {
      matchesSearch = (p.desc + p.internalId + p.sku + (p.supplier || '')).toLowerCase().includes(search.toLowerCase());
    } else {
      if (searchQueryEmbedding && p.embedding) {
        const sim = cosineSimilarity(searchQueryEmbedding, p.embedding);
        (p as any)._sim = sim;
        matchesSearch = sim > 0.4; // Threshold to filter completely irrelevant
      } else if (search.trim()) {
        // Fallback or cold start if no embedding on product
        matchesSearch = (p.desc + p.internalId + p.sku + (p.supplier || '')).toLowerCase().includes(search.toLowerCase());
      }
    }

    const matchesCategory = categoryFilter === '' || p.category === categoryFilter;
    const matchesSupplier = supplierFilter === '' || p.supplier === supplierFilter;
    
    let matchesDate = true;
    if (dateStartFilter || dateEndFilter) {
       const pDate = parseDate(p.created || '');
       if (dateStartFilter) {
          const sDate = new Date(dateStartFilter);
          sDate.setHours(0,0,0,0);
          if (pDate < sDate) matchesDate = false;
       }
       if (dateEndFilter) {
          const eDate = new Date(dateEndFilter);
          eDate.setHours(23,59,59,999);
          if (pDate > eDate) matchesDate = false;
       }
    }
    
    return matchesSearch && matchesCategory && matchesSupplier && matchesDate;
  });

  const isFiltering = search.length > 0 || categoryFilter !== '' || supplierFilter !== '';

  const sortedProducts = React.useMemo(() => {
    let sortableProducts = [...filteredProducts];
    if (searchMode === 'ai' && searchQueryEmbedding) {
      sortableProducts.sort((a, b) => ((b as any)._sim || 0) - ((a as any)._sim || 0));
    } else if (sortConfig !== null) {
      sortableProducts.sort((a, b) => {
        let aVal = a[sortConfig.key];
        let bVal = b[sortConfig.key];
        
        // Handle undefined or null comparison safety
        if (aVal === undefined || aVal === null) aVal = '';
        if (bVal === undefined || bVal === null) bVal = '';

        if (sortConfig.key === 'created') {
          const d1 = parseDate(a.created || '').getTime();
          const d2 = parseDate(b.created || '').getTime();
          if (d1 < d2) return sortConfig.direction === 'asc' ? -1 : 1;
          if (d1 > d2) return sortConfig.direction === 'asc' ? 1 : -1;
          return 0;
        }

        if (aVal < bVal) return sortConfig.direction === 'asc' ? -1 : 1;
        if (aVal > bVal) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }
    return sortableProducts;
  }, [filteredProducts, sortConfig, searchMode, searchQueryEmbedding]);

  const requestSort = (key: keyof Product) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  const savePriceLists = async (newLists: PriceListDef[]) => {
    try {
      await setDoc(doc(db, 'settings', 'pricelists'), { lists: newLists });
      showToast('✓ רשימת המחירונים נשמרה');
    } catch (err) {
      handleFirestoreError(err);
    }
  };

  const saveCompany = async () => {
    if (!isAdmin) { showToast('⚠️ פעולה זו דורשת הרשאת מנהל'); return; }
    try {
      await setDoc(doc(db, 'settings', 'company'), { 
        name: cName.substring(0, 200), 
        hp: cHp.substring(0, 50),
        addr: cAddr.substring(0, 500), 
        phone: cPhone.substring(0, 50), 
        email: cEmail.substring(0, 200), 
        web: cWeb.substring(0, 200), 
        logo: cLogo ? cLogo.substring(0, 1999990) : '',
        updatedAt: serverTimestamp()
      });
      showToast('✓ פרטי החברה נשמרו בענן');
    } catch (err) {
      handleFirestoreError(err);
    }
  };

  // --- Export Functions ---
  const exportExcel = async (itemsToExport: Product[]) => {
    if (!itemsToExport.length) { showToast('⚠️ אין פריטים לייצוא'); return; }

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('קטלוג מוצרים', { views: [{ rightToLeft: true }] });

      // Calculate heights and widths
      worksheet.getColumn(1).width = 15; // מק"ט פנימי
      worksheet.getColumn(2).width = 15; // SKU
      worksheet.getColumn(3).width = 15; // קטגוריה
      worksheet.getColumn(4).width = 40; // תיאור
      worksheet.getColumn(5).width = 15; // ספק
      worksheet.getColumn(6).width = 15; // תאריך
      worksheet.getColumn(7).width = 15; // תמונה (roughly 100px)

      // Add Company Header
      const coNameRow = worksheet.addRow([company.name || 'קטלוג מוצרים']);
      coNameRow.font = { size: 16, bold: true };
      const coInfoRow = worksheet.addRow([[company.addr, company.phone, company.email, company.web].filter(Boolean).join(' | ')]);
      coInfoRow.font = { size: 10, color: { argb: 'FF555555' } };
      worksheet.addRow([]);

      let startRow = 4;

      if (company.logo) {
        worksheet.mergeCells('A1:B3');
        const [prefix, b64] = company.logo.split(',');
        const ext = prefix.includes('jpeg') || prefix.includes('jpg') ? 'jpeg' : 'png';
        const logoId = workbook.addImage({ base64: b64, extension: ext as any });
        worksheet.addImage(logoId, {
          tl: { col: 0, row: 0 },
          ext: { width: 60, height: 60 }
        });
        startRow = 5;
      }

      // Headers
      const plHeaders = priceLists.map(pl => pl.name);
      
      const headerRow = worksheet.addRow(['מק"ט פנימי', 'מק"ט כללי (SKU)', 'קטגוריה', 'תיאור מוצר', 'ספק', 'תאריך קליטה', 'מחיר עלות', 'מלאי', ...plHeaders, 'תמונה']);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16A34A' } };
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
      });

      // Data Rows
      itemsToExport.forEach((p, index) => {
        const plValues = priceLists.map(pl => (p.prices && p.prices[pl.id]) || '');
        const row = worksheet.addRow([p.internalId, p.sku, p.category || '', p.desc, p.supplier || '', p.created || '', p.price || '', p.stock || '', ...plValues]);
        row.height = 70; // Make height ~90px
        
        row.eachCell((cell) => {
            cell.alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
            cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });

        if (p.img) {
          const [prefix, b64] = p.img.split(',');
          const ext = prefix.includes('jpeg') || prefix.includes('jpg') ? 'jpeg' : 'png';
          try {
            const imgId = workbook.addImage({ base64: b64, extension: ext as any });
            worksheet.addImage(imgId, {
              tl: { col: 8 + priceLists.length, row: row.number - 1 },
              ext: { width: 80, height: 80 }
            });
          } catch(e) { console.error("Error adding image inside excel", e); }
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `catalog_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast('✓ קובץ Excel הופק והורד בהצלחה!');
    } catch (err) {
      console.error(err);
      showToast('⚠️ שגיאה בהפקת אקסל');
    }
  };

  const exportPDF = (itemsToExport: Product[]) => {
    if (!itemsToExport.length) { showToast('⚠️ אין פריטים לייצוא'); return; }
    
    const username = auth.currentUser?.email || 'אורח';
    const dateObj = new Date();
    const dateStr = dateObj.toLocaleDateString('he-IL');
    const timeStr = dateObj.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' });

    const cNameStr = company.name || 'קטלוג מוצרים';
    const cInfo = [company.hp ? 'ח.פ: ' + company.hp : '', company.addr, company.phone, company.email, company.web].filter(Boolean).join(' | ');

    // Chunk into pages (14 items = 7 rows of 2)
    const ITEMS_PER_PAGE = 14;
    const pages = [];
    for (let i = 0; i < itemsToExport.length; i += ITEMS_PER_PAGE) {
      pages.push(itemsToExport.slice(i, i + ITEMS_PER_PAGE));
    }
    const totalPages = pages.length;

    const dispPriceStr = (p: Product): string => {
      if (displayPriceId === 'none') return '';
      if (displayPriceId === 'all') {
        let allPrices = '';
        if (isAdmin && p.price !== undefined && p.price !== '') allPrices += `<tr><td class="lbl">מחיר עלות</td><td>${p.price} ₪</td></tr>`;
        priceLists.forEach(pl => {
           if (p.prices && p.prices[pl.id] !== undefined && p.prices[pl.id] !== '') {
             allPrices += `<tr><td class="lbl">${pl.name}</td><td>${p.prices[pl.id]} ₪</td></tr>`;
           }
        });
        return allPrices;
      }
      if (displayPriceId === 'cost') {
        if (isAdmin && p.price !== undefined && p.price !== '') return `<tr><td class="lbl">מחיר עלות</td><td>${p.price} ₪</td></tr>`;
      } else {
        const pl = priceLists.find(x => x.id === displayPriceId);
        if (pl && p.prices && p.prices[displayPriceId] !== undefined && p.prices[displayPriceId] !== '') {
           return `<tr><td class="lbl">${pl.name}</td><td>${p.prices[displayPriceId]} ₪</td></tr>`;
        }
      }
      return '';
    };

    const printHtml = pages.map((pageItems, pageIndex) => {
      const cards = pageItems.map(p => `
      <div class="prod-card">
        <div class="prod-img">${p.img ? `<img src="${p.img}">` : `<div class="no-img">אין תמונה</div>`}</div>
        <div class="prod-info">
          <div class="prod-name">${p.desc}</div>
          <table class="info-table">
            <tr><td class="lbl">מק"ט פנימי</td><td>${p.internalId}</td></tr>
            <tr><td class="lbl">מק"ט כללי</td><td>${p.sku}</td></tr>
            ${p.category ? `<tr><td class="lbl">קטגוריה</td><td>${p.category}</td></tr>` : ''}
            ${p.supplier ? `<tr><td class="lbl">ספק</td><td>${p.supplier}</td></tr>` : ''}
            <tr><td class="lbl">מלאי</td><td>${p.stock || '0'}</td></tr>
            ${dispPriceStr(p)}
          </table>
        </div>
      </div>`).join('');

      return `
      <div class="page">
        <div class="header">
          <div class="header-content">
            ${company.logo ? `<img src="${company.logo}" class="logo">` : ''}
            <div>
              <div class="co-name">${cNameStr}</div>
              <div class="co-info">${cInfo}</div>
            </div>
          </div>
          <div class="header-meta">
            <div>תאריך: ${dateStr} ${timeStr}</div>
            <div>הופק ע"י: ${username}</div>
          </div>
        </div>
        <div class="grid">${cards}</div>
        <div class="footer">
          <div class="page-num">עמוד ${pageIndex + 1} מתוך ${totalPages}</div>
        </div>
      </div>
      `;
    }).join('');

    const html = `<!DOCTYPE html><html dir="rtl" lang="he"><head><meta charset="UTF-8"><title>קטלוג</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
@page { size: A4 portrait; margin: 0; }
body{font-family:Arial,'David',Tahoma,sans-serif;direction:rtl;color:#1a1a2e;background:#e5e7eb;font-size:13px;}
.page{background:#fff;width:210mm;height:296mm;padding:12mm 15mm;margin:0 auto;position:relative;page-break-after:always;overflow:hidden;}
.header{display:flex;justify-content:space-between;align-items:flex-start;border:1px solid #e5e7eb;padding:12px 18px;border-radius:6px;margin-bottom:14px;}
.header-content{display:flex;align-items:center;gap:14px;}
.logo{width:80px;height:80px;object-fit:contain;border-radius:6px;background:#fff;padding:2px}
.co-name{font-size:17px;font-weight:700}
.co-info{font-size:11px;opacity:.75;margin-top:3px}
.header-meta{text-align:left;font-size:11px;color:#6b7280;line-height:1.5;}
.footer{position:absolute;bottom:12mm;left:15mm;right:15mm;display:flex;justify-content:flex-end;align-items:center;padding-top:10px;}
.page-num{font-weight:bold;color:#1a1a2e;font-size:12px;}
.grid{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-start;align-content:flex-start;}
.prod-card{width:calc(50% - 6px);border:1px solid #d1d5db;border-radius:8px;padding:8px;display:flex;gap:10px;background:#fafafa;height:113px;overflow:hidden;}
.prod-img{width:95px;height:95px;flex-shrink:0;border-radius:6px;overflow:hidden;background:#e5e7eb;display:flex;align-items:center;justify-content:center}
.prod-img img{width:100%;height:100%;object-fit:contain;background:#fff}
.no-img{font-size:10px;color:#9ca3af;text-align:center;padding:8px}
.prod-info{flex:1;min-width:0;display:flex;flex-direction:column;justify-content:center;}
.prod-name{font-weight:700;font-size:13px;color:#111827;margin-bottom:5px;line-height:1.2;max-height:30px;overflow:hidden;}
.info-table{width:100%;font-size:11px;border-collapse:collapse}
.info-table td{padding:1px 0;vertical-align:top}
.info-table .lbl{color:#6b7280;font-weight:700;white-space:nowrap;width:70px}
.print-btn{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);padding:11px 26px;background:#1a1a2e;color:#fff;border:none;border-radius:8px;font-size:15px;cursor:pointer;font-family:inherit;box-shadow:0 4px 16px rgba(0,0,0,.2);z-index:999;}
@media print{
  body{background:#fff;}
  .page{margin:0;box-shadow:none;width:100%;height:100vh;}
  .print-btn{display:none}
}
</style></head>
<body>
${printHtml}
<button class="print-btn" onclick="window.print()">🖨️ הדפס / שמור כ-PDF</button>
<script>setTimeout(()=>window.print(),600);</script>
</body></html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.target = '_blank';
    a.click();
    showToast('✓ הקטלוג נפתח בכרטיסייה חדשה');
  };

  // --- JSON Backup / Restore ---
  const handleExportBackup = () => {
    const data = { products, company };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `catalog_backup_${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    showToast('✓ קובץ גיבוי נשמר בהצלחה!');
  };

  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const data = JSON.parse(ev.target?.result as string);
        if (data.products && Array.isArray(data.products)) {
            for (const p of data.products) {
              const cleaned: any = {
                internalId: p.internalId || '',
                sku: p.sku || '',
                desc: p.desc || '',
                supplier: p.supplier || '',
                category: p.category || '',
                price: p.price ?? '',
                stock: p.stock ?? '',
                prices: p.prices || {},
                img: p.img || '',
                created: p.created || new Date().toLocaleDateString('he-IL'),
                updatedAt: serverTimestamp()
              };
              await setDoc(doc(db, 'products', p.id || Date.now().toString() + Math.random().toString()), cleaned);
           }
        }
        if (data.company) {
           const comp = data.company;
           const cleanCompany = {
             name: comp.name || '',
             addr: comp.addr || '',
             phone: comp.phone || '',
             email: comp.email || '',
             web: comp.web || '',
             logo: comp.logo || '',
             updatedAt: serverTimestamp()
           }
           await setDoc(doc(db, 'settings', 'company'), cleanCompany);
        }
        showToast('✓ שחזור גיבוי בוצע בהצלחה, מרענן נתונים במחזור הבא של Firestore.');
      } catch (err) {
        handleFirestoreError(err);
        showToast('⚠️ שגיאה בשחזור או בפענוח קובץ הגיבוי');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  const handlePlExcelPick = (e: React.ChangeEvent<HTMLInputElement>, pl: PriceListDef) => {
    if (!isEditor) { showToast('⚠️ אין לך הרשאת ייבוא/עריכה'); e.target.value=''; return; }
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rawData: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        if (rawData.length > 0) {
          const headers = Object.keys(rawData[0]);
          setPlImportHeaders(headers);
          
          const cleanHeader = (h: string) => h.toLowerCase().trim();
          const autoMap: Record<string, string> = { internalId: '', sku: '', price: '' };
          headers.forEach((h) => {
            const lowerH = cleanHeader(h);
            if (lowerH.includes('פנימי') || lowerH.includes('internal')) autoMap.internalId = h;
            else if (lowerH.includes('sku') || lowerH.includes('כללי') || lowerH.includes('ברקוד')) autoMap.sku = h;
            else if (lowerH.includes('מחיר') || lowerH.includes('price') || lowerH.includes(pl.name.toLowerCase())) autoMap.price = h;
          });
          
          setPlFieldMapping(autoMap);
          setPlImportData(rawData);
          setImportingPl(pl);
          setImportPlModalOpen(true);
        } else {
          showToast('⚠️ לא נמצאו נתונים בקובץ');
        }
      } catch (err) {
        console.error(err);
        showToast('⚠️ שגיאה בפענוח קובץ Excel');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset input
  };

  const executePlImport = async () => {
    if (!plFieldMapping.internalId && !plFieldMapping.sku) {
      showToast('⚠️ חובה לשייך לפחות שדה מזהה אחד (מק"ט כללי או מק"ט פנימי)');
      return;
    }
    if (!plFieldMapping.price) {
      showToast('⚠️ חובה לשייך עמודת מחיר');
      return;
    }
    if (!importingPl) return;

    let updatedCount = 0;
    const updates: Promise<void>[] = [];

    plImportData.forEach((row) => {
      const parsedInternalId = plFieldMapping.internalId && row[plFieldMapping.internalId] !== undefined ? String(row[plFieldMapping.internalId] || '').trim().substring(0, 100) : '';
      const parsedSku = plFieldMapping.sku && row[plFieldMapping.sku] !== undefined ? String(row[plFieldMapping.sku] || '').trim().substring(0, 100) : '';
      
      const existingProduct = products.find(p => 
        (parsedSku && p.sku === parsedSku) || 
        (parsedInternalId && p.internalId === parsedInternalId)
      );

      if (existingProduct) {
        const rawPriceValue = row[plFieldMapping.price];
        if (rawPriceValue !== undefined && rawPriceValue !== '') {
          const newPrice = Number(rawPriceValue);
          const currentPrices = existingProduct.prices ? { ...existingProduct.prices } : {};
          currentPrices[importingPl.id] = newPrice;
          
          updates.push(
            setDoc(doc(db, 'products', existingProduct.id), { prices: currentPrices, updatedAt: serverTimestamp() }, { merge: true })
          );
          updatedCount++;
        }
      }
    });

    if (updates.length > 0) {
      try {
        await Promise.all(updates);
        showToast(`✓ עודכנו מחירי "${importingPl.name}" עבור ${updatedCount} פריטים!`);
        setImportPlModalOpen(false);
        setImportingPl(null);
      } catch (err) {
        console.error(err);
        showToast('⚠️ שגיאה בעדכון מחירים במסד הנתונים');
      }
    } else {
      showToast('⚠️ לא נמצאו פריטים תואמים בקטלוג לעדכון');
    }
  };

  const handleExcelImportPick = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!isEditor) { showToast('⚠️ אין לך הרשאת ייבוא/עריכה'); e.target.value=''; return; }
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rawData: any[] = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        if (rawData.length > 0) {
          const headers = Object.keys(rawData[0]);
          setImportHeaders(headers);
          
          const cleanHeader = (h: string) => h.toLowerCase().trim();
          const autoMap: Record<string, string> = { desc: '', internalId: '', sku: '', supplier: '', category: '', img: '', price: '', stock: '' };
          priceLists.forEach(pl => autoMap[`priceList_${pl.id}`] = '');
          headers.forEach((h) => {
            const lowerH = cleanHeader(h);
            if (lowerH.includes('תיאור') || lowerH.includes('desc')) autoMap.desc = h;
            else if (lowerH.includes('פנימי') || lowerH.includes('internal')) autoMap.internalId = h;
            else if (lowerH.includes('sku') || lowerH.includes('כללי') || lowerH.includes('ברקוד')) autoMap.sku = h;
            else if (lowerH.includes('ספק') || lowerH.includes('supplier')) autoMap.supplier = h;
            else if (lowerH.includes('קטגור') || lowerH.includes('category')) autoMap.category = h;
            else if (lowerH.includes('תמונ') || lowerH.includes('img') || lowerH.includes('image')) autoMap.img = h;
            else if (lowerH.includes('מחיר עלות') || lowerH.includes('price')) autoMap.price = h;
            else if (lowerH.includes('מלאי') || lowerH.includes('stock')) autoMap.stock = h;
            else {
              priceLists.forEach(pl => {
                if (lowerH.includes(pl.name.toLowerCase())) autoMap[`priceList_${pl.id}`] = h;
              });
            }
          });
          setFieldMapping(autoMap);
          setImportData(rawData);
          setImportModalOpen(true);
        } else {
          showToast('⚠️ לא נמצאו נתונים בקובץ');
        }
      } catch (err) {
        console.error(err);
        showToast('⚠️ שגיאה בפענוח קובץ Excel');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset input
  };

  const executeImport = async () => {
    if (!fieldMapping.internalId && !fieldMapping.sku) {
      showToast('⚠️ חובה לשייך לפחות שדה מזהה אחד (מק"ט כללי או מק"ט פנימי)');
      return;
    }

    const newProducts = importData.map((row, idx) => {
      const parsedInternalId = fieldMapping.internalId && row[fieldMapping.internalId] !== undefined ? String(row[fieldMapping.internalId] || '').trim().substring(0, 100) : '';
      const parsedSku = fieldMapping.sku && row[fieldMapping.sku] !== undefined ? String(row[fieldMapping.sku] || '').trim().substring(0, 100) : '';
      
      let existingProduct = products.find(p => 
        (parsedSku && p.sku === parsedSku) || 
        (parsedInternalId && p.internalId === parsedInternalId)
      );

      const pId = existingProduct ? existingProduct.id : Date.now().toString() + idx;

      const pricesMap: Record<string, number | string> = existingProduct?.prices ? { ...existingProduct.prices } : {};
      priceLists.forEach(pl => {
        const h = fieldMapping[`priceList_${pl.id}`];
        if (h && row[h] !== undefined && row[h] !== '') {
          pricesMap[pl.id] = Number(row[h]);
        }
      });
      return {
      id: pId,
      internalId: parsedInternalId || existingProduct?.internalId || '',
      sku: parsedSku || existingProduct?.sku || '',
      desc: fieldMapping.desc && row[fieldMapping.desc] !== undefined ? String(row[fieldMapping.desc] || '').trim().substring(0, 5000) : existingProduct?.desc || '',
      supplier: fieldMapping.supplier && row[fieldMapping.supplier] !== undefined ? String(row[fieldMapping.supplier] || '').trim().substring(0, 200) : existingProduct?.supplier || '',
      category: fieldMapping.category && row[fieldMapping.category] !== undefined ? String(row[fieldMapping.category] || '').trim().substring(0, 200) : existingProduct?.category || '',
      price: fieldMapping.price && row[fieldMapping.price] !== undefined && row[fieldMapping.price] !== '' ? Number(row[fieldMapping.price]) : existingProduct?.price !== undefined ? existingProduct?.price : '',
      stock: fieldMapping.stock && row[fieldMapping.stock] !== undefined && row[fieldMapping.stock] !== '' ? Number(row[fieldMapping.stock]) : existingProduct?.stock !== undefined ? existingProduct?.stock : '',
      prices: pricesMap,
      img: fieldMapping.img && row[fieldMapping.img] !== undefined ? String(row[fieldMapping.img]).trim().substring(0, 1999990) : existingProduct?.img || '',
      created: existingProduct?.created || new Date().toLocaleDateString('he-IL'),
      updatedAt: serverTimestamp()
    }}).filter(p => !!p.sku || !!p.internalId);

    const validProducts = newProducts.filter(p => !!p.desc); // Require description ultimately (either from excel or existing)

    if (validProducts.length > 0) {
      try {
        for(let p of validProducts) {
          const {id, ...rest} = p;
          await setDoc(doc(db, 'products', id), rest, { merge: true });
        }
        showToast(`✓ התווספו/עודכנו ${validProducts.length} פריטים לקטלוג!`);
        setImportModalOpen(false);
        setImportData([]);
      } catch(err) {
        handleFirestoreError(err);
      }
    } else {
      showToast('⚠️ לא נמצאו פריטים תקינים. יש לוודא שלפריטים חדשים יש "תיאור".');
    }
  };

  return (
    <div className="min-h-screen h-screen bg-slate-100 text-slate-900 font-sans flex flex-col overflow-hidden" dir="rtl">
      {/* Topbar */}
      <header className="bg-white text-slate-900 border-b border-slate-200 px-4 md:px-5 h-14 flex items-center justify-between shadow-sm z-40 shrink-0 relative flex-wrap md:flex-nowrap">
        <div className="text-lg font-bold flex items-center gap-2 min-w-max">
          <button className="md:hidden p-1 mr-[-8px] text-slate-500 hover:bg-slate-100 rounded-md" onClick={() => setMobileMenuOpen(!mobileMenuOpen)}>
             <Menu size={20} />
          </button>
          <span className="text-rose-500">◈</span> CatalogPro
        </div>
        
        <div className="hidden md:flex items-center gap-3 text-xs ml-auto pl-4 border-l border-slate-100">
           {syncStatus === 'synced' && <><CloudLightning size={14} className="text-emerald-500" /><span className="text-emerald-700 font-medium hidden lg:inline">מסונכרן</span></>}
           {syncStatus === 'syncing' && <><CloudCog size={14} className="text-amber-500 animate-spin" /><span className="text-amber-700 font-medium hidden lg:inline">מסנכרן...</span></>}
           {syncStatus === 'error' && <><Shield size={14} className="text-red-500" /><span className="text-red-700 font-medium hidden lg:inline">שגיאת סנכרון</span></>}
        </div>

        <div className="flex gap-1 overflow-x-auto hide-scrollbars items-center w-full md:w-auto mt-2 md:mt-0 justify-start md:justify-end">
          {canSeeTab('catalog') && (
            <div className="flex gap-1.5 items-center mr-2 md:mr-3 border-r border-slate-200 pr-2 md:pr-3">
              <select
                value={displayPriceId}
                onChange={e => setDisplayPriceId(e.target.value)}
                className="bg-slate-50 border border-slate-200 text-slate-700 text-xs md:text-sm rounded-md px-2 py-1.5 focus:outline-none focus:border-indigo-400 font-medium ml-1"
              >
                <option value="none">ללא מחיר</option>
                <option value="all">כל המחירים</option>
                {isAdmin && <option value="cost">מחיר עלות</option>}
                {priceLists.map(pl => (
                  <option key={pl.id} value={pl.id}>{pl.name}</option>
                ))}
              </select>
              <button
                onClick={() => setCatalogView('grid')}
                className={`p-1.5 rounded-md transition-colors ${catalogView === 'grid' ? 'bg-slate-200 text-slate-800' : 'text-slate-400 hover:bg-slate-100 hover:text-slate-600'}`}
                title="תצוגת כרטיסיות"
              >
                <LayoutGrid size={18} />
              </button>
              <button
                onClick={() => setCatalogView('list')}
                className={`p-1.5 rounded-md transition-colors ${catalogView === 'list' ? 'bg-slate-200 text-slate-800' : 'text-slate-400 hover:bg-slate-100 hover:text-slate-600'}`}
                title="תצוגת רשימה"
              >
                <List size={18} />
              </button>
            </div>
          )}
          {canSeeTab('catalog') && (
          <NavLink
            to="/catalog"
            className={({ isActive }) => `px-2 md:px-3 py-1.5 rounded-md text-xs md:text-sm font-medium flex items-center gap-1.5 transition-colors whitespace-nowrap ${isActive ? 'bg-rose-500 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100 hover:text-slate-900'}`}
          >
            <LayoutGrid size={16} /> קטלוג
          </NavLink>
          )}
          {canSeeTab('add') && (
            <NavLink
              to="/add"
              onClick={() => cancelEdit()}
              className={({ isActive }) => `px-2 md:px-3 py-1.5 rounded-md text-xs md:text-sm font-medium flex items-center gap-1.5 transition-colors whitespace-nowrap ${isActive ? 'bg-rose-500 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100 hover:text-slate-900'}`}
            >
              <Plus size={16} /> <span className="hidden sm:inline">הוסף / ערוך</span>
            </NavLink>
          )}
          {canSeeTab('company') && (
            <NavLink
              to="/company"
              className={({ isActive }) => `px-2 md:px-3 py-1.5 rounded-md text-xs md:text-sm font-medium flex items-center gap-1.5 transition-colors whitespace-nowrap ${isActive ? 'bg-rose-500 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100 hover:text-slate-900'}`}
            >
              <Building2 size={16} /> <span className="hidden sm:inline"> חברה</span>
            </NavLink>
          )}
          {canSeeTab('pricelists') && (
            <NavLink
              to="/pricelists"
              className={({ isActive }) => `px-2 md:px-3 py-1.5 rounded-md text-xs md:text-sm font-medium flex items-center gap-1.5 transition-colors whitespace-nowrap ${isActive ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-500 hover:bg-slate-100 hover:text-slate-900'}`}
            >
              <List size={16} /> <span className="hidden sm:inline">מחירונים</span>
            </NavLink>
          )}
          {canSeeTab('admin') && (
            <NavLink
              to="/admin"
              className={({ isActive }) => `px-2 md:px-3 py-1.5 rounded-md text-xs md:text-sm font-medium flex items-center gap-1.5 transition-colors whitespace-nowrap ${isActive ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-500 hover:bg-indigo-50 hover:text-indigo-600'}`}
            >
              <Shield size={16} /> <span className="hidden sm:inline">ניהול</span>
            </NavLink>
          )}
        </div>
      </header>

      <div className="flex-1 flex flex-col md:flex-row overflow-hidden relative">
        {/* Sidebar */}
        <aside className={`${mobileMenuOpen ? 'block' : 'hidden'} md:flex w-full md:w-72 bg-white border-b md:border-b-0 md:border-l border-slate-200 flex-col overflow-y-auto shrink-0 shadow-[2px_0_10px_rgba(0,0,0,0.02)] z-20 absolute md:relative inset-0 md:inset-auto h-full`}>
          <div className="p-4 flex-1 flex flex-col">
            <div className="bg-white rounded-lg p-4 text-slate-800 mb-4 shadow-sm border border-slate-200 flex flex-col items-center gap-3 text-center">
              <div className="w-24 h-24 shrink-0 rounded-md bg-white flex items-center justify-center overflow-hidden border border-slate-100 p-2 shadow-sm">
                {company.logo || cLogo ? (
                  <img src={company.logo || cLogo || ''} alt="" className="w-full h-full object-contain" />
                ) : (
                  <Building2 size={32} className="text-slate-400" />
                )}
              </div>
              <div className="w-full overflow-hidden flex-1">
                <div className="font-bold text-[16px] text-slate-800">{company.name || cName || 'שם החברה שלך'}</div>
                {(company.hp || cHp) && <div className="text-[12px] text-slate-500 font-medium">ח.פ: {company.hp || cHp}</div>}
                
                <div className="text-[12px] text-slate-500 mt-2 flex flex-col items-center gap-0.5">
                  {(company.addr || cAddr) && <div>{company.addr || cAddr}</div>}
                  {(company.email || cEmail) && <div>{company.email || cEmail}</div>}
                  {(company.phone || cPhone) && <div>{company.phone || cPhone}</div>}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs text-indigo-700 bg-indigo-50 border border-indigo-200 rounded-md px-3 py-2 mb-4 justify-between">
              <div className="flex items-center gap-2">
                <Shield size={14} className="shrink-0 text-indigo-500" />
                <span>הרשאתך: <b>{userRole}</b></span>
              </div>
              <div className="md:hidden flex items-center gap-1.5 pl-2 border-l border-indigo-200">
                 {syncStatus === 'synced' && <CloudLightning size={14} className="text-emerald-500" />}
                 {syncStatus === 'syncing' && <CloudCog size={14} className="text-amber-500 animate-spin" />}
                 {syncStatus === 'error' && <Shield size={14} className="text-red-500" />}
              </div>
            </div>


          <div className="grid grid-cols-3 gap-2 mb-4">
            <div className="bg-slate-50 border border-slate-200 rounded-md p-2 text-center">
              <div className="text-lg font-bold text-slate-800">{products.length}</div>
              <div className="text-[10px] text-slate-500 mt-0.5">פריטים</div>
            </div>
            <div className="bg-slate-50 border border-slate-200 rounded-md p-2 text-center">
              <div className="text-lg font-bold text-slate-800">{new Set(products.map(p => p.supplier).filter(Boolean)).size}</div>
              <div className="text-[10px] text-slate-500 mt-0.5">ספקים</div>
            </div>
            <div className="bg-slate-50 border border-slate-200 rounded-md p-2 text-center">
              <div className="text-lg font-bold text-slate-800">{products.filter(p => p.img).length}</div>
              <div className="text-[10px] text-slate-500 mt-0.5">תמונות</div>
            </div>
          </div>

          <div className="relative mb-4 pb-4 border-b border-slate-200 space-y-2">
            {uniqueCategories.length > 0 && (
              <select
                value={categoryFilter}
                onChange={e => setCategoryFilter(e.target.value)}
                className="w-full py-2 px-3 bg-white border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium text-slate-700 cursor-pointer appearance-none"
              >
                <option value="">כל הקטגוריות</option>
                {uniqueCategories.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            )}
            {uniqueSuppliers.length > 0 && (
              <select
                value={supplierFilter}
                onChange={e => setSupplierFilter(e.target.value)}
                className="w-full py-2 px-3 bg-white border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium text-slate-700 cursor-pointer appearance-none"
              >
                <option value="">כל הספקים</option>
                {uniqueSuppliers.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            )}
            
            <div className="flex gap-2">
              <div className="flex-1">
                <label className="block text-[11px] text-slate-500 mb-1">תאריך תחילה (יצירה)</label>
                <input
                  type="date"
                  className="w-full py-1.5 px-2.5 bg-white border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium text-slate-700"
                  value={dateStartFilter}
                  onChange={e => setDateStartFilter(e.target.value)}
                />
              </div>
              <div className="flex-1">
                <label className="block text-[11px] text-slate-500 mb-1">תאריך סיום (יצירה)</label>
                <input
                  type="date"
                  className="w-full py-1.5 px-2.5 bg-white border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium text-slate-700"
                  value={dateEndFilter}
                  onChange={e => setDateEndFilter(e.target.value)}
                />
              </div>
            </div>

            <div className="pt-2 border-t border-slate-100">
               <label className="block text-[11px] text-slate-500 mb-1">מיון</label>
               <select
                value={sortConfig ? `${sortConfig.key}_${sortConfig.direction}` : ''}
                onChange={e => {
                  if (!e.target.value) { setSortConfig(null); return; }
                  const [k, d] = e.target.value.split('_');
                  setSortConfig({ key: k as any, direction: d as 'asc'|'desc' });
                }}
                className="w-full py-1.5 px-2 bg-white border border-slate-200 rounded-md text-[12px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium text-slate-700 cursor-pointer appearance-none"
               >
                 <option value="">ללא מיון</option>
                 <option value="created_desc">תאריך יצירה (מהחדש לישן)</option>
                 <option value="created_asc">תאריך יצירה (מהישן לחדש)</option>
                 <option value="desc_asc">תיאור (א-ת)</option>
                 <option value="desc_desc">תיאור (ת-א)</option>
                 <option value="sku_asc">מק"ט כללי (עולה)</option>
                 <option value="sku_desc">מק"ט כללי (יורד)</option>
               </select>
            </div>

            <div className="relative">
              <Search size={16} className={`absolute right-3 top-1/2 -translate-y-1/2 ${isSearchingSemantic ? 'text-indigo-400 animate-pulse' : 'text-slate-400'}`} />
              <input
                type="text"
                placeholder={searchMode === 'ai' ? 'חיפוש סמנטי (RAG)...' : 'חיפוש רגיל (תיאור, מק"ט, ספק...)'}
                className={`w-full py-2 pr-9 pl-3 bg-slate-50 border rounded-md text-[13px] outline-none transition-all font-medium placeholder:font-normal ${searchMode === 'ai' ? 'border-indigo-300 focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 placeholder:text-indigo-300 text-indigo-900 bg-indigo-50/30' : 'border-slate-200 focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100'}`}
                value={search}
                onChange={e => setSearch(e.target.value)}
              />
              <button
                onClick={() => setSearchMode(prev => prev === 'classic' ? 'ai' : 'classic')}
                className={`absolute left-2 top-1/2 -translate-y-1/2 px-2 py-0.5 rounded text-[10px] font-bold transition-colors ${searchMode === 'ai' ? 'bg-indigo-600 text-white shadow-sm' : 'bg-slate-200 text-slate-500 hover:bg-slate-300'}`}
                title="החלף בין חיפוש רגיל לחיפוש עם בינה מלאכותית"
              >
                {searchMode === 'ai' ? '✨ IA' : 'רגיל'}
              </button>
            </div>
          </div>

          <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 mb-4 space-y-2">
            <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5 mb-2">
              <Download size={14} /> ייצוא והפקה
            </div>
            <button onClick={() => exportExcel(products)} className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white py-1.5 px-3 rounded-md text-[13px] font-semibold transition-colors">
              <FileSpreadsheet size={16} /> Excel הכל
            </button>
            <button onClick={() => exportPDF(products)} className="w-full flex items-center justify-center gap-2 bg-indigo-700 hover:bg-indigo-800 text-white py-1.5 px-3 rounded-md text-[13px] font-semibold transition-colors">
              <FileText size={16} /> PDF הכל
            </button>
            {isFiltering && (
              <div className="pt-2 border-t border-slate-200 mt-2 flex flex-col gap-2">
                <div className="text-[10px] text-center text-slate-500 font-bold mb-0.5">תוצאות סינון ({filteredProducts.length})</div>
                <button onClick={() => exportExcel(filteredProducts)} className="w-full flex items-center justify-center gap-2 bg-emerald-500 hover:bg-emerald-600 text-white py-1.5 px-3 rounded-md text-[13px] font-semibold transition-colors">
                  <FileSpreadsheet size={16} /> Excel מסונן
                </button>
                <button onClick={() => exportPDF(filteredProducts)} className="w-full flex items-center justify-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white py-1.5 px-3 rounded-md text-[13px] font-semibold transition-colors">
                  <FileText size={16} /> PDF מסונן
                </button>
              </div>
            )}
          </div>

          {isAdmin && (
            <div className="bg-orange-50 border border-orange-200 rounded-lg p-3 mb-4 space-y-2">
              <div className="text-[11px] font-bold text-orange-700 uppercase tracking-wider flex items-center gap-1.5 mb-2">
                <DatabaseBackup size={14} /> גיבוי וייבוא מתקדם
              </div>

              <label className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white py-1.5 px-3 rounded-md text-[12px] font-semibold transition-colors cursor-pointer shadow-sm">
                <FileUp size={14} /> ייבוא ועדכון פריטים/מלאי מ-Excel
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleExcelImportPick} />
              </label>

              <button onClick={handleExportBackup} className="w-full flex items-center justify-center gap-2 bg-white border border-orange-300 hover:bg-orange-100 text-orange-800 py-1.5 px-3 rounded-md text-[12px] font-semibold transition-colors mt-2">
                <Download size={14} /> הורד קובץ גיבוי לוקאלי
              </button>
              <label className="w-full flex items-center justify-center gap-2 bg-white border border-orange-300 hover:bg-orange-100 text-orange-800 py-1.5 px-3 rounded-md text-[12px] font-semibold transition-colors mt-2 cursor-pointer cursor-allowed">
                <Upload size={14} /> שחזר מקובץ גיבוי
                <input type="file" accept=".json" className="hidden" onChange={handleImportBackup} />
              </label>
            </div>
          )}

            <div className="md:hidden mt-auto pt-4 text-center">
               <button onClick={() => setMobileMenuOpen(false)} className="bg-slate-100 hover:bg-slate-200 px-4 py-2 rounded-md font-medium text-[13px] w-full text-slate-600">סגור תפריט צד</button>
            </div>
          </div>
        </aside>

        {/* Content Area */}
        <main className={`flex-1 p-4 md:p-6 overflow-y-auto bg-slate-100 relative ${mobileMenuOpen ? 'opacity-50 pointer-events-none md:opacity-100 md:pointer-events-auto' : ''}`}>
          
          <Routes>
          {/* CATALOG TAB */}
          <Route path="/catalog" element={canSeeTab('catalog') ? (
            <div className={sortedProducts.length > 0 && catalogView === 'grid' ? "grid grid-cols-[repeat(auto-fill,minmax(260px,1fr))] gap-4 items-start" : (sortedProducts.length > 0 && catalogView === 'list' ? "overflow-x-auto bg-white rounded-xl shadow-sm border border-slate-200" : "h-full flex items-center justify-center")}>
              {sortedProducts.length === 0 ? (
                <div className="text-center p-12 text-slate-400 flex flex-col items-center w-full">
                  <PackageSearch size={48} className="mb-4 opacity-50 stroke-[1.5]" />
                  <p className="font-semibold text-slate-600 mb-1">הקטלוג ריק או לא נמצאו תוצאות</p>
                  {isEditor && <p className="text-sm">לחץ "הוסף / ערוך" בכדי להוסיף פריטים חדשים</p>}
                </div>
              ) : catalogView === 'grid' ? (
                sortedProducts.map(p => (
                  <div key={p.id} className="bg-white border border-slate-200 rounded-xl overflow-hidden hover:shadow-lg transition-transform hover:-translate-y-0.5 duration-200 group">
                    <div className="w-full h-40 bg-slate-100 flex items-center justify-center overflow-hidden relative">
                      {p.img ? (
                         <img src={p.img} alt="" className="w-full h-full object-cover" />
                      ) : (
                         <ImagePlus size={36} className="text-slate-300 opacity-60" />
                      )}
                      
                      {isEditor && (
                        <div className="absolute top-2 left-2 flex gap-1.5 opacity-0 group-hover:opacity-100 transition-opacity">
                          <label className="w-8 h-8 rounded-md bg-white/95 text-indigo-600 hover:text-indigo-800 hover:bg-white shadow-sm flex items-center justify-center transition-colors cursor-pointer" title="שנה תמונה">
                            <ImagePlus size={15} />
                            <input type="file" accept="image/*" className="hidden" onChange={(e) => { if(e.target.files?.[0]) handleInlineImageUpload(p.id, e.target.files[0]); }} />
                          </label>
                          <button onClick={() => startEdit(p.id)} className="w-8 h-8 rounded-md bg-white/95 text-indigo-600 hover:text-indigo-800 hover:bg-white shadow-sm flex items-center justify-center transition-colors">
                            <Pencil size={15} />
                          </button>
                          <button onClick={() => delProduct(p.id)} className="w-8 h-8 rounded-md bg-white/95 text-red-600 hover:text-red-800 hover:bg-white shadow-sm flex items-center justify-center transition-colors">
                            <Trash2 size={15} />
                          </button>
                        </div>
                      )}
                    </div>
                    
                    <div className="bg-gradient-to-br from-slate-900 to-indigo-950 px-3 py-2 flex items-center justify-between text-[11px] text-indigo-200 font-medium">
                       <span>מק"ט: <span className="text-white">{p.internalId}</span></span>
                       <span>SKU: <span className="text-white">{p.sku}</span></span>
                    </div>
                    
                    <div className="p-0">
                      <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                        <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">תיאור</div>
                        <div className="text-[13.5px] text-slate-900 font-semibold leading-snug break-words">{p.desc}</div>
                      </div>
                      <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                        <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">פנימי</div>
                        <div className="text-[13px] text-slate-700 font-medium break-all">{p.internalId}</div>
                      </div>
                      <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                        <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">כללי SKU</div>
                        <div className="text-[13px] text-slate-700 font-medium break-all">{p.sku}</div>
                      </div>
                      {p.category && (
                        <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                          <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">קטגוריה</div>
                          <div className="text-[13px] text-slate-700 font-medium break-words">{p.category}</div>
                        </div>
                      )}
                      <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                        <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">ספק</div>
                        <div className="text-[13px] text-slate-700 font-medium break-words">{p.supplier || '—'}</div>
                      </div>
                      {((displayPriceId === 'all' && isAdmin) || displayPriceId === 'cost') && (
                        <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                          <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">מחיר עלות</div>
                          <div className="text-[13px] text-slate-700 font-medium break-words">{p.price || '—'} ₪</div>
                        </div>
                      )}
                      <div className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                        <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">מלאי</div>
                        <div className="text-[13px] text-slate-700 font-medium break-words">{p.stock || '0'} יח'</div>
                      </div>
                      {priceLists.filter(pl => displayPriceId === 'all' || displayPriceId === pl.id).map(pl => (
                        <div key={pl.id} className="flex items-start px-3 py-2 border-b border-slate-50 gap-2">
                          <div className="text-[10.5px] font-bold text-slate-500 min-w-[70px] uppercase tracking-wide shrink-0 pt-0.5">{pl.name}</div>
                          <div className="text-[13px] text-slate-700 font-medium break-words">{(p.prices && p.prices[pl.id]) || '—'} ₪</div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))
              ) : (
                <table className="w-full text-right text-sm">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-600">
                    <tr>
                      <th className="px-4 py-3 font-semibold w-16 text-center">תמונה</th>
                      <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('internalId')}>
                        <div className="flex items-center gap-1">מק"ט פנימי {sortConfig?.key === 'internalId' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                      </th>
                      <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('sku')}>
                        <div className="flex items-center gap-1">SKU {sortConfig?.key === 'sku' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                      </th>
                      <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('desc')}>
                        <div className="flex items-center gap-1">תיאור {sortConfig?.key === 'desc' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                      </th>
                      <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('category')}>
                        <div className="flex items-center gap-1">קטגוריה {sortConfig?.key === 'category' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                      </th>
                      <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('supplier')}>
                        <div className="flex items-center gap-1">ספק {sortConfig?.key === 'supplier' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                      </th>
                      {((displayPriceId === 'all' && isAdmin) || displayPriceId === 'cost') && (
                        <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('price')}>
                          <div className="flex items-center gap-1">מחיר {sortConfig?.key === 'price' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                        </th>
                      )}
                      <th className="px-4 py-3 font-semibold cursor-pointer hover:bg-slate-100 transition-colors group" onClick={() => requestSort('stock')}>
                        <div className="flex items-center gap-1">מלאי {sortConfig?.key === 'stock' && (sortConfig.direction === 'asc' ? <ArrowUpNarrowWide size={14} className="text-indigo-500" /> : <ArrowDownWideNarrow size={14} className="text-indigo-500" />)}</div>
                      </th>
                      {priceLists.filter(pl => displayPriceId === 'all' || displayPriceId === pl.id).map(pl => (
                        <th key={pl.id} className="px-4 py-3 font-semibold">{pl.name}</th>
                      ))}
                      {isEditor && <th className="px-4 py-3 font-semibold text-center">פעולות</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {sortedProducts.map(p => (
                      <tr key={p.id} className="hover:bg-slate-50 transition-colors group">
                        <td className="px-4 py-2">
                          <div className="w-10 h-10 bg-slate-100 rounded-md overflow-hidden flex items-center justify-center shrink-0 border border-slate-200 mx-auto relative group/img">
                            {p.img ? <img src={p.img} alt="" className="w-full h-full object-cover" /> : <ImagePlus size={16} className="text-slate-300" />}
                            {isEditor && (
                              <label className="absolute inset-0 bg-black/50 text-white flex items-center justify-center opacity-0 group-hover/img:opacity-100 cursor-pointer transition-opacity">
                                <ImagePlus size={16} />
                                <input type="file" accept="image/*" className="hidden" onChange={(e) => { if(e.target.files?.[0]) handleInlineImageUpload(p.id, e.target.files[0]); }} />
                              </label>
                            )}
                          </div>
                        </td>
                        <td className="px-4 py-3 font-medium text-slate-700">{p.internalId}</td>
                        <td className="px-4 py-3 text-slate-600">{p.sku}</td>
                        <td className="px-4 py-3 font-semibold text-slate-800 max-w-[200px] truncate" title={p.desc}>{p.desc}</td>
                        <td className="px-4 py-3 text-slate-600">{p.category || '—'}</td>
                        <td className="px-4 py-3 text-slate-600">{p.supplier || '—'}</td>
                        {((displayPriceId === 'all' && isAdmin) || displayPriceId === 'cost') && (
                          <td className="px-4 py-3 font-semibold text-slate-700">{p.price || '—'} ₪</td>
                        )}
                        <td className="px-4 py-3 font-semibold text-slate-700">{p.stock || '0'}</td>
                        {priceLists.filter(pl => displayPriceId === 'all' || displayPriceId === pl.id).map(pl => (
                           <td key={pl.id} className="px-4 py-3 text-slate-600">{(p.prices && p.prices[pl.id]) || '—'} ₪</td>
                        ))}
                        {isEditor && (
                          <td className="px-4 py-3">
                            <div className="flex items-center justify-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button onClick={() => startEdit(p.id)} className="w-8 h-8 rounded-md bg-white text-indigo-600 hover:text-indigo-800 hover:bg-indigo-50 border border-slate-200 flex items-center justify-center transition-colors shadow-sm">
                                <Pencil size={14} />
                              </button>
                              <button onClick={() => delProduct(p.id)} className="w-8 h-8 rounded-md bg-white text-red-600 hover:text-red-800 hover:bg-red-50 border border-slate-200 flex items-center justify-center transition-colors shadow-sm">
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          ) : <Navigate to="/catalog" replace />} />

          {/* ADD / EDIT TAB */}
          <Route path="/add" element={canSeeTab('add') ? (
            <div className="max-w-[540px] mx-auto bg-white rounded-xl shadow-sm border border-slate-200 p-6">
              {editId !== null && (
                <div className="bg-amber-50 border border-amber-300 text-amber-800 px-4 py-3 rounded-lg mb-5 flex items-center gap-2 text-[13px] font-medium">
                  <Pencil size={16} /> <span>מצב עריכה: עורך את מק"ט {fId}</span>
                </div>
              )}
              
              <div className="flex items-center gap-2 mb-5">
                <div className="text-xs font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2 flex-1">
                  <FileEdit size={16} /> פרטי הפריט
                </div>
                {fImg && fImg.startsWith('data:') && (
                  <button 
                    onClick={handleGenerateAI}
                    disabled={isAnalyzing}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 rounded-md text-[13px] font-bold transition-colors disabled:opacity-50"
                  >
                    {isAnalyzing ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                    {isAnalyzing ? 'מנתח...' : "✨ Générer avec l'IA"}
                  </button>
                )}
              </div>

              <div className="space-y-4">
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 flex items-center gap-2">
                    מק"ט פנימי (Internal ID) *
                  </label>
                  <input type="text" value={fId} onChange={e => setFId(e.target.value)} placeholder="לדוגמה: PRD-001" className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 flex items-center gap-2">
                    מק"ט כללי (SKU) *
                    {aiGeneratedFields.sku && <span className="bg-indigo-100 text-indigo-700 text-[10px] px-1.5 py-0.5 rounded flex items-center gap-1"><Sparkles size={10} /> IA</span>}
                  </label>
                  <input type="text" value={fSku} onChange={e => setFSku(e.target.value)} placeholder="לדוגמה: SKU-12345" className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 flex items-center gap-2">
                    תיאור מוצר *
                    {aiGeneratedFields.desc && <span className="bg-indigo-100 text-indigo-700 text-[10px] px-1.5 py-0.5 rounded flex items-center gap-1"><Sparkles size={10} /> IA</span>}
                  </label>
                  <textarea value={fDesc} onChange={e => setFDesc(e.target.value)} placeholder="תיאור מפורט של המוצר..." className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium min-h-[80px] resize-y"></textarea>
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 flex items-center gap-2">
                    קטגוריה
                    {aiGeneratedFields.category && <span className="bg-indigo-100 text-indigo-700 text-[10px] px-1.5 py-0.5 rounded flex items-center gap-1"><Sparkles size={10} /> IA</span>}
                  </label>
                  <input type="text" list="category-options" value={fCategory} onChange={e => setFCategory(e.target.value)} placeholder="בחר קטגוריה או הזן חדשה..." className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                  <datalist id="category-options">
                    {uniqueCategories.map(c => <option key={c} value={c} />)}
                  </datalist>
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 flex items-center gap-2">
                    ספק
                    {aiGeneratedFields.supplier && <span className="bg-indigo-100 text-indigo-700 text-[10px] px-1.5 py-0.5 rounded flex items-center gap-1"><Sparkles size={10} /> IA</span>}
                  </label>
                  <input type="text" value={fSup} onChange={e => setFSup(e.target.value)} placeholder="שם הספק" className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">מחיר</label>
                  <input type="number" dir="ltr" value={fPrice} onChange={e => setFPrice(e.target.value)} placeholder="0.00" className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">מלאי</label>
                  <input type="number" dir="ltr" value={fStock} onChange={e => setFStock(e.target.value)} placeholder="0" className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                </div>
                
                {priceLists.map(pl => (
                  <div key={pl.id}>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">מחירון: {pl.name}</label>
                    <input type="number" dir="ltr" value={fPrices[pl.id] || ''} onChange={e => setFPrices({...fPrices, [pl.id]: e.target.value ? Number(e.target.value) : ''})} placeholder="0.00" className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 transition-all font-medium" />
                  </div>
                ))}
                
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">תמונת מוצר (עד 5MB)</label>
                  <div className="relative border-2 border-dashed border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/50 rounded-lg p-6 text-center transition-colors cursor-pointer group bg-slate-50">
                    <input type="file" accept="image/*" onChange={(e) => onImgPick(e, setFImg, setFImgFile)} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                    
                    {fImg ? (
                      <img src={fImg} alt="Preview" className="w-full max-h-[160px] object-contain rounded-md mx-auto" />
                    ) : (
                      <div className="flex flex-col items-center">
                        <Upload className="text-slate-400 mb-2 group-hover:text-indigo-500 transition-colors" size={32} />
                        <span className="text-sm text-slate-500 font-medium">לחץ או גרור להעלאת תמונה</span>
                      </div>
                    )}
                  </div>
                  {fImg && (
                    <button onClick={() => { setFImg(null); setFImgFile(null); }} className="mt-2 text-[12px] text-red-500 hover:text-red-700 font-medium flex items-center gap-1">
                      <Trash2 size={12} /> הסר תמונה
                    </button>
                  )}
                </div>

                <div className="flex flex-wrap gap-2 pt-3 border-t border-slate-100">
                  <button onClick={saveProduct} disabled={isUploading} className="flex items-center gap-2 bg-rose-500 hover:bg-rose-600 disabled:bg-rose-300 text-white px-5 py-2.5 rounded-md text-[13px] font-bold transition-colors">
                    {isUploading ? <Loader2 size={16} className="animate-spin" /> : <Check size={16} />} 
                    {isUploading ? 'שומר...' : (editId !== null ? 'עדכן פריט' : 'הוסף לקטלוג')}
                  </button>
                  <button onClick={cancelEdit} className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 px-5 py-2.5 rounded-md text-[13px] font-bold transition-colors border border-slate-200">
                    <X size={16} /> נקה טופס
                  </button>
                </div>
              </div>
            </div>
          ) : <Navigate to="/catalog" replace />} />

          {/* COMPANY TAB (Admin Only / if canSeeTab) */}
          <Route path="/company" element={canSeeTab('company') ? (
            <div className="max-w-[540px] mx-auto bg-white rounded-xl shadow-sm border border-slate-200 p-6">
              <div className="text-xs font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2 mb-5">
                <Building2 size={16} /> פרטי החברה
              </div>
              
              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">לוגו חברה</label>
                  <div className="flex items-center gap-4">
                    <div className="relative w-20 h-20 rounded-lg bg-slate-50 border-2 border-dashed border-slate-200 flex items-center justify-center overflow-hidden hover:border-indigo-400 transition-colors cursor-pointer group">
                      <input type="file" accept="image/*" onChange={(e) => onImgPick(e, setCLogo)} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
                      {cLogo ? (
                        <img src={cLogo} alt="Logo" className="w-full h-full object-contain bg-white" />
                      ) : (
                        <Upload className="text-slate-400 group-hover:text-indigo-400" size={24} />
                      )}
                    </div>
                  </div>
                  {cLogo && (
                    <button onClick={() => setCLogo(null)} className="mt-2 text-[12px] text-red-500 hover:text-red-700 font-medium flex items-center gap-1">
                      <Trash2 size={12} /> הסר לוגו
                    </button>
                  )}
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">שם החברה</label>
                  <input type="text" value={cName} onChange={e => setCName(e.target.value)} className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">ח.פ / ת.ז</label>
                  <input type="text" value={cHp} onChange={e => setCHp(e.target.value)} className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">כתובת מלאה</label>
                  <input type="text" value={cAddr} onChange={e => setCAddr(e.target.value)} className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">טלפון / נייד</label>
                  <input type="text" value={cPhone} onChange={e => setCPhone(e.target.value)} className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">אימייל</label>
                  <input type="email" value={cEmail} onChange={e => setCEmail(e.target.value)} className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 transition-all font-medium" />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">אתר אינטרנט</label>
                  <input type="text" value={cWeb} onChange={e => setCWeb(e.target.value)} className="w-full px-3 py-2 border border-slate-200 rounded-md bg-slate-50 text-[14px] outline-none focus:border-indigo-400 focus:ring-2 transition-all font-medium" />
                </div>

                <div className="pt-3 border-t border-slate-100">
                  <button onClick={saveCompany} className="flex items-center gap-2 bg-rose-500 hover:bg-rose-600 text-white px-5 py-2.5 rounded-md text-[13px] font-bold transition-colors shadow-sm">
                    <Save size={16} /> שמור פרטים מזהים
                  </button>
                </div>
              </div>
            </div>
          ) : <Navigate to="/catalog" replace />} />

          {/* PRICE LISTS */}
          <Route path="/pricelists" element={canSeeTab('pricelists') ? (
            <div className="max-w-3xl mx-auto space-y-6">
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-100 bg-slate-50 flex flex-col md:flex-row gap-4 justify-between items-start md:items-center">
                  <div>
                    <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                       <List size={16} className="text-indigo-600" /> ניהול מחירונים
                    </h3>
                    <p className="text-xs text-slate-500 mt-1">נהל מחירונים נוספים עבור הקטלוג (לדוגמה: מחיר קמעונאי, סיטונאי וכו').</p>
                  </div>
                  <div className="flex w-full md:w-auto items-center gap-2">
                    <input 
                      type="text" 
                      placeholder="שם מחירון חדש..." 
                      className="flex-1 w-full md:w-48 px-3 py-1.5 border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-1 focus:ring-indigo-100 transition-all font-medium bg-white"
                      value={newPriceListName}
                      onChange={(e) => setNewPriceListName(e.target.value)}
                    />
                    <button onClick={() => {
                        if (newPriceListName.trim()) {
                           savePriceLists([...priceLists, { id: 'pl_' + Date.now().toString(), name: newPriceListName.trim() }]);
                           setNewPriceListName('');
                        } else {
                           showToast('⚠️ יש להזין שם מחירון');
                        }
                    }} className="text-xs shrink-0 bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-md font-medium transition-colors flex items-center gap-1 h-[34px]">
                      <Plus size={14} /> מחירון חדש
                    </button>
                  </div>
                </div>
                <div className="px-6 py-4">
                  {priceLists.length === 0 ? (
                    <div className="text-center py-6 text-slate-500 text-sm">אין מחירונים מוגדרים.</div>
                  ) : (
                    <div className="space-y-3">
                      {priceLists.map((pl, idx) => (
                        <div key={pl.id} className="flex justify-between items-center bg-slate-50 border border-slate-200 p-3 rounded-lg flex-wrap gap-3">
                          {editingPriceListId === pl.id ? (
                            <div className="flex flex-1 items-center gap-2 min-w-[200px]">
                              <input 
                                type="text" 
                                className="w-full px-3 py-1.5 border border-indigo-300 rounded-md text-[13px] outline-none focus:ring-1 focus:ring-indigo-100 transition-all font-medium bg-white"
                                value={editingPriceListName}
                                onChange={(e) => setEditingPriceListName(e.target.value)}
                                autoFocus
                              />
                            </div>
                          ) : (
                            <span className="font-bold text-slate-800 text-sm">{pl.name}</span>
                          )}
                          <div className="flex gap-2 relative">
                            {editingPriceListId === pl.id ? (
                               <>
                                 <button onClick={() => {
                                    if (editingPriceListName.trim()) {
                                       const updated = [...priceLists];
                                       updated[idx].name = editingPriceListName.trim();
                                       savePriceLists(updated);
                                       setEditingPriceListId(null);
                                    }
                                 }} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-md text-xs font-semibold transition-colors flex items-center gap-1">שמור</button>
                                 <button onClick={() => setEditingPriceListId(null)} className="bg-slate-200 hover:bg-slate-300 text-slate-700 px-3 py-1.5 rounded-md text-xs font-semibold transition-colors">בטל</button>
                               </>
                            ) : (
                               <>
                                <div className="relative line-height-none">
                                  <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handlePlExcelPick(e, pl)} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer text-[0]" title="ייבוא מחירון מ-Excel" />
                                  <button className="text-emerald-600 hover:bg-emerald-100 p-1.5 rounded-md transition-colors" title="ייבוא מחירון מ-Excel">
                                     <FileSpreadsheet size={14} />
                                  </button>
                                </div>
                                <button onClick={() => {
                                    setEditingPriceListId(pl.id);
                                    setEditingPriceListName(pl.name);
                                }} className="text-indigo-600 hover:bg-indigo-100 p-1.5 rounded-md transition-colors" title="ערוך שם">
                                   <Pencil size={14} />
                                </button>
                                {priceListToDelete === pl.id ? (
                                   <div className="absolute bg-white border border-red-200 shadow-md rounded-md p-2 flex flex-col gap-2 z-10 bottom-full mb-1 ml-4 whitespace-nowrap">
                                     <div className="text-[11px] font-bold text-red-700">למחוק מחירון?</div>
                                     <div className="flex gap-1">
                                        <button onClick={() => {
                                          savePriceLists(priceLists.filter(x => x.id !== pl.id));
                                          setPriceListToDelete(null);
                                        }} className="bg-red-600 text-white px-2 py-1 text-[10px] rounded hover:bg-red-700">מחק</button>
                                        <button onClick={() => setPriceListToDelete(null)} className="bg-slate-200 text-slate-800 px-2 py-1 text-[10px] rounded hover:bg-slate-300">בטל</button>
                                     </div>
                                   </div>
                                ) : null}
                                <button onClick={() => {
                                    setPriceListToDelete(pl.id);
                                }} className="text-red-500 hover:bg-red-50 p-1.5 rounded-md transition-colors" title="מחק">
                                   <Trash2 size={14} />
                                </button>
                               </>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          ) : <Navigate to="/catalog" replace />} />

          {/* SYSTEM SETTINGS / ADMIN */}
          <Route path="/admin" element={canSeeTab('admin') ? (
            <Admin onNavigateToCompany={() => navigate('/company')} />
          ) : <Navigate to="/catalog" replace />} />
          
          <Route path="/" element={<Navigate to="/catalog" replace />} />
          <Route path="*" element={<Navigate to="/catalog" replace />} />
          </Routes>
          
        </main>
      </div>



      {/* Import Mapping Modal */}
      {importModalOpen && (
        <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center z-[200] p-4 transition-opacity">
          <div className="bg-white rounded-xl shadow-xl max-w-lg w-full p-6 animate-in fade-in zoom-in-95 duration-200 overflow-y-auto max-h-[90vh]">
            <h3 className="text-xl font-bold text-slate-800 mb-2 flex items-center gap-2">
              <FileSpreadsheet size={20} className="text-emerald-600" /> הגדרת שדות לייבוא
            </h3>
            <p className="text-slate-600 mb-6 text-sm">
              לחץ על כל שדה במערכת ובחר את העמודה התואמת מקובץ האקסל שהעלית.
            </p>
            
            <div className="space-y-4 mb-6 relative">
              {[
                { key: 'internalId', label: 'מק"ט פנימי', required: false },
                { key: 'sku', label: 'מק"ט כללי (SKU)', required: false },
                { key: 'desc', label: 'תיאור מוצר (חובה למוצר חדש)', required: false },
                { key: 'supplier', label: 'ספק', required: false },
                { key: 'category', label: 'קטגוריה', required: false },
                { key: 'img', label: 'תמונה (קישור יחסי או מלא)', required: false },
                { key: 'price', label: 'מחיר עלות', required: false },
                { key: 'stock', label: 'מלאי', required: false },
                ...priceLists.map(pl => ({ key: `priceList_${pl.id}`, label: `מחירון: ${pl.name}`, required: false }))
              ].map(({ key, label, required }) => (
                <div key={key} className="flex flex-col gap-1.5">
                  <label className="text-xs font-bold text-slate-700">
                    {label} {required && <span className="text-red-500">*</span>}
                  </label>
                  <select
                    value={fieldMapping[key as keyof typeof fieldMapping]}
                    onChange={e => setFieldMapping(prev => ({ ...prev, [key]: e.target.value }))}
                    className="w-full py-2 px-3 bg-slate-50 border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-2 transition-colors cursor-pointer appearance-none"
                  >
                    <option value="">-- אל תייבא שדה --</option>
                    {importHeaders.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>

            <div className="flex gap-3 justify-end items-center" dir="ltr">
              {(!fieldMapping.internalId && !fieldMapping.sku) && (
                 <div className="text-red-500 text-xs font-semibold px-4 text-right flex-1" dir="rtl">
                    * חובה לשייך לפחות מק"ט אחד (כללי או פנימי)
                 </div>
              )}
              <button 
                onClick={executeImport}
                disabled={(!fieldMapping.internalId && !fieldMapping.sku)}
                className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-sm font-semibold rounded-md transition-colors shadow-sm flex items-center gap-2"
              >
                <Check size={16} /> בצע ייבוא רשומות לסנכרון
              </button>
              <button 
                onClick={() => { setImportModalOpen(false); setImportData([]); }}
                className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-semibold rounded-md transition-colors"
               >
                ביטול
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Price List Import Mapping Modal */}
      {importPlModalOpen && importingPl && (
        <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center z-[200] p-4 transition-opacity">
          <div className="bg-white rounded-xl shadow-xl max-w-lg w-full p-6 animate-in fade-in zoom-in-95 duration-200 overflow-y-auto max-h-[90vh]">
            <h3 className="text-xl font-bold text-slate-800 mb-2 flex items-center gap-2">
              <FileSpreadsheet size={20} className="text-indigo-600" /> עדכון מחירים "מחירון: {importingPl.name}"
            </h3>
            <p className="text-sm text-slate-500 mb-6 font-medium">
              נמצאו {plImportData.length} שורות. אנא שייך את העמודות הרלוונטיות. המערכת תעדכן רק מחירי מוצרים שקיימים במערכת לפי מק"ט.
            </p>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 auto-rows-max mb-8 border-t border-slate-100 pt-4">
              {[
                { key: 'internalId', label: 'מק"ט פנימי', required: true },
                { key: 'sku', label: 'מק"ט יצרן (Barcode)', required: true },
                { key: 'price', label: 'שורת מחיר בעדכון', required: true }
              ].map(({ key, label, required }) => (
                <div key={key} className="flex flex-col gap-1.5">
                  <label className="text-xs font-bold text-slate-700">
                    {label} {required && <span className="text-red-500">*</span>}
                  </label>
                  <select
                    value={plFieldMapping[key as keyof typeof plFieldMapping]}
                    onChange={e => setPlFieldMapping(prev => ({ ...prev, [key]: e.target.value }))}
                    className="w-full py-2 px-3 bg-slate-50 border border-slate-200 rounded-md text-[13px] outline-none focus:border-indigo-400 focus:ring-2 transition-colors cursor-pointer appearance-none"
                  >
                    <option value="">-- אל תייבא שדה --</option>
                    {plImportHeaders.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>

            <div className="flex gap-3 justify-end items-center" dir="ltr">
              {(!plFieldMapping.internalId && !plFieldMapping.sku) && (
                 <div className="text-red-500 text-xs font-semibold px-4 text-right flex-1" dir="rtl">
                    * חובה לשייך לפחות מק"ט אחד (כללי או פנימי)
                 </div>
              )}
              {(!plFieldMapping.price) && (
                 <div className="text-red-500 text-xs font-semibold px-4 text-right flex-1" dir="rtl">
                    * חובה לשייך שדה מחיר ערוך
                 </div>
              )}
              <button 
                onClick={executePlImport}
                disabled={(!plFieldMapping.internalId && !plFieldMapping.sku) || (!plFieldMapping.price)}
                className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-sm font-semibold rounded-md transition-colors shadow-sm flex items-center gap-2"
              >
                <Check size={16} /> בצע עדכון מחירון
              </button>
              <button 
                onClick={() => { setImportPlModalOpen(false); setPlImportData([]); setImportingPl(null); }}
                className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-semibold rounded-md transition-colors"
               >
                ביטול
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Global Toast Notification */}
      <div className={`fixed bottom-6 left-1/2 -translate-x-1/2 bg-slate-900 text-white px-5 py-3 rounded-lg shadow-xl text-[13px] font-medium flex items-center gap-2 z-[100] transition-all duration-300 pointer-events-none transform ${toastMessage ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4'}`}>
        {toastMessage}
      </div>
    </div>
  );
}

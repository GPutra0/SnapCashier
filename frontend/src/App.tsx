/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { 
  LayoutDashboard, 
  ShoppingCart, 
  Package, 
  Wallet, 
  Plus, 
  Search, 
  Trash2, 
  Edit, 
  AlertCircle, 
  TrendingUp, 
  TrendingDown, 
  ChevronRight,
  X,
  CreditCard,
  History,
  PackageSearch,
  CheckCircle2,
  Minus,
  Printer,
  Users,
  Clock,
  BookOpen,
  ChevronDown,
  QrCode,
  Check,
  Mail,
  Lock,
  EyeOff,
  User,
  ChevronLeft
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import * as XLSX from 'xlsx';
import { auth, googleProvider, signInWithPopup, signOut, onAuthStateChanged, db, updateProfile } from './firebase';
import { collection, doc, writeBatch, onSnapshot, query, setDoc, deleteDoc, getDocs, getDoc } from 'firebase/firestore';

// --- Types ---

interface Product {
  id: string;
  nama: string;
  harga: number;
  stok: number;
  url_gambar: string;
  kategori: string;
}

interface FinanceRecord {
  id: string;
  tipe: 'pemasukan' | 'pengeluaran';
  jumlah: number;
  keterangan: string;
  tanggal: string;
}

interface CartItem {
  id: string;
  productId: string;
  nama: string;
  harga: number;
  jumlah: number;
}

interface DebtItem {
  nama: string;
  harga: number;
  jumlah: number;
}

interface Debt {
  id: string;
  pemberiHutang: string; // Nama pelanggan yang berhutang
  items: DebtItem[];
  total: number;
  tanggal: string;
  status: 'belum_lunas' | 'lunas';
}

// --- Categories ---
const CATEGORIES = [
  "Sembako",
  "Bahan Pangan",
  "Bumbu Dapur",
  "Produk Susu & Olahannya",
  "Minuman (Beverages)",
  "Makanan Ringan (Snacks)",
  "Mie & Instan",
  "Perawatan Tubuh (Toiletries)",
  "Kebutuhan Rumah Tangga",
  "Lain-lain"
];

// --- Mock Data ---
const INITIAL_PRODUCTS: Product[] = [
  { id: 'B001', nama: 'Masako Ayam 10g', harga: 500, stok: 100, kategori: 'Bumbu Dapur', url_gambar: 'https://images.unsplash.com/photo-1598103442097-8b74394b95c6?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'B002', nama: 'Royco Sapi 10g', harga: 500, stok: 80, kategori: 'Bumbu Dapur', url_gambar: 'https://images.unsplash.com/photo-1598103442097-8b74394b95c6?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'B003', nama: 'Ladaku Merica Bubuk', harga: 1000, stok: 45, kategori: 'Bumbu Dapur', url_gambar: 'https://images.unsplash.com/photo-1599940824399-b87987cb9721?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'S001', nama: 'Beras Pandan Wangi 1kg', harga: 15000, stok: 20, kategori: 'Bahan Pangan', url_gambar: 'https://images.unsplash.com/photo-1586201375761-83865001e31c?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'S002', nama: 'Gula Pasir 1kg', harga: 16000, stok: 15, kategori: 'Sembako', url_gambar: 'https://images.unsplash.com/photo-1581441363689-1f3c3c414635?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'S003', nama: 'Minyak Goreng 1L', harga: 18000, stok: 10, kategori: 'Sembako', url_gambar: 'https://images.unsplash.com/photo-1474979266404-7eaacbcd87c5?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'M001', nama: 'Indomie Goreng Original', harga: 3500, stok: 40, kategori: 'Mie & Instan', url_gambar: 'https://images.unsplash.com/photo-1569718212165-3a8278d5f624?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'M002', nama: 'Mie Sedaap Kari Spesial', harga: 3200, stok: 35, kategori: 'Mie & Instan', url_gambar: 'https://images.unsplash.com/photo-1569718212165-3a8278d5f624?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'J001', nama: 'Chiki Balls Keju', harga: 2000, stok: 15, kategori: 'Makanan Ringan (Snacks)', url_gambar: 'https://images.unsplash.com/photo-1621447509323-570a17e47dcd?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'J002', nama: 'Roma Kelapa 300g', harga: 10000, stok: 12, kategori: 'Makanan Ringan (Snacks)', url_gambar: 'https://images.unsplash.com/photo-1558961363-fa8fdf82db35?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'D001', nama: 'Susu UHT Full Cream', harga: 6000, stok: 12, kategori: 'Produk Susu & Olahannya', url_gambar: 'https://images.unsplash.com/photo-1563636619-e910ef4ff55d?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'T001', nama: 'Sabun Lifebuoy Merah', harga: 4500, stok: 20, kategori: 'Perawatan Tubuh (Toiletries)', url_gambar: 'https://images.unsplash.com/photo-1600857062241-98e5dba7f214?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'L001', nama: 'Tisu Wajah 250s', harga: 12000, stok: 4, kategori: 'Lain-lain', url_gambar: 'https://images.unsplash.com/photo-1584622781464-556f82798e4d?q=80&w=250&h=250&auto=format&fit=crop' },
  { id: 'R001', nama: 'Rokok Filter 16', harga: 30000, stok: 10, kategori: 'Lain-lain', url_gambar: 'https://images.unsplash.com/photo-1527137342181-19aab11a8ee1?q=80&w=250&h=250&auto=format&fit=crop' },
];

const INITIAL_FINANCE: FinanceRecord[] = [
  { id: 'f1', tipe: 'pemasukan', jumlah: 550000, keterangan: 'Saldo Awal Toko', tanggal: new Date().toISOString().split('T')[0] },
  { id: 'f2', tipe: 'pengeluaran', jumlah: 200000, keterangan: 'Belanja Stok Bahan', tanggal: new Date().toISOString().split('T')[0] },
];

export default function App() {
  const [user, setUser] = useState<any>(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authMode, setAuthMode] = useState<'login' | 'signup' | 'forgot'>('login');

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const [activeTab, setActiveTab] = useState<'dashboard' | 'pos' | 'inventory' | 'finance' | 'debts'>('dashboard');
  const [products, setProducts] = useState<Product[]>([]);
  const [finance, setFinance] = useState<FinanceRecord[]>(INITIAL_FINANCE);
  const [debts, setDebts] = useState<Debt[]>([]);
  const [hasLoadedData, setHasLoadedData] = useState(false);
  const [cart, setCart] = useState<CartItem[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedMethod, setSelectedMethod] = useState<'Tunai' | 'M-Banking' | 'DANA' | 'OVO' | 'QRIS' | 'Hutang'>('Tunai');
  const [debtorName, setDebtorName] = useState('');
  const [selectedCategoryFilter, setSelectedCategoryFilter] = useState<string>('Semua');
  const [showMobileCart, setShowMobileCart] = useState(false);
  const [importingInfo, setImportingInfo] = useState('');

  const dynamicCategories = useMemo(() => {
    const cats = new Set(CATEGORIES);
    products.forEach(p => cats.add(p.kategori));
    return Array.from(cats);
  }, [products]);

  useEffect(() => {
    const wipeData = async () => {
      if (!user || localStorage.getItem('wiped_products_once')) return;
      try {
        const q = query(collection(db, 'users', user.uid, 'products'));
        const snapshot = await getDocs(q);
        if (snapshot.empty) return;
        
        const batch = writeBatch(db);
        snapshot.forEach(docSnap => {
          batch.delete(docSnap.ref);
        });
        await batch.commit();
        localStorage.setItem('wiped_products_once', 'true');
        console.log('Successfully wiped products');
      } catch (err) {
        console.error('Failed to wipe', err);
      }
    };
    wipeData();
  }, [user]);

  useEffect(() => {
    if (!user) {
      setProducts(INITIAL_PRODUCTS); // fallback to initial if not logged in
      setStoreInfo({ storeName: 'Admin Toko', address: '', profileImage: '' });
      setFinance(INITIAL_FINANCE);
      setDebts([]);
      setHasLoadedData(false);
      return;
    }

    const docRefProfile = doc(db, 'users', user.uid);
    const unsubscribeProfile = onSnapshot(docRefProfile, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        setStoreInfo({ storeName: data.storeName || 'Admin Toko', address: data.address || '', profileImage: data.profileImage || '' });
      } else {
        setStoreInfo({ storeName: 'Admin Toko', address: '', profileImage: '' });
      }
    });

    const initData = async () => {
      try {
        const pFin = getDoc(doc(db, 'users', user.uid, 'data', 'finance_list'));
        const pDeb = getDoc(doc(db, 'users', user.uid, 'data', 'debts_list'));
        const [finSnap, debSnap] = await Promise.all([pFin, pDeb]);
        
        if (finSnap.exists()) {
          setFinance(finSnap.data().records || []);
        } else {
          setFinance([]);
        }
        
        if (debSnap.exists()) {
          setDebts(debSnap.data().records || []);
        } else {
          setDebts([]);
        }
      } catch (err) {
        console.error("Gagal load data", err);
      } finally {
        setHasLoadedData(true);
      }
    };
    initData();

    const q = query(collection(db, 'users', user.uid, 'products'));
    const unsubscribeProd = onSnapshot(q, (snapshot) => {
      const prods: Product[] = [];
      snapshot.forEach((docSnap) => {
        prods.push(docSnap.data() as Product);
      });
      setProducts(prods);
    });

    return () => {
      unsubscribeProd();
      unsubscribeProfile();
    };
  }, [user]);

  useEffect(() => {
    if (!user || !hasLoadedData) return;
    setDoc(doc(db, 'users', user.uid, 'data', 'finance_list'), { records: finance });
  }, [finance, user, hasLoadedData]);

  useEffect(() => {
    if (!user || !hasLoadedData) return;
    setDoc(doc(db, 'users', user.uid, 'data', 'debts_list'), { records: debts });
  }, [debts, user, hasLoadedData]);

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !user) return;
    
    setImportingInfo('Membaca file Excel...');
    
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      // Convert to json
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];
      
      setImportingInfo('Menyimpan ke database...');
      const batch = writeBatch(db);
      
      jsonData.forEach((row, index) => {
        const id = row.id || row.Id || row.ID || `PROD-${Date.now()}-${index}`;
        const prod: Product = {
          id: String(id),
          nama: String(row['Nama Barang'] || row['nama barang'] || row.nama || row.Nama || row.Name || 'Produk Baru'),
          harga: Number(row.Harga || row.harga || row.Price || 0),
          stok: Number(row.Stok || row.stok || row.Stock || 0),
          kategori: String(row['Jenis Barang'] || row['jenis barang'] || row.kategori || row.Kategori || row.Category || 'Lain-lain'),
          url_gambar: String(row.url_gambar || row.Url_Gambar || row.Image || 'https://images.unsplash.com/photo-1556910103-1c02745aae4d?q=80&w=250&h=250&auto=format&fit=crop')
        };
        const docRef = doc(db, 'users', user.uid, 'products', prod.id);
        batch.set(docRef, prod);
      });
      
      await batch.commit();
      setImportingInfo('Berhasil!');
      setTimeout(() => setImportingInfo(''), 3000);
    } catch (error) {
      console.error(error);
      setImportingInfo('Gagal import file.');
      setTimeout(() => setImportingInfo(''), 3000);
    }
    
    // Reset file input
    if (e.target) e.target.value = '';
  };

  const handleDeleteAllProducts = async () => {
    if (!user) return;
    if (confirm('Apakah Anda yakin ingin menghapus SEMUA produk? Tindakan ini tidak dapat dibatalkan.')) {
      setImportingInfo('Menghapus semua...');
      try {
        const batch = writeBatch(db);
        products.forEach(product => {
          const docRef = doc(db, 'users', user.uid, 'products', product.id);
          batch.delete(docRef);
        });
        await batch.commit();
        setImportingInfo('Berhasil dihapus!');
        setTimeout(() => setImportingInfo(''), 3000);
      } catch (error) {
        console.error(error);
        setImportingInfo('Gagal menghapus.');
        setTimeout(() => setImportingInfo(''), 3000);
      }
    }
  };
  
  // Last transaction for invoice
  const [lastTransaction, setLastTransaction] = useState<{
    id: string;
    items: CartItem[];
    total: number;
    date: string;
    method: string;
    debtor?: string | null;
  } | null>(null);
  const [showInvoice, setShowInvoice] = useState(false);
  const [showToast, setShowToast] = useState(false);
  const [isProcessingPayment, setIsProcessingPayment] = useState(false);
  const [selectedBank, setSelectedBank] = useState<string>('BCA');

  // Modals state
  const [showProductModal, setShowProductModal] = useState(false);
  const [showProfileModal, setShowProfileModal] = useState(false);
  const [profileName, setProfileName] = useState('');
  const [profileImage, setProfileImage] = useState('');
  const [profileStoreName, setProfileStoreName] = useState('');
  const [profileAddress, setProfileAddress] = useState('');
  const [storeInfo, setStoreInfo] = useState({ storeName: 'Admin Toko', address: '', profileImage: '' });
  const [editingProduct, setEditingProduct] = useState<Product | null>(null);
  const [productToDelete, setProductToDelete] = useState<Product | null>(null);
  const [formImage, setFormImage] = useState<string>('');
  const [showDebtModal, setShowDebtModal] = useState(false);

  const handleProfileImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const MAX_WIDTH = 250;
        const MAX_HEIGHT = 250;
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
        
        const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
        setProfileImage(dataUrl);
      };
      img.src = event.target?.result as string;
    };
    reader.readAsDataURL(file);
  };

  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        const MAX_WIDTH = 500;
        const MAX_HEIGHT = 500;
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
        
        const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
        setFormImage(dataUrl);
      };
      img.src = event.target?.result as string;
    };
    reader.readAsDataURL(file);
  };

  // --- Helpers ---
  const addToCart = (product: Product) => {
    if (product.stok <= 0) return;
    
    // Trigger feedback visual (opsional: bisa tambah sound effect)
    setCart(prev => {
      const existing = prev.find(item => item.productId === product.id);
      if (existing) {
        return prev.map(item => 
          item.productId === product.id 
            ? { ...item, jumlah: item.jumlah + 1 } 
            : item
        );
      }
      return [...prev, { 
        id: Math.random().toString(36).substr(2, 9), 
        productId: product.id, 
        nama: product.nama, 
        harga: product.harga, 
        jumlah: 1 
      }];
    });
  };

  const updateCartQuantity = (id: string, delta: number) => {
    setCart(prev => prev.map(item => {
      if (item.id === id) {
        const newQty = Math.max(1, item.jumlah + delta);
        // Cek stok asli
        const product = products.find(p => p.id === item.productId);
        if (product && newQty > product.stok && delta > 0) return item;
        return { ...item, jumlah: newQty };
      }
      return item;
    }));
  };

  const removeFromCart = (id: string) => {
    setCart(prev => prev.filter(item => item.id !== id));
  };

  const clearCart = () => setCart([]);

  const initiatePayment = () => {
    const total = cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0);
    if (total === 0) return;

    if (selectedMethod === 'Hutang' && !debtorName.trim()) {
      alert('Mohon masukkan nama penghutang');
      return;
    }

    setIsProcessingPayment(true);
    setShowMobileCart(false);
  };

  const processPayment = () => {
    const total = cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0);
    if (total === 0) return;

    if (selectedMethod === 'Hutang' && !debtorName.trim()) {
      alert('Mohon masukkan nama penghutang');
      return;
    }

    const transactionId = Math.random().toString(36).substr(2, 9).toUpperCase();
    const date = new Date().toLocaleString('id-ID', { 
      day: 'numeric', 
      month: 'long', 
      year: 'numeric',
      hour: '2-digit', 
      minute: '2-digit' 
    });

    if (selectedMethod === 'Hutang') {
      setDebts(prev => {
        const existingDebtIndex = prev.findIndex(d => d.pemberiHutang.toLowerCase() === debtorName.toLowerCase() && d.status === 'belum_lunas');
        
        if (existingDebtIndex !== -1) {
          const updatedDebts = [...prev];
          const existingDebt = updatedDebts[existingDebtIndex];
          updatedDebts[existingDebtIndex] = {
            ...existingDebt,
            items: [...existingDebt.items, ...cart.map(item => ({ nama: item.nama, harga: item.harga, jumlah: item.jumlah }))],
            total: existingDebt.total + total,
            // Perbarui tanggal ke transaksi terakhir
            tanggal: new Date().toISOString().split('T')[0]
          };
          return updatedDebts;
        } else {
          const newDebt: Debt = {
            id: transactionId,
            pemberiHutang: debtorName,
            items: cart.map(item => ({ nama: item.nama, harga: item.harga, jumlah: item.jumlah })),
            total: total,
            tanggal: new Date().toISOString().split('T')[0],
            status: 'belum_lunas'
          };
          return [newDebt, ...prev];
        }
      });
      setDebtorName('');
    } else {
      // Add to finance only if not debt (debt is not income yet)
      const newRecord: FinanceRecord = {
        id: transactionId,
        tipe: 'pemasukan',
        jumlah: total,
        keterangan: `Penjualan #${transactionId}`,
        tanggal: new Date().toISOString().split('T')[0]
      };
      setFinance(prev => [newRecord, ...prev]);
    }

    // Save for invoice
    setLastTransaction({
      id: transactionId,
      items: [...cart],
      total: total,
      date: date,
      method: selectedMethod,
      debtor: selectedMethod === 'Hutang' ? debtorName : null
    });

    // Update inventory
    setProducts(prev => prev.map(p => {
      const cartItem = cart.find(ci => ci.productId === p.id);
      return cartItem ? { ...p, stok: p.stok - cartItem.jumlah } : p;
    }));
    if (user) {
      const batch = writeBatch(db);
      cart.forEach(ci => {
        const prod = products.find(p => p.id === ci.productId);
        if (prod) {
          batch.update(doc(db, 'users', user.uid, 'products', prod.id), { stok: prod.stok - ci.jumlah });
        }
      });
      batch.commit().catch(e => console.error(e));
    }
    
    clearCart();
    setShowInvoice(true);
  };

  const totalBelanja = cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0);
  
  const handleEditOrder = () => {
    if (!lastTransaction) return;
    setCart(lastTransaction.items);
    setProducts(prev => prev.map(p => {
      const item = lastTransaction.items.find(ci => ci.productId === p.id);
      return item ? { ...p, stok: p.stok + item.jumlah } : p;
    }));
    if (user) {
      const batch = writeBatch(db);
      lastTransaction.items.forEach(ci => {
        const prod = products.find(p => p.id === ci.productId);
        if (prod) {
          batch.update(doc(db, 'users', user.uid, 'products', prod.id), { stok: prod.stok + ci.jumlah });
        }
      });
      batch.commit().catch(e => console.error(e));
    }
    setFinance(prev => prev.filter(f => f.id !== lastTransaction.id));
    setShowInvoice(false);
    setActiveTab('pos');
  };

  // Financial Stats
  const totalIncome = useMemo(() => 
    finance.filter(f => f.tipe === 'pemasukan').reduce((sum, f) => sum + f.jumlah, 0)
  , [finance]);

  const totalExpense = useMemo(() => 
    finance.filter(f => f.tipe === 'pengeluaran').reduce((sum, f) => sum + f.jumlah, 0)
  , [finance]);

  const totalPiutang = useMemo(() => 
    debts.filter(d => d.status === 'belum_lunas').reduce((sum, d) => sum + d.total, 0)
  , [debts]);

  const monthlyChartData = useMemo(() => {
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];
    const monthlyMap: { [key: string]: { month: string, sortKey: string, Pemasukan: number, Pengeluaran: number } } = {};

    finance.forEach(record => {
      const date = new Date(record.tanggal);
      if (isNaN(date.getTime())) return;
      
      const monthYear = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
      const sortKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
      
      if (!monthlyMap[monthYear]) {
        monthlyMap[monthYear] = { month: monthYear, sortKey, Pemasukan: 0, Pengeluaran: 0 };
      }
      
      if (record.tipe === 'pemasukan') {
        monthlyMap[monthYear].Pemasukan += record.jumlah;
      } else if (record.tipe === 'pengeluaran') {
        monthlyMap[monthYear].Pengeluaran += record.jumlah;
      }
    });

    return Object.values(monthlyMap).sort((a, b) => a.sortKey.localeCompare(b.sortKey));
  }, [finance]);

  const balance = totalIncome - totalExpense;

  // Low Stock Items
  const lowStockItems = useMemo(() => 
    products.filter(p => p.stok < 5)
  , [products]);

  // --- Views ---

  const Sidebar = () => (
    <aside className="w-72 bg-white border-r border-slate-200/60 h-screen fixed left-0 top-0 hidden md:flex flex-col shadow-[1px_0_10px_rgba(0,0,0,0.02)] z-50">
      <div className="p-8 border-b border-slate-100/80 flex items-center gap-3">
        <div className="w-12 h-12 bg-gradient-to-br from-blue-600 to-indigo-700 rounded-2xl flex items-center justify-center text-white shadow-xl shadow-blue-200/50">
          <CreditCard className="w-7 h-7" />
        </div>
        <div>
          <h1 className="font-display font-extrabold text-slate-900 text-xl tracking-tight leading-none">SnapCashier</h1>
          <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">Sistem Pintar</p>
        </div>
      </div>
      
      <nav className="flex-1 p-5 space-y-2.5 mt-4">
        {[
          { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
          { id: 'inventory', label: 'Stok Barang', icon: Package },
          { id: 'debts', label: 'Buku Hutang', icon: BookOpen },
          { id: 'finance', label: 'Laporan Keuangan', icon: Wallet },
          { id: 'pos', label: 'Kasir Pintar', icon: ShoppingCart },
        ].map((item) => (
          <button
            key={item.id}
            onClick={() => setActiveTab(item.id as any)}
            className={`w-full flex items-center gap-4 px-5 py-3.5 rounded-2xl transition-all duration-300 font-semibold text-sm group ${
              activeTab === item.id 
                ? 'bg-blue-600 text-white shadow-lg shadow-blue-200/40 translate-x-1' 
                : 'text-slate-500 hover:bg-slate-50 hover:text-blue-600 hover:translate-x-1'
            }`}
          >
            <item.icon className={`w-5 h-5 transition-colors ${activeTab === item.id ? 'text-white' : 'text-slate-400 group-hover:text-blue-600'}`} />
            {item.label}
          </button>
        ))}
      </nav>

      <div className="p-6 border-t border-slate-100/80">
        <div className="bg-slate-50 rounded-2xl p-4 border border-slate-100 flex flex-col gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-100 to-indigo-100 flex items-center justify-center text-blue-600 font-extrabold text-sm shadow-inner shrink-0 overflow-hidden">
              {(storeInfo.profileImage || user?.photoURL) ? <img src={storeInfo.profileImage || user?.photoURL} alt="Avatar" className="w-full h-full object-cover" /> : user?.email?.charAt(0).toUpperCase()}
            </div>
            <div className="overflow-hidden">
              <p className="text-sm font-bold text-slate-800 truncate">{storeInfo.storeName}</p>
              <p className="text-[10px] font-bold text-slate-500 truncate">{storeInfo.address}</p>
              <p className="text-[10px] font-bold text-blue-500 uppercase truncate">{user?.displayName || user?.email}</p>
            </div>
          </div>
          <div className="flex gap-2">
            <button 
              onClick={() => {
                setProfileName(user?.displayName || '');
                setProfileImage(storeInfo.profileImage || user?.photoURL || '');
                setProfileStoreName(storeInfo.storeName);
                setProfileAddress(storeInfo.address);
                setShowProfileModal(true);
              }}
              className="flex-1 py-2 text-xs font-black text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-xl transition-all border border-blue-100/50 uppercase tracking-widest"
            >
              Profil
            </button>
            <button 
              onClick={() => signOut(auth)}
              className="flex-1 py-2 text-xs font-black text-rose-500 bg-rose-50 hover:bg-rose-100 rounded-xl transition-all border border-rose-100/50 uppercase tracking-widest"
            >
              Keluar
            </button>
          </div>
        </div>
      </div>
    </aside>
  );

  if (authLoading) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center">
        <div className="w-16 h-16 border-4 border-blue-500 border-t-transparent rounded-full animate-spin"></div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4 sm:p-8 bg-sky-100/80 bg-[url('https://images.unsplash.com/photo-1534088568595-a066f410cbda?auto=format&fit=crop&w=2000&q=80')] bg-cover bg-center">
        <div className="absolute inset-0 bg-white/30 backdrop-blur-[2px]"></div>
        <motion.div 
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          className="bg-white rounded-[2rem] sm:rounded-[3rem] w-full max-w-5xl shadow-[0_20px_60px_-15px_rgba(0,0,0,0.3)] flex flex-col md:flex-row p-3 sm:p-4 gap-4 sm:gap-8 relative z-10"
        >
          {/* Left Graphic */}
          <div className="relative w-full md:w-[45%] lg:w-1/2 aspect-square md:aspect-auto rounded-[1.5rem] sm:rounded-[2.5rem] overflow-hidden hidden sm:block">
            <img 
              src="https://images.unsplash.com/photo-1600880292203-757bb62b4baf?auto=format&fit=crop&w=1000&q=80" 
              alt="Cashier" 
              className="absolute inset-0 w-full h-full object-cover"
            />
            <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-black/20 to-transparent"></div>
            <div className="absolute bottom-10 left-8 right-8 lg:bottom-12 lg:left-12">
              <h2 className="text-4xl lg:text-5xl font-black text-white leading-[1.1] tracking-tighter">
                SELL.<br />MANAGE.<br />SCALE.
              </h2>
            </div>
          </div>

          {/* Right Form */}
          <div className="w-full md:w-[55%] lg:w-1/2 flex flex-col justify-center py-8 sm:py-12 md:py-16 px-6 sm:px-10 lg:px-14 relative bg-white">
            {authMode !== 'login' && (
              <button 
                onClick={() => setAuthMode('login')}
                className="absolute top-6 left-6 sm:top-8 sm:left-8 w-10 h-10 bg-slate-50 hover:bg-slate-100 rounded-full flex items-center justify-center text-slate-700 transition-colors shadow-sm"
              >
                <ChevronLeft className="w-5 h-5" />
              </button>
            )}

            {/* Login Form */}
            {authMode === 'login' && (
              <div className="w-full max-w-[360px] mx-auto transition-all animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="text-center mb-8">
                  <h1 className="text-3xl font-black text-slate-900 tracking-tight mb-3">Log in</h1>
                  <p className="text-slate-500 text-xs sm:text-sm font-medium leading-relaxed px-4">
                    Enter your email and password to securely access your account and manage your services.
                  </p>
                </div>

                <form className="space-y-4" onSubmit={(e) => { e.preventDefault(); signInWithPopup(auth, googleProvider); }}>
                  <div className="space-y-4">
                    <div className="relative">
                      <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                        <Mail className="w-5 h-5" />
                      </div>
                      <input 
                        type="email" 
                        placeholder="Email address" 
                        className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-5 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                      />
                    </div>
                    <div className="relative">
                      <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                        <Lock className="w-5 h-5" />
                      </div>
                      <input 
                        type="password" 
                        placeholder="Password" 
                        className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-12 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                      />
                      <button type="button" className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors">
                        <EyeOff className="w-5 h-5" />
                      </button>
                    </div>
                  </div>

                  <div className="flex items-center justify-between py-2 px-1">
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <div className="w-4 h-4 rounded border border-slate-300 flex items-center justify-center group-hover:border-emerald-500 transition-colors bg-white">
                        <Check className="w-3 h-3 text-transparent group-active:text-slate-300 transition-colors" />
                      </div>
                      <span className="text-xs font-bold text-slate-600">Remember me</span>
                    </label>
                    <button 
                      type="button" 
                      onClick={() => setAuthMode('forgot')}
                      className="text-xs font-bold text-slate-800 hover:text-emerald-600 transition-colors"
                    >
                      Forgot Password
                    </button>
                  </div>

                  <div className="pt-2">
                    <button 
                      type="submit"
                      className="w-full bg-emerald-500 text-white font-bold py-4 rounded-full hover:bg-emerald-600 transition-all active:scale-[0.98] text-sm shadow-lg shadow-emerald-500/30"
                    >
                      Login
                    </button>
                  </div>

                  <div className="text-center pt-3 pb-8 border-b border-slate-100">
                    <p className="text-xs font-bold text-slate-500">
                      Don't have an account? <button type="button" onClick={() => setAuthMode('signup')} className="text-emerald-500 hover:text-emerald-600 ml-1 transition-colors">Sign Up here</button>
                    </p>
                  </div>

                  <div className="text-center pt-2">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-6">Or Continue With Account</p>
                    <div className="flex items-center justify-center gap-4">
                      {/* Placeholder Apple */}
                      <button type="button" className="w-12 h-12 rounded-full bg-white border border-slate-100 shadow-sm flex items-center justify-center hover:bg-slate-50 transition-colors active:scale-95 text-slate-800">
                         <svg className="w-5 h-5 fill-current" viewBox="0 0 24 24">
                           <path d="M17.05 20.28c-.98.95-2.05.8-3.08.35-1.09-.46-2.09-.48-3.24 0-1.44.62-2.2.44-3.06-.35C2.79 15.25 3.51 7.59 9.05 7.31c1.35.07 2.29.74 3.08.8 1.18-.19 2.31-.88 3.5-.8 1.49.09 2.59.57 3.25 1.5-3.02 1.74-2.52 5.58.37 6.78-.66 1.83-1.6 3.8-2.2 4.69zM12.03 7.35c-.15-2.2 1.63-4.04 3.65-4.35.34 2.31-1.76 4.19-3.65 4.35z"/>
                         </svg>
                      </button>
                      {/* Google */}
                      <button type="button" onClick={() => signInWithPopup(auth, googleProvider)} className="w-12 h-12 rounded-full bg-white border border-slate-100 shadow-sm flex items-center justify-center hover:bg-slate-50 transition-colors active:scale-95">
                         <svg className="w-5 h-5" viewBox="0 0 24 24">
                          <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4"/>
                          <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853"/>
                          <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" fill="#FBBC05"/>
                          <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335"/>
                        </svg>
                      </button>
                    </div>
                  </div>
                </form>
              </div>
            )}

            {/* Signup Form */}
            {authMode === 'signup' && (
              <div className="w-full max-w-[360px] mx-auto transition-all animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="text-center mb-8">
                  <h1 className="text-3xl font-black text-slate-900 tracking-tight mb-3">Create Account</h1>
                  <p className="text-slate-500 text-xs sm:text-sm font-medium leading-relaxed px-4">
                    Create a new account to get started and enjoy seamless access to our features.
                  </p>
                </div>

                <form className="space-y-4" onSubmit={(e) => { e.preventDefault(); signInWithPopup(auth, googleProvider); }}>
                  <div className="space-y-4">
                    <div className="relative">
                      <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                        <User className="w-5 h-5" />
                      </div>
                      <input 
                        type="text" 
                        placeholder="Name" 
                        className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-5 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                      />
                    </div>
                    <div className="relative">
                      <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                        <Mail className="w-5 h-5" />
                      </div>
                      <input 
                        type="email" 
                        placeholder="Email address" 
                        className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-5 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                      />
                    </div>
                    <div className="relative">
                      <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                        <Lock className="w-5 h-5" />
                      </div>
                      <input 
                        type="password" 
                        placeholder="Password" 
                        className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-12 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                      />
                      <button type="button" className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors">
                        <EyeOff className="w-5 h-5" />
                      </button>
                    </div>
                    <div className="relative">
                      <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                        <Lock className="w-5 h-5" />
                      </div>
                      <input 
                        type="password" 
                        placeholder="Confirm Password" 
                        className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-12 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                      />
                      <button type="button" className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors">
                        <EyeOff className="w-5 h-5" />
                      </button>
                    </div>
                  </div>

                  <div className="pt-4">
                    <button 
                      type="submit"
                      className="w-full bg-emerald-500 text-white font-bold py-4 rounded-full hover:bg-emerald-600 transition-all active:scale-[0.98] text-sm shadow-lg shadow-emerald-500/30"
                    >
                      Create Account
                    </button>
                  </div>

                  <div className="text-center pt-3 pb-8 border-b border-slate-100">
                    <p className="text-xs font-bold text-slate-500">
                      Already have an account? <button type="button" onClick={() => setAuthMode('login')} className="text-emerald-500 hover:text-emerald-600 ml-1 transition-colors">Sign In here</button>
                    </p>
                  </div>

                  <div className="text-center pt-2">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-6">Or Continue With Account</p>
                    <div className="flex items-center justify-center gap-4">
                      {/* Apple */}
                      <button type="button" className="w-12 h-12 rounded-full bg-white border border-slate-100 shadow-sm flex items-center justify-center hover:bg-slate-50 transition-colors active:scale-95 text-slate-800">
                         <svg className="w-5 h-5 fill-current" viewBox="0 0 24 24">
                           <path d="M17.05 20.28c-.98.95-2.05.8-3.08.35-1.09-.46-2.09-.48-3.24 0-1.44.62-2.2.44-3.06-.35C2.79 15.25 3.51 7.59 9.05 7.31c1.35.07 2.29.74 3.08.8 1.18-.19 2.31-.88 3.5-.8 1.49.09 2.59.57 3.25 1.5-3.02 1.74-2.52 5.58.37 6.78-.66 1.83-1.6 3.8-2.2 4.69zM12.03 7.35c-.15-2.2 1.63-4.04 3.65-4.35.34 2.31-1.76 4.19-3.65 4.35z"/>
                         </svg>
                      </button>
                      {/* Google */}
                      <button type="button" onClick={() => signInWithPopup(auth, googleProvider)} className="w-12 h-12 rounded-full bg-white border border-slate-100 shadow-sm flex items-center justify-center hover:bg-slate-50 transition-colors active:scale-95">
                         <svg className="w-5 h-5" viewBox="0 0 24 24">
                          <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4"/>
                          <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853"/>
                          <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" fill="#FBBC05"/>
                          <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335"/>
                        </svg>
                      </button>
                    </div>
                  </div>
                </form>
              </div>
            )}

            {/* Forgot Password */}
            {authMode === 'forgot' && (
              <div className="w-full max-w-[360px] mx-auto transition-all animate-in fade-in slide-in-from-bottom-4 duration-500">
                <div className="text-center mb-8">
                  <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight mb-3">Forgot Password</h1>
                  <p className="text-slate-500 text-xs sm:text-sm font-medium leading-relaxed px-4">
                    Enter your email address to receive a reset link and regain access to your account.
                  </p>
                </div>

                <form className="space-y-8" onSubmit={(e) => { e.preventDefault(); }}>
                  <div className="relative">
                    <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                      <Mail className="w-5 h-5" />
                    </div>
                    <input 
                      type="email" 
                      placeholder="Email address" 
                      className="w-full bg-slate-50 shadow-sm border border-slate-100/60 rounded-full py-4 pl-12 pr-5 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400 font-medium text-slate-700"
                    />
                  </div>

                  <button 
                    type="submit"
                    className="w-full bg-emerald-500 text-white font-bold py-4 rounded-full hover:bg-emerald-600 transition-all active:scale-[0.98] text-sm shadow-lg shadow-emerald-500/30"
                  >
                    Continue
                  </button>
                </form>
              </div>
            )}
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex w-full overflow-x-hidden">
      <Sidebar />
      
      <main className="flex-1 md:ml-72 p-4 sm:p-8 md:p-12 min-h-screen max-w-[100vw] overflow-x-hidden md:max-w-none md:overflow-x-visible">
        <header className="mb-10 md:mb-16 flex flex-col md:flex-row md:items-center justify-between gap-8">
          <div className="flex items-center gap-4">
            <div className="md:hidden w-12 h-12 bg-blue-600 rounded-2xl flex items-center justify-center text-white shadow-xl shadow-blue-200">
              <CreditCard className="w-6 h-6 stroke-[2.5]" />
            </div>
            <div>
              <h2 className="text-3xl md:text-4xl font-display font-black text-slate-900 tracking-tight">
                {activeTab === 'dashboard' && 'Ringkasan Bisnis'}
                {activeTab === 'pos' && 'Kasir Pintar'}
                {activeTab === 'inventory' && 'Stok Barang'}
                {activeTab === 'debts' && 'Manajemen Hutang'}
                {activeTab === 'finance' && 'Dashboard Dana'}
              </h2>
              <p className="text-slate-400 text-sm md:text-base mt-2 font-bold uppercase tracking-widest flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                Live Status • {new Date().toLocaleDateString('id-ID', { weekday: 'long' })}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-4">
            {lowStockItems.length > 0 && (
              <motion.div 
                initial={{ scale: 0.9, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                className="flex items-center gap-3 bg-rose-50 text-rose-600 px-6 py-3 rounded-[1.5rem] border border-rose-100 text-xs font-black shadow-lg shadow-rose-200/20"
              >
                <div className="relative">
                  <AlertCircle className="w-5 h-5 stroke-[3]" />
                  <span className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-rose-500 rounded-full animate-ping" />
                </div>
                {lowStockItems.length} STOK KRITIS
              </motion.div>
            )}
            <div className="bg-white px-6 py-3 rounded-[1.5rem] text-[10px] text-slate-500 font-black uppercase tracking-[0.2em] hidden md:flex items-center gap-3 shadow-sm border border-slate-100">
              {new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
            </div>
            
            <button 
              onClick={() => signOut(auth)}
              className="md:hidden w-12 h-12 rounded-[1.2rem] bg-white border border-slate-200 flex items-center justify-center text-rose-500 shadow-sm shrink-0"
              title="Keluar"
            >
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"></path><polyline points="16 17 21 12 16 7"></polyline><line x1="21" y1="12" x2="9" y2="12"></line></svg>
            </button>
          </div>
        </header>

        <AnimatePresence mode="wait">
          {activeTab === 'dashboard' && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, scale: 0.98, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.98, y: -10 }}
              transition={{ duration: 0.4, ease: [0.23, 1, 0.32, 1] }}
              className="space-y-10 pb-24 md:pb-0"
            >
              {/* Stats Grid */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-200">
                  <div className="flex items-center justify-between mb-4">
                    <div className="p-3 bg-emerald-100 text-emerald-600 rounded-2xl">
                      <TrendingUp className="w-6 h-6" />
                    </div>
                    <span className="text-xs font-bold text-emerald-600 bg-emerald-50 px-2 py-1 rounded-lg">+12.5%</span>
                  </div>
                  <p className="text-slate-500 text-sm font-medium">Total Pemasukan</p>
                  <h3 className="text-2xl font-bold text-slate-800 mt-1">Rp {totalIncome.toLocaleString()}</h3>
                </div>

                <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-200">
                  <div className="flex items-center justify-between mb-4">
                    <div className="p-3 bg-rose-100 text-rose-600 rounded-2xl">
                      <TrendingDown className="w-6 h-6" />
                    </div>
                    <span className="text-xs font-bold text-rose-600 bg-rose-50 px-2 py-1 rounded-lg">-2.4%</span>
                  </div>
                  <p className="text-slate-500 text-sm font-medium">Total Pengeluaran</p>
                  <h3 className="text-2xl font-bold text-slate-800 mt-1">Rp {totalExpense.toLocaleString()}</h3>
                </div>

                <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-200">
                  <div className="flex items-center justify-between mb-4">
                    <div className="p-3 bg-blue-100 text-blue-600 rounded-2xl">
                      <Wallet className="w-6 h-6" />
                    </div>
                  </div>
                  <p className="text-slate-500 text-sm font-medium">Saldo Bersih</p>
                  <h3 className="text-2xl font-bold text-slate-800 mt-1">Rp {balance.toLocaleString()}</h3>
                </div>

                <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-200">
                  <div className="flex items-center justify-between mb-4">
                    <div className="p-3 bg-rose-50 text-rose-500 rounded-2xl">
                      <BookOpen className="w-6 h-6" />
                    </div>
                  </div>
                  <p className="text-slate-500 text-sm font-medium">Total Piutang (Hutang)</p>
                  <h3 className="text-2xl font-bold text-rose-600 mt-1">Rp {totalPiutang.toLocaleString()}</h3>
                </div>
              </div>

                {/* Charts & Table */}
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 md:gap-8">
                  {/* Financial Overview Line Chart */}
                <div className="bg-white p-7 md:p-10 rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.03)] border border-slate-100/80">
                  <h4 className="font-display font-extrabold text-slate-900 mb-8 flex items-center gap-3">
                    <div className="w-1 h-6 bg-blue-600 rounded-full" />
                    Tren Finansial
                  </h4>
                  <div className="h-64 sm:h-72 w-full">
                    {monthlyChartData.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={monthlyChartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis 
                            dataKey="month" 
                            axisLine={false} 
                            tickLine={false} 
                            tick={{ fill: '#64748b', fontSize: 12, fontWeight: 500 }} 
                            dy={10}
                          />
                          <YAxis 
                            axisLine={false} 
                            tickLine={false} 
                            tick={{ fill: '#64748b', fontSize: 10 }}
                            tickFormatter={(value) => `Rp ${(value / 1000)}k`}
                          />
                          <Tooltip 
                            contentStyle={{ borderRadius: '1rem', border: 'none', boxShadow: '0 20px 50px rgba(0,0,0,0.1)' }}
                            formatter={(value: number) => [`Rp ${value.toLocaleString()}`, '']}
                            labelStyle={{ fontWeight: 'bold', color: '#0f172a', marginBottom: '0.5rem' }}
                          />
                          <Legend iconType="circle" wrapperStyle={{ fontSize: '12px', fontWeight: 'bold', paddingTop: '1rem' }} />
                          <Line 
                            type="monotone" 
                            dataKey="Pemasukan" 
                            stroke="#10b981" 
                            strokeWidth={3} 
                            dot={{ r: 4, strokeWidth: 2 }} 
                            activeDot={{ r: 6, stroke: '#10b981', strokeWidth: 2, fill: '#fff' }} 
                            animationDuration={1500} 
                          />
                          <Line 
                            type="monotone" 
                            dataKey="Pengeluaran" 
                            stroke="#f43f5e" 
                            strokeWidth={3} 
                            dot={{ r: 4, strokeWidth: 2 }} 
                            activeDot={{ r: 6, stroke: '#f43f5e', strokeWidth: 2, fill: '#fff' }} 
                            animationDuration={1500} 
                          />
                        </LineChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="w-full h-full flex items-center justify-center text-slate-400 font-medium text-sm">
                        Belum ada data finansial
                      </div>
                    )}
                  </div>
                </div>

                {/* Stock by Category Chart */}
                <div className="bg-white p-7 md:p-10 rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.03)] border border-slate-100/80">
                  <h4 className="font-display font-extrabold text-slate-900 mb-8 flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className="w-1 h-6 bg-blue-600 rounded-full" />
                      Produk per Kategori
                    </div>
                    <span className="text-[10px] font-black bg-blue-50 text-blue-600 px-3 py-1.5 rounded-xl uppercase tracking-wider border border-blue-100">Statistik</span>
                  </h4>
                  <div className="space-y-3 overflow-y-auto max-h-[300px] pr-2 custom-scrollbar">
                    {dynamicCategories.map(cat => {
                      const count = products.filter(p => p.kategori === cat).length;
                      const percentage = (count / products.length) * 100;
                      if (count === 0) return null;
                      return (
                        <div key={cat} className="space-y-1">
                          <div className="flex justify-between text-xs font-bold">
                            <span className="text-slate-600">{cat}</span>
                            <span className="text-blue-600">{count} Item</span>
                          </div>
                          <div className="h-2 w-full bg-slate-100 rounded-full overflow-hidden">
                            <motion.div 
                              initial={{ width: 0 }}
                              animate={{ width: `${percentage}%` }}
                              className="h-full bg-blue-500 rounded-full shadow-sm"
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Low Stock Alerts */}
              <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-200 mt-8 mb-8">
                <div className="flex items-center justify-between mb-6">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-amber-100 text-amber-600 rounded-lg">
                      <AlertCircle className="w-5 h-5" />
                    </div>
                    <h4 className="font-bold text-slate-800">Stok Hampir Habis</h4>
                  </div>
                  <button 
                    onClick={() => setActiveTab('inventory')}
                    className="text-xs font-bold text-blue-600 hover:text-blue-700 transition-colors"
                  >
                    Atur Semua Stok
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {lowStockItems.length === 0 ? (
                    <div className="col-span-full py-10 flex flex-col items-center justify-center text-center text-slate-400 bg-slate-50/50 rounded-2xl border-2 border-dashed border-slate-100">
                      <Package className="w-10 h-10 mb-2 opacity-20" />
                      <p className="text-sm">Semua stok barang mencukupi.</p>
                    </div>
                  ) : (
                    lowStockItems.map(item => (
                      <div key={item.id} className="flex items-center justify-between p-4 bg-amber-50/50 rounded-2xl border border-amber-100">
                        <div className="flex items-center gap-3">
                          <img src={item.url_gambar} className="w-12 h-12 rounded-xl object-cover" />
                          <div>
                            <p className="font-bold text-slate-700 text-sm">{item.nama}</p>
                            <p className="text-[10px] text-amber-500 font-black">STOK: {item.stok}</p>
                          </div>
                        </div>
                        <ChevronRight className="w-4 h-4 text-amber-600" />
                      </div>
                    ))
                  )}
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'pos' && (
            <motion.div 
              key="pos"
              initial={{ opacity: 0, scale: 0.98, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.98, y: -10 }}
              transition={{ duration: 0.4, ease: [0.23, 1, 0.32, 1] }}
              className="flex flex-col xl:flex-row gap-10 pb-32 md:pb-0"
            >
              {/* Product Grid - Visual Payment */}
              <div className="flex-1 space-y-6">
                <div className="space-y-4">
                  <div className="flex gap-4 items-center">
                    <div className="relative flex-1 group">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400 group-focus-within:text-blue-500 transition-colors" />
                      <input 
                        type="text" 
                        placeholder="Cari bumbu, mie, atau jajan..." 
                        className="w-full bg-white border border-slate-200 rounded-2xl py-3.5 pl-12 pr-4 outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all shadow-sm text-slate-700"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                  </div>

                  <div className="relative">
                    <select 
                      className="w-full appearance-none bg-white border border-slate-200 text-slate-700 py-3 pl-4 pr-10 rounded-2xl font-bold text-sm focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all shadow-sm"
                      value={selectedCategoryFilter}
                      onChange={(e) => setSelectedCategoryFilter(e.target.value)}
                    >
                      <option value="Semua">Semua Barang</option>
                      {dynamicCategories.map(cat => (
                        <option key={cat} value={cat}>{cat}</option>
                      ))}
                    </select>
                    <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400 pointer-events-none" />
                  </div>
                </div>

                  <div className="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-4 md:gap-8">
                  <AnimatePresence mode="popLayout">
                    {products
                      .filter(p => p.nama.toLowerCase().includes(searchQuery.toLowerCase()))
                      .filter(p => selectedCategoryFilter === 'Semua' || p.kategori === selectedCategoryFilter)
                      .map(product => (
                        <motion.div 
                          layout
                          initial={{ opacity: 0, y: 30 }}
                          animate={{ opacity: 1, y: 0 }}
                          exit={{ opacity: 0, scale: 0.9 }}
                          whileHover={{ y: -10 }}
                          whileTap={{ scale: 0.96 }}
                          key={product.id}
                          onClick={() => addToCart(product)}
                          className={`group bg-white rounded-[2rem] p-3 md:p-4 border border-slate-100/80 cursor-pointer transition-all duration-500 hover:shadow-2xl hover:shadow-blue-600/10 relative overflow-hidden flex flex-col ${product.stok <= 0 ? 'opacity-60 grayscale' : ''}`}
                        >
                          <div className="aspect-square rounded-2xl overflow-hidden relative shadow-inner">
                            <img 
                              src={product.url_gambar} 
                              alt={product.nama} 
                              className="w-full h-full object-cover transition-transform duration-1000 group-hover:scale-110"
                            />
                            {product.stok < 5 && product.stok > 0 && (
                              <div className="absolute top-3 left-3 glass text-orange-600 text-[9px] font-black px-2.5 py-1.5 rounded-xl uppercase shadow-md border border-orange-100/50">Stok Tipis</div>
                            )}
                            {product.stok <= 0 && (
                              <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-[4px] flex items-center justify-center">
                                <span className="bg-white/90 text-slate-900 px-4 py-1.5 rounded-2xl text-[10px] font-black uppercase tracking-widest shadow-xl">Out of Stock</span>
                              </div>
                            )}
                            <div className="absolute inset-0 ring-1 ring-inset ring-black/5 rounded-2xl" />
                            
                            {/* Hover Overlay */}
                            <div className="absolute inset-0 bg-blue-600/0 group-hover:bg-blue-600/10 transition-colors duration-500" />
                            
                            {/* Tap Indicator */}
                            <div className="absolute bottom-3 right-3 opacity-0 translate-y-4 group-hover:opacity-100 group-hover:translate-y-0 transition-all duration-500 ease-out">
                              <div className="p-3 bg-blue-600 text-white rounded-2xl shadow-xl shadow-blue-400/40">
                                <Plus className="w-5 h-5 stroke-[2.5]" />
                              </div>
                            </div>
                          </div>
                          
                          <div className="mt-4 px-1 flex-1 flex flex-col">
                            <div className="flex items-center gap-2 mb-1.5">
                              <span className="w-1.5 h-1.5 rounded-full bg-blue-500" />
                              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{product.kategori}</p>
                            </div>
                            <h5 className="font-display font-bold text-slate-800 text-sm md:text-base truncate leading-snug">{product.nama}</h5>
                            
                            <div className="mt-auto pt-4 flex items-center justify-between border-t border-slate-50/80">
                              <p className="text-slate-900 font-extrabold text-base md:text-lg tracking-tight">Rp {product.harga.toLocaleString()}</p>
                              <div className={`px-2.5 py-1 rounded-xl text-[10px] font-black shadow-sm ${product.stok < 5 ? 'bg-orange-50 text-orange-600' : 'bg-slate-50 text-slate-500'}`}>
                                {product.stok}
                              </div>
                            </div>
                          </div>
                        </motion.div>
                      ))}
                  </AnimatePresence>
                </div>
              </div>

              {/* Cart Section - Desktop Sidebar */}
              <div className="hidden xl:flex w-[420px] bg-white rounded-[3rem] border border-slate-200/60 shadow-[0_20px_50px_rgba(0,0,0,0.02)] flex-col overflow-hidden max-h-[calc(100vh-140px)] sticky top-28 group">
                <div className="p-8 border-b border-slate-100/60 flex items-center justify-between bg-slate-50/30 shrink-0">
                  <div>
                    <h4 className="font-display font-black text-slate-900 flex items-center gap-3 text-lg">
                      <div className="p-2 bg-blue-600 text-white rounded-xl shadow-lg shadow-blue-200/50">
                        <ShoppingCart className="w-5 h-5" />
                      </div>
                      Keranjang
                    </h4>
                    <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest mt-1">{cart.length} Pesanan</p>
                  </div>
                  <button onClick={clearCart} className="p-3 bg-white text-rose-500 hover:bg-rose-50 rounded-xl transition-all border border-slate-100 shadow-sm active:scale-95">
                    <Trash2 className="w-5 h-5" />
                  </button>
                </div>

                <div className="flex-1 overflow-y-auto p-6 space-y-4 no-scrollbar scroll-smooth">
                  <AnimatePresence mode="popLayout" initial={false}>
                    {cart.length === 0 ? (
                      <motion.div 
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        className="h-full flex flex-col items-center justify-center text-center text-slate-300 py-16"
                      >
                        <div className="w-24 h-24 bg-slate-50/80 rounded-full flex items-center justify-center mb-6 ring-8 ring-slate-50/30 relative">
                          <PackageSearch className="w-12 h-12 opacity-10" />
                          <div className="absolute inset-0 bg-gradient-to-tr from-slate-100/50 to-transparent blur-xl" />
                        </div>
                        <p className="text-sm font-bold text-slate-400">Keranjang Kosong</p>
                        <p className="text-[10px] text-slate-300 mt-2 font-medium px-4">Pilih barang di sebelah kiri untuk memulai transaksi.</p>
                      </motion.div>
                    ) : (
                      cart.map(item => (
                        <motion.div 
                          layout
                          initial={{ x: 20, opacity: 0 }}
                          animate={{ x: 0, opacity: 1 }}
                          exit={{ x: -20, opacity: 0 }}
                          key={item.id} 
                          className="flex flex-col bg-white p-4 rounded-2xl border border-slate-100/80 gap-3 group/item hover:border-blue-200 hover:shadow-lg hover:shadow-blue-600/5 transition-all duration-300"
                        >
                          <div className="flex items-center justify-between">
                            <div className="flex-1 min-w-0 pr-4">
                              <p className="font-bold text-slate-800 text-sm truncate tracking-tight">{item.nama}</p>
                              <p className="text-xs text-blue-600 font-black mt-0.5">Rp {item.harga.toLocaleString()}</p>
                            </div>
                            <div className="text-right">
                              <p className="font-extrabold text-slate-900 text-sm">Rp {(item.harga * item.jumlah).toLocaleString()}</p>
                            </div>
                          </div>
                          <div className="flex items-center justify-between pt-3 border-t border-slate-50">
                            <div className="flex items-center bg-slate-50 rounded-xl border border-slate-100 p-1">
                              <button 
                                onClick={() => updateCartQuantity(item.id, -1)}
                                className="p-1.5 hover:bg-white hover:shadow-sm rounded-lg text-slate-400 transition-all active:scale-90"
                              >
                                <Minus className="w-3.5 h-3.5" />
                              </button>
                              <span className="w-10 text-center text-xs font-black text-slate-800">{item.jumlah}</span>
                              <button 
                                onClick={() => updateCartQuantity(item.id, 1)}
                                className="p-1.5 hover:bg-white hover:shadow-sm rounded-lg text-blue-600 transition-all active:scale-90"
                              >
                                <Plus className="w-3.5 h-3.5" />
                              </button>
                            </div>
                            <button 
                              onClick={() => removeFromCart(item.id)}
                              className="w-8 h-8 flex items-center justify-center bg-rose-50 text-rose-500 rounded-lg opacity-0 group-hover/item:opacity-100 transition-all hover:bg-rose-100 active:scale-90"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </motion.div>
                      ))
                    )}
                  </AnimatePresence>
                </div>

                <div className="p-8 bg-slate-50/80 backdrop-blur-md border-t border-slate-200/60 space-y-6 shrink-0 shadow-[0_-10px_30px_rgba(0,0,0,0.02)]">
                  <div className="space-y-4">
                    <div className="flex items-center justify-between">
                      <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-1">Metode Bayar</p>
                      <span className={`text-[10px] font-black px-2 py-0.5 rounded-lg border ${selectedMethod === 'Hutang' ? 'bg-rose-50 text-rose-600 border-rose-100' : 'bg-blue-50 text-blue-600 border-blue-100'}`}>{selectedMethod}</span>
                    </div>
                    <div className="flex flex-wrap gap-2 pb-1">
                      {['Tunai', 'QRIS', 'M-Banking', 'DANA', 'OVO', 'Hutang'].map(m => (
                        <button
                          key={m}
                          onClick={() => setSelectedMethod(m as any)}
                          className={`py-2.5 px-4 rounded-xl text-[10px] font-black transition-all border shrink-0 ${
                            selectedMethod === m 
                              ? 'bg-blue-600 border-blue-600 text-white shadow-lg shadow-blue-200/30' 
                              : 'bg-white border-slate-200 text-slate-500 hover:border-blue-300'
                          }`}
                        >
                          {m}
                        </button>
                      ))}
                    </div>
                  </div>
                  
                  {selectedMethod === 'Hutang' && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="space-y-3"
                    >
                      <label className="text-[10px] font-black text-rose-500 uppercase tracking-widest pl-1">Nama Penghutang *</label>
                      <input 
                        type="text" 
                        placeholder="Contoh: Bpk. Slamet..." 
                        value={debtorName}
                        onChange={(e) => setDebtorName(e.target.value)}
                        className="w-full bg-white border border-slate-200 rounded-2xl py-3.5 px-5 text-sm font-bold outline-none focus:ring-4 focus:ring-rose-500/10 focus:border-rose-500 transition-all shadow-sm"
                      />
                    </motion.div>
                  )}

                  <div className="flex items-center justify-between py-1 px-1">
                    <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Total Bayar</p>
                    <motion.p 
                      key={totalBelanja}
                      initial={{ scale: 1.1 }}
                      animate={{ scale: 1 }}
                      className="text-3xl font-display font-black text-slate-900 tracking-tighter"
                    >
                      Rp {totalBelanja.toLocaleString()}
                    </motion.p>
                  </div>

                  <button 
                    disabled={cart.length === 0}
                    onClick={initiatePayment}
                    className={`w-full text-white rounded-2xl py-5 font-black uppercase text-xs tracking-[0.2em] shadow-xl transition-all flex items-center justify-center gap-2 group/pay transform active:scale-95 ${selectedMethod === 'Hutang' ? 'bg-rose-500 hover:bg-rose-600 shadow-rose-200' : 'bg-blue-600 hover:bg-blue-700 shadow-blue-200'} disabled:opacity-30 disabled:shadow-none`}
                  >
                    {selectedMethod === 'Hutang' ? 'Catat Hutang' : 'Proses Pembayaran'}
                    <ChevronRight className="w-4 h-4 group-hover/pay:translate-x-1.5 transition-transform" />
                  </button>
                </div>
              </div>
            </motion.div>
          )}
          
          {activeTab === 'inventory' && (
            <motion.div 
              key="inventory"
              initial={{ opacity: 0, scale: 0.98, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.98, y: -10 }}
              transition={{ duration: 0.4, ease: [0.23, 1, 0.32, 1] }}
              className="space-y-8 pb-24 md:pb-0"
            >
              <div className="bg-white rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.02)] border border-slate-100 overflow-hidden">
                <div className="p-6 md:p-10 border-b border-slate-50 space-y-6">
                  <div className="flex flex-row items-center justify-between gap-3 md:gap-4">
                    <div className="relative flex-1 group">
                      <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 md:w-5 md:h-5 text-slate-300 group-focus-within:text-blue-500 transition-colors" />
                      <input 
                        type="text" 
                        placeholder="Cari ID atau nama..." 
                        className="w-full bg-slate-50 border border-slate-200/60 rounded-xl md:rounded-2xl py-3.5 md:py-4 pl-10 md:pl-14 pr-4 outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all text-xs md:text-sm font-medium"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                      />
                    </div>
                    <div className="flex shrink-0 gap-2">
                      <button 
                        onClick={() => { setEditingProduct(null); setFormImage(''); setShowProductModal(true); }}
                        className="bg-blue-600 text-white shrink-0 w-12 h-12 md:w-auto md:h-auto md:px-10 md:py-4 rounded-xl md:rounded-2xl font-bold flex items-center justify-center gap-2 hover:bg-blue-700 transition-all text-xs md:text-sm shadow-xl shadow-blue-200/50 active:scale-95"
                      >
                        <Plus className="w-5 h-5 md:w-5 md:h-5 stroke-[3]" />
                        <span className="hidden md:inline">Tambah Produk</span>
                      </button>
                    </div>
                  </div>

                  <div className="relative">
                    <select 
                      className="w-full appearance-none bg-slate-50 border border-slate-200/60 text-slate-700 py-3.5 pl-4 pr-10 rounded-xl font-bold text-sm focus:outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all shadow-sm"
                      value={selectedCategoryFilter}
                      onChange={(e) => setSelectedCategoryFilter(e.target.value)}
                    >
                      <option value="Semua">Semua Produk</option>
                      {dynamicCategories.map(cat => (
                        <option key={cat} value={cat}>{cat}</option>
                      ))}
                    </select>
                    <ChevronDown className="absolute right-4 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400 pointer-events-none" />
                  </div>
                </div>

                <div className="hidden md:block overflow-x-auto no-scrollbar">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50/40 text-slate-400 text-[10px] uppercase tracking-[0.2em] font-black italic">
                        <th className="px-10 py-6">Detail Produk</th>
                        <th className="px-10 py-6">Kategori</th>
                        <th className="px-10 py-6">SKU ID</th>
                        <th className="px-10 py-6 text-right">Harga Satuan</th>
                        <th className="px-10 py-6 text-center">Status Stok</th>
                        <th className="px-10 py-6 text-center">Tindakan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {products
                        .filter(p => 
                          p.nama.toLowerCase().includes(searchQuery.toLowerCase()) || 
                          p.id.toLowerCase().includes(searchQuery.toLowerCase())
                        )
                        .filter(p => selectedCategoryFilter === 'Semua' || p.kategori === selectedCategoryFilter)
                        .map((product) => (
                          <tr key={product.id} className="hover:bg-blue-50/20 transition-all group">
                            <td className="px-10 py-5">
                              <div className="flex items-center gap-5">
                                <div className="relative shrink-0">
                                  <img src={product.url_gambar} className="w-14 h-14 rounded-2xl object-cover border border-slate-100 group-hover:scale-105 transition-transform duration-500 shadow-sm" alt={product.nama} />
                                  <div className="absolute inset-0 ring-1 ring-inset ring-black/5 rounded-2xl" />
                                </div>
                                <div>
                                  <span className="font-bold text-slate-800 block leading-tight">{product.nama}</span>
                                  <span className="text-[10px] text-slate-400 mt-1 block">In Stock Ready</span>
                                </div>
                              </div>
                            </td>
                            <td className="px-10 py-5">
                              <span className="text-[10px] font-black text-blue-500 bg-blue-50/50 px-3 py-1.5 rounded-lg uppercase tracking-tight border border-blue-100/50">{product.kategori}</span>
                            </td>
                            <td className="px-10 py-5 text-sm font-mono font-bold text-slate-300">#{product.id}</td>
                            <td className="px-10 py-5 text-right">
                              <span className="text-base font-extrabold text-slate-800 italic">Rp {product.harga.toLocaleString()}</span>
                            </td>
                            <td className="px-10 py-5 text-center">
                              <div className="flex flex-col items-center gap-1">
                                <div className="h-1.5 w-20 bg-slate-100 rounded-full overflow-hidden">
                                  <div 
                                    className={`h-full rounded-full ${product.stok < 5 ? 'bg-orange-500' : 'bg-emerald-500'}`}
                                    style={{ width: `${Math.min(100, (product.stok / 50) * 100)}%` }}
                                  />
                                </div>
                                <span className={`text-[10px] font-black uppercase ${product.stok < 5 ? 'text-orange-600' : 'text-emerald-600'}`}>
                                  {product.stok} Unit
                                </span>
                              </div>
                            </td>
                            <td className="px-10 py-5">
                              <div className="flex items-center justify-center gap-3">
                                <button 
                                  onClick={() => { setEditingProduct(product); setFormImage(product.url_gambar); setShowProductModal(true); }}
                                  className="p-3 text-slate-400 hover:text-blue-600 hover:bg-white hover:shadow-lg rounded-xl transition-all active:scale-90 border border-transparent hover:border-blue-100"
                                >
                                  <Edit className="w-5 h-5" />
                                </button>
                                <button 
                                  onClick={() => setProductToDelete(product)}
                                  className="p-3 text-slate-400 hover:text-rose-600 hover:bg-white hover:shadow-lg rounded-xl transition-all active:scale-90 border border-transparent hover:border-rose-100"
                                >
                                  <Trash2 className="w-5 h-5" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>

                {/* Card View - Mobile Only */}
                <div className="md:hidden divide-y divide-slate-50 pb-20">
                  {products
                    .filter(p => 
                      p.nama.toLowerCase().includes(searchQuery.toLowerCase()) || 
                      p.id.toLowerCase().includes(searchQuery.toLowerCase())
                    )
                    .filter(p => selectedCategoryFilter === 'Semua' || p.kategori === selectedCategoryFilter)
                    .map((product) => (
                      <div key={product.id} className="p-6 space-y-4">
                        <div className="flex items-center gap-4">
                          <img src={product.url_gambar} className="w-16 h-16 rounded-2xl object-cover border border-slate-100 shadow-sm" alt={product.nama} />
                          <div className="flex-1 min-w-0">
                            <div className="flex justify-between items-start">
                              <h5 className="font-bold text-slate-800 text-sm md:text-base truncate leading-tight">{product.nama}</h5>
                              <p className="text-[10px] font-mono font-bold text-slate-300 ml-2">#{product.id}</p>
                            </div>
                            <div className="flex items-center justify-between mt-1">
                              <p className="text-sm font-extrabold text-blue-600">Rp {product.harga.toLocaleString()}</p>
                              <span className={`text-[9px] font-black px-2 py-1 rounded-lg uppercase tracking-widest ${product.stok < 5 ? 'bg-orange-50 text-orange-600' : 'bg-emerald-50 text-emerald-600'}`}>
                                {product.stok} Unit
                              </span>
                            </div>
                          </div>
                        </div>
                        <div className="flex gap-3">
                          <button 
                            onClick={() => { setEditingProduct(product); setFormImage(product.url_gambar); setShowProductModal(true); }}
                            className="flex-1 bg-white border border-slate-100 shadow-sm text-slate-600 py-3 rounded-xl text-xs font-bold flex items-center justify-center gap-2 active:bg-slate-50 active:scale-95 transition-all"
                          >
                            <Edit className="w-4 h-4" />
                            Ubah
                          </button>
                          <button 
                            onClick={() => setProductToDelete(product)}
                            className="flex-1 bg-white border border-slate-100 shadow-sm text-rose-500 py-3 rounded-xl text-xs font-bold flex items-center justify-center gap-2 active:bg-rose-50 active:scale-95 transition-all"
                          >
                            <Trash2 className="w-4 h-4" />
                            Hapus
                          </button>
                        </div>
                      </div>
                    ))}
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'finance' && (
            <motion.div 
              key="finance"
              initial={{ opacity: 0, scale: 0.98, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.98, y: -10 }}
              transition={{ duration: 0.4, ease: [0.23, 1, 0.32, 1] }}
              className="space-y-12 pb-24 md:pb-0"
            >
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-10">
                {/* Input Finansial */}
                <div className="lg:col-span-1">
                  <div className="bg-white p-10 rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.02)] border border-slate-100 sticky top-28">
                    <h4 className="font-display font-extrabold text-slate-900 mb-8 flex items-center gap-3">
                      <div className="p-2 bg-blue-50 text-blue-600 rounded-xl">
                        <Plus className="w-5 h-5 stroke-[3]" />
                      </div>
                      Tambah Catatan
                    </h4>
                    <form className="space-y-6" onSubmit={(e) => {
                      e.preventDefault();
                      const formData = new FormData(e.currentTarget);
                      const newRecord: FinanceRecord = {
                        id: Math.random().toString(36).substr(2, 9),
                        tipe: formData.get('tipe') as any,
                        jumlah: Number(formData.get('jumlah')),
                        keterangan: formData.get('keterangan') as string,
                        tanggal: new Date().toISOString().split('T')[0]
                      };
                      setFinance(prev => [newRecord, ...prev]);
                      (e.target as any).reset();
                    }}>
                      <div>
                        <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Tipe Transaksi</label>
                        <select name="tipe" className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all">
                          <option value="pemasukan">Dana Masuk (+)</option>
                          <option value="pengeluaran">Dana Keluar (-)</option>
                        </select>
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Jumlah Nominal</label>
                        <div className="relative">
                          <span className="absolute left-6 top-1/2 -translate-y-1/2 font-black text-slate-400">Rp</span>
                          <input name="jumlah" type="number" required placeholder="0" className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 pl-14 pr-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all text-slate-800" />
                        </div>
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Memo / Keterangan</label>
                        <input name="keterangan" type="text" required placeholder="Contoh: Belanja Minyak..." className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all text-slate-800 placeholder:normal-case placeholder:font-medium" />
                      </div>
                      <button className="w-full bg-slate-900 text-white font-black py-5 rounded-[1.5rem] shadow-2xl shadow-slate-200 hover:bg-slate-800 transition-all mt-6 uppercase tracking-[0.2em] text-xs transform active:scale-95">
                        Simpan Record
                      </button>
                    </form>
                  </div>
                </div>

                {/* History Laporan */}
                <div className="lg:col-span-2">
                  <div className="bg-white rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.02)] border border-slate-100 flex flex-col min-h-[600px] overflow-hidden">
                    <div className="p-10 border-b border-slate-50 flex items-center justify-between bg-slate-50/20">
                      <h4 className="font-display font-extrabold text-slate-900 flex items-center gap-3">
                        <div className="p-2 bg-blue-600 text-white rounded-xl shadow-lg shadow-blue-200/50">
                          <History className="w-5 h-5" />
                        </div>
                        Aktivitas Lokal
                      </h4>
                      <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest bg-white px-4 py-2 rounded-xl border border-slate-100">{new Date().toLocaleString('id-ID', { month: 'long', year: 'numeric' })}</p>
                    </div>
                    <div className="flex-1 p-6 md:p-10 space-y-5 no-scrollbar scroll-smooth">
                      {finance.length === 0 ? (
                        <div className="h-full flex flex-col items-center justify-center text-center text-slate-300 py-32">
                          <div className="w-24 h-24 mb-6 rounded-full bg-slate-50/80 flex items-center justify-center ring-8 ring-slate-50/30">
                            <History className="w-12 h-12 opacity-10" />
                          </div>
                          <p className="text-base font-bold text-slate-400">Belum Ada Transaksi</p>
                          <p className="text-xs text-slate-300 mt-2 font-medium">Semua catatan transaksi akan muncul di sini.</p>
                        </div>
                      ) : (
                        finance.map(record => (
                          <div key={record.id} className="flex flex-col sm:flex-row sm:items-center justify-between p-6 rounded-[2rem] bg-white border border-slate-100/80 group gap-5 hover:border-blue-100 hover:shadow-xl hover:shadow-blue-500/5 transition-all duration-300">
                            <div className="flex items-center gap-5">
                              <div className={`w-14 h-14 rounded-2xl flex items-center justify-center shrink-0 shadow-sm ${record.tipe === 'pemasukan' ? 'bg-emerald-50 text-emerald-600 border border-emerald-100/50' : 'bg-rose-50 text-rose-600 border border-rose-100/50'}`}>
                                {record.tipe === 'pemasukan' ? <TrendingUp className="w-6 h-6" /> : <TrendingDown className="w-6 h-6" />}
                              </div>
                              <div className="min-w-0">
                                <p className="font-bold text-slate-800 text-base md:text-lg truncate tracking-tight">{record.keterangan}</p>
                                <div className="flex items-center gap-2 mt-1">
                                  <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{record.tanggal}</span>
                                  <span className="w-1 h-1 bg-slate-200 rounded-full" />
                                  <span className="text-[10px] font-bold text-blue-500 uppercase">ID: #{record.id.toUpperCase()}</span>
                                </div>
                              </div>
                            </div>
                            <div className="flex items-center justify-between sm:justify-end gap-6 border-t sm:border-t-0 pt-4 sm:pt-0">
                              <div className="text-left sm:text-right">
                                <p className={`font-display font-black text-lg md:text-xl italic ${record.tipe === 'pemasukan' ? 'text-emerald-600' : 'text-rose-600'}`}>
                                  {record.tipe === 'pemasukan' ? '+' : '-'} Rp {record.jumlah.toLocaleString()}
                                </p>
                                <span className="text-[9px] font-black text-slate-300 uppercase tracking-[0.2em] block mt-1">{record.tipe}</span>
                              </div>
                              <button 
                                onClick={() => setFinance(prev => prev.filter(f => f.id !== record.id))}
                                className="w-10 h-10 flex items-center justify-center text-rose-400 hover:text-white hover:bg-rose-500 rounded-xl transition-all sm:opacity-0 group-hover:opacity-100 active:scale-90 shadow-sm border border-transparent hover:border-rose-600"
                              >
                                <Trash2 className="w-5 h-5" />
                              </button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === 'debts' && (
            <motion.div 
              key="debts"
              initial={{ opacity: 0, scale: 0.98, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.98, y: -10 }}
              transition={{ duration: 0.4, ease: [0.23, 1, 0.32, 1] }}
              className="space-y-12 pb-24 md:pb-0"
            >
              <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                <div className="lg:col-span-1 bg-white p-8 rounded-3xl shadow-sm border border-slate-200 h-fit">
                  <div className="flex items-center gap-4 mb-6">
                    <div className="p-3 bg-rose-100 text-rose-600 rounded-2xl">
                      <BookOpen className="w-6 h-6" />
                    </div>
                    <div>
                      <p className="text-slate-500 text-[10px] font-black uppercase tracking-widest">Total Hutang Aktif</p>
                      <h3 className="text-2xl font-black text-slate-900">Rp {debts.filter(d => d.status === 'belum_lunas').reduce((sum, d) => sum + d.total, 0).toLocaleString()}</h3>
                    </div>
                  </div>
                  <div className="space-y-4">
                    <div className="flex items-center justify-between p-4 bg-slate-50 rounded-2xl">
                      <span className="text-xs font-bold text-slate-500">Jumlah Orang</span>
                      <span className="text-sm font-black text-slate-900">{new Set(debts.filter(d => d.status === 'belum_lunas').map(d => d.pemberiHutang)).size} Orang</span>
                    </div>
                    <div className="flex items-center justify-between p-4 bg-emerald-50 rounded-2xl">
                      <span className="text-xs font-bold text-emerald-600">Lunas Bulan Ini</span>
                      <span className="text-sm font-black text-emerald-700">{debts.filter(d => d.status === 'lunas').length} Transaksi</span>
                    </div>
                  </div>
                </div>

                <div className="lg:col-span-3 bg-white rounded-[2.5rem] shadow-[0_20px_50px_rgba(0,0,0,0.02)] border border-slate-100 min-h-[600px] flex flex-col overflow-hidden">
                  <div className="p-10 border-b border-slate-50 flex flex-col md:flex-row md:items-center justify-between gap-6 bg-slate-50/20">
                    <h4 className="font-display font-extrabold text-slate-900 flex items-center gap-3">
                      <div className="p-2 bg-rose-600 text-white rounded-xl shadow-lg shadow-rose-200/50">
                        <Users className="w-5 h-5" />
                      </div>
                      Daftar Penghutang
                    </h4>
                    <div className="flex flex-col sm:flex-row items-center gap-4 w-full md:w-auto">
                      <div className="relative group w-full md:w-auto md:min-w-[300px]">
                        <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                        <input 
                          type="text" 
                          placeholder="Cari nama penghutang..." 
                          className="w-full bg-white border border-slate-200 rounded-xl py-2.5 pl-10 pr-4 text-sm font-medium outline-none focus:ring-4 focus:ring-blue-500/5 transition-all"
                          value={searchQuery}
                          onChange={(e) => setSearchQuery(e.target.value)}
                        />
                      </div>
                      <button 
                        onClick={() => setShowDebtModal(true)}
                        className="w-full sm:w-auto bg-rose-600 text-white px-6 py-2.5 rounded-xl font-bold flex items-center justify-center gap-2 hover:bg-rose-700 transition-all text-sm shadow-lg shadow-rose-200/50 active:scale-95 shrink-0"
                      >
                        <Plus className="w-4 h-4 stroke-[3]" />
                        Catat Hutang
                      </button>
                    </div>
                  </div>

                  <div className="flex-1 p-6 md:p-10 space-y-6 overflow-y-auto no-scrollbar">
                    {debts.filter(d => d.pemberiHutang.toLowerCase().includes(searchQuery.toLowerCase())).length === 0 ? (
                      <div className="h-full flex flex-col items-center justify-center text-center text-slate-300 py-32">
                        <div className="w-24 h-24 mb-6 rounded-full bg-slate-50/80 flex items-center justify-center ring-8 ring-slate-50/30">
                          <Users className="w-12 h-12 opacity-10" />
                        </div>
                        <p className="text-base font-bold text-slate-400">Belum Ada Hutang</p>
                        <p className="text-xs text-slate-300 mt-2 font-medium">Data orang yang berhutang akan muncul di sini.</p>
                      </div>
                    ) : (
                      debts
                        .filter(d => d.pemberiHutang.toLowerCase().includes(searchQuery.toLowerCase()))
                        .map(debt => (
                        <motion.div 
                          layout
                          key={debt.id} 
                          className={`flex flex-col p-8 rounded-[2.5rem] border transition-all duration-300 group ${debt.status === 'lunas' ? 'bg-slate-50/50 border-slate-100 opacity-70' : 'bg-white border-slate-100 hover:border-rose-100 hover:shadow-2xl hover:shadow-rose-500/5'}`}
                        >
                          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                            <div className="flex items-center gap-5">
                              <div className={`w-16 h-16 rounded-[1.5rem] flex items-center justify-center shrink-0 shadow-sm font-black text-xl ${debt.status === 'lunas' ? 'bg-slate-100 text-slate-400' : 'bg-rose-50 text-rose-600'}`}>
                                {debt.pemberiHutang.charAt(0).toUpperCase()}
                              </div>
                              <div className="min-w-0">
                                <h5 className="font-display font-black text-xl text-slate-900 tracking-tight">{debt.pemberiHutang}</h5>
                                <div className="flex items-center gap-3 mt-1.5">
                                  <span className="flex items-center gap-1 text-[10px] font-black text-slate-400 uppercase tracking-widest bg-slate-100 px-2 py-1 rounded-lg">
                                    <Clock className="w-3 h-3" />
                                    {debt.tanggal}
                                  </span>
                                  <span className={`text-[10px] font-black uppercase tracking-widest px-2 py-1 rounded-lg ${debt.status === 'lunas' ? 'bg-emerald-50 text-emerald-600' : 'bg-amber-50 text-amber-600'}`}>
                                    {debt.status === 'lunas' ? 'LUNAS' : 'BELUM BAYAR'}
                                  </span>
                                </div>
                              </div>
                            </div>
                            <div className="flex items-center justify-between md:justify-end gap-10">
                              <div className="text-right">
                                <p className="text-[10px] font-black text-slate-300 uppercase tracking-widest mb-1">Total Nominal</p>
                                <p className={`font-display font-black text-2xl italic ${debt.status === 'lunas' ? 'text-slate-400' : 'text-rose-600'}`}>
                                  Rp {debt.total.toLocaleString()}
                                </p>
                              </div>
                              {debt.status === 'belum_lunas' && (
                                <button 
                                  onClick={() => {
                                    if (confirm(`Tandai hutang ${debt.pemberiHutang} sebagai lunas? Data hutang ini akan dihapus dari daftar karena otomatis masuk ke Riwayat Keuangan.`)) {
                                      setDebts(prev => prev.filter(d => d.id !== debt.id));
                                      // Add to finance as income
                                      setFinance(prev => [{
                                        id: `PL-${debt.id}`,
                                        tipe: 'pemasukan',
                                        jumlah: debt.total,
                                        keterangan: `Pelunasan Hutang: ${debt.pemberiHutang}`,
                                        tanggal: new Date().toISOString().split('T')[0]
                                      }, ...prev]);
                                    }
                                  }}
                                  className="bg-emerald-600 text-white font-black py-4 px-8 rounded-2xl text-[10px] uppercase tracking-widest shadow-xl shadow-emerald-200 hover:bg-emerald-700 transition-all transform active:scale-95"
                                >
                                  Tandai Lunas
                                </button>
                              )}
                              <button 
                                onClick={() => {
                                  if(confirm('Hapus records hutang ini?')) {
                                    setDebts(prev => prev.filter(d => d.id !== debt.id));
                                  }
                                }}
                                className="w-12 h-12 flex items-center justify-center text-slate-300 hover:text-rose-500 hover:bg-rose-50 rounded-2xl transition-all active:scale-90"
                              >
                                <Trash2 className="w-6 h-6" />
                              </button>
                            </div>
                          </div>
                          
                          <div className="mt-8 pt-6 border-t border-slate-50 flex flex-wrap gap-3">
                            <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mr-2">Daftar Barang:</span>
                            {debt.items.map((item, idx) => (
                              <div key={idx} className="bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-100 flex items-center gap-2">
                                <span className="text-[10px] font-bold text-slate-600">{item.nama}</span>
                                <span className="w-1 h-1 bg-slate-300 rounded-full" />
                                <span className="text-[10px] font-black text-blue-600">x{item.jumlah}</span>
                              </div>
                            ))}
                          </div>
                        </motion.div>
                      ))
                    )}
                  </div>
                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Debt Modal */}
      {showDebtModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[1000] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0, y: 10 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            className="bg-white w-full max-w-xl rounded-[3rem] shadow-[0_25px_80px_rgba(0,0,0,0.15)] overflow-hidden"
          >
            <div className="p-6 sm:p-10 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div>
                <h3 className="text-xl sm:text-2xl font-display font-black text-slate-900 tracking-tight">
                  Catat Hutang Baru
                </h3>
                <p className="text-[10px] sm:text-xs text-slate-400 font-bold uppercase tracking-[0.2em] mt-1">Rekam data hutang manual</p>
              </div>
              <button 
                onClick={() => setShowDebtModal(false)}
                className="w-10 h-10 sm:w-12 sm:h-12 flex shrink-0 items-center justify-center bg-white border border-slate-100 shadow-sm rounded-2xl text-slate-400 hover:text-rose-500 transition-all hover:rotate-90"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <form className="p-6 sm:p-10 space-y-6 sm:space-y-8" onSubmit={(e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const data = {
                nama: formData.get('nama') as string,
                keterangan: formData.get('keterangan') as string,
                nominal: parseInt(formData.get('nominal') as string),
              };
              
              if (!data.nama || isNaN(data.nominal) || data.nominal <= 0) {
                alert('Mohon isi data dengan lengkap dan benar!');
                return;
              }

              const transactionId = Math.random().toString(36).substr(2, 9).toUpperCase();
              
              setDebts(prev => {
                const existingDebtIndex = prev.findIndex(d => d.pemberiHutang.toLowerCase() === data.nama.toLowerCase() && d.status === 'belum_lunas');
                
                if (existingDebtIndex !== -1) {
                  const updatedDebts = [...prev];
                  const existingDebt = updatedDebts[existingDebtIndex];
                  updatedDebts[existingDebtIndex] = {
                    ...existingDebt,
                    items: [...existingDebt.items, { nama: data.keterangan || 'Hutang Manual', harga: data.nominal, jumlah: 1 }],
                    total: existingDebt.total + data.nominal,
                    tanggal: new Date().toISOString().split('T')[0]
                  };
                  return updatedDebts;
                } else {
                  const newDebt: Debt = {
                    id: transactionId,
                    pemberiHutang: data.nama,
                    items: [{ nama: data.keterangan || 'Hutang Manual', harga: data.nominal, jumlah: 1 }],
                    total: data.nominal,
                    tanggal: new Date().toISOString().split('T')[0],
                    status: 'belum_lunas'
                  };
                  return [newDebt, ...prev];
                }
              });
              setShowDebtModal(false);
            }}>
              <div className="space-y-5">
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-1 mb-2 block">Nama Penghutang</label>
                  <input 
                    name="nama" 
                    required 
                    placeholder="Contoh: Bpk. Budi..."
                    className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-4 px-5 text-sm font-bold outline-none focus:ring-4 focus:ring-rose-500/10 focus:border-rose-500 transition-all"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-1 mb-2 block">Keterangan Barang/Tujuan</label>
                  <input 
                    name="keterangan" 
                    placeholder="Contoh: Hutang sembako..."
                    className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-4 px-5 text-sm font-bold outline-none focus:ring-4 focus:ring-rose-500/10 focus:border-rose-500 transition-all"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest pl-1 mb-2 block">Total Nominal (Rp)</label>
                  <input 
                    name="nominal" 
                    type="number" 
                    required 
                    placeholder="100000"
                    className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-4 px-5 text-sm font-bold outline-none focus:ring-4 focus:ring-rose-500/10 focus:border-rose-500 transition-all"
                  />
                </div>
              </div>
              <button 
                type="submit"
                className="w-full bg-rose-600 text-white rounded-2xl py-4 font-bold tracking-wide hover:bg-rose-700 transition-all shadow-xl shadow-rose-200/50 active:scale-95"
              >
                Simpan Hutang
              </button>
            </form>
          </motion.div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {productToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[1100] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="bg-white w-full max-w-sm rounded-[2rem] shadow-2xl overflow-hidden p-6 sm:p-8 text-center"
          >
            <div className="w-16 h-16 bg-rose-100 text-rose-500 rounded-full flex items-center justify-center mx-auto mb-4">
              <Trash2 className="w-8 h-8" />
            </div>
            <h3 className="text-xl font-black text-slate-900 mb-2">Hapus Produk?</h3>
            <p className="text-sm text-slate-500 mb-8 leading-relaxed">
              Apakah Anda yakin ingin menghapus produk <span className="font-bold text-slate-700">{productToDelete.nama}</span>? Tindakan ini tidak dapat dibatalkan.
            </p>
            
            <div className="flex gap-3">
              <button 
                onClick={() => setProductToDelete(null)}
                className="flex-1 bg-slate-100 text-slate-600 font-bold py-3.5 rounded-xl hover:bg-slate-200 transition-all text-sm"
              >
                Batal
              </button>
              <button 
                onClick={async () => {
                   const p = productToDelete;
                   setProductToDelete(null);
                   if (user) {
                     try {
                       await deleteDoc(doc(db, 'users', user.uid, 'products', p.id));
                       setImportingInfo('Berhasil menghapus produk!');
                       setTimeout(() => setImportingInfo(''), 3000);
                     } catch(err) {
                       console.error(err);
                     }
                   } else {
                     setProducts(prev => prev.filter(item => item.id !== p.id));
                     setImportingInfo('Berhasil menghapus produk!');
                     setTimeout(() => setImportingInfo(''), 3000);
                   }
                }}
                className="flex-1 bg-rose-500 text-white font-bold py-3.5 rounded-xl hover:bg-rose-600 transition-all shadow-lg shadow-rose-200 text-sm"
              >
                Hapus
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Profile Modal */}
      {showProfileModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[1000] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0, y: 10 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            className="bg-white w-full max-w-sm rounded-[2rem] shadow-[0_25px_80px_rgba(0,0,0,0.15)] overflow-hidden"
          >
            <div className="p-6 sm:p-8 bg-slate-50/50 border-b border-slate-100 flex items-center justify-between">
              <h3 className="text-xl font-black text-slate-900 tracking-tight">Profil Pengguna</h3>
              <button 
                onClick={() => setShowProfileModal(false)}
                className="w-10 h-10 flex items-center justify-center bg-white border border-slate-200 rounded-full text-slate-400 hover:text-rose-500 hover:border-rose-200 hover:bg-rose-50 transition-all"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <form className="p-6 sm:p-8 space-y-4" onSubmit={async (e) => {
              e.preventDefault();
              if (user) {
                try {
                  // If profileImage is a URL (starts with http), we can save it to Firebase Auth
                  // Base64 strings can be too large for Firebase auth photoURL, so skip it there
                  const canSaveToAuth = profileImage.startsWith('http://') || profileImage.startsWith('https://');
                  
                  await updateProfile(user, {
                    displayName: profileName,
                    ...(canSaveToAuth ? { photoURL: profileImage } : {})
                  });
                  // force re-render by updating user state inline
                  setUser({ ...user, displayName: profileName, ...(canSaveToAuth ? { photoURL: profileImage } : {}) });
                  
                  const docRefProf = doc(db, 'users', user.uid);
                  await setDoc(docRefProf, {
                    storeName: profileStoreName,
                    address: profileAddress,
                    profileImage: profileImage
                  }, { merge: true });

                  setImportingInfo('Profil berhasil diperbarui!');
                  setTimeout(() => setImportingInfo(''), 3000);
                  setShowProfileModal(false);
                } catch (error) {
                  console.error(error);
                  setImportingInfo('Gagal memperbarui profil');
                  setTimeout(() => setImportingInfo(''), 3000);
                }
              }
            }}>
              <div className="flex flex-col sm:flex-row gap-4">
                <div className="w-full">
                  <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Nama Pengguna</label>
                  <input 
                    type="text" 
                    value={profileName}
                    onChange={(e) => setProfileName(e.target.value)}
                    placeholder="Nama Lengkap"
                    className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all"
                  />
                </div>
                <div className="w-full">
                  <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Nama Toko</label>
                  <input 
                    type="text" 
                    value={profileStoreName}
                    onChange={(e) => setProfileStoreName(e.target.value)}
                    placeholder="Nama Toko"
                    className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all"
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Alamat Toko</label>
                <input 
                  type="text" 
                  value={profileAddress}
                  onChange={(e) => setProfileAddress(e.target.value)}
                  placeholder="Alamat Lengkap"
                  className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all"
                />
              </div>

              <div>
                <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Foto Profil</label>
                <div className="flex border-2 border-dashed border-slate-200 rounded-2xl p-4 md:p-6 bg-slate-50/50 hover:bg-slate-50 relative group transition-all">
                   <div className="w-full text-center">
                     <div className="w-12 h-12 bg-white rounded-xl flex items-center justify-center mx-auto mb-3 shadow-sm border border-slate-100 group-hover:scale-110 transition-transform">
                       <Plus className="w-6 h-6 text-slate-400" />
                     </div>
                     <span className="text-xs font-bold text-slate-500 block">Pilih Gambar Profil</span>
                     <span className="text-[10px] font-medium text-slate-400 mt-1 block">Atau drag & drop ke sini</span>
                   </div>
                   <input type="file" accept="image/*" onChange={handleProfileImageUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                </div>
                
                {profileImage && (
                  <div className="mt-4 flex justify-center">
                    <img src={profileImage} alt="Preview" className="w-16 h-16 rounded-full object-cover border-4 border-white shadow-lg" />
                  </div>
                )}
              </div>

              <div className="pt-4 border-t border-slate-100 flex gap-3">
                <button 
                  type="button"
                  onClick={() => setShowProfileModal(false)}
                  className="w-full py-4 bg-slate-100 text-slate-500 font-bold rounded-2xl hover:bg-slate-200 transition-all text-sm"
                >
                  Batal
                </button>
                <button 
                  type="submit"
                  className="w-full py-4 bg-blue-600 text-white font-bold rounded-2xl shadow-lg shadow-blue-200 hover:bg-blue-700 hover:scale-[1.02] active:scale-95 transition-all text-sm"
                >
                  Simpan
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {/* Product Modal */}
      {showProductModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[1000] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0, y: 10 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            className="bg-white w-full max-w-xl rounded-[3rem] shadow-[0_25px_80px_rgba(0,0,0,0.15)] overflow-hidden"
          >
            <div className="p-6 sm:p-10 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div>
                <h3 className="text-xl sm:text-2xl font-display font-black text-slate-900 tracking-tight">
                  {editingProduct ? 'Edit Informasi Produk' : 'Registrasi Produk Baru'}
                </h3>
                <p className="text-[10px] sm:text-xs text-slate-400 font-bold uppercase tracking-[0.2em] mt-1">Lengkapi data produk di bawah ini</p>
              </div>
              <button 
                onClick={() => setShowProductModal(false)}
                className="w-10 h-10 sm:w-12 sm:h-12 flex shrink-0 items-center justify-center bg-white border border-slate-100 shadow-sm rounded-2xl text-slate-400 hover:text-rose-500 transition-all hover:rotate-90"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <form className="p-6 sm:p-10 space-y-6 sm:space-y-8" onSubmit={async (e) => {
              e.preventDefault();
              const formData = new FormData(e.currentTarget);
              const data = {
                id: editingProduct ? editingProduct.id : `PROD-${Date.now()}-${Math.random().toString(36).substr(2, 5).toUpperCase()}`,
                nama: formData.get('nama') as string,
                harga: Number(formData.get('harga')),
                stok: Number(formData.get('stok')),
                kategori: formData.get('kategori') as string || 'Lain-lain',
                url_gambar: formImage || `https://ui-avatars.com/api/?name=${encodeURIComponent(formData.get('nama') as string)}&background=random&color=fff&size=200`
              };

              const isEditing = !!editingProduct;

              if (user) {
                try {
                  const docRef = doc(db, 'users', user.uid, 'products', data.id);
                  await setDoc(docRef, data);
                  setImportingInfo(isEditing ? 'Berhasil mengubah produk!' : 'Berhasil menambahkan produk!');
                  setTimeout(() => setImportingInfo(''), 3000);
                } catch (err) {
                  console.error(err);
                  alert('Gagal menyimpan produk');
                }
              } else {
                if (isEditing) {
                  setProducts(prev => prev.map(p => p.id === editingProduct!.id ? data : p));
                  setImportingInfo('Berhasil mengubah produk!');
                } else {
                  setProducts(prev => [...prev, data]);
                  setImportingInfo('Berhasil menambahkan produk!');
                }
                setTimeout(() => setImportingInfo(''), 3000);
              }

              setShowProductModal(false);
              setEditingProduct(null);
            }}>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div className="space-y-8 md:col-span-2">
                  <div>
                    <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Nama Produk</label>
                    <input name="nama" defaultValue={editingProduct?.nama} required className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all placeholder:font-medium placeholder:text-slate-300" placeholder="Contoh: Indomie Goreng Spesial..." />
                  </div>
                </div>

                <div>
                  <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Harga Jual (Rp)</label>
                  <div className="relative">
                    <span className="absolute left-6 top-1/2 -translate-y-1/2 font-black text-slate-400">Rp</span>
                    <input name="harga" type="number" defaultValue={editingProduct?.harga} required className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 pl-14 pr-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all" placeholder="0" />
                  </div>
                </div>

                <div>
                  <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Jumlah Stok</label>
                  <input name="stok" type="number" defaultValue={editingProduct?.stok} required className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all" placeholder="0" />
                </div>

                <div>
                  <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block">Klasifikasi Kategori</label>
                  <input 
                    name="kategori" 
                    list="kategori_default_list"
                    defaultValue={editingProduct?.kategori || ""} 
                    required 
                    placeholder="Contoh: Sembako"
                    className="w-full bg-slate-50 border border-slate-200/60 rounded-2xl py-4 px-6 text-sm font-bold outline-none focus:ring-4 focus:ring-blue-500/5 focus:border-blue-500 transition-all"
                  />
                  <datalist id="kategori_default_list">
                    {dynamicCategories.map(cat => (
                      <option key={cat} value={cat}>{cat}</option>
                    ))}
                  </datalist>
                </div>

                <div className="min-w-0">
                  <label className="text-[10px] uppercase tracking-[0.2em] font-black text-slate-300 mb-3 block truncate">Foto Produk (Opsional)</label>
                  <div className="flex gap-3 sm:gap-4 items-center w-full min-w-0">
                    {formImage && (
                      <img src={formImage} alt="Preview" className="w-14 h-14 sm:w-16 sm:h-16 rounded-2xl object-cover shrink-0 border border-slate-200" />
                    )}
                    <label className="flex-1 min-w-0 bg-slate-50 border border-slate-200/60 rounded-xl sm:rounded-2xl py-3 px-3 sm:py-4 sm:px-6 text-xs sm:text-sm font-bold flex items-center justify-center cursor-pointer hover:bg-slate-100 transition-all text-blue-600 text-center">
                      <span className="truncate">{formImage ? 'Ubah Gambar' : 'Pilih Gambar'}</span>
                      <input type="file" accept="image/*" onChange={handleImageUpload} className="hidden" />
                    </label>
                  </div>
                </div>
              </div>

              <div className="pt-6 flex gap-4">
                <button 
                  type="button"
                  onClick={() => setShowProductModal(false)}
                  className="flex-1 bg-rose-500 text-white font-black py-5 rounded-[1.5rem] hover:bg-rose-600 shadow-xl shadow-rose-200/50 transition-all uppercase tracking-widest text-[10px]"
                >
                  Batalkan
                </button>
                <button className="flex-[2] bg-blue-600 text-white font-black py-5 rounded-[1.5rem] shadow-2xl shadow-blue-200 hover:bg-blue-700 transition-all uppercase tracking-[0.2em] text-[10px] transform active:scale-95">
                  {editingProduct ? 'Simpan Perubahan' : 'Terbitkan Produk'}
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}

      {/* Invoice Modal */}
      {showInvoice && lastTransaction && (
        <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-xl z-[1001] flex items-center justify-center p-4 print:p-0 print:bg-white print:block print:relative overflow-y-auto">
          <motion.div 
            initial={{ y: 100, opacity: 0, scale: 0.9 }}
            animate={{ y: 0, opacity: 1, scale: 1 }}
            id="printable-invoice"
            className="bg-white w-full max-w-sm rounded-[3.5rem] shadow-[0_40px_100px_rgba(0,0,0,0.3)] overflow-hidden relative print:shadow-none print:max-w-none print:w-full print:rounded-none flex flex-col max-h-[90vh]"
          >
            {/* Header Struk - Premium Look */}
            <div className="bg-slate-900 p-6 sm:p-8 text-center text-white relative print:bg-white print:text-slate-800 print:border-b-4 print:border-double print:border-slate-800 shrink-0">
              <div className="w-14 h-14 bg-blue-600 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-xl shadow-blue-500/40 border-4 border-slate-800 print:hidden translate-y-2">
                <CreditCard className="w-6 h-6 text-white stroke-[2.5]" />
              </div>
              <h3 className="text-xl font-display font-black tracking-tighter mt-2">SnapCashier</h3>
              <p className="text-slate-400 text-[10px] font-black uppercase tracking-[0.3em] mt-1">Official Receipt</p>
              
              <button 
                onClick={() => setShowInvoice(false)}
                className="absolute top-6 right-6 w-10 h-10 flex items-center justify-center bg-slate-800 hover:bg-rose-500 text-slate-400 hover:text-white rounded-full transition-all print:hidden shrink-0 z-10"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-6 sm:p-8 space-y-6 bg-[radial-gradient(circle_at_top_right,_var(--tw-gradient-stops))] from-blue-50/20 via-white to-white flex-1 overflow-y-auto no-scrollbar print:overflow-visible relative">
              <div className="flex justify-between items-start text-[10px] border-b border-slate-100 pb-4 print:border-slate-300">
                <div className="space-y-4">
                  <div>
                    <p className="text-slate-300 font-black uppercase tracking-widest mb-1">Receipt ID</p>
                    <p className="text-slate-900 font-mono font-bold text-sm tracking-tight italic">#{lastTransaction.id.toUpperCase()}</p>
                  </div>
                  <div>
                    <p className="text-slate-300 font-black uppercase tracking-widest mb-1">Mtd. Bayar</p>
                    <span className="text-blue-600 font-black flex items-center gap-1 text-[11px] leading-none">
                      <div className="w-1.5 h-1.5 rounded-full bg-blue-500 animate-pulse" />
                      {lastTransaction.method}
                      {lastTransaction.debtor && (
                        <span className="bg-rose-50 text-rose-600 px-2 py-1 rounded ml-1.5 font-bold uppercase text-[9px] tracking-wider border border-rose-100">Hutang: {lastTransaction.debtor}</span>
                      )}
                    </span>
                  </div>
                </div>
                <div className="text-right">
                  <p className="text-slate-300 font-black uppercase tracking-widest mb-1">Waktu Selesai</p>
                  <p className="text-slate-800 font-bold text-sm">{lastTransaction.date}</p>
                  <p className="text-[10px] text-slate-400 mt-1 font-medium">{new Date().toLocaleTimeString('id-ID')}</p>
                </div>
              </div>

              {/* Items Table Style */}
              <div className="space-y-4">
                <p className="text-[10px] font-black text-slate-300 uppercase tracking-[0.2em] italic border-l-2 border-blue-500 pl-3 sticky top-0 bg-white z-10 w-full py-1">Order Details</p>
                <div className="space-y-3 pr-2">
                  {lastTransaction.items.map((item, idx) => (
                    <div key={idx} className="flex justify-between items-center group">
                      <div className="flex-1 min-w-0 pr-4">
                        <p className="font-bold text-slate-800 text-sm truncate tracking-tight">{item.nama}</p>
                        <p className="text-[10px] text-slate-400 font-medium">Rp {item.harga.toLocaleString()} × {item.jumlah}</p>
                      </div>
                      <p className="font-black text-slate-900 text-sm">Rp {(item.harga * item.jumlah).toLocaleString()}</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="p-6 sm:p-8 bg-white border-t border-slate-100 shrink-0 print:border-t-0 space-y-6">
              {/* Calculation */}
              <div className="border-t-2 border-dashed border-slate-100 pt-4 space-y-4 print:border-slate-800">
                <div className="flex justify-between items-center bg-slate-50 p-3 sm:p-4 rounded-2xl border border-slate-100/50">
                  <p className="text-slate-500 text-[10px] sm:text-xs font-bold">Subtotal Pesanan</p>
                  <p className="text-slate-800 font-black text-sm sm:text-base">Rp {lastTransaction.total.toLocaleString()}</p>
                </div>
                <div className="flex justify-between items-center px-1 sm:px-2">
                  <div>
                    <p className="text-blue-600 font-black text-lg sm:text-xl italic tracking-tighter">TOTAL TRANSFER</p>
                    <p className="text-[8px] sm:text-[9px] text-blue-400 font-bold uppercase tracking-widest">Nett Amount</p>
                  </div>
                  <p className="text-blue-600 font-display font-black text-2xl sm:text-3xl tracking-tighter italic">Rp {lastTransaction.total.toLocaleString()}</p>
                </div>
              </div>

              {/* Thank You Note */}
              <div className="text-center pt-2 space-y-6">
                <div>
                  <p className="text-[9px] sm:text-[10px] text-slate-300 font-bold uppercase tracking-[0.2em] mb-4 print:text-slate-800">
                    Sistem Kasir Pintar • SnapCashier V2.0
                  </p>
                  <div className="w-16 h-1 mx-auto flex justify-center gap-1.5 mb-4">
                    <div className="w-2 h-1 bg-slate-100 rounded-full" />
                    <div className="w-8 h-1 bg-blue-500 rounded-full" />
                    <div className="w-2 h-1 bg-slate-100 rounded-full" />
                  </div>
                </div>

                <div className="flex flex-col gap-3 print:hidden">
                  <button 
                    onClick={() => {
                      setShowToast(true);
                      setTimeout(() => setShowToast(false), 3000);
                      window.print();
                    }}
                    className="w-full bg-slate-900 text-white font-black py-4 sm:py-5 rounded-[1.5rem] hover:bg-slate-800 transition-all text-xs tracking-[0.2em] flex items-center justify-center gap-3 shadow-2xl shadow-slate-200"
                  >
                    <Printer className="w-5 h-5 stroke-[2.5]" />
                    CETAK STRUK SEKARANG
                  </button>
                  <div className="grid grid-cols-2 gap-3 sm:gap-4">
                    <button 
                      onClick={handleEditOrder}
                      className="bg-white border border-slate-200 text-slate-500 font-black py-3 sm:py-4 rounded-2xl hover:bg-slate-50 transition-all text-[10px] tracking-widest uppercase flex items-center justify-center gap-2"
                    >
                      <Plus className="w-4 h-4" />
                      Ubah Item
                    </button>
                    <button 
                      onClick={() => setShowInvoice(false)}
                      className="bg-blue-600 text-white font-black py-3 sm:py-4 rounded-2xl hover:bg-blue-700 transition-all text-[10px] tracking-widest uppercase shadow-xl shadow-blue-100"
                    >
                      Transaksi Baru
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </motion.div>
        </div>
      )}

      {/* Payment Processing Modal */}
      {isProcessingPayment && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[1000] flex items-center justify-center p-4">
          <motion.div 
            initial={{ scale: 0.95, opacity: 0, y: 10 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            className="bg-white w-full max-w-md rounded-[3rem] shadow-[0_25px_80px_rgba(0,0,0,0.15)] overflow-hidden flex flex-col max-h-[90vh]"
          >
            <div className="p-6 sm:p-10 border-b border-slate-100 text-centershrink-0 bg-slate-50/50">
              <h3 className="text-xl sm:text-2xl font-display font-black text-slate-900 tracking-tight mb-2 text-center">
                Pembayaran {selectedMethod}
              </h3>
              <p className="text-[10px] sm:text-xs text-slate-400 font-bold uppercase tracking-[0.2em] text-center">
                Selesaikan pembayaran
              </p>
            </div>
            
            <div className="p-4 sm:p-6 flex-1 space-y-4 flex flex-col items-center overflow-y-auto no-scrollbar">
              {selectedMethod === 'QRIS' && (
                <>
                  <div className="text-center w-full">
                    <p className="text-[10px] font-black text-slate-400 mb-1 uppercase tracking-widest">Total Belanja</p>
                    <p className="text-3xl font-display font-black text-blue-600 tracking-tighter">
                      Rp {(cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0)).toLocaleString()}
                    </p>
                  </div>
                  <div className="p-4 bg-white rounded-2xl border border-slate-100 shadow-lg shadow-slate-200/50 flex flex-col items-center justify-center space-y-4 w-full">
                    <div className="w-40 h-40 bg-slate-50 rounded-2xl flex items-center justify-center border-2 border-dashed border-slate-200">
                      <QrCode className="w-20 h-20 text-slate-300" />
                    </div>
                    <p className="text-[10px] font-bold text-slate-400 text-center max-w-[200px]">Scan QR code ini dengan aplikasi E-Wallet pelanggan.</p>
                  </div>
                </>
              )}

              {selectedMethod === 'M-Banking' && (
                <div className="w-full space-y-4">
                  <div className="text-center w-full">
                    <p className="text-[10px] font-black text-slate-400 mb-1 uppercase tracking-widest">Total Belanja</p>
                    <p className="text-3xl font-display font-black text-blue-600 tracking-tighter">
                      Rp {(cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0)).toLocaleString()}
                    </p>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    {['BCA', 'Mandiri', 'BRI', 'BNI'].map(bank => (
                      <button
                        key={bank}
                        onClick={() => setSelectedBank(bank)}
                        className={`py-3 rounded-xl font-black text-xs uppercase tracking-wide border-2 transition-all active:scale-95 ${selectedBank === bank ? 'border-blue-600 text-blue-600 bg-blue-50/50 shadow-md shadow-blue-500/10' : 'border-slate-100 text-slate-400 hover:border-slate-300 hover:text-slate-600 bg-white'}`}
                      >
                        {bank}
                      </button>
                    ))}
                  </div>
                  
                  <motion.div 
                    key={selectedBank}
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="p-4 bg-slate-50 rounded-2xl border border-slate-200/60 text-center relative mt-4 text-[10px]"
                  >
                    <p className="font-black text-slate-400 uppercase tracking-widest mb-1 shadow-sm">No. Rekening {selectedBank}</p>
                    <p className="text-xl font-mono font-black text-slate-700 tracking-tight">
                      {selectedBank === 'BCA' && '8234 5678 90'}
                      {selectedBank === 'Mandiri' && '137 000 1234'}
                      {selectedBank === 'BRI' && '0123 4567 890'}
                      {selectedBank === 'BNI' && '0987 6543 21'}
                    </p>
                  </motion.div>
                </div>
              )}

              {(selectedMethod === 'DANA' || selectedMethod === 'OVO') && (
                <>
                  <div className="text-center w-full">
                    <p className="text-[10px] font-black text-slate-400 mb-1 uppercase tracking-widest">Total Belanja</p>
                    <p className="text-3xl font-display font-black text-blue-600 tracking-tighter">
                      Rp {(cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0)).toLocaleString()}
                    </p>
                  </div>
                  <div className="p-4 bg-white rounded-2xl border border-slate-100 shadow-lg shadow-slate-200/50 flex flex-col items-center justify-center w-full relative overflow-hidden">
                    <div className="absolute inset-0 bg-blue-50/30"></div>
                    <div className="relative z-10 flex flex-col items-center space-y-4">
                      <div className="w-32 h-32 bg-white rounded-2xl flex items-center justify-center shadow-sm border border-slate-100">
                        <QrCode className="w-16 h-16 text-slate-300" />
                      </div>
                      
                      <div className="w-full flex items-center gap-2">
                         <div className="h-px bg-slate-200 flex-1"></div>
                         <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Atau Transfer</span>
                         <div className="h-px bg-slate-200 flex-1"></div>
                      </div>

                      <div className="text-center w-full">
                        <p className="text-[10px] font-black text-blue-500 mb-1 uppercase tracking-widest">{selectedMethod} Number</p>
                        <p className="text-xl font-mono font-black text-slate-700 tracking-tight">0812 3456 7890</p>
                      </div>
                    </div>
                  </div>
                </>
              )}

              {selectedMethod === 'Tunai' && (
                <div className="w-full p-6 bg-emerald-50 rounded-3xl border border-emerald-100 flex flex-col items-center justify-center text-center space-y-4">
                  <div className="text-center w-full">
                    <p className="text-[10px] font-black text-emerald-600/60 mb-1 uppercase tracking-widest">Total Belanja</p>
                    <p className="text-4xl font-display font-black text-emerald-600 tracking-tighter">
                      Rp {(cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0)).toLocaleString()}
                    </p>
                  </div>
                  <div className="h-px w-full max-w-[200px] bg-emerald-200/50"></div>
                  <p className="text-xs font-bold text-emerald-700/60 leading-relaxed max-w-[250px]">
                    Pastikan Anda menerima uang tunai sesuai tagihan sebelum menekan Selesai.
                  </p>
                </div>
              )}

              {selectedMethod === 'Hutang' && (
                <div className="w-full p-6 bg-rose-50 rounded-3xl border border-rose-100 flex flex-col items-center justify-center text-center space-y-4">
                  <div className="text-center w-full">
                    <p className="text-[10px] font-black text-rose-600/60 mb-1 uppercase tracking-widest">Total Hutang</p>
                    <p className="text-4xl font-display font-black text-rose-600 tracking-tighter">
                      Rp {(cart.reduce((sum, item) => sum + (item.harga * item.jumlah), 0)).toLocaleString()}
                    </p>
                  </div>
                  <div className="h-px w-full max-w-[200px] bg-rose-200/50"></div>
                  <div>
                     <p className="text-[10px] font-black text-rose-700/60 uppercase tracking-widest mb-1">Dicatat atas nama</p>
                     <p className="font-black text-rose-700 uppercase">{debtorName}</p>
                  </div>
                </div>
              )}
            </div>
            
            <div className="p-6 border-t border-slate-100 flex gap-4 shrink-0 bg-slate-50/50">
              <button 
                onClick={() => setIsProcessingPayment(false)}
                className="flex-1 bg-rose-500 border border-rose-500 text-white rounded-2xl py-4 font-black uppercase text-xs tracking-widest hover:bg-rose-600 transition-all active:scale-95 shadow-xl shadow-rose-200/50"
              >
                Batal
              </button>
              <button 
                onClick={() => {
                  processPayment();
                  setIsProcessingPayment(false);
                }}
                className="flex-[2] bg-blue-600 text-white rounded-2xl py-4 font-black uppercase text-xs tracking-widest hover:bg-blue-700 transition-all shadow-xl shadow-blue-200/50 active:scale-95 flex items-center justify-center gap-2 group"
              >
                Selesai Bayar
                <Check className="w-5 h-5 group-hover:scale-110 transition-transform" />
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {/* Floating Checkout Button for Tablet (where right sidebar is hidden) */}
      <AnimatePresence>
        {activeTab === 'pos' && cart.length > 0 && !showInvoice && (
          <motion.div 
            initial={{ y: 20, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: 50, opacity: 0 }}
            className="fixed bottom-6 right-6 z-[450] hidden md:block xl:hidden"
          >
            <button 
              onClick={() => setShowMobileCart(true)}
              className={`flex items-center gap-4 ${selectedMethod === 'Hutang' ? 'bg-rose-500 shadow-rose-200' : 'bg-blue-600 shadow-blue-200'} text-white rounded-2xl py-4 px-8 shadow-2xl font-bold transition-all hover:scale-105 active:scale-95`}
            >
              <div className="flex flex-col items-start leading-none gap-1">
                <span className="text-[10px] font-black uppercase opacity-70 tracking-widest">{cart.length} Item</span>
                <span className="text-xl">Rp {totalBelanja.toLocaleString()}</span>
              </div>
              <div className="h-10 w-px bg-white/20 mx-2" />
              <div className="flex items-center gap-2">
                <span className="text-sm font-black uppercase tracking-widest">{selectedMethod === 'Hutang' ? 'HUTANG' : 'BAYAR'}</span>
                <ChevronRight className="w-5 h-5 stroke-[3]" />
              </div>
            </button>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Mobile-friendly Bottom Nav & Actions */}
      <div className="fixed bottom-0 left-0 right-0 z-[500] md:hidden">
        {/* Floating Cart Button for Mobile */}
        <AnimatePresence>
          {activeTab === 'pos' && cart.length > 0 && !showInvoice && (
            <motion.div 
              initial={{ y: 20, opacity: 0 }}
              animate={{ y: 0, opacity: 1 }}
              exit={{ y: 50, opacity: 0 }}
              className="absolute -top-20 left-1/2 -translate-x-1/2 w-full px-4"
            >
              <button 
                onClick={() => setShowMobileCart(true)}
                className={`w-full ${selectedMethod === 'Hutang' ? 'bg-rose-500 shadow-rose-200' : 'bg-blue-600 shadow-blue-200'} text-white rounded-2xl py-4 shadow-2xl flex items-center justify-between px-6 font-bold transition-all active:scale-95`}
              >
                <div className="flex items-center gap-3">
                  <div className="bg-white/20 px-2.5 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest">{cart.length} Item</div>
                  <span>{selectedMethod === 'Hutang' ? 'HUTANG' : 'BAYAR'}</span>
                </div>
                <div className="text-right">
                  <p className="text-[10px] opacity-70 uppercase font-black">Total</p>
                  <p className="text-lg">Rp {totalBelanja.toLocaleString()}</p>
                </div>
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Bottom Nav */}
        <div className="bg-white/90 backdrop-blur-2xl border-t border-slate-100 px-6 py-4 flex justify-around items-center shadow-[0_-15px_40px_rgba(0,0,0,0.06)] rounded-t-[3rem] ring-1 ring-black/5">
          {[
            { id: 'dashboard', icon: LayoutDashboard, label: 'Beranda' },
            { id: 'inventory', icon: Package, label: 'Gudang' },
            { id: 'debts', icon: BookOpen, label: 'Hutang' },
            { id: 'finance', icon: Wallet, label: 'Laporan' },
            { id: 'pos', icon: ShoppingCart, label: 'Kasir' },
          ].map(item => (
            <button 
              key={item.id} 
              onClick={() => setActiveTab(item.id as any)}
              className={`flex flex-col items-center gap-1.5 transition-all duration-300 relative ${activeTab === item.id ? 'text-blue-600' : 'text-slate-400'}`}
            >
              <div className={`p-3 rounded-2xl transition-all duration-500 ${activeTab === item.id ? 'bg-blue-600 text-white shadow-xl shadow-blue-200 -translate-y-2' : ''}`}>
                <item.icon className={`w-6 h-6 ${activeTab === item.id ? 'stroke-[2.5]' : 'stroke-2'}`} />
              </div>
              <span className={`text-[10px] font-black uppercase tracking-widest transition-all ${activeTab === item.id ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-2'}`}>{item.label}</span>
              {activeTab === item.id && (
                <motion.div layoutId="nav-dot" className="absolute -top-1 w-1 h-1 bg-blue-600 rounded-full" />
              )}
            </button>
          ))}
        </div>
      </div>

      {/* Mobile Cart Drawer/Modal */}
      <AnimatePresence>
        {showMobileCart && activeTab === 'pos' && (
          <div className="fixed inset-0 z-[600] md:hidden">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowMobileCart(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ y: '100%' }}
              animate={{ y: 0 }}
              exit={{ y: '100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="absolute bottom-0 left-0 right-0 bg-white rounded-t-[3rem] max-h-[85vh] flex flex-col overflow-hidden"
            >
              <div className="w-12 h-1.5 bg-slate-200 rounded-full mx-auto mt-4 mb-2" />
              
              <div className="p-6 sm:p-8 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                <div className="flex items-center gap-3 sm:gap-4">
                  <div className="p-2 sm:p-3 bg-blue-600 text-white rounded-xl sm:rounded-2xl shadow-lg shadow-blue-200">
                    <ShoppingCart className="w-5 h-5 sm:w-6 sm:h-6 stroke-[2.5]" />
                  </div>
                  <div>
                    <h4 className="font-display font-black text-slate-900 text-lg sm:text-xl tracking-tight">Rincian Belanja</h4>
                    <p className="text-[9px] sm:text-[10px] font-black text-slate-400 uppercase tracking-widest mt-1">{cart.length} Jenis Produk</p>
                  </div>
                </div>
                <button onClick={clearCart} className="p-2 sm:p-3 text-rose-500 hover:bg-rose-50 rounded-xl sm:rounded-2xl transition-all active:scale-90">
                  <Trash2 className="w-5 h-5 sm:w-6 sm:h-6" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-4 sm:p-8 space-y-4 sm:space-y-6 no-scrollbar scroll-smooth">
                <AnimatePresence mode="popLayout">
                  {cart.map(item => (
                    <motion.div 
                      layout
                      initial={{ opacity: 0, x: 20 }}
                      animate={{ opacity: 1, x: 0 }}
                      exit={{ opacity: 0, x: -20 }}
                      key={item.id} 
                      className="flex flex-col bg-white p-4 sm:p-6 rounded-2xl sm:rounded-[2.5rem] border border-slate-100 group shadow-sm hover:border-blue-100 hover:shadow-xl hover:shadow-blue-500/5 transition-all duration-300"
                    >
                      <div className="flex items-center justify-between mb-3 sm:mb-4">
                        <div className="flex-1 min-w-0 pr-4">
                          <p className="font-bold text-slate-800 text-sm sm:text-lg truncate tracking-tight">{item.nama}</p>
                          <p className="text-xs sm:text-sm text-blue-600 font-extrabold mt-0.5">Rp {item.harga.toLocaleString()}</p>
                        </div>
                        <p className="font-black text-slate-900 text-base sm:text-lg">Rp {(item.harga * item.jumlah).toLocaleString()}</p>
                      </div>
                      
                      <div className="flex items-center justify-between pt-3 sm:pt-4 border-t border-slate-50">
                        <div className="flex items-center bg-slate-50 rounded-xl sm:rounded-2xl border border-slate-200 p-1 sm:p-1.5 px-2 sm:px-3">
                          <button 
                            onClick={() => updateCartQuantity(item.id, -1)}
                            className="p-1.5 sm:p-2 hover:bg-white hover:shadow-sm rounded-lg sm:rounded-xl text-slate-400 transition-all active:scale-90"
                          >
                            <Minus className="w-4 h-4 sm:w-5 sm:h-5" />
                          </button>
                          <span className="w-10 sm:w-12 text-center text-sm sm:text-base font-black text-slate-800">{item.jumlah}</span>
                          <button 
                            onClick={() => updateCartQuantity(item.id, 1)}
                            className="p-1.5 sm:p-2 hover:bg-white hover:shadow-sm rounded-lg sm:rounded-xl text-blue-600 transition-all active:scale-90"
                          >
                            <Plus className="w-4 h-4 sm:w-5 sm:h-5" />
                          </button>
                        </div>
                        <button 
                          onClick={() => removeFromCart(item.id)}
                          className="w-10 h-10 sm:w-12 sm:h-12 flex items-center justify-center bg-rose-50 text-rose-500 rounded-xl sm:rounded-2xl opacity-100 transition-all hover:bg-rose-100 active:scale-90"
                        >
                          <Trash2 className="w-6 h-6" />
                        </button>
                      </div>
                    </motion.div>
                  ))}
                </AnimatePresence>
              </div>

              <div className="p-4 sm:p-6 bg-slate-50 border-t border-slate-100 space-y-4 pb-10">
                <div className="space-y-3">
                  <label className="text-[10px] uppercase font-black text-slate-400 tracking-[0.2em] block">Sistem Pembayaran</label>
                  <div className="relative">
                    <select 
                      value={selectedMethod}
                      onChange={(e) => setSelectedMethod(e.target.value as any)}
                      className="w-full bg-white border border-slate-200 rounded-xl py-3 px-4 text-xs font-black text-slate-700 outline-none focus:border-blue-500/30 appearance-none shadow-sm"
                    >
                      {['Tunai', 'QRIS', 'M-Banking', 'DANA', 'OVO', 'Hutang'].map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </select>
                    <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none">
                      <ChevronDown className="w-4 h-4 text-slate-400" />
                    </div>
                  </div>
                </div>

                {selectedMethod === 'Hutang' && (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="space-y-2"
                  >
                    <input 
                      type="text" 
                      placeholder="Identitas Penghutang..." 
                      value={debtorName}
                      onChange={(e) => setDebtorName(e.target.value)}
                      className="w-full bg-white border border-slate-200 rounded-xl py-3 px-4 text-xs font-black outline-none focus:border-rose-500/30 transition-all shadow-sm"
                    />
                  </motion.div>
                )}

                <div className="flex flex-col gap-4">
                  <div className="flex justify-between items-end border-b border-dashed border-slate-200 pb-4">
                    <div>
                      <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Total</p>
                      <motion.p 
                        key={totalBelanja}
                        initial={{ scale: 1.1, y: 5 }}
                        animate={{ scale: 1, y: 0 }}
                        className="text-2xl font-display font-black text-slate-900 tracking-tighter italic"
                      >
                        Rp {totalBelanja.toLocaleString()}
                      </motion.p>
                    </div>
                    <div className="text-right pb-1">
                       <p className="text-[9px] font-black text-blue-500 bg-blue-50 px-2 py-1 rounded-md border border-blue-100 uppercase tracking-widest">{selectedMethod}</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => {
                      initiatePayment();
                    }}
                    className="w-full bg-blue-600 text-white py-4 rounded-xl font-black uppercase text-xs tracking-widest shadow-xl shadow-blue-500/20 active:scale-95 transition-all flex items-center justify-center gap-2 group"
                  >
                    Kirim & Bayar
                    <ChevronRight className="w-4 h-4 group-active:translate-x-2 transition-transform" />
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Payment Success Toast */}
      <AnimatePresence>
        {showToast && (
          <motion.div 
            initial={{ y: -100, x: '-50%', opacity: 0 }}
            animate={{ y: 20, x: '-50%', opacity: 1 }}
            exit={{ y: -100, x: '-50%', opacity: 0 }}
            className="fixed top-0 left-1/2 z-[2000] bg-white text-slate-900 px-6 py-4 rounded-[2rem] shadow-2xl border border-slate-100 flex items-center gap-4 font-black"
          >
            <div className="w-10 h-10 bg-emerald-500 text-white rounded-2xl flex items-center justify-center shadow-lg shadow-emerald-200">
              <CheckCircle2 className="w-6 h-6" />
            </div>
            <div className="pr-4">
              <p className="text-sm">Pembayaran Berhasil!</p>
              <p className="text-[10px] text-slate-400 uppercase tracking-widest mt-0.5">Struk siap dicetak</p>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Global Toast Notification */}
      <AnimatePresence>
        {importingInfo && (
          <motion.div
            initial={{ opacity: 0, y: 50, x: '-50%' }}
            animate={{ opacity: 1, y: 0, x: '-50%' }}
            exit={{ opacity: 0, y: 20, x: '-50%' }}
            className="fixed bottom-10 left-1/2 z-[9999] bg-slate-900 text-white px-6 py-3 sm:py-4 rounded-xl sm:rounded-2xl shadow-2xl font-bold flex items-center gap-3 text-xs sm:text-sm"
          >
            <div className="w-2 h-2 sm:w-2.5 sm:h-2.5 bg-emerald-400 rounded-full animate-pulse" />
            {importingInfo}
          </motion.div>
        )}
      </AnimatePresence>    
    </div>
  );
}


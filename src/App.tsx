import { useState, useEffect, useCallback, useRef, useMemo } from 'react';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell, PieChart, Pie, Legend
} from 'recharts';
import * as XLSX from 'xlsx';

// ─────────────────────────────────────────────
// TYPES — matching your real Excel structure
// ─────────────────────────────────────────────
type Role = 'admin' | 'user';
interface User {
  id: string; email: string; name: string; createdAt: string; role: Role;
}
interface StoredUser extends User { passwordHash: string; }

/** Emails qui sont automatiquement administrateurs. */
const ADMIN_EMAILS = ['admin@mrp.com'];

/** PDP = Atelier / Chaine / Marque / Produit-fini / S23-S27 / Total */
interface PDPRow {
  id: string;
  atelier: string;
  chaine: string;
  marque: string;
  produitFini: string;
  s23: number; s24: number; s25: number; s26: number; s27: number;
}

/** BOM = Code P-F / Code ERP / Désignation / Quantité */
interface BOMRow {
  id: string;
  codePF: string;
  codeERP: string;
  designation: string;
  quantite: number;
}

/** Stock = Code ERP / Designation / Stock d'aujourd'hui */
interface StockRow {
  id: string;
  codeERP: string;
  designation: string;
  stock: number;
}

/** Computed surplus per codeERP */
interface SurplusRow {
  id: string;
  codeERP: string;
  codePF: string;        // Code(s) produit-fini liés (depuis BOM)
  designation: string;
  stock: number;
  besoinsNet: { s23: number; s24: number; s25: number; s26: number; s27: number };
  besoinTotal: number;
  surplusStock: number;
}

interface KitComponent {
  codeERP: string;
  designation: string;
  quantite: number;        // Qté par kit (coefficient BOM)
  stock: number;           // Stock total du composant
  besoinPDP: number;       // Besoin total consommé par le PDP (SOMME.SI)
  surplus: number;         // Restant = Stock - Besoin PDP
  kitsPossibles: number;   // ENT(surplus / quantite)
  manquant: number;        // Qté à ajouter pour atteindre le potentiel max
  estLimitant: boolean;    // composant qui bloque la production
}

interface KitRow {
  id: string;
  atelier: string;
  chaine: string;
  marque: string;
  produitFini: string;
  maxKits: number;         // kits productibles MAINTENANT avec le surplus
  potentiel: number;       // kits productibles si on complète les composants manquants
  status: 'MANQUE' | 'BOM MANQUANTE';
  composantLimitant: string;
  components: KitComponent[];
}

// ─────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 9);
const hashPwd = (p: string) => btoa(p + ':MRP2024');
const verifyPwd = (p: string, h: string) => hashPwd(p) === h;
const fmt = (n: number) => n.toLocaleString('fr-FR');
/** Robust number parser: handles spaces, non-breaking spaces, French comma decimals, €, etc. */
const parseNum = (raw: string | undefined): number => {
  if (raw === undefined || raw === null) return 0;
  let s = String(raw).trim();
  if (!s) return 0;
  // remove all kinds of spaces (incl. non-breaking / thin) + apostrophe thousands sep + currency/percent
  s = s.replace(/[\s\u00A0\u202F\u2009']/g, '').replace(/[€$%]/g, '');
  if (s.includes('.') && s.includes(',')) {
    // '.' = thousands, ',' = decimal
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};

// ─────────────────────────────────────────────
// STORAGE
// ─────────────────────────────────────────────
const USERS_KEY = 'mrp_users';
const SESSION_KEY = 'mrp_session';
const dataKey = (u: string) => `mrp_data_${u}`;

const getUsers = (): StoredUser[] => JSON.parse(localStorage.getItem(USERS_KEY) || '[]');
const saveUsers = (u: StoredUser[]) => localStorage.setItem(USERS_KEY, JSON.stringify(u));
const getSession = (): User | null => {
  const s = localStorage.getItem(SESSION_KEY);
  if (!s) return null;
  try {
    const d = JSON.parse(s);
    if (d.expires <= Date.now()) return null;
    const u = d.user as User;
    // Compatibilité : anciens comptes sans rôle
    if (!u.role) u.role = ADMIN_EMAILS.includes(u.email.toLowerCase()) ? 'admin' : 'user';
    return u;
  }
  catch { return null; }
};
const saveSession = (u: User) => localStorage.setItem(SESSION_KEY, JSON.stringify({ user: u, expires: Date.now() + 86400000 * 7 }));
const clearSession = () => localStorage.removeItem(SESSION_KEY);

interface UserData { pdp: PDPRow[]; bom: BOMRow[]; stock: StockRow[]; }
const loadData = (uid: string): UserData | null => {
  const s = localStorage.getItem(dataKey(uid));
  return s ? JSON.parse(s) : null;
};
const saveData = (uid: string, d: UserData) => localStorage.setItem(dataKey(uid), JSON.stringify(d));

// ─────────────────────────────────────────────
// COMPTES PRÉ-ENREGISTRÉS
// ─────────────────────────────────────────────
// Admin  : admin@mrp.com / admin123
// User   : user@mrp.com  / user1234
const seedDefaultUsers = () => {
  const users = getUsers();
  if (users.length > 0) return;
  const defaults: StoredUser[] = [
    { id: 'admin001', email: 'admin@mrp.com', name: 'Administrateur', createdAt: new Date().toISOString(), passwordHash: hashPwd('admin123'), role: 'admin' },
    { id: 'user001',  email: 'user@mrp.com',  name: 'Utilisateur',   createdAt: new Date().toISOString(), passwordHash: hashPwd('user1234'), role: 'user' },
  ];
  saveUsers(defaults);
};

// ─────────────────────────────────────────────
// DEMO DATA
// ─────────────────────────────────────────────
const demoPDP = (): PDPRow[] => [
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'D071', s23:141200, s24:152000, s25:190000, s26:152000, s27:190000 },
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'D080', s23:0, s24:600, s25:0, s26:0, s27:0 },
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'D081', s23:0, s24:260, s25:0, s26:0, s27:0 },
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'D082', s23:0, s24:0, s25:0, s26:3980, s27:0 },
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'D083', s23:0, s24:0, s25:0, s26:0, s27:0 },
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'FA02', s23:0, s24:0, s25:15000, s26:15000, s27:27245 },
  { id: uid(), atelier:'A1', chaine:'C1', marque:'M1', produitFini:'FA04', s23:0, s24:0, s25:0, s26:0, s27:16525 },
  { id: uid(), atelier:'A2', chaine:'C2', marque:'M2', produitFini:'FE26', s23:0, s24:1197, s25:0, s26:0, s27:0 },
  { id: uid(), atelier:'A2', chaine:'C2', marque:'M2', produitFini:'FEV26', s23:0, s24:0, s25:0, s26:471, s27:0 },
  { id: uid(), atelier:'A2', chaine:'C2', marque:'M2', produitFini:'DS:BP', s23:0, s24:0, s25:0, s26:3100, s27:0 },
  { id: uid(), atelier:'A2', chaine:'C2', marque:'M2', produitFini:'DS:DA', s23:5000, s24:5000, s25:0, s26:3000, s27:0 },
  { id: uid(), atelier:'A2', chaine:'C2', marque:'M2', produitFini:'DS:PT', s23:500, s24:20000, s25:0, s26:0, s27:0 },
];

const demoBOM = (): BOMRow[] => [
  { id: uid(), codePF:'D071', codeERP:'SP050-12117', designation:'Corps Fiche Mâle 16A D071', quantite:1 },
  { id: uid(), codePF:'D080', codeERP:'SP050-12117', designation:'Corps Fiche Mâle 16A D071', quantite:1 },
  { id: uid(), codePF:'FA02', codeERP:'SP214-79', designation:'Fond + Couvercle 02 Départs APPARENT', quantite:1 },
  { id: uid(), codePF:'FA04', codeERP:'SP194-57', designation:'Fond + Couvercle 04 Départs APPARENT', quantite:1 },
  { id: uid(), codePF:'FE26', codeERP:'SP239-94', designation:'Fond TDB ENCASTRÉ 26 dép.', quantite:1 },
  { id: uid(), codePF:'DS:BP', codeERP:'SPP011-326', designation:'Cage Interrupteur D/L', quantite:1 },
  { id: uid(), codePF:'DS:PT', codeERP:'SPP005-328', designation:'Porte Mécanisme Prise D/L', quantite:1 },
];

const demoStock = (): StockRow[] => [
  { id: uid(), codeERP:'SP050-12117', designation:'Corps Fiche Mâle 16A D071', stock:10800 },
  { id: uid(), codeERP:'SP214-79', designation:'Fond + Couvercle 02 Départs APPARENT', stock:1320 },
  { id: uid(), codeERP:'SP194-57', designation:'Fond + Couvercle 04 Départs APPARENT', stock:1120 },
  { id: uid(), codeERP:'SP239-94', designation:'Fond TDB ENCASTRÉ 26 dép.', stock:203 },
  { id: uid(), codeERP:'SP237-93.1', designation:'Couvercle TDB 26 départs', stock:335 },
  { id: uid(), codeERP:'SPP011-326', designation:'Cage Interrupteur D/L', stock:45000 },
  { id: uid(), codeERP:'SPP005-328', designation:'Porte Mécanisme Prise D/L', stock:19500 },
];

// ═════════════════════════════════════════════════════════════════
// 🧮 MOTEUR DE CALCUL — SURPLUS
// ═════════════════════════════════════════════════════════════════
// FORMULES (par Code ERP, regroupées avec une logique SOMME.SI) :
//
//   Stock total        = SOMME.SI(stock ; même Code ERP)                  ← on additionne
//                        toutes les lignes de stock ayant le même Code ERP (pas de doublon)
//
//   Besoin PDP (Sxx)   = SOMME pour chaque produit-fini qui utilise ce composant
//                        de ( quantité PDP du produit en semaine Sxx  ×  Quantité BOM )
//
//   Besoin total       = Besoin S23 + S24 + S25 + S26 + S27
//
//   SURPLUS STOCK      = Stock total − Besoin total PDP
//                        (= ce qui RESTE après avoir fabriqué tout le PDP)
//
const computeSurplus = (pdp: PDPRow[], bom: BOMRow[], stock: StockRow[]): SurplusRow[] => {
  // 1 ligne unique par Code ERP (regroupement / dédoublonnage)
  const erps = [...new Set(stock.map(s => s.codeERP).filter(Boolean))];

  return erps.map(erp => {
    // SOMME.SI sur le stock : on additionne tous les stocks du même Code ERP
    const stockRows = stock.filter(s => s.codeERP === erp);
    const currentStock = stockRows.reduce((a, s) => a + s.stock, 0);
    const designation = stockRows[0]?.designation ?? '';

    const bomRows = bom.filter(b => b.codeERP === erp);
    const codePF = [...new Set(bomRows.map(b => b.codePF).filter(Boolean))].join(', ');

    // Besoin PDP par semaine = SOMME ( PDP produit-fini × coefficient BOM )
    let s23 = 0, s24 = 0, s25 = 0, s26 = 0, s27 = 0;
    bomRows.forEach(b => {
      pdp.filter(p => p.produitFini === b.codePF).forEach(p => {
        s23 += p.s23 * b.quantite;
        s24 += p.s24 * b.quantite;
        s25 += p.s25 * b.quantite;
        s26 += p.s26 * b.quantite;
        s27 += p.s27 * b.quantite;
      });
    });
    const besoinTotal = s23 + s24 + s25 + s26 + s27;

    return {
      id: erp, codeERP: erp,
      codePF: codePF || '-',
      designation,
      stock: currentStock,
      besoinsNet: { s23, s24, s25, s26, s27 },
      besoinTotal,
      surplusStock: currentStock - besoinTotal,   // ← SURPLUS = Stock − Besoin PDP
    };
  });
};

// ═════════════════════════════════════════════════════════════════
// 🧩 MOTEUR DE CALCUL — KITS (production opportuniste avec le SURPLUS)
// ═════════════════════════════════════════════════════════════════
// IDÉE : avec le SURPLUS de stock (ce qui reste après le PDP), on propose
// les produits-finis qu'on peut encore fabriquer.
//
// FORMULES (pour chaque composant d'un produit-fini) :
//
//   Stock            = stock total du composant
//   Besoin PDP       = besoin total consommé par le PDP (SOMME.SI, identique à Surplus)
//   Surplus (restant)= Stock − Besoin PDP                ← ce qui reste pour les kits
//   Qté / kit        = coefficient BOM (quantité du composant dans 1 kit)
//   Kits possibles   = ENT( Surplus ÷ Qté/kit )          ← combien de kits ce composant permet
//
//   Kits productibles du produit (maxKits) = MIN(Kits possibles de tous les composants)
//   Potentiel max                          = MAX(Kits possibles de tous les composants)
//   À ajouter (par composant) = MAX(0 ; Potentiel × Qté/kit − Surplus)
//
const computeKits = (pdp: PDPRow[], bom: BOMRow[], stock: StockRow[]): KitRow[] => {
  // On réutilise EXACTEMENT le surplus calculé dans l'onglet Surplus (pas de recalcul divergent)
  const surplusList = computeSurplus(pdp, bom, stock);
  const surplusByErp = new Map(surplusList.map(s => [s.codeERP, s]));

  const produits = [...new Set(pdp.map(p => p.produitFini).filter(Boolean))];

  return produits.map(produitFini => {
    const pdpRows = pdp.filter(p => p.produitFini === produitFini);
    const firstPdp = pdpRows[0];
    const kitBom = bom.filter(b => b.codePF === produitFini);

    // Pré-calcul par composant
    const base = kitBom.map(b => {
      const sr = surplusByErp.get(b.codeERP);
      const stockTotal = sr?.stock ?? 0;
      const besoinPDP = sr?.besoinTotal ?? 0;
      const surplus = sr ? sr.surplusStock : stockTotal;   // restant = stock − besoin PDP
      const kitsPossibles = b.quantite > 0 ? Math.floor(surplus / b.quantite) : 0;
      return { b, stockTotal, besoinPDP, surplus, kitsPossibles };
    });

    // Kits productibles maintenant = composant le plus contraignant (MIN)
    const maxKits = base.length ? Math.max(0, Math.min(...base.map(x => x.kitsPossibles))) : 0;
    // Potentiel = si on complétait les composants manquants (MAX)
    const potentiel = base.length ? Math.max(0, Math.max(...base.map(x => x.kitsPossibles))) : 0;
    // Composant limitant = celui avec le moins de kits possibles
    const limitant = base.length ? [...base].sort((a, b) => a.kitsPossibles - b.kitsPossibles)[0] : null;

    const components: KitComponent[] = base.map(x => ({
      codeERP: x.b.codeERP,
      designation: x.b.designation || surplusByErp.get(x.b.codeERP)?.designation || '',
      quantite: x.b.quantite,
      stock: x.stockTotal,
      besoinPDP: x.besoinPDP,
      surplus: x.surplus,
      kitsPossibles: x.kitsPossibles,
      manquant: Math.max(0, potentiel * x.b.quantite - x.surplus),
      estLimitant: limitant ? x.b.codeERP === limitant.b.codeERP : false,
    }));

    // Statut : BOM MANQUANTE si pas de nomenclature, sinon MANQUE si un composant bloque
    const hasMissing = components.some(c => c.surplus < c.quantite);
    const status: KitRow['status'] = components.length === 0 ? 'BOM MANQUANTE' : 'MANQUE';

    return {
      id: produitFini,
      atelier: firstPdp?.atelier ?? '',
      chaine: firstPdp?.chaine ?? '',
      marque: firstPdp?.marque ?? '',
      produitFini,
      maxKits,
      potentiel,
      status,
      composantLimitant: limitant?.b.codeERP ?? '-',
      components,
      _show: components.length === 0 || hasMissing,
    } as KitRow & { _show: boolean };
  })
  // On n'affiche que les produits à compléter (manque composant ou BOM absente)
  .filter(k => (k as KitRow & { _show: boolean })._show)
  .map(({ ...k }) => { delete (k as Record<string, unknown>)._show; return k as KitRow; });
};

// ═════════════════════════════════════════════════════════════════
// 📥 EXPORT EXCEL (CSV) & PDF (impression)
// ═════════════════════════════════════════════════════════════════
const downloadExcel = (filename: string, headers: string[], rows: (string | number)[][]) => {
  const data = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(data);
  // Largeur auto des colonnes
  ws['!cols'] = headers.map((h, i) => {
    const maxLen = Math.max(h.length, ...rows.map(r => String(r[i] ?? '').length));
    return { wch: Math.min(maxLen + 2, 40) };
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Données');
  XLSX.writeFile(wb, filename);
};

// ─────────────────────────────────────────────
// CSV IMPORT
// ─────────────────────────────────────────────
const detectSep = (line: string) => line.includes('\t') ? '\t' : line.includes(';') ? ';' : ',';

const parsePDPCsv = (text: string): PDPRow[] => {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const sep = detectSep(lines[0]);
  const rows: PDPRow[] = [];
  for (let i = 1; i < lines.length; i++) {
    const c = lines[i].split(sep).map(x => x.trim().replace(/^"|"$/g, ''));
    if (c.length < 5) continue;
    // skip if it's an empty/total row with no produit-fini
    if (!c[3] && !c[0] && !c[1]) continue;
    rows.push({
      id: uid(), atelier: c[0]||'', chaine: c[1]||'', marque: c[2]||'', produitFini: c[3]||'',
      s23: parseNum(c[4]), s24: parseNum(c[5]), s25: parseNum(c[6]), s26: parseNum(c[7]), s27: parseNum(c[8]),
    });
  }
  return rows;
};

const parseBOMCsv = (text: string): BOMRow[] => {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const sep = detectSep(lines[0]);
  return lines.slice(1).map(line => {
    const c = line.split(sep).map(x => x.trim().replace(/^"|"$/g, ''));
    return { id: uid(), codePF: c[0]||'', codeERP: c[1]||'', designation: c[2]||'', quantite: parseNum(c[3])||1 };
  }).filter(r => r.codeERP);
};

const parseStockCsv = (text: string): StockRow[] => {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const sep = detectSep(lines[0]);
  return lines.slice(1).map(line => {
    const c = line.split(sep).map(x => x.trim().replace(/^"|"$/g, ''));
    return { id: uid(), codeERP: c[0]||'', designation: c[1]||'', stock: parseNum(c[2]) };
  }).filter(r => r.codeERP);
};

// ─────────────────────────────────────────────
// AUTH PAGE
// ─────────────────────────────────────────────
const AuthPage: React.FC<{ onAuth: (u: User) => void }> = ({ onAuth }) => {
  const [tab, setTab] = useState<'login'|'register'>('login');
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [pwd, setPwd] = useState('');
  const [pwd2, setPwd2] = useState('');
  const [err, setErr] = useState('');
  const [loading, setLoading] = useState(false);

  const doLogin = async (e: React.FormEvent) => {
    e.preventDefault(); setErr(''); setLoading(true);
    await new Promise(r => setTimeout(r, 400));
    const users = getUsers();
    const found = users.find(u => u.email.toLowerCase() === email.toLowerCase() && verifyPwd(pwd, u.passwordHash));
    if (found) {
      const role: Role = found.role ?? (ADMIN_EMAILS.includes(found.email.toLowerCase()) ? 'admin' : 'user');
      const u: User = { id: found.id, email: found.email, name: found.name, createdAt: found.createdAt, role };
      saveSession(u);
      onAuth(u);
    } else { setErr('Email ou mot de passe incorrect'); }
    setLoading(false);
  };

  const doRegister = async (e: React.FormEvent) => {
    e.preventDefault(); setErr('');
    if (!name.trim()) { setErr('Veuillez saisir votre nom'); return; }
    if (pwd.length < 6) { setErr('Mot de passe : 6 caractères minimum'); return; }
    if (pwd !== pwd2) { setErr('Les mots de passe ne correspondent pas'); return; }
    setLoading(true); await new Promise(r => setTimeout(r, 400));
    const users = getUsers();
    if (users.find(u => u.email.toLowerCase() === email.toLowerCase())) {
      setErr('Cet email est déjà utilisé'); setLoading(false); return;
    }
    // Le premier compte créé OU les emails de la liste ADMIN_EMAILS deviennent admin
    const isFirstUser = users.length === 0;
    const role: Role = (isFirstUser || ADMIN_EMAILS.includes(email.toLowerCase())) ? 'admin' : 'user';
    const nu: StoredUser = { id: uid(), email, name, createdAt: new Date().toISOString(), passwordHash: hashPwd(pwd), role };
    saveUsers([...users, nu]);
    const u: User = { id: nu.id, email: nu.email, name: nu.name, createdAt: nu.createdAt, role };
    saveSession(u); onAuth(u); setLoading(false);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-800 via-blue-900 to-indigo-900 flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-20 h-20 bg-white/10 rounded-2xl mb-4 backdrop-blur"><span className="text-4xl">🏭</span></div>
          <h1 className="text-3xl font-bold text-white">MRP</h1>
          <p className="text-blue-200 mt-1">Gestion du Surstock Semi-Fini Plastique</p>
        </div>
        <div className="bg-white rounded-2xl shadow-2xl overflow-hidden">
          <div className="flex border-b border-gray-100">
            <button onClick={() => { setTab('login'); setErr(''); }}
              className={`flex-1 py-4 text-sm font-semibold transition-colors ${tab==='login' ? 'bg-blue-600 text-white' : 'text-gray-500 hover:bg-gray-50'}`}>Se connecter</button>
            <button onClick={() => { setTab('register'); setErr(''); }}
              className={`flex-1 py-4 text-sm font-semibold transition-colors ${tab==='register' ? 'bg-blue-600 text-white' : 'text-gray-500 hover:bg-gray-50'}`}>Créer un compte</button>
          </div>
          <div className="p-8">
            {err && <div className="mb-4 px-4 py-3 bg-red-50 border border-red-200 text-red-600 rounded-lg text-sm">⚠️ {err}</div>}
            {tab === 'login' ? (
              <form onSubmit={doLogin} className="space-y-4">
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Email</label>
                  <input type="email" value={email} onChange={e => setEmail(e.target.value)} required className="w-full px-4 py-3 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" placeholder="INSERER VOTRE MAIL" /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Mot de passe</label>
                  <input type="password" value={pwd} onChange={e => setPwd(e.target.value)} required className="w-full px-4 py-3 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" placeholder="••••••••" /></div>
                <button type="submit" disabled={loading} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 rounded-xl font-semibold transition-colors disabled:opacity-50">
                  {loading ? ' Connexion…' : 'Se connecter →'}</button>
              </form>
            ) : (
              <form onSubmit={doRegister} className="space-y-4">
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Nom complet</label>
                  <input type="text" value={name} onChange={e => setName(e.target.value)} required className="w-full px-4 py-3 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" placeholder="INSERER VOTRE NOM COMPLET" /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Email</label>
                  <input type="email" value={email} onChange={e => setEmail(e.target.value)} required className="w-full px-4 py-3 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" placeholder="INSERER VOTRE MAIL" /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Mot de passe</label>
                  <input type="password" value={pwd} onChange={e => setPwd(e.target.value)} required className="w-full px-4 py-3 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" placeholder="6 caractères minimum" /></div>
                <div><label className="block text-sm font-medium text-gray-700 mb-1">Confirmer</label>
                  <input type="password" value={pwd2} onChange={e => setPwd2(e.target.value)} required className="w-full px-4 py-3 border border-gray-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" placeholder="••••••••" /></div>
                <button type="submit" disabled={loading} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 rounded-xl font-semibold transition-colors disabled:opacity-50">
                  {loading ? '⏳ Création…' : 'Créer mon compte →'}</button>
              </form>
            )}
            <p className="text-center text-xs text-gray-400 mt-6">🔒 Données privées • Espace isolé par utilisateur</p>
          </div>
        </div>
      </div>
    </div>
  );
};

// ────────────────────────────────────────────
// IMPORT MODAL
// ─────────────────────────────────────────────
const ImportModal: React.FC<{ type: 'PDP'|'BOM'|'STOCK'; onImport: (r: PDPRow[]|BOMRow[]|StockRow[]) => void; onClose: () => void }> = ({ type, onImport, onClose }) => {
  const [text, setText] = useState('');
  const [preview, setPreview] = useState<string[][]>([]);
  const fileRef = useRef<HTMLInputElement>(null);
  const headers: Record<string, string[]> = {
    PDP: ['Atelier','Chaine','Marque','Produit-fini','S23','S24','S25','S26','S27'],
    BOM: ['Code P-F','Code ERP','Désignation','Quantité'],
    STOCK: ['Code ERP','Designation','Stock d\'aujourd\'hui'],
  };
  const handleFile = (f: File) => {
    const r = new FileReader();
    r.onload = e => {
      const t = e.target?.result as string; setText(t);
      const firstLine = t.trim().split(/\r?\n/)[0] || '';
      const sep = firstLine.includes('\t') ? '\t' : firstLine.includes(';') ? ';' : ',';
      setPreview(t.trim().split(/\r?\n/).slice(0,6).map(l => l.split(sep).map(x => x.trim().replace(/^"|"$/g, ''))));
    };
    r.readAsText(f, 'UTF-8');
  };
  const doImport = () => {
    if (!text) return;
    if (type === 'PDP') onImport(parsePDPCsv(text));
    if (type === 'BOM') onImport(parseBOMCsv(text));
    if (type === 'STOCK') onImport(parseStockCsv(text));
    onClose();
  };
  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
        <div className="p-6 border-b border-gray-100 flex justify-between items-center">
          <div><h2 className="text-lg font-bold text-gray-800">📥 Importer {type}</h2><p className="text-sm text-gray-500 mt-1">Depuis Excel → Enregistrer en CSV</p></div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-600 text-2xl">×</button>
        </div>
        <div className="p-6 space-y-4">
          <div className="bg-blue-50 rounded-xl p-4">
            <p className="text-sm font-semibold text-blue-800 mb-2">📋 Colonnes attendues :</p>
            <div className="flex gap-2 flex-wrap">{headers[type].map((h, i) => <span key={i} className="px-2 py-1 bg-blue-200 text-blue-800 rounded text-xs font-mono">{h}</span>)}</div>
          </div>
          <div className="border-2 border-dashed border-gray-300 rounded-xl p-8 text-center cursor-pointer hover:border-blue-400 hover:bg-blue-50 transition-colors"
            onClick={() => fileRef.current?.click()} onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }}>
            <span className="text-4xl"></span><p className="text-gray-600 mt-2 font-medium">Glissez votre CSV ici</p><p className="text-gray-400 text-sm">ou cliquez pour parcourir</p>
            <input ref={fileRef} type="file" accept=".csv,.txt" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
          </div>
          <div><p className="text-sm font-medium text-gray-700 mb-1">Ou collez le CSV :</p>
            <textarea value={text} onChange={e => setText(e.target.value)} className="w-full h-28 px-3 py-2 border border-gray-200 rounded-lg text-xs font-mono focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder={headers[type].join(';')} /></div>
          {preview.length > 0 && <div><p className="text-sm font-medium text-gray-700 mb-1">Aperçu :</p>
            <div className="overflow-x-auto border border-gray-200 rounded-lg"><table className="text-xs w-full"><tbody>
              {preview.map((row, i) => <tr key={i} className={i===0?'bg-gray-100 font-semibold':'border-t border-gray-100'}>{row.slice(0,8).map((c,j)=><td key={j} className="px-2 py-1 truncate max-w-[100px]">{c}</td>)}</tr>)}
            </tbody></table></div></div>}
          <div className="flex gap-3">
            <button onClick={onClose} className="flex-1 py-3 border border-gray-200 text-gray-600 rounded-xl font-medium hover:bg-gray-50">Annuler</button>
            <button onClick={doImport} disabled={!text} className="flex-1 py-3 bg-blue-600 text-white rounded-xl font-semibold hover:bg-blue-700 disabled:opacity-40">✅ Importer</button>
          </div>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// PDP TAB
// ─────────────────────────────────────────────
const PDPTab: React.FC<{ data: PDPRow[]; onChange: (d: PDPRow[]) => void; onImport: () => void }> = ({ data, onChange, onImport }) => {
  const add = () => onChange([...data, { id: uid(), atelier:'', chaine:'', marque:'', produitFini:'', s23:0, s24:0, s25:0, s26:0, s27:0 }]);
  const del = (id: string) => onChange(data.filter(r => r.id !== id));
  const upd = (id: string, f: keyof PDPRow, v: string | number) => {
    onChange(data.map(r => r.id === id ? { ...r, [f]: typeof v==='string' && f!=='atelier' && f!=='chaine' && f!=='marque' && f!=='produitFini' ? Number(v) : v } : r));
  };
  const weeks = [{ k:'s23', l:'S23' },{ k:'s24', l:'S24' },{ k:'s25', l:'S25' },{ k:'s26', l:'S26' },{ k:'s27', l:'S27' }] as const;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-3 justify-between items-center">
        <div><h2 className="text-lg font-bold text-gray-800">📊 Plan Directeur de Production</h2><p className="text-sm text-gray-500">{data.length} lignes</p></div>
        <div className="flex gap-2">
          <button onClick={onImport} className="px-4 py-2 bg-violet-600 text-white rounded-xl text-sm font-semibold hover:bg-violet-700">📥 Importer CSV</button>
          <button onClick={add} className="px-4 py-2 bg-blue-600 text-white rounded-xl text-sm font-semibold hover:bg-blue-700">+ Ajouter</button>
        </div>
      </div>
      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-800 text-white">
              <tr>
                <th className="px-3 py-3 text-left text-xs font-semibold">Atelier</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Chaine</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Marque</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Produit-fini</th>
                {weeks.map(w => <th key={w.k} className="px-3 py-3 text-center text-xs font-semibold bg-blue-700">{w.l}</th>)}
                <th className="px-3 py-3 text-center text-xs font-semibold bg-slate-700">Total</th>
                <th className="px-3 py-3 w-8"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {data.map(row => (
                <tr key={row.id} className="hover:bg-blue-50">
                  <td className="px-3 py-2"><input value={row.atelier} onChange={e => upd(row.id,'atelier',e.target.value)} className="w-20 bg-transparent border border-gray-200 rounded px-1 py-0.5 text-xs focus:outline-none focus:border-blue-400" /></td>
                  <td className="px-3 py-2"><input value={row.chaine} onChange={e => upd(row.id,'chaine',e.target.value)} className="w-20 bg-transparent border border-gray-200 rounded px-1 py-0.5 text-xs focus:outline-none focus:border-blue-400" /></td>
                  <td className="px-3 py-2"><input value={row.marque} onChange={e => upd(row.id,'marque',e.target.value)} className="w-20 bg-transparent border border-gray-200 rounded px-1 py-0.5 text-xs focus:outline-none focus:border-blue-400" /></td>
                  <td className="px-3 py-2"><input value={row.produitFini} onChange={e => upd(row.id,'produitFini',e.target.value)} className="w-24 bg-transparent border border-gray-200 rounded px-1 py-0.5 text-xs focus:outline-none focus:border-blue-400" /></td>
                  {weeks.map(w => (
                    <td key={w.k} className="px-2 py-2 bg-blue-50"><input type="number" value={row[w.k]} onChange={e => upd(row.id, w.k, e.target.value)}
                      className="w-20 bg-white border border-blue-200 rounded px-1 py-0.5 text-xs text-right focus:outline-none focus:border-blue-500 font-mono" /></td>
                  ))}
                  <td className="px-3 py-2 text-center font-mono font-bold text-slate-700">{fmt(row.s23+row.s24+row.s25+row.s26+row.s27)}</td>
                  <td className="px-2 py-2 text-center"><button onClick={() => del(row.id)} className="text-red-400 hover:text-red-600">✕</button></td>
                </tr>
              ))}
              {data.length===0 && <tr><td colSpan={11} className="text-center py-12 text-gray-400">Importez votre PDP depuis Excel</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// BOM TAB
// ─────────────────────────────────────────────
const BOMTab: React.FC<{ data: BOMRow[]; onChange: (d: BOMRow[]) => void; onImport: () => void; isAdmin: boolean }> = ({ data, onChange, onImport, isAdmin }) => {
  const [search, setSearch] = useState('');
  const add = () => onChange([...data, { id: uid(), codePF:'', codeERP:'', designation:'', quantite:1 }]);
  const del = (id: string) => onChange(data.filter(r => r.id !== id));
  const upd = (id: string, f: keyof BOMRow, v: string | number) => onChange(data.map(r => r.id === id ? { ...r, [f]: f==='quantite' ? Number(v) : v } : r));

  const q = search.toLowerCase();
  const filtered = data.filter(r =>
    !q ||
    r.codePF.toLowerCase().includes(q) ||
    r.codeERP.toLowerCase().includes(q) ||
    r.designation.toLowerCase().includes(q)
  );

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-3 justify-between items-center">
        <div>
          <h2 className="text-lg font-bold text-gray-800">🗂️ Nomenclature (BOM)</h2>
          <p className="text-sm text-gray-500">{data.length} lignes {!isAdmin && <span className="ml-2 px-2 py-0.5 bg-gray-100 text-gray-500 rounded-full text-xs font-medium">🔒 Lecture seule</span>}</p>
        </div>
        <div className="flex gap-2 flex-wrap">
          {isAdmin && (
            <>
              <button onClick={onImport} className="px-4 py-2 bg-violet-600 text-white rounded-xl text-sm font-semibold hover:bg-violet-700">📥 Importer CSV</button>
              <button onClick={add} className="px-4 py-2 bg-blue-600 text-white rounded-xl text-sm font-semibold hover:bg-blue-700">+ Ajouter</button>
            </>
          )}
        </div>
      </div>
      {!isAdmin && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-4 py-3 text-sm text-amber-700 flex items-center gap-2">
          <span>🔒</span> Seul l'administrateur peut modifier ou importer la nomenclature (BOM).
        </div>
      )}
      {/* Barre de recherche — disponible pour tout le monde */}
      <input
        value={search}
        onChange={e => setSearch(e.target.value)}
        className="w-full max-w-lg px-4 py-2 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        placeholder="🔍 Rechercher par Code P-F, Code ERP ou Désignation…"
      />
      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-slate-800 text-white">
            <tr>
              <th className="px-4 py-3 text-left text-xs font-semibold">Code P-F</th>
              <th className="px-4 py-3 text-left text-xs font-semibold">Code ERP</th>
              <th className="px-4 py-3 text-left text-xs font-semibold">Désignation</th>
              <th className="px-4 py-3 text-center text-xs font-semibold">Quantité</th>
              {isAdmin && <th className="px-4 py-3 w-8"></th>}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filtered.map(row => (
              <tr key={row.id} className="hover:bg-gray-50">
                {isAdmin ? (
                  <>
                    <td className="px-4 py-2"><input value={row.codePF} onChange={e => upd(row.id,'codePF',e.target.value)} className="w-24 bg-transparent border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-blue-400" /></td>
                    <td className="px-4 py-2"><input value={row.codeERP} onChange={e => upd(row.id,'codeERP',e.target.value)} className="w-32 bg-transparent border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-blue-400" /></td>
                    <td className="px-4 py-2"><input value={row.designation} onChange={e => upd(row.id,'designation',e.target.value)} className="w-64 bg-transparent border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-blue-400" /></td>
                    <td className="px-4 py-2 text-center"><input type="number" value={row.quantite} onChange={e => upd(row.id,'quantite',e.target.value)} className="w-16 bg-transparent border border-gray-200 rounded px-2 py-1 text-xs text-center focus:outline-none focus:border-blue-400" min="0" step="0.01" /></td>
                    <td className="px-4 py-2 text-center"><button onClick={() => del(row.id)} className="text-red-400 hover:text-red-600">✕</button></td>
                  </>
                ) : (
                  <>
                    <td className="px-4 py-2 font-mono text-xs">{row.codePF}</td>
                    <td className="px-4 py-2 font-mono text-xs">{row.codeERP}</td>
                    <td className="px-4 py-2 text-xs text-gray-600">{row.designation}</td>
                    <td className="px-4 py-2 text-center font-mono text-xs">{row.quantite}</td>
                  </>
                )}
              </tr>
            ))}
            {filtered.length===0 && <tr><td colSpan={5} className="text-center py-12 text-gray-400">{search ? 'Aucun résultat' : (isAdmin ? 'Importez votre BOM depuis Excel' : 'Aucune nomenclature disponible')}</td></tr>}
          </tbody>
        </table>
        {search && filtered.length > 0 && (
          <div className="px-4 py-2 bg-gray-50 text-xs text-gray-500 border-t">{filtered.length} résultat(s) sur {data.length}</div>
        )}
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// STOCK TAB
// ─────────────────────────────────────────────
const StockTab: React.FC<{ data: StockRow[]; onChange: (d: StockRow[]) => void; onImport: () => void }> = ({ data, onChange, onImport }) => {
  const del = (id: string) => onChange(data.filter(r => r.id !== id));
  const upd = (id: string, f: keyof StockRow, v: string | number) => onChange(data.map(r => r.id === id ? { ...r, [f]: f==='stock' ? Number(v) : v } : r));
  const total = data.reduce((a, r) => a + r.stock, 0);

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-3 justify-between items-center">
        <div>
          <h2 className="text-lg font-bold text-gray-800">📦 Semi Fini Plastique</h2>
          <p className="text-sm text-gray-500">{data.length} réf. • Total: {fmt(total)} pcs <span className="ml-2 px-2 py-0.5 bg-gray-100 text-gray-500 rounded-full text-xs font-medium">🔒 Code & Désignation figés</span></p>
        </div>
        <div className="flex gap-2">
          <button onClick={onImport} className="px-4 py-2 bg-violet-600 text-white rounded-xl text-sm font-semibold hover:bg-violet-700">📥 Importer CSV</button>
        </div>
      </div>
      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-slate-800 text-white">
            <tr><th className="px-4 py-3 text-left text-xs font-semibold">Code ERP</th><th className="px-4 py-3 text-left text-xs font-semibold">Désignation</th><th className="px-4 py-3 text-right text-xs font-semibold">Stock d'aujourd'hui</th><th className="px-4 py-3 w-8"></th></tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {data.map(row => (
              <tr key={row.id} className="hover:bg-gray-50">
                <td className="px-4 py-2 font-mono text-xs text-gray-700 bg-gray-50">{row.codeERP || <span className="text-gray-300">—</span>}</td>
                <td className="px-4 py-2 text-xs text-gray-600 bg-gray-50">{row.designation || <span className="text-gray-300">—</span>}</td>
                <td className="px-4 py-2"><input type="number" value={row.stock} onChange={e => upd(row.id,'stock',e.target.value)} className="w-28 bg-transparent border border-gray-200 rounded px-2 py-1 text-xs text-right focus:outline-none focus:border-blue-400 font-mono" min="0" /></td>
                <td className="px-4 py-2 text-center"><button onClick={() => del(row.id)} className="text-red-400 hover:text-red-600">✕</button></td>
              </tr>
            ))}
            {data.length===0 && <tr><td colSpan={4} className="text-center py-12 text-gray-400">Importez votre stock depuis Excel</td></tr>}
            {data.length>0 && <tr className="bg-gray-50 font-semibold"><td colSpan={2} className="px-4 py-3">TOTAL</td><td className="px-4 py-3 text-right font-mono">{fmt(total)}</td><td></td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// KITS TAB — production opportuniste basée sur le SURPLUS
// ─────────────────────────────────────────────
const STATUS_STYLE: Record<KitRow['status'], string> = {
  'MANQUE': 'bg-red-100 text-red-700',
  'BOM MANQUANTE': 'bg-orange-100 text-orange-700',
};

const KitsTab: React.FC<{ kits: KitRow[] }> = ({ kits }) => {
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'MANQUE' | 'BOM MANQUANTE'>('all');
  const [expanded, setExpanded] = useState<string | null>(null);

  const filtered = kits
    .filter(k => statusFilter === 'all' ? true : k.status === statusFilter)
    .filter(k =>
      !search ||
      k.produitFini.toLowerCase().includes(search.toLowerCase()) ||
      k.atelier.toLowerCase().includes(search.toLowerCase()) ||
      k.chaine.toLowerCase().includes(search.toLowerCase()) ||
      k.marque.toLowerCase().includes(search.toLowerCase())
    );
  const manqueCount = kits.filter(k => k.status === 'MANQUE').length;
  const bomMissingCount = kits.filter(k => k.status === 'BOM MANQUANTE').length;

  const exportExcel = () => {
    const headers = ['Produit-fini','Atelier','Chaine','Marque','Kits productibles','Potentiel max','Composant limitant','Statut','Composant','Désignation','Qté/kit','Stock','Besoin PDP','Surplus','Kits possibles','À ajouter'];
    const rows: (string|number)[][] = [];
    filtered.forEach(k => {
      if (k.components.length === 0) {
        rows.push([k.produitFini, k.atelier, k.chaine, k.marque, k.maxKits, k.potentiel, k.composantLimitant, k.status, '', 'BOM manquante', '', '', '', '', '', '']);
      } else {
        k.components.forEach(c => {
          rows.push([k.produitFini, k.atelier, k.chaine, k.marque, k.maxKits, k.potentiel, k.composantLimitant, k.status, c.codeERP, c.designation, c.quantite, c.stock, c.besoinPDP, c.surplus, c.kitsPossibles, c.manquant]);
        });
      }
    });
    downloadExcel(`kits_${new Date().toISOString().slice(0,10)}.xlsx`, headers, rows);
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-3 justify-between items-center">
        <div>
          <h2 className="text-lg font-bold text-gray-800">🧩 Kits à compléter</h2>
          <p className="text-sm text-gray-500">Produits dont un composant manque ou dont la nomenclature est absente</p>
        </div>
        <div className="flex gap-2">
          <input
            value={search}
            onChange={e => setSearch(e.target.value)}
            className="min-w-[180px] px-4 py-2 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            placeholder="Rechercher un kit..."
          />
          <button onClick={exportExcel} className="px-4 py-2 bg-emerald-600 text-white rounded-xl text-sm font-semibold hover:bg-emerald-700 whitespace-nowrap">📊 Excel</button>
          <button onClick={() => window.print()} className="px-4 py-2 bg-rose-600 text-white rounded-xl text-sm font-semibold hover:bg-rose-700 whitespace-nowrap">🖨️ PDF</button>
        </div>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <div className="bg-red-50 rounded-xl p-4 border border-red-100 shadow-sm">
          <p className="text-xs text-red-700 uppercase font-medium">Manque composant</p>
          <p className="text-2xl font-bold text-red-600 mt-1">{manqueCount}</p>
        </div>
        <div className="bg-orange-50 rounded-xl p-4 border border-orange-100 shadow-sm">
          <p className="text-xs text-orange-700 uppercase font-medium">BOM manquante</p>
          <p className="text-2xl font-bold text-orange-600 mt-1">{bomMissingCount}</p>
        </div>
        <div className="bg-white rounded-xl p-4 border border-gray-100 shadow-sm">
          <p className="text-xs text-gray-500 uppercase font-medium">Total à compléter</p>
          <p className="text-2xl font-bold text-gray-800 mt-1">{kits.length}</p>
        </div>
      </div>

      {/* Filtre par statut */}
      <div className="flex rounded-xl overflow-hidden border border-gray-200 bg-white w-fit">
        {([
          { id: 'all' as const, label: 'Tout' },
          { id: 'MANQUE' as const, label: '❌ Manque' },
          { id: 'BOM MANQUANTE' as const, label: '⚠️ BOM manquante' },
        ]).map(f => (
          <button key={f.id} onClick={() => setStatusFilter(f.id)}
            className={`px-4 py-2 text-sm font-medium transition-colors ${statusFilter===f.id ? 'bg-slate-800 text-white' : 'text-gray-600 hover:bg-gray-50'}`}>
            {f.label}
          </button>
        ))}
      </div>

      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-800 text-white">
              <tr>
                <th className="px-3 py-3 text-left text-xs font-semibold">Produit-fini</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Atelier</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Chaine</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Marque</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-green-800">Kits productibles</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">Potentiel max</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Composant limitant</th>
                <th className="px-3 py-3 text-center text-xs font-semibold">Statut</th>
                <th className="px-3 py-3 w-8"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.length === 0 && <tr><td colSpan={9} className="text-center py-12 text-gray-400">Aucun kit trouvé</td></tr>}
              {filtered.map(kit => {
                const manquants = kit.components.filter(c => c.manquant > 0);
                return (
                <>
                  <tr key={kit.id} className={`transition-colors ${kit.status==='BOM MANQUANTE'?'bg-orange-50 hover:bg-orange-100':'bg-red-50 hover:bg-red-100'}`}>
                    <td className="px-3 py-3 font-mono text-xs font-bold text-gray-800">{kit.produitFini}</td>
                    <td className="px-3 py-3 text-xs text-gray-600">{kit.atelier}</td>
                    <td className="px-3 py-3 text-xs text-gray-600">{kit.chaine}</td>
                    <td className="px-3 py-3 text-xs text-gray-600">{kit.marque}</td>
                    <td className="px-3 py-3 text-right font-mono text-sm font-bold text-green-700 bg-green-50">{fmt(kit.maxKits)}</td>
                    <td className="px-3 py-3 text-right font-mono text-xs text-blue-700">{fmt(kit.potentiel)}</td>
                    <td className="px-3 py-3 text-xs font-mono text-gray-600">{kit.composantLimitant}</td>
                    <td className="px-3 py-3 text-center">
                      <span className={`px-2 py-1 rounded-full text-xs font-bold ${STATUS_STYLE[kit.status]}`}>{kit.status}</span>
                    </td>
                    <td className="px-3 py-3 text-center">
                      <button onClick={() => setExpanded(expanded === kit.id ? null : kit.id)} className="text-blue-600 hover:text-blue-800 text-xs font-bold">
                        {expanded === kit.id ? '−' : '+'}
                      </button>
                    </td>
                  </tr>
                  {expanded === kit.id && (
                    <tr key={`${kit.id}-details`} className="bg-slate-50">
                      <td colSpan={9} className="px-6 py-4">
                        {kit.components.length === 0 ? (
                          <p className="text-xs text-orange-600">Aucune ligne BOM trouvée pour ce Produit-fini.</p>
                        ) : (
                          <>
                            {/* Composants manquants en évidence */}
                            {manquants.length > 0 && (
                              <div className="mb-4 bg-red-50 border border-red-200 rounded-lg p-3">
                                <p className="text-xs font-bold text-red-700 mb-2">⚠️ Pour produire le potentiel de {fmt(kit.potentiel)} kits, il faut ajouter :</p>
                                <div className="flex flex-wrap gap-2">
                                  {manquants.map(c => (
                                    <span key={c.codeERP} className="px-2 py-1 bg-red-100 text-red-700 rounded text-xs font-mono font-semibold">
                                      {c.codeERP} : +{fmt(c.manquant)}
                                    </span>
                                  ))}
                                </div>
                              </div>
                            )}
                            <p className="text-xs font-bold text-gray-600 mb-2">Composants du kit — Surplus = Stock − Besoin PDP</p>
                            <div className="overflow-x-auto">
                              <table className="w-full text-xs">
                                <thead>
                                  <tr className="text-gray-500 border-b border-gray-200">
                                    <th className="py-2 text-left">Code ERP</th>
                                    <th className="py-2 text-left">Désignation</th>
                                    <th className="py-2 text-right">Qté / kit</th>
                                    <th className="py-2 text-right">Stock</th>
                                    <th className="py-2 text-right">Besoin PDP</th>
                                    <th className="py-2 text-right">Surplus (restant)</th>
                                    <th className="py-2 text-right">Kits possibles</th>
                                    <th className="py-2 text-right">À ajouter</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {kit.components.map(c => (
                                    <tr key={c.codeERP} className={`border-b border-gray-100 ${c.estLimitant ? 'bg-red-50' : ''}`}>
                                      <td className="py-2 font-mono font-semibold">
                                        {c.estLimitant && <span title="Composant limitant">🔴 </span>}{c.codeERP}
                                      </td>
                                      <td className="py-2 text-gray-600">{c.designation}</td>
                                      <td className="py-2 text-right font-mono">{c.quantite}</td>
                                      <td className="py-2 text-right font-mono">{fmt(c.stock)}</td>
                                      <td className="py-2 text-right font-mono text-blue-700">{fmt(c.besoinPDP)}</td>
                                      <td className={`py-2 text-right font-mono font-semibold ${c.surplus < 0 ? 'text-red-600' : 'text-green-600'}`}>{c.surplus > 0 ? '+' : ''}{fmt(c.surplus)}</td>
                                      <td className="py-2 text-right font-mono font-bold">{fmt(c.kitsPossibles)}</td>
                                      <td className={`py-2 text-right font-mono font-bold ${c.manquant > 0 ? 'text-red-600' : 'text-gray-300'}`}>{c.manquant > 0 ? `+${fmt(c.manquant)}` : '—'}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </>
                        )}
                      </td>
                    </tr>
                  )}
                </>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// BESOIN SF TAB — Besoin en Semi-Fini par composant
// ─────────────────────────────────────────────
const BesoinSFTab: React.FC<{ surplus: SurplusRow[] }> = ({ surplus }) => {
  const [search, setSearch] = useState('');
  // Seuls les composants qui ont un besoin > 0
  const data = surplus.filter(r => r.besoinTotal > 0);
  const filtered = data.filter(r =>
    !search ||
    r.codeERP.toLowerCase().includes(search.toLowerCase()) ||
    r.codePF.toLowerCase().includes(search.toLowerCase()) ||
    r.designation.toLowerCase().includes(search.toLowerCase())
  );
  const totalBesoin = data.reduce((a, r) => a + r.besoinTotal, 0);

  const exportExcel = () => {
    const headers = ['Code ERP','Code Produit-fini','Désignation','Besoin S-1','Besoin S-2','Besoin S-3','Besoin S-4','Besoin S-5','Besoin total'];
    const rows = filtered.map(r => [
      r.codeERP, r.codePF, r.designation,
      r.besoinsNet.s23, r.besoinsNet.s24, r.besoinsNet.s25, r.besoinsNet.s26, r.besoinsNet.s27,
      r.besoinTotal,
    ]);
    downloadExcel(`besoin_sf_${new Date().toISOString().slice(0,10)}.xlsx`, headers, rows);
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-3 justify-between items-center">
        <div>
          <h2 className="text-lg font-bold text-gray-800">📋 Besoin Semi-Fini Plastique</h2>
          <p className="text-sm text-gray-500">Besoin net en composants calculé à partir du PDP × BOM (SOMME.SI par Code ERP)</p>
        </div>
        <div className="flex gap-2">
          <button onClick={exportExcel} className="px-4 py-2 bg-emerald-600 text-white rounded-xl text-sm font-semibold hover:bg-emerald-700 flex items-center gap-1">📊 Excel</button>
          <button onClick={() => window.print()} className="px-4 py-2 bg-rose-600 text-white rounded-xl text-sm font-semibold hover:bg-rose-700 flex items-center gap-1">🖨️ PDF</button>
        </div>
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <div className="bg-gradient-to-br from-blue-500 to-indigo-600 text-white rounded-xl p-4 shadow">
          <p className="text-blue-100 text-xs font-medium uppercase">Besoin total</p>
          <p className="text-2xl font-bold mt-1">{fmt(totalBesoin)}</p>
          <p className="text-blue-200 text-xs mt-1">pcs à produire</p>
        </div>
        <div className="bg-white rounded-xl p-4 border border-gray-100 shadow-sm">
          <p className="text-xs text-gray-500 uppercase font-medium">Composants concernés</p>
          <p className="text-2xl font-bold text-gray-800 mt-1">{data.length}</p>
        </div>
        <div className="bg-white rounded-xl p-4 border border-gray-100 shadow-sm">
          <p className="text-xs text-gray-500 uppercase font-medium">Produits-finis liés</p>
          <p className="text-2xl font-bold text-gray-800 mt-1">{new Set(data.flatMap(r => r.codePF.split(',').map(s => s.trim())).filter(Boolean)).size}</p>
        </div>
      </div>

      <input value={search} onChange={e => setSearch(e.target.value)} className="w-full max-w-md px-4 py-2 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="🔍 Rechercher Code ERP, Produit-fini, Désignation…" />

      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-800 text-white">
              <tr>
                <th className="px-3 py-3 text-left text-xs font-semibold sticky left-0 bg-slate-800">Code ERP</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Code Produit-fini</th>
                <th className="px-3 py-3 text-left text-xs font-semibold min-w-[160px]">Désignation</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">S-1</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">S-2</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">S-3</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">S-4</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">S-5</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-indigo-800">BESOIN TOTAL</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.length === 0 && <tr><td colSpan={9} className="text-center py-12 text-gray-400">Aucun besoin</td></tr>}
              {filtered.map(row => (
                <tr key={row.id} className="hover:bg-blue-50 transition-colors">
                  <td className="px-3 py-3 font-mono text-xs font-semibold sticky left-0 bg-inherit">{row.codeERP}</td>
                  <td className="px-3 py-3 font-mono text-xs text-gray-700">{row.codePF}</td>
                  <td className="px-3 py-3 text-xs text-gray-700 max-w-[200px] truncate">{row.designation}</td>
                  {(['s23','s24','s25','s26','s27'] as const).map(k => (
                    <td key={k} className="px-3 py-3 text-right font-mono text-xs text-blue-700">{row.besoinsNet[k] ? fmt(row.besoinsNet[k]) : <span className="text-gray-300">—</span>}</td>
                  ))}
                  <td className="px-3 py-3 text-right">
                    <span className="px-2 py-1 bg-indigo-100 text-indigo-800 rounded-full text-xs font-bold font-mono">{fmt(row.besoinTotal)}</span>
                  </td>
                </tr>
              ))}
            </tbody>
            {filtered.length > 0 && (
              <tfoot className="bg-gray-50 font-semibold text-sm">
                <tr>
                  <td colSpan={3} className="px-3 py-3">TOTAL</td>
                  {(['s23','s24','s25','s26','s27'] as const).map(k => (
                    <td key={k} className="px-3 py-3 text-right font-mono text-blue-700">{fmt(filtered.reduce((a, r) => a + r.besoinsNet[k], 0))}</td>
                  ))}
                  <td className="px-3 py-3 text-right font-mono text-indigo-800 font-bold">{fmt(filtered.reduce((a, r) => a + r.besoinTotal, 0))}</td>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// SURPLUS TAB
// ─────────────────────────────────────────────
const SurplusTab: React.FC<{ surplus: SurplusRow[] }> = ({ surplus }) => {
  const [filter, setFilter] = useState<'all'|'surplus'|'manque'>('all');
  const [search, setSearch] = useState('');
  const filtered = surplus
    .filter(r => filter==='all' ? true : filter==='surplus' ? r.surplusStock>0 : r.surplusStock<0)
    .filter(r => !search || r.codeERP.toLowerCase().includes(search.toLowerCase()) || r.codePF.toLowerCase().includes(search.toLowerCase()) || r.designation.toLowerCase().includes(search.toLowerCase()));

  const totalSurplus = surplus.filter(r=>r.surplusStock>0).reduce((a,r)=>a+r.surplusStock,0);
  const totalManque  = surplus.filter(r=>r.surplusStock<0).reduce((a,r)=>a+r.surplusStock,0);

  const exportExcel = () => {
    const headers = ['Code ERP','Code Produit-fini','Désignation','Stock','Besoin S-1','Besoin S-2','Besoin S-3','Besoin S-4','Besoin S-5','Besoin total','Surplus stock'];
    const rows = filtered.map(r => [
      r.codeERP, r.codePF, r.designation, r.stock,
      r.besoinsNet.s23, r.besoinsNet.s24, r.besoinsNet.s25, r.besoinsNet.s26, r.besoinsNet.s27,
      r.besoinTotal, r.surplusStock,
    ]);
    downloadExcel(`surplus_${new Date().toISOString().slice(0,10)}.xlsx`, headers, rows);
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap gap-3 justify-between items-center">
        <div><h2 className="text-lg font-bold text-gray-800">⚡ Analyse Surplus / Manque</h2><p className="text-sm text-gray-500">Stock − Besoin total PDP = Surplus stock</p></div>
        <div className="flex gap-2">
          <button onClick={exportExcel} className="px-4 py-2 bg-emerald-600 text-white rounded-xl text-sm font-semibold hover:bg-emerald-700 flex items-center gap-1">📊 Excel</button>
          <button onClick={() => window.print()} className="px-4 py-2 bg-rose-600 text-white rounded-xl text-sm font-semibold hover:bg-rose-700 flex items-center gap-1">🖨️ PDF</button>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-3">
        <div className="bg-gradient-to-br from-green-500 to-emerald-600 text-white rounded-xl p-4 shadow">
          <p className="text-green-100 text-xs font-medium uppercase">Surplus total</p>
          <p className="text-2xl font-bold mt-1">+{fmt(totalSurplus)}</p>
          <p className="text-green-200 text-xs mt-1">{surplus.filter(r=>r.surplusStock>0).length} réf.</p>
        </div>
        <div className="bg-gradient-to-br from-red-500 to-rose-600 text-white rounded-xl p-4 shadow">
          <p className="text-red-100 text-xs font-medium uppercase">Manque total</p>
          <p className="text-2xl font-bold mt-1">{fmt(totalManque)}</p>
          <p className="text-red-200 text-xs mt-1">{surplus.filter(r=>r.surplusStock<0).length} réf.</p>
        </div>
      </div>

      <div className="flex flex-wrap gap-3">
        <div className="flex rounded-xl overflow-hidden border border-gray-200 bg-white">
          {(['all','surplus','manque'] as const).map(f => (
            <button key={f} onClick={() => setFilter(f)}
              className={`px-4 py-2 text-sm font-medium transition-colors ${filter===f ? 'bg-slate-800 text-white' : 'text-gray-600 hover:bg-gray-50'}`}>
              {f==='all'?'Tout':f==='surplus'?'✅ Surplus':'❌ Manque'}
            </button>
          ))}
        </div>
        <input value={search} onChange={e => setSearch(e.target.value)} className="flex-1 min-w-[200px] px-4 py-2 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder=" Rechercher…" />
      </div>

      <div className="bg-white rounded-xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead className="bg-slate-800 text-white">
              <tr>
                <th className="px-3 py-3 text-left text-xs font-semibold sticky left-0 bg-slate-800">Code ERP</th>
                <th className="px-3 py-3 text-left text-xs font-semibold">Code Produit-fini</th>
                <th className="px-3 py-3 text-left text-xs font-semibold min-w-[160px]">Désignation</th>
                <th className="px-3 py-3 text-right text-xs font-semibold">Stock</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">Besoin S-1</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">Besoin S-2</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">Besoin S-3</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">Besoin S-4</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-blue-800">Besoin S-5</th>
                <th className="px-3 py-3 text-right text-xs font-semibold bg-green-800">SURPLUS STOCK</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filtered.length===0 && <tr><td colSpan={10} className="text-center py-12 text-gray-400">Aucun résultat</td></tr>}
              {filtered.map(row => {
                const isManque = row.surplusStock < 0;
                const isSurplus = row.surplusStock > 0;
                return (
                  <tr key={row.id} className={`transition-colors ${isManque?'bg-red-50 hover:bg-red-100':isSurplus?'bg-green-50 hover:bg-green-100':'hover:bg-gray-50'}`}>
                    <td className="px-3 py-3 font-mono text-xs font-semibold sticky left-0 bg-inherit">{row.codeERP}</td>
                    <td className="px-3 py-3 font-mono text-xs text-gray-700">{row.codePF}</td>
                    <td className="px-3 py-3 text-xs text-gray-700 max-w-[200px] truncate">{row.designation}</td>
                    <td className="px-3 py-3 text-right font-mono text-xs font-semibold">{fmt(row.stock)}</td>
                    {(['s23','s24','s25','s26','s27'] as const).map(k => (
                      <td key={k} className="px-3 py-3 text-right font-mono text-xs text-blue-700">{row.besoinsNet[k] ? fmt(row.besoinsNet[k]) : <span className="text-gray-300">—</span>}</td>
                    ))}
                    <td className="px-3 py-3 text-right">
                      <span className={`px-2 py-1 rounded-full text-xs font-bold font-mono ${isManque?'bg-red-500 text-white':isSurplus?'bg-green-500 text-white':'bg-gray-200 text-gray-600'}`}>
                        {isSurplus?'+':''}{fmt(row.surplusStock)}
                      </span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// DASHBOARD TAB
// ─────────────────────────────────────────────
const DashboardTab: React.FC<{ surplus: SurplusRow[]; stock: StockRow[]; kits: KitRow[] }> = ({ surplus, stock, kits }) => {
  const surplusItems = surplus.filter(r=>r.surplusStock>0).sort((a,b)=>b.surplusStock-a.surplusStock).slice(0,8);
  const manqueItems  = surplus.filter(r=>r.surplusStock<0).sort((a,b)=>a.surplusStock-b.surplusStock).slice(0,8);
  const totalStock   = stock.reduce((a,r)=>a+r.stock,0);
  const totalSurplus = surplus.filter(r=>r.surplusStock>0).reduce((a,r)=>a+r.surplusStock,0);

  const pieData = [
    { name:'Surplus', value: surplus.filter(r=>r.surplusStock>0).length, color:'#059669' },
    { name:'Équilibré', value: surplus.filter(r=>r.surplusStock===0).length, color:'#94A3B8' },
    { name:'Manque', value: surplus.filter(r=>r.surplusStock<0).length, color:'#DC2626' },
  ].filter(d=>d.value>0);

  const periodTotals = (['s23','s24','s25','s26','s27'] as const).map(k => ({
    name: k.toUpperCase(),
    besoin: surplus.reduce((a,r)=>a+r.besoinsNet[k],0),
  }));

  return (
    <div className="space-y-6">
      <div><h2 className="text-lg font-bold text-gray-800"> Tableau de Bord</h2></div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white rounded-xl p-5 shadow-sm border border-gray-100">
          <p className="text-xs text-gray-500 uppercase tracking-wide font-medium">Stock total</p>
          <p className="text-2xl font-bold text-blue-700 mt-1">{fmt(totalStock)}</p>
          <p className="text-xs text-gray-400 mt-1">pcs en stock</p>
        </div>
        <div className="bg-gradient-to-br from-green-50 to-emerald-50 rounded-xl p-5 shadow-sm border border-green-100">
          <p className="text-xs text-green-700 uppercase tracking-wide font-medium">Surplus total</p>
          <p className="text-2xl font-bold text-green-600 mt-1">+{fmt(totalSurplus)}</p>
          <p className="text-xs text-green-500 mt-1">{surplus.filter(r=>r.surplusStock>0).length} réf. excédentaires</p>
        </div>
        <div className="bg-white rounded-xl p-5 shadow-sm border border-gray-100">
          <p className="text-xs text-gray-500 uppercase tracking-wide font-medium">Références</p>
          <p className="text-2xl font-bold text-gray-800 mt-1">{surplus.length}</p>
          <p className="text-xs text-gray-400 mt-1">codes ERP distincts</p>
        </div>
        <div className="bg-white rounded-xl p-5 shadow-sm border border-gray-100">
          <p className="text-xs text-gray-500 uppercase tracking-wide font-medium">Kits à compléter</p>
          <p className="text-2xl font-bold text-violet-700 mt-1">{kits.length}</p>
          <p className="text-xs text-gray-400 mt-1">{kits.filter(k => k.status === 'MANQUE').length} manque · {kits.filter(k => k.status === 'BOM MANQUANTE').length} BOM</p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl shadow-sm p-5">
          <h3 className="text-sm font-bold text-gray-700 mb-4">📅 Besoins nets par semaine</h3>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={periodTotals} barSize={40}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
              <XAxis dataKey="name" tick={{fontSize:13,fontWeight:600}} />
              <YAxis tickFormatter={v=>`${(v/1000).toFixed(0)}k`} tick={{fontSize:11}} />
              <Tooltip formatter={(v:unknown)=>[fmt(Number(v)),'Besoin']} />
              <Bar dataKey="besoin" fill="#1E40AF" radius={[8,8,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="bg-white rounded-xl shadow-sm p-5">
          <h3 className="text-sm font-bold text-gray-700 mb-4">🎯 Répartition références</h3>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={pieData} cx="50%" cy="50%" outerRadius={80} dataKey="value" label={({name,value})=>`${name}: ${value}`} labelLine={false}>
                {pieData.map((d,i)=><Cell key={i} fill={d.color} />)}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-xl shadow-sm overflow-hidden">
          <div className="px-5 py-4 bg-green-50 border-b border-green-100"><h3 className="text-sm font-bold text-green-800">✅ Top Surplus</h3></div>
          <div className="divide-y divide-gray-100">
            {surplusItems.length===0 && <p className="text-center py-8 text-gray-400 text-sm">Aucun surplus</p>}
            {surplusItems.map(r => (
              <div key={r.id} className="px-5 py-3 flex justify-between items-center hover:bg-green-50">
                <div><p className="text-xs font-semibold text-gray-800 font-mono">{r.codeERP}</p><p className="text-xs text-gray-500 truncate max-w-[180px]">{r.designation}</p></div>
                <span className="text-sm font-bold text-green-600 font-mono">+{fmt(r.surplusStock)}</span>
              </div>
            ))}
          </div>
        </div>
        <div className="bg-white rounded-xl shadow-sm overflow-hidden">
          <div className="px-5 py-4 bg-red-50 border-b border-red-100"><h3 className="text-sm font-bold text-red-800">❌ Top Manques</h3></div>
          <div className="divide-y divide-gray-100">
            {manqueItems.length===0 && <p className="text-center py-8 text-gray-400 text-sm">Aucun manque 🎉</p>}
            {manqueItems.map(r => (
              <div key={r.id} className="px-5 py-3 flex justify-between items-center hover:bg-red-50">
                <div><p className="text-xs font-semibold text-gray-800 font-mono">{r.codeERP}</p><p className="text-xs text-gray-500 truncate max-w-[180px]">{r.designation}</p></div>
                <span className="text-sm font-bold text-red-600 font-mono">{fmt(r.surplusStock)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

// ─────────────────────────────────────────────
// MAIN APP
// ────────────────────────────────────────────
type Tab = 'dashboard'|'pdp'|'bom'|'stock'|'besoinSF'|'surplus'|'kits';

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [tab, setTab] = useState<Tab>('dashboard');
  const [pdp, setPdp] = useState<PDPRow[]>([]);
  const [bom, setBom] = useState<BOMRow[]>([]);
  const [stock, setStock] = useState<StockRow[]>([]);
  const [syncing, setSyncing] = useState(false);
  const [lastSync, setLastSync] = useState('—');
  const [importTarget, setImportTarget] = useState<'PDP'|'BOM'|'STOCK'|null>(null);

  const saveTimer = useRef<ReturnType<typeof setTimeout>|null>(null);

  useEffect(() => {
    seedDefaultUsers();
    const u = getSession();
    if (u) { setUser(u); restoreData(u.id); }
  }, []);

  const restoreData = (userId: string) => {
    const d = loadData(userId);
    if (d) {
      setPdp(d.pdp||[]); setBom(d.bom||[]); setStock(d.stock||[]);
    } else {
      const pp = demoPDP(); const bb = demoBOM(); const ss = demoStock();
      setPdp(pp); setBom(bb); setStock(ss);
      saveData(userId, { pdp:pp, bom:bb, stock:ss });
    }
    setLastSync(new Date().toLocaleTimeString('fr-FR'));
  };

  const onAuth = (u: User) => { setUser(u); restoreData(u.id); };

  const persist = useCallback((pp: PDPRow[], bb: BOMRow[], ss: StockRow[]) => {
    if (!user) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    setSyncing(true);
    saveTimer.current = setTimeout(() => {
      saveData(user.id, { pdp:pp, bom:bb, stock:ss });
      setSyncing(false); setLastSync(new Date().toLocaleTimeString('fr-FR'));
    }, 600);
  }, [user]);

  const handlePDP   = (d: PDPRow[])   => { setPdp(d);   persist(d, bom, stock); };
  const handleBOM   = (d: BOMRow[])   => { setBom(d);   persist(pdp, d, stock); };
  const handleStock = (d: StockRow[]) => { setStock(d); persist(pdp, bom, d); };
  const handleImport = (rows: PDPRow[]|BOMRow[]|StockRow[]) => {
    if (importTarget==='PDP')   handlePDP(rows as PDPRow[]);
    if (importTarget==='BOM'  && user?.role==='admin') handleBOM(rows as BOMRow[]);
    if (importTarget==='STOCK') handleStock(rows as StockRow[]);
  };
  const logout = () => { clearSession(); setUser(null); setPdp([]); setBom([]); setStock([]); };

  const surplus = useMemo(() => computeSurplus(pdp, bom, stock), [pdp, bom, stock]);
  const kits = useMemo(() => computeKits(pdp, bom, stock), [pdp, bom, stock]);

  const surplusNonZeroCount = surplus.filter((r: SurplusRow) => r.surplusStock !== 0).length;

  if (!user) return <AuthPage onAuth={onAuth} />;

  const besoinCount = surplus.filter(r => r.besoinTotal > 0).length;

  const navTabs: { id: Tab; label: string; badge?: number }[] = [
    { id:'dashboard', label:'📈 Dashboard' },
    { id:'pdp',       label:' PDP',       badge: pdp.length },
    { id:'bom',       label:'🗂️ BOM',      badge: bom.length },
    { id:'stock',     label:'📦 Stock',     badge: stock.length },
    { id:'besoinSF',  label:'📋 Besoin SF', badge: besoinCount },
    { id:'surplus',   label:'⚡ Surplus',   badge: surplusNonZeroCount },
    { id:'kits',      label:'🧩 Kits',      badge: kits.length },
  ];

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header */}
      <header className="bg-slate-800 text-white sticky top-0 z-40 shadow-lg">
        <div className="max-w-screen-xl mx-auto px-4 py-3 flex items-center gap-4">
          <span className="text-2xl">🏭</span>
          <div className="flex-1 min-w-0">
            <h1 className="text-base font-bold leading-tight">MRP</h1>
            <div className="flex items-center gap-2 text-xs text-gray-400">
              {syncing
                ? <><div className="w-2 h-2 rounded-full bg-blue-400 animate-ping" /><span>Sync…</span></>
                : <><div className="w-2 h-2 rounded-full bg-green-500" /><span>Sauvegardé {lastSync}</span></>
              }
            </div>
          </div>
          <div className="flex items-center gap-2">
            <div className={`w-8 h-8 rounded-full flex items-center justify-center text-white text-sm font-bold ${user.role==='admin' ? 'bg-amber-500' : 'bg-blue-500'}`}>{user.name.charAt(0).toUpperCase()}</div>
            <div className="hidden sm:block">
              <p className="text-xs font-semibold leading-none flex items-center gap-1">{user.name}{user.role==='admin' && <span className="px-1.5 py-0.5 bg-amber-400 text-amber-900 rounded text-[10px] font-bold">ADMIN</span>}</p>
              <p className="text-xs text-slate-400 leading-none mt-0.5">{user.email}</p>
            </div>
            <button onClick={logout} className="ml-1 px-2 py-1 bg-red-600/20 hover:bg-red-600/40 text-red-300 text-xs rounded-lg">⏻</button>
          </div>
        </div>
      </header>

      {/* Nav */}
      <nav className="bg-white border-b border-gray-200 sticky top-14 z-30">
        <div className="max-w-screen-xl mx-auto px-2 flex overflow-x-auto gap-0.5 py-1">
          {navTabs.map(t => (
            <button key={t.id} onClick={()=>setTab(t.id)}
              className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-sm font-medium whitespace-nowrap transition-colors ${tab===t.id?'bg-slate-800 text-white':'text-gray-600 hover:bg-gray-100'}`}>
              {t.label}
              {t.badge!==undefined && t.badge>0 && <span className={`px-1.5 py-0.5 rounded-full text-xs font-bold ${tab===t.id?'bg-white/20 text-white':'bg-blue-100 text-blue-600'}`}>{t.badge}</span>}
            </button>
          ))}
        </div>
      </nav>

      {/* Content */}
      <main className="flex-1 max-w-screen-xl mx-auto w-full px-3 sm:px-6 py-5">
        {tab==='dashboard' && <DashboardTab surplus={surplus} stock={stock} kits={kits} />}
        {tab==='pdp'       && <PDPTab data={pdp} onChange={handlePDP} onImport={()=>setImportTarget('PDP')} />}
        {tab==='bom'       && <BOMTab data={bom} onChange={handleBOM} onImport={()=>setImportTarget('BOM')} isAdmin={user.role==='admin'} />}
        {tab==='stock'     && <StockTab data={stock} onChange={handleStock} onImport={()=>setImportTarget('STOCK')} />}
        {tab==='besoinSF'  && <BesoinSFTab surplus={surplus} />}
        {tab==='surplus'   && <SurplusTab surplus={surplus} />}
        {tab==='kits'      && <KitsTab kits={kits} />}
      </main>

      <footer className="border-t border-gray-200 bg-white py-3 text-center text-xs text-gray-400">
        🏭 MRP Semi Fini plastique — Données sauvegardées localement • {user.name} • {new Date().toLocaleDateString('fr-FR')}
      </footer>

      {importTarget && <ImportModal type={importTarget} onImport={handleImport} onClose={()=>setImportTarget(null)} />}
    </div>
  );
}

import { useEffect, useMemo, useState } from "react";
import { initializeApp } from "firebase/app";
import {
  getAuth,
  onAuthStateChanged,
  signOut,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
} from "firebase/auth";
import {
  getFirestore,
  collection,
  addDoc,
  getDocs,
  updateDoc,
  deleteDoc,
  doc,
  query,
  where,
} from "firebase/firestore";
import jsPDF from "jspdf";
import * as XLSX from "xlsx";

const firebaseConfig = {
  apiKey: "AIzaSyCX5L65n6MAsrfTb2thhO4Nuqmq4bNYSiI",
  authDomain: "emlakapp-827a2.firebaseapp.com",
  projectId: "emlakapp-827a2",
  storageBucket: "emlakapp-827a2.firebasestorage.app",
  messagingSenderId: "910928133044",
  appId: "1:910928133044:web:50f2b0f0e6884c959010e1",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const emptyCustomer = () => ({
  fullName: "",
  phone: "",
  email: "",
  customerType: "alıcı",
  interestedLocation: "",
  interestedPropertyType: "",
  budgetMin: "",
  budgetMax: "",
  note: "",
});

const emptyPortfolio = () => ({
  portfolioNo: "",
  title: "",
  portfolioType: "satılık",
  propertyType: "",
  price: "",
  roomCount: "",
  grossArea: "",
  city: "",
  district: "",
  address: "",
  ownerName: "",
  ownerPhone: "",
  imageUrl: "",
  description: "",
  status: "aktif",
  note: "",
});

const emptyMatch = () => ({
  customerId: "",
  portfolioId: "",
  status: "önerildi",
  note: "",
});

const emptyAppointment = () => ({
  title: "",
  customerId: "",
  portfolioId: "",
  date: "",
  time: "",
  location: "",
  status: "planlandı",
  note: "",
});

const emptyContract = () => ({
  contractNo: "",
  companyName: "",
  agentName: "",
  owner: "",
  phone: "",
  address: "",
  propertyType: "",
  salePrice: "",
  commission: "",
  startDate: "",
  endDate: "",
  status: "aktif",
  note: "",
});


const DEFAULT_YETKI_TERMS = `1. Mülk sahibi, satış/kiralama yetkisini münhasıran bu sözleşmede belirtilen emlak ofisine vermiştir.
2. Yetki süresi boyunca başka bir aracı kurum veya kişiyle işlem yapılamaz.
3. Emlak ofisi, mülkü ilan sitelerinde ve kendi kanallarında tanıtma hakkına sahiptir.
4. Satış/kiralama gerçekleştiğinde, belirtilen komisyon oranı peşin olarak ödenecektir.
5. Sözleşme süresi dolmadan mülk sahibi tarafından feshedilmesi halinde, o güne kadar yapılan masraflar talep edilebilir.
6. Taraflar arasındaki uyuşmazlıklarda sözleşmenin imzalandığı yer mahkemeleri yetkilidir.`;

const DEFAULT_SATIS_TERMS = `1. Satıcı, mülkün tüm yasal yükümlülüklerden ari olduğunu beyan ve taahhüt eder.
2. Alıcı, belirlenen kapora bedelini sözleşme imzası ile birlikte öder; cayma halinde kapora iade edilmez.
3. Satıcı caydığı takdirde aldığı kaporanın iki katını alıcıya iade eder.
4. Tapu devri, tüm ödemelerin tamamlanmasının ardından gerçekleştirilir.
5. Tapu harcı ve ilgili masraflar yasal oranlarda taraflarca paylaşılır.
6. Emlak komisyonu, tapu devri gerçekleştiğinde belirlenen oran üzerinden ödenir.
7. Bu sözleşmeden doğan uyuşmazlıklarda tarafların yerleşim yeri mahkemeleri yetkilidir.`;

const DEFAULT_KIRA_TERMS = `1. Kiracı, kiralananı özenle kullanmak ve komşulara saygılı davranmak zorundadır.
2. Kiralanan, yazılı izin alınmadan üçüncü kişilere devredilemez veya alt kiraya verilemez.
3. Kira bedeli, her ayın belirlenen gününde eksiksiz ödenecektir; geç ödemelerde yasal faiz uygulanır.
4. Kiracı, abonelik sözleşmelerini kendi adına yapacak ve sözleşme sonunda kapatacaktır.
5. Kira artışı, Türkiye İstatistik Kurumu tarafından açıklanan TÜFE oranında yapılacaktır.
6. Sözleşme sonunda kiralanan, teslim alındığı hâliyle iade edilecektir.
7. Kiralananın mülkiyet vergisi kiralayana, kullanım vergisi ve harçlar kiracıya aittir.
8. Bu sözleşmeden doğan uyuşmazlıklarda tarafların yerleşim yeri mahkemeleri yetkilidir.`;

const emptyContractTemplates = () => ({
  yetkiTerms: DEFAULT_YETKI_TERMS,
  satisTerms: DEFAULT_SATIS_TERMS,
  kiraTerms: DEFAULT_KIRA_TERMS,
});

const emptyOfficeProfile = () => ({
  officeName: "",
  agentName: "",
  phone: "",
  email: "",
  website: "",
  address: "",
  city: "",
  taxOffice: "",
  taxNo: "",
  logoDataUrl: "",
  logoName: "",
});

const emptySaleAgreement = () => ({
  contractNo: "",
  contractDate: new Date().toISOString().slice(0, 10),
  sellerName: "",
  sellerPhone: "",
  buyerName: "",
  buyerPhone: "",
  propertyTitle: "",
  propertyAddress: "",
  titleDeedInfo: "",
  salePrice: "",
  deposit: "",
  commissionRate: "",
  commissionAmount: "",
  deliveryDate: "",
  note: "",
});

const emptyRentAgreement = () => ({
  contractNo: "",
  contractDate: new Date().toISOString().slice(0, 10),
  propertyNeighborhood: "",
  propertyStreet: "",
  propertyDoorNo: "",
  propertyKind: "",
  propertyTitle: "",
  propertyAddress: "",
  landlordName: "",
  landlordPhone: "",
  landlordTcNo: "",
  landlordAddress: "",
  tenantName: "",
  tenantPhone: "",
  tenantTcNo: "",
  tenantAddress: "",
  monthlyRent: "",
  annualRent: "",
  deposit: "",
  paymentDay: "",
  paymentMethod: "",
  usageType: "",
  propertyCondition: "",
  fixtures: "",
  rentDuration: "",
  startDate: new Date().toISOString().slice(0, 10),
  endDate: "",
  increaseRate: "",
  guarantorName: "",
  guarantorTcNo: "",
  guarantorAddress: "",
  note: "",
});

const RENT_GENERAL_CONDITIONS = [
  "1. Kiracı, kiralananı özenle kullanmak zorundadır.",
  "2. Kiracı, kiralananda ve çevrede oturanlara iyi niyet kuralları içinde davranmaya zorunludur.",
  "3. Kiracı, kiralananı kısmen veya tamamen üçüncü kişilere kiralayamaz, alt kiraya veremez; devir ve temlik edemez.",
  "4. Kiracı, kiralayanın yazılı izni olmadıkça, kiralananda değişiklik yapamaz; aksi halde doğacak zararı karşılamak zorundadır.",
  "5. Üçüncü kişilerin kiralanan üzerinde hak iddia etmeleri halinde, kiracı durumu derhal kiralayana haber vermek zorundadır.",
  "6. Kiracı, kiralananda yapılması gereken onarımları derhal kiralayana bildirmek zorundadır; aksi halde doğacak zarardan sorumludur.",
  "7. Kiracı, kat malikleri kurulunca kendisine tebliğ edilen hususları kiralayana haber vermek zorundadır.",
  "8. Kiracı, kat malikleri kurulu kararı uyarınca yapılması gereken işlere izin vermek zorundadır.",
  "9. Kiracı, kiralanandaki olağan kullanımdan dolayı yapılması gereken onarımları yapmak veya yaptırmak ve giderlerini karşılamak zorundadır.",
  "10. Kiralananın mülkiyet hakkından doğan vergileri kiralayana, kullanımdan doğan vergi, resim ve harçları kiracıya aittir.",
  "11. Kiracı, kira sözleşmesinin sonunda kiralananı aldığı şekilde kiralayana teslim etmek zorundadır.",
  "12. Kiralananın iyi ve kullanılmaya elverişli halde teslim edildiği asıldır.",
  "13. Kiracı, kira sözleşmesinin sona ermesi veya satışa çıkartılması halinde kiralananın gezilmesine ve incelenmesine izin vermek zorundadır.",
  "14. Kiralananın boşaltılması gereken hallerde, kiralananın boşaltılmaması durumunda ortaya çıkacak zararlardan kiracı sorumlu olacaktır.",
  "15. Kiracı, ciddi tehlike oluşturmayan kusurlardan dolayı kiralayanı teslim almaktan kaçınamaz ve kiradan indirim talep edemez.",
  "16. Kiracı, kiralana yaptığı faydalı ve lüks şeylerin bedelini kiralayandan isteyemez.",
  "17. Kiracı, kiralayanın yazılı olurunu almak ve giderleri kendisine ait olmak üzere teknik donanımları yaptırabilir.",
  "18. İşbu kira sözleşmesinde yer almayan hususlar hakkında 6098 sayılı Türk Borçlar Kanunu hükümleri geçerlidir.",
];

const RENT_SPECIAL_CONDITIONS = [
  "1. Kiralanan alt kiraya verilemez, ortak alınamaz; devir ve temlik edilemez.",
  "2. Kiralanan, sözleşmede belirtilen kullanım amacı dışında kullanılamaz.",
  "3. Kira bedelleri, tarafların belirlediği ödeme şekline göre zamanında ödenecektir.",
  "4. Kiralananın aidat, yakıt ve genel giderleri aksi kararlaştırılmadıkça kiracı tarafından ödenecektir.",
  "5. Kapılar, pencereler ve tesisat araçları sağlam, tam ve kullanılmaya elverişli olarak teslim edilmiştir.",
  "6. Kiracı, elektrik, su, doğalgaz ve benzeri abonelikleri kendi adına yaptıracak ve sözleşme sonunda kapatacaktır.",
  "7. Kiracı, kiralananı özenle kullanacak; kiralayan da gerekli onarımları makul sürede yaptıracaktır.",
  "8. Sözleşmeden doğacak uyuşmazlıklarda tarafların yerleşim yeri mahkemeleri ve icra müdürlükleri yetkilidir.",
];

export default function App() {
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [dataLoading, setDataLoading] = useState(false);
  const [migrationDone, setMigrationDone] = useState(false);

  const [authMode, setAuthMode] = useState("login");
  const [authForm, setAuthForm] = useState({ email: "", password: "" });

  const [theme, setTheme] = useState(() => localStorage.getItem("ekey-theme") || "light");
  const [isMobile, setIsMobile] = useState(
    typeof window !== "undefined" ? window.innerWidth < 1000 : false
  );
  const [activeTab, setActiveTab] = useState("dashboard");

  const [customers, setCustomers] = useState([]);
  const [portfolios, setPortfolios] = useState([]);
  const [matches, setMatches] = useState([]);
  const [appointments, setAppointments] = useState([]);
  const [contracts, setContracts] = useState([]);

  const [customerForm, setCustomerForm] = useState(emptyCustomer());
  const [portfolioForm, setPortfolioForm] = useState(emptyPortfolio());
  const [matchForm, setMatchForm] = useState(emptyMatch());
  const [appointmentForm, setAppointmentForm] = useState(emptyAppointment());
  const [contractForm, setContractForm] = useState(emptyContract());
  const [officeProfile, setOfficeProfile] = useState(emptyOfficeProfile());
  const [contractTemplates, setContractTemplates] = useState(emptyContractTemplates());
  const [templateScanLoading, setTemplateScanLoading] = useState(false);
  const [templateTab, setTemplateTab] = useState("yetki");
  const [saleAgreementForm, setSaleAgreementForm] = useState(emptySaleAgreement());
  const [rentAgreementForm, setRentAgreementForm] = useState(emptyRentAgreement());

  const [editingCustomerId, setEditingCustomerId] = useState(null);
  const [editingPortfolioId, setEditingPortfolioId] = useState(null);
  const [editingAppointmentId, setEditingAppointmentId] = useState(null);
  const [editingContractId, setEditingContractId] = useState(null);

  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [selectedPortfolio, setSelectedPortfolio] = useState(null);

  const [customerSearch, setCustomerSearch] = useState("");
  const [portfolioSearch, setPortfolioSearch] = useState("");
  const [appointmentSearch, setAppointmentSearch] = useState("");
  const [contractSearch, setContractSearch] = useState("");

  const [customerTypeFilter, setCustomerTypeFilter] = useState("tümü");
  const [portfolioTypeFilter, setPortfolioTypeFilter] = useState("tümü");
  const [appointmentStatusFilter, setAppointmentStatusFilter] = useState("tümü");
  const [contractStatusFilter, setContractStatusFilter] = useState("tümü");

  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < 1000);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  useEffect(() => {
    localStorage.setItem("ekey-theme", theme);
  }, [theme]);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => {
      setUser(u || null);
      setAuthLoading(false);
    });
    return () => unsub();
  }, []);


  useEffect(() => {
    if (!user?.uid) return;
    try {
      const savedOfficeProfile = localStorage.getItem(`ekey-office-profile-${user.uid}`);
      const savedSaleAgreement = localStorage.getItem(`ekey-sale-agreement-${user.uid}`);
      const savedRentAgreement = localStorage.getItem(`ekey-rent-agreement-${user.uid}`);
      if (savedOfficeProfile) setOfficeProfile({ ...emptyOfficeProfile(), ...JSON.parse(savedOfficeProfile) });
      if (savedSaleAgreement) setSaleAgreementForm({ ...emptySaleAgreement(), ...JSON.parse(savedSaleAgreement) });
      if (savedRentAgreement) setRentAgreementForm({ ...emptyRentAgreement(), ...JSON.parse(savedRentAgreement) });
      const savedTemplates = localStorage.getItem(`ekey-contract-templates-${user.uid}`);
      if (savedTemplates) setContractTemplates({ ...emptyContractTemplates(), ...JSON.parse(savedTemplates) });
    } catch (e) {
      console.warn("Yerel sözleşme ayarları okunamadı", e);
    }
  }, [user?.uid]);

  useEffect(() => {
    if (!user?.uid) return;
    try {
      localStorage.setItem(`ekey-office-profile-${user.uid}`, JSON.stringify(officeProfile));
      localStorage.setItem(`ekey-sale-agreement-${user.uid}`, JSON.stringify(saleAgreementForm));
      localStorage.setItem(`ekey-rent-agreement-${user.uid}`, JSON.stringify(rentAgreementForm));
      localStorage.setItem(`ekey-contract-templates-${user.uid}`, JSON.stringify(contractTemplates));
    } catch (e) {
      console.warn("Yerel sözleşme ayarları yazılamadı", e);
    }
  }, [user?.uid, officeProfile, saleAgreementForm, rentAgreementForm, contractTemplates]);

  useEffect(() => {
    if (!user?.uid) return;

    let cancelled = false;

    async function initUserData() {
      try {
        setDataLoading(true);
        await migrateLegacyContracts(user.uid, user.email || "");
        if (!cancelled) {
          await loadAll(user.uid);
          setMigrationDone(true);
        }
      } catch (e) {
        console.error("Veri yükleme hatası:", e);
        alert("Veriler yüklenirken hata oluştu: " + e.message);
      } finally {
        if (!cancelled) setDataLoading(false);
      }
    }

    initUserData();

    return () => {
      cancelled = true;
    };
  }, [user?.uid]);

  const dark = theme === "dark";
  const todayStr = new Date().toISOString().slice(0, 10);

  const colors = {
    bg: dark ? "#07111f" : "#eef4ff",
    bg2: dark ? "#0b1728" : "#f8fbff",
    panel: dark ? "rgba(16,26,43,0.92)" : "rgba(255,255,255,0.94)",
    panel2: dark ? "rgba(20,32,51,0.9)" : "#f7faff",
    panel3: dark ? "#0d1728" : "#edf3ff",
    text: dark ? "#f8fafc" : "#0f172a",
    muted: dark ? "#cbd5e1" : "#475569",
    sub: dark ? "#94a3b8" : "#64748b",
    border: dark ? "#22324b" : "#dbe4f0",
    border2: dark ? "#31425d" : "#c7d4e5",
    primary: dark ? "#60a5fa" : "#2563eb",
    primarySoft: dark ? "rgba(96,165,250,0.16)" : "rgba(37,99,235,0.12)",
    primaryText: "#ffffff",
    success: dark ? "#22c55e" : "#15803d",
    successSoft: dark ? "rgba(34,197,94,0.16)" : "rgba(34,197,94,0.12)",
    warning: dark ? "#f59e0b" : "#d97706",
    warningSoft: dark ? "rgba(245,158,11,0.16)" : "rgba(245,158,11,0.12)",
    danger: dark ? "#f87171" : "#dc2626",
    dangerSoft: dark ? "rgba(248,113,113,0.16)" : "rgba(220,38,38,0.12)",
    shadow: dark
      ? "0 22px 55px rgba(0,0,0,0.38)"
      : "0 20px 48px rgba(15, 23, 42, 0.08)",
    shadowSoft: dark
      ? "0 12px 26px rgba(2,6,23,0.22)"
      : "0 10px 24px rgba(15, 23, 42, 0.05)",
    ring: dark
      ? "0 0 0 1px rgba(255,255,255,0.03) inset"
      : "0 0 0 1px rgba(255,255,255,0.72) inset",
  };

  const navItems = [
    { key: "dashboard", label: "Anasayfa", icon: "🏠" },
    { key: "reports", label: "Raporlar", icon: "📊" },
    { key: "customers", label: "Müşteriler", icon: "👥" },
    { key: "portfolios", label: "Portföyler", icon: "🏢" },
    { key: "matches", label: "Eşleştirmeler", icon: "🔗" },
    { key: "appointments", label: "Randevular", icon: "📅" },
    { key: "contracts", label: "Sözleşmeler", icon: "📄" },
    { key: "saleContracts", label: "Satış Sözleşmesi", icon: "🧾" },
    { key: "rentalContracts", label: "Kira Sözleşmesi", icon: "🏘️" },
    { key: "settings", label: "Ayarlar", icon: "⚙️" },
  ];

  async function loadAll(uid) {
    await Promise.all([
      loadCollection("customers", setCustomers, uid),
      loadCollection("portfolios", setPortfolios, uid),
      loadCollection("matches", setMatches, uid),
      loadCollection("appointments", setAppointments, uid),
      loadCollection("contracts", setContracts, uid),
    ]);
  }

  async function loadCollection(name, setter, uid) {
    if (!uid) {
      setter([]);
      return;
    }

    const q = query(collection(db, name), where("userId", "==", uid));
    const snap = await getDocs(q);
    const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

    rows.sort((a, b) => {
      const aDate = a.updatedAt || a.createdAt || "";
      const bDate = b.updatedAt || b.createdAt || "";
      return String(bDate).localeCompare(String(aDate));
    });

    setter(rows);
  }

  async function migrateLegacyContracts(uid, userEmail) {
    try {
      const migrationKey = `ekey-contract-migrated-${uid}`;
      if (localStorage.getItem(migrationKey) === "1") return;

      const possibleKeys = [
        "contracts",
        "savedContracts",
        "ekey-contracts",
        "e-key-contracts",
        "sozlesmeler",
      ];

      let legacyContracts = [];

      for (const key of possibleKeys) {
        const raw = localStorage.getItem(key);
        if (!raw) continue;
        try {
          const parsed = JSON.parse(raw);
          if (Array.isArray(parsed) && parsed.length) {
            legacyContracts = parsed;
            break;
          }
        } catch (err) {
          console.warn(`LocalStorage parse hatası (${key})`, err);
        }
      }

      if (!legacyContracts.length) {
        localStorage.setItem(migrationKey, "1");
        return;
      }

      const existingSnap = await getDocs(
        query(collection(db, "contracts"), where("userId", "==", uid))
      );

      const existingNumbers = new Set(
        existingSnap.docs.map((d) => d.data()?.contractNo).filter(Boolean)
      );

      for (const item of legacyContracts) {
        const contractNo = item?.contractNo || "";
        if (contractNo && existingNumbers.has(contractNo)) continue;

        await addDoc(collection(db, "contracts"), {
          contractNo: item?.contractNo || "",
          companyName: item?.companyName || "",
          agentName: item?.agentName || "",
          owner: item?.owner || "",
          phone: item?.phone || "",
          address: item?.address || "",
          propertyType: item?.propertyType || "",
          salePrice: item?.salePrice || "",
          commission: item?.commission || "",
          startDate: item?.startDate || "",
          endDate: item?.endDate || "",
          status: item?.status || "aktif",
          note: item?.note || "",
          userId: uid,
          userEmail: userEmail || "",
          createdAt: item?.createdAt || new Date().toISOString(),
          migratedFromLocal: true,
        });
      }

      localStorage.setItem(migrationKey, "1");
    } catch (e) {
      console.error("Legacy contract migration error:", e);
    }
  }

  async function handleRegister() {
    try {
      if (!authForm.email || !authForm.password) {
        alert("E-posta ve şifre gerekli");
        return;
      }

      await createUserWithEmailAndPassword(auth, authForm.email, authForm.password);
      setAuthForm({ email: "", password: "" });
    } catch (e) {
      alert("Kayıt hatası: " + e.message);
    }
  }

  async function handleLogin() {
    try {
      if (!authForm.email || !authForm.password) {
        alert("E-posta ve şifre gerekli");
        return;
      }

      await signInWithEmailAndPassword(auth, authForm.email, authForm.password);
      setAuthForm({ email: "", password: "" });
    } catch (e) {
      alert("Giriş hatası: " + e.message);
    }
  }

  async function handleLogout() {
    try {
      await signOut(auth);
      setCustomers([]);
      setPortfolios([]);
      setMatches([]);
      setAppointments([]);
      setContracts([]);
      setMigrationDone(false);
      setSelectedCustomer(null);
      setSelectedPortfolio(null);
    } catch (e) {
      alert("Çıkış hatası: " + e.message);
    }
  }


  async function saveCustomer() {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (!customerForm.fullName || !customerForm.phone || !customerForm.customerType) {
        alert("Ad soyad, telefon ve müşteri tipi zorunlu");
        return;
      }

      if (editingCustomerId) {
        await updateDoc(doc(db, "customers", editingCustomerId), {
          ...customerForm,
          updatedAt: new Date().toISOString(),
        });
      } else {
        await addDoc(collection(db, "customers"), {
          ...customerForm,
          userId: user.uid,
          userEmail: user.email || "",
          createdAt: new Date().toISOString(),
        });
      }

      setCustomerForm(emptyCustomer());
      setEditingCustomerId(null);
      await loadCollection("customers", setCustomers, user.uid);
    } catch (e) {
      alert("Müşteri kayıt hatası: " + e.message);
    }
  }

  async function savePortfolio() {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (
        !portfolioForm.title ||
        !portfolioForm.portfolioType ||
        !portfolioForm.propertyType ||
        !portfolioForm.price ||
        !portfolioForm.city ||
        !portfolioForm.district
      ) {
        alert("Başlık, tip, mülk tipi, fiyat, şehir ve ilçe zorunlu");
        return;
      }

      const payload = {
        ...portfolioForm,
        portfolioNo: portfolioForm.portfolioNo || nextPortfolioNo(),
      };

      if (editingPortfolioId) {
        await updateDoc(doc(db, "portfolios", editingPortfolioId), {
          ...payload,
          updatedAt: new Date().toISOString(),
        });
      } else {
        await addDoc(collection(db, "portfolios"), {
          ...payload,
          userId: user.uid,
          userEmail: user.email || "",
          createdAt: new Date().toISOString(),
        });
      }

      setPortfolioForm(emptyPortfolio());
      setEditingPortfolioId(null);
      await loadCollection("portfolios", setPortfolios, user.uid);
    } catch (e) {
      alert("Portföy kayıt hatası: " + e.message);
    }
  }

  async function saveMatch() {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (!matchForm.customerId || !matchForm.portfolioId) {
        alert("Müşteri ve portföy seçmelisin");
        return;
      }

      const exists = matches.some(
        (m) => m.customerId === matchForm.customerId && m.portfolioId === matchForm.portfolioId
      );

      if (exists) {
        alert("Bu müşteri ile bu portföy zaten eşleşmiş");
        return;
      }

      await addDoc(collection(db, "matches"), {
        ...matchForm,
        userId: user.uid,
        userEmail: user.email || "",
        createdAt: new Date().toISOString(),
      });

      setMatchForm(emptyMatch());
      await loadCollection("matches", setMatches, user.uid);
    } catch (e) {
      alert("Eşleştirme hatası: " + e.message);
    }
  }

  async function saveSuggestedMatch(customerId, portfolioId, score = 0, reasons = [], mismatchReasons = []) {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (!customerId || !portfolioId) {
        alert("Müşteri ve portföy seçmelisin");
        return;
      }

      const exists = matches.some(
        (m) => m.customerId === customerId && m.portfolioId === portfolioId
      );

      if (exists) {
        alert("Bu müşteri ile bu portföy zaten eşleşmiş");
        return;
      }

      const noteParts = [];
      if (score) noteParts.push(`Akıllı eşleştirme puanı: ${score}/100`);
      if (reasons.length) noteParts.push(reasons.join(" • "));
      if (mismatchReasons.length) noteParts.push(`Dikkat: ${mismatchReasons.join(" • ")}`);

      await addDoc(collection(db, "matches"), {
        customerId,
        portfolioId,
        status: "önerildi",
        note: noteParts.join(" • "),
        userId: user.uid,
        userEmail: user.email || "",
        createdAt: new Date().toISOString(),
      });

      setMatchForm({
        customerId,
        portfolioId,
        status: "önerildi",
        note: noteParts.join(" • "),
      });
      await loadCollection("matches", setMatches, user.uid);
    } catch (e) {
      alert("Akıllı eşleştirme hatası: " + e.message);
    }
  }

  async function saveAppointment() {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (!appointmentForm.title || !appointmentForm.date || !appointmentForm.time) {
        alert("Başlık, tarih ve saat zorunlu");
        return;
      }

      if (editingAppointmentId) {
        await updateDoc(doc(db, "appointments", editingAppointmentId), {
          ...appointmentForm,
          updatedAt: new Date().toISOString(),
        });
      } else {
        await addDoc(collection(db, "appointments"), {
          ...appointmentForm,
          userId: user.uid,
          userEmail: user.email || "",
          createdAt: new Date().toISOString(),
        });
      }

      setAppointmentForm(emptyAppointment());
      setEditingAppointmentId(null);
      await loadCollection("appointments", setAppointments, user.uid);
    } catch (e) {
      alert("Randevu kayıt hatası: " + e.message);
    }
  }

  async function saveContract() {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (
        !contractForm.companyName ||
        !contractForm.agentName ||
        !contractForm.owner ||
        !contractForm.phone ||
        !contractForm.address ||
        !contractForm.propertyType ||
        !contractForm.salePrice ||
        !contractForm.commission ||
        !contractForm.startDate ||
        !contractForm.endDate
      ) {
        alert("Zorunlu alanları doldur");
        return;
      }

      const payload = {
        ...contractForm,
        contractNo: contractForm.contractNo || nextContractNo(),
        updatedAt: new Date().toISOString(),
      };

      if (editingContractId) {
        await updateDoc(doc(db, "contracts", editingContractId), payload);
      } else {
        await addDoc(collection(db, "contracts"), {
          ...payload,
          userId: user.uid,
          userEmail: user.email || "",
          createdAt: new Date().toISOString(),
        });
      }

      setContractForm(emptyContract());
      setEditingContractId(null);
      await loadCollection("contracts", setContracts, user.uid);
    } catch (e) {
      alert("Sözleşme kayıt hatası: " + e.message);
    }
  }

  async function removeItem(colName, id, setter) {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      if (!window.confirm("Bu kaydı silmek istediğine emin misin?")) return;

      await deleteDoc(doc(db, colName, id));
      await loadCollection(colName, setter, user.uid);

      if (colName === "customers" && selectedCustomer?.id === id) {
        setSelectedCustomer(null);
      }

      if (colName === "portfolios" && selectedPortfolio?.id === id) {
        setSelectedPortfolio(null);
      }
    } catch (e) {
      alert("Silme hatası: " + e.message);
    }
  }

  async function updateMatchStatus(id, newStatus) {
    try {
      if (!user?.uid) {
        alert("Önce giriş yapmalısın");
        return;
      }

      await updateDoc(doc(db, "matches", id), {
        status: newStatus,
        updatedAt: new Date().toISOString(),
      });

      await loadCollection("matches", setMatches, user.uid);
    } catch (e) {
      alert("Durum güncellenemedi: " + e.message);
    }
  }

  function formatCurrency(value) {
    if (!value) return "-";
    const n = Number(String(value).replace(/[^\d]/g, ""));
    if (!n) return value;
    return `${n.toLocaleString("tr-TR")} ₺`;
  }

  function sanitizePhone(phoneValue) {
    let phone = String(phoneValue || "").replace(/\D/g, "");
    if (phone.startsWith("0")) phone = phone.slice(1);
    if (phone && !phone.startsWith("90")) phone = `90${phone}`;
    return phone;
  }

  function shareOnWhatsApp(phoneValue, message) {
    const phone = sanitizePhone(phoneValue);
    if (!phone) {
      alert("Telefon numarası yok");
      return;
    }

    window.open(`https://wa.me/${phone}?text=${encodeURIComponent(message)}`, "_blank");
  }

  function getDaysRemaining(endDate) {
    if (!endDate) return null;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const end = new Date(endDate);
    end.setHours(0, 0, 0, 0);

    return Math.ceil((end - today) / (1000 * 60 * 60 * 24));
  }

  function getContractStatus(item) {
    if (item.status === "iptal edildi") return "iptal edildi";
    const d = getDaysRemaining(item.endDate);
    if (d == null) return item.status || "aktif";
    if (d < 0) return "süresi doldu";
    if (d === 0) return "bugün bitiyor";
    if (d <= 7) return "yakında bitecek";
    return "aktif";
  }

  function statusBadge(status) {
    let bg = dark ? "#334155" : "#e2e8f0";
    let color = dark ? "#f8fafc" : "#334155";

    if (["aktif", "planlandı"].includes(status)) {
      bg = dark ? "rgba(34,197,94,0.18)" : "#dcfce7";
      color = dark ? "#bbf7d0" : "#166534";
    } else if (["yakında bitecek", "ilgileniyor"].includes(status)) {
      bg = dark ? "rgba(245,158,11,0.18)" : "#fef3c7";
      color = dark ? "#fde68a" : "#92400e";
    } else if (["bugün bitiyor", "gösterildi"].includes(status)) {
      bg = dark ? "rgba(45,212,191,0.18)" : "#ccfbf1";
      color = dark ? "#99f6e4" : "#0f766e";
    } else if (["süresi doldu", "iptal", "iptal edildi"].includes(status)) {
      bg = dark ? "rgba(248,113,113,0.18)" : "#fee2e2";
      color = dark ? "#fecaca" : "#991b1b";
    } else if (["satıldı", "kiralandı", "kapandı", "tamamlandı"].includes(status)) {
      bg = dark ? "rgba(96,165,250,0.18)" : "#dbeafe";
      color = dark ? "#bfdbfe" : "#1d4ed8";
    } else if (status === "önerildi") {
      bg = dark ? "rgba(129,140,248,0.18)" : "#e0e7ff";
      color = dark ? "#c7d2fe" : "#3730a3";
    } else if (status === "pazarlık") {
      bg = dark ? "rgba(249,115,22,0.18)" : "#ffedd5";
      color = dark ? "#fed7aa" : "#9a3412";
    } else if (["alıcı", "satıcı", "kiracı", "kiraya veren"].includes(status)) {
      bg = dark ? "rgba(96,165,250,0.18)" : "#e0f2fe";
      color = dark ? "#bfdbfe" : "#075985";
    } else if (["satılık", "kiralık"].includes(status)) {
      bg = dark ? "rgba(168,85,247,0.18)" : "#f3e8ff";
      color = dark ? "#e9d5ff" : "#7e22ce";
    }

    return {
      background: bg,
      color,
      padding: "7px 12px",
      borderRadius: 999,
      fontSize: 12,
      fontWeight: 800,
      display: "inline-flex",
      alignItems: "center",
      gap: 6,
      whiteSpace: "nowrap",
      border: `1px solid ${dark ? "rgba(255,255,255,0.04)" : "rgba(15,23,42,0.04)"}`,
    };
  }

  function nextPortfolioNo() {
    const year = new Date().getFullYear();
    const prefix = `PRT-${year}-`;
    let max = 0;

    portfolios.forEach((p) => {
      if (typeof p.portfolioNo === "string" && p.portfolioNo.startsWith(prefix)) {
        const n = Number(p.portfolioNo.split("-")[2]);
        if (n > max) max = n;
      }
    });

    return `${prefix}${String(max + 1).padStart(3, "0")}`;
  }

  function nextContractNo() {
    const year = new Date().getFullYear();
    const prefix = `KEY-${year}-`;
    let max = 0;

    contracts.forEach((c) => {
      if (typeof c.contractNo === "string" && c.contractNo.startsWith(prefix)) {
        const n = Number(c.contractNo.split("-")[2]);
        if (n > max) max = n;
      }
    });

    return `${prefix}${String(max + 1).padStart(3, "0")}`;
  }

  function getCustomerNameById(id) {
    return customers.find((c) => c.id === id)?.fullName || "-";
  }

  function getPortfolioTitleById(id) {
    return portfolios.find((p) => p.id === id)?.title || "-";
  }

  function startEditCustomer(item) {
    setEditingCustomerId(item.id);
    setCustomerForm({
      fullName: item.fullName || "",
      phone: item.phone || "",
      email: item.email || "",
      customerType: item.customerType || "alıcı",
      interestedLocation: item.interestedLocation || "",
      interestedPropertyType: item.interestedPropertyType || "",
      budgetMin: item.budgetMin || "",
      budgetMax: item.budgetMax || "",
      note: item.note || "",
    });
    setActiveTab("customers");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function startEditPortfolio(item) {
    setEditingPortfolioId(item.id);
    setPortfolioForm({
      portfolioNo: item.portfolioNo || "",
      title: item.title || "",
      portfolioType: item.portfolioType || "satılık",
      propertyType: item.propertyType || "",
      price: item.price || "",
      roomCount: item.roomCount || "",
      grossArea: item.grossArea || "",
      city: item.city || "",
      district: item.district || "",
      address: item.address || "",
      ownerName: item.ownerName || "",
      ownerPhone: item.ownerPhone || "",
      imageUrl: item.imageUrl || "",
      description: item.description || "",
      status: item.status || "aktif",
      note: item.note || "",
    });
    setActiveTab("portfolios");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function startEditAppointment(item) {
    setEditingAppointmentId(item.id);
    setAppointmentForm({
      title: item.title || "",
      customerId: item.customerId || "",
      portfolioId: item.portfolioId || "",
      date: item.date || "",
      time: item.time || "",
      location: item.location || "",
      status: item.status || "planlandı",
      note: item.note || "",
    });
    setActiveTab("appointments");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function startEditContract(item) {
    setEditingContractId(item.id);
    setContractForm({
      contractNo: item.contractNo || "",
      companyName: item.companyName || "",
      agentName: item.agentName || "",
      owner: item.owner || "",
      phone: item.phone || "",
      address: item.address || "",
      propertyType: item.propertyType || "",
      salePrice: item.salePrice || "",
      commission: item.commission || "",
      startDate: item.startDate || "",
      endDate: item.endDate || "",
      status: item.status || "aktif",
      note: item.note || "",
    });
    setActiveTab("contracts");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  const filteredCustomers = useMemo(() => {
    const q = customerSearch.toLowerCase().trim();

    return customers.filter((c) => {
      const hay = [
        c.fullName,
        c.phone,
        c.email,
        c.customerType,
        c.interestedLocation,
        c.interestedPropertyType,
        c.budgetMin,
        c.budgetMax,
        c.note,
      ]
        .join(" ")
        .toLowerCase();

      const okSearch = q ? hay.includes(q) : true;
      const okType = customerTypeFilter === "tümü" ? true : c.customerType === customerTypeFilter;
      return okSearch && okType;
    });
  }, [customers, customerSearch, customerTypeFilter]);

  const filteredPortfolios = useMemo(() => {
    const q = portfolioSearch.toLowerCase().trim();

    return portfolios.filter((p) => {
      const hay = [
        p.portfolioNo,
        p.title,
        p.portfolioType,
        p.propertyType,
        p.price,
        p.city,
        p.district,
        p.address,
        p.ownerName,
        p.ownerPhone,
        p.description,
        p.status,
        p.note,
      ]
        .join(" ")
        .toLowerCase();

      const okSearch = q ? hay.includes(q) : true;
      const okType = portfolioTypeFilter === "tümü" ? true : p.portfolioType === portfolioTypeFilter;
      return okSearch && okType;
    });
  }, [portfolios, portfolioSearch, portfolioTypeFilter]);

  const filteredAppointments = useMemo(() => {
    const q = appointmentSearch.toLowerCase().trim();

    return appointments
      .filter((a) => {
        const hay = [
          a.title,
          a.location,
          a.note,
          a.status,
          a.date,
          a.time,
          getCustomerNameById(a.customerId),
          getPortfolioTitleById(a.portfolioId),
        ]
          .join(" ")
          .toLowerCase();

        const okSearch = q ? hay.includes(q) : true;
        const okStatus =
          appointmentStatusFilter === "tümü" ? true : a.status === appointmentStatusFilter;

        return okSearch && okStatus;
      })
      .sort((a, b) => `${a.date} ${a.time}`.localeCompare(`${b.date} ${b.time}`));
  }, [appointments, appointmentSearch, appointmentStatusFilter, customers, portfolios]);

  const filteredContracts = useMemo(() => {
    const q = contractSearch.toLowerCase().trim();

    return contracts
      .map((c) => ({ ...c, calculatedStatus: getContractStatus(c) }))
      .filter((c) => {
        const hay = [
          c.contractNo,
          c.companyName,
          c.agentName,
          c.owner,
          c.phone,
          c.address,
          c.propertyType,
          c.salePrice,
          c.note,
          c.calculatedStatus,
        ]
          .join(" ")
          .toLowerCase();

        const okSearch = q ? hay.includes(q) : true;
        const okStatus =
          contractStatusFilter === "tümü" ? true : c.calculatedStatus === contractStatusFilter;

        return okSearch && okStatus;
      });
  }, [contracts, contractSearch, contractStatusFilter]);

  const customerOptions = [
    { value: "", label: "Müşteri seç" },
    ...customers.map((c) => ({
      value: c.id,
      label: `${c.fullName} • ${c.phone || "-"}`,
    })),
  ];

  const portfolioOptions = [
    { value: "", label: "Portföy seç" },
    ...portfolios.map((p) => ({
      value: p.id,
      label: `${p.title} • ${p.portfolioNo || "-"}`,
    })),
  ];

  const totalCustomers = customers.length;
  const totalPortfolios = portfolios.length;
  const totalMatches = matches.length;
  const activeContracts = contracts.filter((c) => getContractStatus(c) === "aktif").length;
  const expiringContracts = contracts.filter((c) =>
    ["yakında bitecek", "bugün bitiyor", "süresi doldu"].includes(getContractStatus(c))
  ).length;
  const todayAppointments = appointments.filter((a) => a.date === todayStr).length;
  const upcomingAppointments = appointments.filter((a) => a.date >= todayStr).length;
  const nextFiveAppointments = filteredAppointments.filter((a) => a.date >= todayStr).slice(0, 5);
  const dashboardAgendaItems = nextFiveAppointments.slice(0, 4).map((item) => ({
    ...item,
    visualStatus: getAppointmentVisualStatus(item),
  }));
  const criticalContracts = filteredContracts
    .filter((c) => ["yakında bitecek", "bugün bitiyor", "süresi doldu"].includes(c.calculatedStatus))
    .slice(0, 5);
  const totalPortfolioValue = portfolios.reduce((sum, item) => sum + parsePrice(item.price), 0);
  const closedPortfolioCount = portfolios.filter((item) => ["satıldı", "kiralandı"].includes(item.status)).length;
  const completedMatches = matches.filter((item) => item.status === "tamamlandı").length;
  const completedAppointments = appointments.filter((item) => item.status === "tamamlandı").length;
  const appointmentRows = appointments
    .slice()
    .sort((a, b) => `${a.date || ""} ${a.time || ""}`.localeCompare(`${b.date || ""} ${b.time || ""}`));
  const overdueAppointments = appointments.filter((item) => getAppointmentVisualStatus(item) === "gecikti").length;
  const appointmentThisMonth = appointments.filter((item) => {
    if (!item.date) return false;
    const date = new Date(`${item.date}T00:00:00`);
    const now = new Date();
    return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
  }).length;

  function parsePrice(value) {
    const number = Number(String(value || "").replace(/[^\d]/g, ""));
    return Number.isFinite(number) ? number : 0;
  }

  function formatShortDate(dateStr) {
    if (!dateStr) return "-";
    const date = new Date(dateStr);
    if (Number.isNaN(date.getTime())) return dateStr;
    return date.toLocaleDateString("tr-TR", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  }

  function formatLongDateTime(dateStr, timeStr = "") {
    if (!dateStr) return "-";
    const date = new Date(`${dateStr}T${timeStr || "00:00"}:00`);
    if (Number.isNaN(date.getTime())) return `${dateStr} ${timeStr}`.trim();
    return (
      date.toLocaleDateString("tr-TR", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      }) + (timeStr ? ` • ${timeStr}` : "")
    );
  }

  function getRentAgreementFieldValues() {
    const annualRentValue =
      rentAgreementForm.annualRent ||
      (parsePrice(rentAgreementForm.monthlyRent)
        ? String(parsePrice(rentAgreementForm.monthlyRent) * 12)
        : "");
    const paymentBits = [
      rentAgreementForm.paymentMethod,
      rentAgreementForm.paymentDay ? `Her ay ${rentAgreementForm.paymentDay}. gün` : "",
    ].filter(Boolean);
    const specialConditions = [
      rentAgreementForm.note,
      rentAgreementForm.increaseRate ? `Kira artış oranı: %${rentAgreementForm.increaseRate}` : "",
      rentAgreementForm.endDate ? `Kira bitiş tarihi: ${formatShortDate(rentAgreementForm.endDate)}` : "",
      rentAgreementForm.guarantorAddress ? `Kefil adresi: ${rentAgreementForm.guarantorAddress}` : "",
    ].filter(Boolean).join("\n");

    return {
      mahalle: rentAgreementForm.propertyNeighborhood || "-",
      cadde: rentAgreementForm.propertyStreet || "-",
      numara: rentAgreementForm.propertyDoorNo || "-",
      cins: rentAgreementForm.propertyKind || rentAgreementForm.propertyTitle || "-",
      kiralayan: rentAgreementForm.landlordName || "-",
      kirayaVerenTc: rentAgreementForm.landlordTcNo || "-",
      kirayaVerenAdres: rentAgreementForm.landlordAddress || officeProfile.address || "-",
      kiraci: rentAgreementForm.tenantName || "-",
      kiraciTc: rentAgreementForm.tenantTcNo || "-",
      kiraciAdres: rentAgreementForm.tenantAddress || "-",
      kiraBaslangic: formatShortDate(rentAgreementForm.startDate || rentAgreementForm.contractDate),
      kiraSuresi: rentAgreementForm.rentDuration || "-",
      yillikKira: formatCurrency(annualRentValue),
      aylikKira: formatCurrency(rentAgreementForm.monthlyRent),
      odemeSekli: paymentBits.join(" • ") || "-",
      kullanimSekli: rentAgreementForm.usageType || "-",
      durumu: rentAgreementForm.propertyCondition || "-",
      demirbaslar: rentAgreementForm.fixtures || "-",
      kirayaVerenImza: rentAgreementForm.landlordName || "Kiraya Veren",
      kiraciImza: rentAgreementForm.tenantName || "Kiracı",
      kefilImza: rentAgreementForm.guarantorName || "Kefil",
      ozelKosullar: specialConditions,
    };
  }

  function combineDateTime(dateStr, timeStr = "00:00") {
    return new Date(`${dateStr || todayStr}T${timeStr || "00:00"}:00`);
  }

  function getAppointmentVisualStatus(item) {
    if (!item) return "-";
    if (item.status === "tamamlandı" || item.status === "iptal") return item.status;
    if (!item.date) return item.status || "planlandı";
    const now = new Date();
    const appointmentDate = combineDateTime(item.date, item.time || "00:00");
    if (appointmentDate < now) return "gecikti";
    if (item.date === todayStr) return "bugün";
    return item.status || "planlandı";
  }

  function openCustomerDetail(item) {
    setSelectedCustomer(item);
    setActiveTab("customers");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function openPortfolioDetail(item) {
    setSelectedPortfolio(item);
    setActiveTab("portfolios");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function getPortfolioMatchCount(portfolioId) {
    return matches.filter((item) => item.portfolioId === portfolioId).length;
  }

  function getPortfolioAppointmentCount(portfolioId) {
    return appointments.filter((item) => item.portfolioId === portfolioId).length;
  }

  function getPortfolioDaysActive(item) {
    const createdAt = item?.createdAt ? new Date(item.createdAt) : null;
    if (!createdAt || Number.isNaN(createdAt.getTime())) return 0;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return Math.max(0, Math.ceil((today - createdAt) / (1000 * 60 * 60 * 24)));
  }

  function getCustomerMatchCount(customerId) {
    return matches.filter((item) => item.customerId === customerId).length;
  }

  function getCustomerAppointmentCount(customerId) {
    return appointments.filter((item) => item.customerId === customerId).length;
  }

  function normalizeMatchText(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/ı/g, "i")
      .replace(/ğ/g, "g")
      .replace(/ü/g, "u")
      .replace(/ş/g, "s")
      .replace(/ö/g, "o")
      .replace(/ç/g, "c")
      .replace(/\s+/g, " ")
      .trim();
  }

  function getCustomerTargetPortfolioType(customer) {
    if (customer?.customerType === "alıcı") return "satılık";
    if (customer?.customerType === "kiracı") return "kiralık";
    return "";
  }

  function alreadyMatchedCustomerPortfolio(customerId, portfolioId) {
    return matches.some((item) => item.customerId === customerId && item.portfolioId === portfolioId);
  }

  function getMatchScoreAppearance(score = 0) {
    if (score >= 85) {
      return {
        label: "Çok güçlü",
        background: dark ? "rgba(34,197,94,0.18)" : "#dcfce7",
        color: dark ? "#bbf7d0" : "#166534",
        borderColor: dark ? "rgba(34,197,94,0.32)" : "#86efac",
      };
    }
    if (score >= 70) {
      return {
        label: "Güçlü",
        background: dark ? "rgba(59,130,246,0.18)" : "#dbeafe",
        color: dark ? "#bfdbfe" : "#1d4ed8",
        borderColor: dark ? "rgba(59,130,246,0.32)" : "#93c5fd",
      };
    }
    if (score >= 55) {
      return {
        label: "Orta",
        background: dark ? "rgba(245,158,11,0.18)" : "#fef3c7",
        color: dark ? "#fde68a" : "#92400e",
        borderColor: dark ? "rgba(245,158,11,0.32)" : "#fcd34d",
      };
    }
    return {
      label: "Zayıf",
      background: dark ? "rgba(248,113,113,0.18)" : "#fee2e2",
      color: dark ? "#fecaca" : "#991b1b",
      borderColor: dark ? "rgba(248,113,113,0.32)" : "#fca5a5",
    };
  }

  function getReasonChipPalette(kind = "positive") {
    if (kind === "negative") {
      return {
        background: dark ? "rgba(248,113,113,0.14)" : "#fef2f2",
        color: dark ? "#fecaca" : "#991b1b",
        borderColor: dark ? "rgba(248,113,113,0.24)" : "#fecaca",
      };
    }
    if (kind === "neutral") {
      return {
        background: dark ? "rgba(148,163,184,0.14)" : "#f8fafc",
        color: dark ? "#cbd5e1" : "#475569",
        borderColor: dark ? "rgba(148,163,184,0.22)" : "#cbd5e1",
      };
    }
    return {
      background: dark ? "rgba(34,197,94,0.14)" : "#f0fdf4",
      color: dark ? "#bbf7d0" : "#166534",
      borderColor: dark ? "rgba(34,197,94,0.22)" : "#bbf7d0",
    };
  }

  function renderReasonChips(items = [], kind = "positive") {
    if (!items?.length) return null;
    const palette = getReasonChipPalette(kind);
    return (
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {items.map((reason, index) => (
          <span
            key={`${kind}-${index}-${reason}`}
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              padding: "6px 9px",
              borderRadius: 999,
              fontSize: 11.5,
              fontWeight: 800,
              lineHeight: 1.2,
              background: palette.background,
              color: palette.color,
              border: `1px solid ${palette.borderColor}`,
            }}
          >
            {kind === "negative" ? "⚠" : kind === "neutral" ? "•" : "✓"} {reason}
          </span>
        ))}
      </div>
    );
  }

  function getCustomerPortfolioCompatibility(customer, portfolio) {
    const targetPortfolioType = getCustomerTargetPortfolioType(customer);
    if (!targetPortfolioType || portfolio?.portfolioType !== targetPortfolioType) {
      return { score: 0, reasons: [], mismatchReasons: ["İşlem tipi uyuşmuyor"] };
    }

    if (["satıldı", "kiralandı", "kapandı"].includes(portfolio?.status)) {
      return { score: 0, reasons: [], mismatchReasons: ["Portföy pasif durumda"] };
    }

    let score = 18;
    const reasons = [];
    const mismatchReasons = [];

    const customerLocation = normalizeMatchText(customer?.interestedLocation);
    const portfolioLocation = normalizeMatchText(
      [portfolio?.city, portfolio?.district, portfolio?.address, portfolio?.title].filter(Boolean).join(" ")
    );

    if (!customerLocation) {
      score += 8;
      mismatchReasons.push("Müşteri bölge tercihi girmemiş");
    } else {
      const locationTokens = customerLocation.split(/[\s,/-]+/).filter((token) => token.length > 2);
      const locationHits = locationTokens.filter((token) => portfolioLocation.includes(token)).length;
      if (!locationHits) return { score: 0, reasons: [], mismatchReasons: ["Bölge uyuşmuyor"] };
      score += locationHits >= 2 ? 28 : 18;
      reasons.push(locationHits >= 2 ? "Bölge uyumu güçlü" : "Bölge uyumu");
      if (locationHits === 1) mismatchReasons.push("Bölge eşleşmesi kısmi kaldı");
    }

    const portfolioPrice = parsePrice(portfolio?.price);
    const budgetMin = parsePrice(customer?.budgetMin);
    const budgetMax = parsePrice(customer?.budgetMax || customer?.budgetMin);

    if (portfolioPrice && (budgetMin || budgetMax)) {
      const effectiveMin = budgetMin || 0;
      const effectiveMax = budgetMax || budgetMin || 0;
      const insideRange = portfolioPrice >= effectiveMin && (!effectiveMax || portfolioPrice <= effectiveMax);
      const nearRange = portfolioPrice >= effectiveMin * 0.85 && (!effectiveMax || portfolioPrice <= effectiveMax * 1.1);
      if (insideRange) {
        score += 34;
        reasons.push("Bütçe uyumu");
      } else if (nearRange) {
        score += 18;
        reasons.push("Bütçeye yakın");
        mismatchReasons.push("Fiyat bütçeye yakın ama tam içinde değil");
      } else {
        return { score: 0, reasons: [], mismatchReasons: ["Fiyat bütçe aralığının dışında"] };
      }
    } else {
      score += 8;
      mismatchReasons.push("Bütçe bilgisi eksik");
    }

    const requestedPropertyType = normalizeMatchText(customer?.interestedPropertyType);
    const portfolioPropertyType = normalizeMatchText(portfolio?.propertyType);
    if (requestedPropertyType && portfolioPropertyType) {
      if (
        portfolioPropertyType.includes(requestedPropertyType) ||
        requestedPropertyType.includes(portfolioPropertyType)
      ) {
        score += 24;
        reasons.push("Mülk tipi uyumu");
      } else {
        return { score: 0, reasons: [], mismatchReasons: ["Mülk tipi uyuşmuyor"] };
      }
    } else if (requestedPropertyType || portfolioPropertyType) {
      score += 6;
      mismatchReasons.push(requestedPropertyType ? "Portföyde mülk tipi detayı zayıf" : "Müşteri mülk tipi belirtmemiş");
    } else {
      mismatchReasons.push("Mülk tipi bilgisi eksik");
    }

    return {
      score: Math.min(100, score),
      reasons: reasons.length ? reasons : ["Temel uyum"],
      mismatchReasons,
    };
  }

  function getRecommendedPortfoliosForCustomer(customer) {
    return portfolios
      .map((portfolio) => {
        const compatibility = getCustomerPortfolioCompatibility(customer, portfolio);
        return {
          ...portfolio,
          _matchScore: compatibility.score,
          _matchReasons: compatibility.reasons,
          _matchMismatchReasons: compatibility.mismatchReasons || [],
        };
      })
      .filter((portfolio) => portfolio._matchScore > 0 && !alreadyMatchedCustomerPortfolio(customer?.id, portfolio.id))
      .sort((a, b) => b._matchScore - a._matchScore || parsePrice(b.price) - parsePrice(a.price))
      .slice(0, 5);
  }

  function getRecommendedCustomersForPortfolio(item) {
    return customers
      .map((customer) => {
        const compatibility = getCustomerPortfolioCompatibility(customer, item);
        return {
          ...customer,
          _matchScore: compatibility.score,
          _matchReasons: compatibility.reasons,
          _matchMismatchReasons: compatibility.mismatchReasons || [],
        };
      })
      .filter((customer) => customer._matchScore > 0 && !alreadyMatchedCustomerPortfolio(customer.id, item?.id))
      .sort((a, b) => b._matchScore - a._matchScore || parsePrice(b.budgetMax || b.budgetMin) - parsePrice(a.budgetMax || a.budgetMin))
      .slice(0, 5);
  }

  function getAutoMatchSuggestions(limit = 8) {
    const pairs = [];
    customers.forEach((customer) => {
      getRecommendedPortfoliosForCustomer(customer).forEach((portfolio) => {
        pairs.push({
          customer,
          portfolio,
          score: portfolio._matchScore || 0,
          reasons: portfolio._matchReasons || [],
          mismatchReasons: portfolio._matchMismatchReasons || [],
        });
      });
    });

    return pairs
      .sort((a, b) => b.score - a.score || parsePrice(b.portfolio?.price) - parsePrice(a.portfolio?.price))
      .slice(0, limit);
  }

  function getCurrentWeekDates() {

    const now = new Date();
    const mondayOffset = (now.getDay() + 6) % 7;
    const monday = new Date(now);
    monday.setDate(now.getDate() - mondayOffset);
    monday.setHours(0, 0, 0, 0);
    return Array.from({ length: 7 }, (_, index) => {
      const date = new Date(monday);
      date.setDate(monday.getDate() + index);
      return date;
    });
  }

  function renewContractFromItem(item) {
    const start = new Date();
    const end = new Date();
    end.setDate(end.getDate() + 90);
    setEditingContractId(null);
    setContractForm({
      contractNo: "",
      companyName: item.companyName || officeProfile.officeName || "",
      agentName: item.agentName || officeProfile.agentName || "",
      owner: item.owner || "",
      phone: item.phone || "",
      address: item.address || "",
      propertyType: item.propertyType || "",
      salePrice: item.salePrice || "",
      commission: item.commission || "",
      startDate: start.toISOString().slice(0, 10),
      endDate: end.toISOString().slice(0, 10),
      status: "aktif",
      note: `${item.contractNo || ""} sözleşmesinden yenileme taslağı`,
    });
    setActiveTab("contracts");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function getOfficeDisplayName() {
    return officeProfile.officeName || contractForm.companyName || user?.email || "e-key Emlak";
  }

  function handleOfficeLogoUpload(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      setOfficeProfile((prev) => ({
        ...prev,
        logoDataUrl: String(reader.result || ""),
        logoName: file.name || "logo",
      }));
    };
    reader.readAsDataURL(file);
  }

  async function addPdfLogo(pdf, x = 15, y = 12, width = 24, height = 24) {
    if (!officeProfile.logoDataUrl) return;
    try {
      const format = officeProfile.logoDataUrl.includes("image/png") ? "PNG" : "JPEG";
      pdf.addImage(officeProfile.logoDataUrl, format, x, y, width, height);
    } catch (e) {
      console.warn("Logo PDF'e eklenemedi", e);
    }
  }

  async function drawPdfHeader(pdf, title, subtitle = "") {
    await addPdfLogo(pdf);
    const startX = officeProfile.logoDataUrl ? 46 : 18;
    pdf.setFontSize(18);
    pdf.setFont(undefined, "bold");
    pdf.text(title, startX, 20);
    if (subtitle) {
      pdf.setFontSize(10);
      pdf.setFont(undefined, "normal");
      pdf.text(subtitle, startX, 27);
    }
    pdf.setFontSize(11);
    const headerLines = [
      getOfficeDisplayName(),
      officeProfile.agentName ? `Danışman: ${officeProfile.agentName}` : "",
      officeProfile.phone ? `Tel: ${officeProfile.phone}` : "",
      officeProfile.email ? `E-posta: ${officeProfile.email}` : "",
      officeProfile.address ? `Adres: ${officeProfile.address}` : "",
    ].filter(Boolean);
    let y = 38;
    headerLines.forEach((line) => {
      const wrapped = pdf.splitTextToSize(line, 170);
      pdf.text(wrapped, 18, y);
      y += wrapped.length * 5;
    });
    pdf.setDrawColor(200, 210, 220);
    pdf.line(18, y + 2, 192, y + 2);
    return y + 10;
  }

  function drawPdfField(pdf, label, value, y) {
    pdf.setFont(undefined, "bold");
    pdf.text(`${label}:`, 18, y);
    pdf.setFont(undefined, "normal");
    const wrapped = pdf.splitTextToSize(String(value || "-"), 135);
    pdf.text(wrapped, 58, y);
    return y + Math.max(6, wrapped.length * 6);
  }

  function buildGenericContractSvg({ title, subtitle, rows, footer, contractNo, contractDate, signerLeftLabel, signerLeft, signerMiddleLabel, signerMiddle, signerRightLabel, signerRight }) {
    const W = 794;
    const H = 1123;
    const mX = 36;
    const usableW = W - mX * 2;
    const labelW = 220;
    const valueW = usableW - labelW;
    const hasLogo = !!officeProfile.logoDataUrl;
    const textStartX = hasLogo ? 128 : mX;
    const parts = [];

    // Arka plan
    parts.push(`<rect width="${W}" height="${H}" fill="#f8faff"/>`);
    // Mavi header bar
    parts.push(`<rect x="0" y="0" width="${W}" height="92" fill="#1e3a8a"/>`);
    parts.push(`<rect x="0" y="88" width="${W}" height="5" fill="#3b82f6"/>`);

    // Ofis adı (beyaz, header içinde)
    const offName = officeProfile.officeName || "e-key Emlak";
    parts.push(`<text x="${textStartX}" y="34" font-family="Arial, sans-serif" font-size="17" font-weight="700" fill="#ffffff">${escapeSvgText(offName)}</text>`);

    let oy = 52;
    if (officeProfile.agentName) {
      parts.push(`<text x="${textStartX}" y="${oy}" font-family="Arial, sans-serif" font-size="11" fill="rgba(255,255,255,0.82)">${escapeSvgText("Danışman: " + officeProfile.agentName)}</text>`);
      oy += 15;
    }
    if (officeProfile.phone) {
      parts.push(`<text x="${textStartX}" y="${oy}" font-family="Arial, sans-serif" font-size="11" fill="rgba(255,255,255,0.82)">${escapeSvgText("Tel: " + officeProfile.phone)}</text>`);
      oy += 15;
    }
    if (officeProfile.email) {
      parts.push(`<text x="${textStartX}" y="${oy}" font-family="Arial, sans-serif" font-size="11" fill="rgba(255,255,255,0.75)">${escapeSvgText(officeProfile.email)}</text>`);
    }

    // Sağ üst: sözleşme no + tarih
    if (contractNo && contractNo !== "-") {
      parts.push(`<text x="${W - mX}" y="34" text-anchor="end" font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="rgba(255,255,255,0.92)">${escapeSvgText("No: " + contractNo)}</text>`);
    }
    if (contractDate && contractDate !== "-") {
      parts.push(`<text x="${W - mX}" y="52" text-anchor="end" font-family="Arial, sans-serif" font-size="11" fill="rgba(255,255,255,0.72)">${escapeSvgText("Tarih: " + contractDate)}</text>`);
    }

    // Başlık
    const titleY = 128;
    parts.push(`<text x="${W / 2}" y="${titleY}" text-anchor="middle" font-family="Arial, sans-serif" font-size="21" font-weight="700" fill="#0f172a">${escapeSvgText(title.toUpperCase())}</text>`);
    if (subtitle) {
      parts.push(`<text x="${W / 2}" y="${titleY + 22}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" fill="#64748b">${escapeSvgText(subtitle)}</text>`);
    }
    // Başlık altı çizgi
    parts.push(`<line x1="${W / 2 - 70}" y1="${titleY + 32}" x2="${W / 2 + 70}" y2="${titleY + 32}" stroke="#2563eb" stroke-width="2"/>`);

    // Tablo
    let y = titleY + 52;
    const skipKeys = ["Sözleşme No", "Sözleşme Tarihi"];
    const tableRows = rows.filter(([l]) => !skipKeys.includes(l));
    const lineH = 18;

    tableRows.forEach(([label, value], i) => {
      const labelLines = wrapSvgText(label, 28);
      const valueLines = wrapSvgText(String(value || "-"), 54);
      const rowH = Math.max(44, Math.max(labelLines.length, valueLines.length) * lineH + 16);
      const bg = i % 2 === 0 ? "#eef4ff" : "#ffffff";
      parts.push(`<rect x="${mX}" y="${y}" width="${usableW}" height="${rowH}" fill="${bg}"/>`);
      parts.push(`<rect x="${mX}" y="${y}" width="${usableW}" height="${rowH}" fill="none" stroke="#d1ddf0" stroke-width="1"/>`);
      parts.push(`<line x1="${mX + labelW}" y1="${y}" x2="${mX + labelW}" y2="${y + rowH}" stroke="#d1ddf0" stroke-width="1"/>`);
      // Etiket sol kenarda renkli çizgi
      parts.push(`<rect x="${mX}" y="${y}" width="4" height="${rowH}" fill="#3b82f6"/>`);
      parts.push(svgTextBlock(labelLines, mX + 14, y + 17, lineH, `font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="#334155"`));
      parts.push(svgTextBlock(valueLines, mX + labelW + 10, y + 17, lineH, `font-family="Arial, sans-serif" font-size="12" fill="#0f172a"`));
      y += rowH;
    });

    // Not kutusu
    if (footer) {
      y += 18;
      const noteLines = wrapSvgText(footer, 82);
      const noteH = noteLines.length * lineH + 28;
      parts.push(`<rect x="${mX}" y="${y}" width="${usableW}" height="${noteH}" fill="#fffbeb" stroke="#f59e0b" stroke-width="1" rx="5"/>`);
      parts.push(`<text x="${mX + 14}" y="${y + 18}" font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="#92400e">Not:</text>`);
      parts.push(svgTextBlock(noteLines, mX + 52, y + 18, lineH, `font-family="Arial, sans-serif" font-size="11" fill="#78350f"`));
      y += noteH;
    }

    // İmza alanı
    const sigY = Math.max(y + 40, 960);
    parts.push(`<line x1="${mX}" y1="${sigY - 10}" x2="${W - mX}" y2="${sigY - 10}" stroke="#cbd5e1" stroke-width="1"/>`);
    const sigCols = [
      [mX + 30, signerLeftLabel || "İmza 1", signerLeft || ""],
      [W / 2 - 90, signerMiddleLabel || "İmza 2", signerMiddle || ""],
      [W - mX - 210, signerRightLabel || "İmza 3", signerRight || ""],
    ];
    sigCols.forEach(([x, lbl, name]) => {
      parts.push(`<text x="${x + 90}" y="${sigY + 10}" text-anchor="middle" font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="#64748b">${escapeSvgText(lbl)}</text>`);
      parts.push(`<line x1="${x}" y1="${sigY + 46}" x2="${x + 180}" y2="${sigY + 46}" stroke="#334155" stroke-width="1.2"/>`);
      if (name) {
        parts.push(`<text x="${x + 90}" y="${sigY + 62}" text-anchor="middle" font-family="Arial, sans-serif" font-size="12" fill="#0f172a">${escapeSvgText(name)}</text>`);
      }
    });

    // Alt footer
    parts.push(`<text x="${W / 2}" y="${H - 14}" text-anchor="middle" font-family="Arial, sans-serif" font-size="10" fill="#94a3b8">Bu belge e-key CRM sistemi tarafından oluşturulmuştur.</text>`);

    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${parts.join("")}</svg>`;
  }

  function buildTermsPageSvg(termsText, title) {
    const W = 794;
    const H = 1123;
    const mX = 36;
    const usableW = W - mX * 2;
    const parts = [];
    parts.push(`<rect width="${W}" height="${H}" fill="#f8faff"/>`);
    parts.push(`<rect x="0" y="0" width="${W}" height="56" fill="#1e3a8a"/>`);
    parts.push(`<rect x="0" y="52" width="${W}" height="4" fill="#3b82f6"/>`);
    const offName = officeProfile.officeName || "e-key Emlak";
    parts.push(`<text x="${mX}" y="34" font-family="Arial, sans-serif" font-size="14" font-weight="700" fill="#ffffff">${escapeSvgText(offName)}</text>`);
    parts.push(`<text x="${W - mX}" y="34" text-anchor="end" font-family="Arial, sans-serif" font-size="12" fill="rgba(255,255,255,0.75)">${escapeSvgText(title)}</text>`);
    let y = 88;
    parts.push(`<text x="${W / 2}" y="${y}" text-anchor="middle" font-family="Arial, sans-serif" font-size="18" font-weight="700" fill="#0f172a">GENEL ŞARTLAR VE KOŞULLAR</text>`);
    parts.push(`<line x1="${W / 2 - 80}" y1="${y + 12}" x2="${W / 2 + 80}" y2="${y + 12}" stroke="#2563eb" stroke-width="2"/>`);
    y += 40;
    const lines = (termsText || "").split("\n").filter((l) => l.trim());
    lines.forEach((line) => {
      const wrapped = wrapSvgText(line.trim(), 88);
      const isNumbered = /^\d+\./.test(line.trim());
      const attr = `font-family="Arial, sans-serif" font-size="${isNumbered ? 12 : 11}" font-weight="${isNumbered ? "700" : "400"}" fill="${isNumbered ? "#0f172a" : "#334155"}"`;
      parts.push(svgTextBlock(wrapped, mX + (isNumbered ? 0 : 12), y, 17, attr));
      y += wrapped.length * 17 + 6;
      if (y > H - 60) return;
    });
    parts.push(`<text x="${W / 2}" y="${H - 14}" text-anchor="middle" font-family="Arial, sans-serif" font-size="10" fill="#94a3b8">Bu belge e-key CRM sistemi tarafından oluşturulmuştur.</text>`);
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${parts.join("")}</svg>`;
  }

  async function exportGenericContractPdf({ fileName, title, subtitle = "", rows = [], footer = "", signerLeft = "", signerMiddle = "", signerRight = "", signerLeftLabel = "Mülk Sahibi", signerMiddleLabel = "Emlakçı", signerRightLabel = "İmza" }) {
    try {
      const contractNoRow = rows.find(([l]) => l === "Sözleşme No");
      const contractDateRow = rows.find(([l]) => l === "Sözleşme Tarihi");
      const svg = buildGenericContractSvg({
        title,
        subtitle,
        rows,
        footer,
        contractNo: contractNoRow ? contractNoRow[1] : "",
        contractDate: contractDateRow ? contractDateRow[1] : "",
        signerLeftLabel,
        signerLeft,
        signerMiddleLabel,
        signerMiddle,
        signerRightLabel,
        signerRight,
      });
      const pdf = new jsPDF({ unit: "mm", format: "a4" });
      await renderSvgPageToPdf(pdf, svg);
      if (contractTemplates.yetkiTerms && contractTemplates.yetkiTerms.trim()) {
        const termsSvg = buildTermsPageSvg(contractTemplates.yetkiTerms, title);
        pdf.addPage();
        await renderSvgPageToPdf(pdf, termsSvg);
      }
      pdf.save(fileName);
    } catch (e) {
      alert("PDF hatası: " + e.message);
    }
  }

  function safePdfText(value) {
    return String(value || "-")
      .replace(/€/g, "EUR")
      .replace(/₺/g, "TL")
      .replace(/[‘’]/g, "'")
      .replace(/[“”]/g, '"');
  }

  function escapeSvgText(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function wrapSvgText(text, maxChars) {
    const normalized = safePdfText(text).replace(/\s+/g, " ").trim();
    if (!normalized) return ["-"];

    const words = normalized.split(" ");
    const lines = [];
    let current = "";

    words.forEach((word) => {
      if (!word) return;
      const candidate = current ? `${current} ${word}` : word;
      if (candidate.length <= maxChars) {
        current = candidate;
      } else {
        if (current) lines.push(current);
        if (word.length <= maxChars) {
          current = word;
        } else {
          const pieces = word.match(new RegExp(`.{1,${Math.max(6, maxChars - 2)}}`, "g")) || [word];
          pieces.forEach((piece, index) => {
            if (index === pieces.length - 1) {
              current = piece;
            } else {
              lines.push(piece);
            }
          });
        }
      }
    });

    if (current) lines.push(current);
    return lines.length ? lines : ["-"];
  }

  function svgTextBlock(lines, x, y, lineHeight, attrs = "") {
    return `<text x="${x}" y="${y}" ${attrs}>${lines
      .map((line, index) => `<tspan x="${x}" dy="${index === 0 ? 0 : lineHeight}">${escapeSvgText(line)}</tspan>`)
      .join("")}</text>`;
  }

  async function renderSvgPageToPdf(pdf, svgMarkup, addPageBefore = false) {
    if (addPageBefore) pdf.addPage();

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();

    await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = async () => {
        try {
          const scale = 2;
          const canvas = document.createElement("canvas");
          canvas.width = Math.round(img.width * scale);
          canvas.height = Math.round(img.height * scale);
          const ctx = canvas.getContext("2d");
          if (!ctx) throw new Error("Canvas context oluşturulamadı");
          ctx.setTransform(scale, 0, 0, scale, 0, 0);
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, img.width, img.height);
          ctx.drawImage(img, 0, 0);
          pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 0, pageWidth, pageHeight, undefined, "FAST");
          await addPdfLogo(pdf, 8, 7, 24, 24);
          resolve();
        } catch (error) {
          reject(error);
        }
      };
      img.onerror = reject;
      img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgMarkup)}`;
    });
  }

  function buildRentAgreementPageOneSvg() {
    const values = getRentAgreementFieldValues();
    const width = 794;
    const height = 1123;
    const left = 42;
    const top = 94;
    const tableWidth = width - left * 2;
    const labelWidth = 240;
    const valueWidth = tableWidth - labelWidth;
    const cellFontSize = 12.8;
    const cellLineHeight = 16;

    const rows = [
      ["Kiralananın Mahallesi", values.mahalle, 44, 24, 34],
      ["Kiralananın Cadde/Sokağı", values.cadde, 44, 26, 34],
      ["Kiralananın Numarası", values.numara, 42, 24, 34],
      ["Kiralananın Cinsi", values.cins, 42, 24, 34],
      ["Kiralayan", values.kiralayan, 42, 24, 34],
      ["Kiraya Verenin T.C. Kimlik No", values.kirayaVerenTc, 36, 24, 50],
      ["Kiraya Verenin Adresi", values.kirayaVerenAdres, 36, 46, 68],
      ["Kiracının Adı-Soyadı", values.kiraci, 42, 24, 34],
      ["Kiracının T.C. Kimlik No", values.kiraciTc, 36, 24, 50],
      ["Kiracının Adresi", values.kiraciAdres, 36, 46, 68],
      ["Kira Başlangıç Tarihi", values.kiraBaslangic, 42, 24, 34],
      ["Kira Süresi", values.kiraSuresi, 42, 28, 34],
      ["Yıllık Kira Bedeli", values.yillikKira, 40, 24, 34],
      ["Aylık Kira Bedeli", values.aylikKira, 40, 24, 34],
      ["Kiranın Ödeme Şekli", values.odemeSekli, 32, 34, 42],
      ["Kiralananı Kullanım Şekli", values.kullanimSekli, 30, 30, 42],
      ["Kiralananın Durumu", values.durumu, 34, 30, 42],
      ["Kiralananla Birlikte Teslim Edilen Demirbaşlar", values.demirbaslar, 30, 42, 106],
    ];

    let y = top;
    const rowSvgs = [];

    rows.forEach(([label, value, labelChars, valueChars, minHeight]) => {
      const labelLines = wrapSvgText(label, labelChars);
      const valueLines = wrapSvgText(value, valueChars);
      const rowHeight = Math.max(minHeight, labelLines.length * cellLineHeight + 18, valueLines.length * cellLineHeight + 18);

      rowSvgs.push(`<rect x="${left}" y="${y}" width="${labelWidth}" height="${rowHeight}" fill="none" stroke="#111" stroke-width="1.5" />`);
      rowSvgs.push(`<rect x="${left + labelWidth}" y="${y}" width="${valueWidth}" height="${rowHeight}" fill="none" stroke="#111" stroke-width="1.5" />`);
      rowSvgs.push(svgTextBlock(labelLines, left + 8, y + 22, cellLineHeight, `font-family="Times New Roman, serif" font-size="${cellFontSize}"`));
      rowSvgs.push(svgTextBlock(valueLines, left + labelWidth + 8, y + 22, cellLineHeight, `font-family="Times New Roman, serif" font-size="${cellFontSize}"`));
      y += rowHeight;
    });

    const officeLine = [officeProfile.officeName, officeProfile.agentName, officeProfile.phone].filter(Boolean).join(" • ");

    const signLineY = Math.max(y + 34, 980);
    const signRoleY = signLineY - 10;
    const signNameY = signLineY + 20;

    return `
      <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
        <rect width="100%" height="100%" fill="#ffffff" />
        <text x="${width / 2}" y="52" text-anchor="middle" font-family="Times New Roman, serif" font-size="25" font-weight="700">KİRA SÖZLEŞMESİ</text>
        <line x1="${width / 2 - 88}" y1="58" x2="${width / 2 + 88}" y2="58" stroke="#111" stroke-width="1.8" />
        ${officeLine ? `<text x="${width - 42}" y="34" text-anchor="end" font-family="Arial, sans-serif" font-size="11" fill="#444">${escapeSvgText(officeLine)}</text>` : ""}
        ${rowSvgs.join("")}
        <line x1="90" y1="${signLineY}" x2="210" y2="${signLineY}" stroke="#111" stroke-width="1" />
        <line x1="337" y1="${signLineY}" x2="457" y2="${signLineY}" stroke="#111" stroke-width="1" />
        <line x1="584" y1="${signLineY}" x2="704" y2="${signLineY}" stroke="#111" stroke-width="1" />
        <text x="150" y="${signRoleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="14" font-weight="700">Kiraya Veren</text>
        <text x="${width / 2}" y="${signRoleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="14" font-weight="700">Kiracı</text>
        <text x="${width - 150}" y="${signRoleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="14" font-weight="700">Kefil</text>
        <text x="150" y="${signNameY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="15">${escapeSvgText(values.kirayaVerenImza)}</text>
        <text x="${width / 2}" y="${signNameY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="15">${escapeSvgText(values.kiraciImza)}</text>
        <text x="${width - 150}" y="${signNameY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="15">${escapeSvgText(values.kefilImza)}</text>
      </svg>
    `.trim();
  }

  function buildRentAgreementTextPageSvg(title, items, extraText = "", includeSignatureArea = false, includeBottomSignatures = false) {
    const width = 794;
    const height = 1123;
    const marginX = 28;
    const usableWidth = width - marginX * 2;
    const titleY = 62;
    const titleFontSize = 24;
    const headerRuleStartX = officeProfile.logoDataUrl ? 104 : marginX;
    const textPresets = includeSignatureArea
      ? [
          { maxWidthChars: 112, bodyFontSize: 14.4, bodyLineHeight: 21, itemGap: 10 },
          { maxWidthChars: 118, bodyFontSize: 13.7, bodyLineHeight: 20, itemGap: 9 },
          { maxWidthChars: 124, bodyFontSize: 13.0, bodyLineHeight: 18.5, itemGap: 8 },
          { maxWidthChars: 130, bodyFontSize: 12.4, bodyLineHeight: 17.2, itemGap: 7 },
        ]
      : [
          { maxWidthChars: 112, bodyFontSize: 14.8, bodyLineHeight: 22.5, itemGap: 11 },
          { maxWidthChars: 118, bodyFontSize: 14.0, bodyLineHeight: 21, itemGap: 10 },
          { maxWidthChars: 124, bodyFontSize: 13.2, bodyLineHeight: 19.2, itemGap: 9 },
          { maxWidthChars: 130, bodyFontSize: 12.6, bodyLineHeight: 18, itemGap: 8 },
        ];

    const signatureLineY = includeSignatureArea ? 1018 : includeBottomSignatures ? 1010 : null;
    const contentBottomLimit = includeSignatureArea ? 350 : includeBottomSignatures ? signatureLineY - 68 : 1010;

    let chosenPreset = textPresets[textPresets.length - 1];
    let chosenWrappedItems = [];

    textPresets.some((preset) => {
      const wrappedItems = items.map((item) => wrapSvgText(item, preset.maxWidthChars));
      const estimatedHeight = wrappedItems.reduce((sum, lines) => sum + lines.length * preset.bodyLineHeight + preset.itemGap, 0);
      if (110 + estimatedHeight <= contentBottomLimit) {
        chosenPreset = preset;
        chosenWrappedItems = wrappedItems;
        return true;
      }
      return false;
    });

    if (!chosenWrappedItems.length) {
      chosenWrappedItems = items.map((item) => wrapSvgText(item, chosenPreset.maxWidthChars));
    }

    let y = 120;
    const blocks = [];
    blocks.push(`<line x1="${headerRuleStartX}" y1="86" x2="${width - marginX}" y2="86" stroke="#111" stroke-width="1.2" />`);

    chosenWrappedItems.forEach((lines) => {
      blocks.push(
        svgTextBlock(
          lines,
          marginX,
          y,
          chosenPreset.bodyLineHeight,
          `font-family="Times New Roman, serif" font-size="${chosenPreset.bodyFontSize}"`
        )
      );
      y += lines.length * chosenPreset.bodyLineHeight + chosenPreset.itemGap;
    });

    let middleBlock = "";
    let signatureBlock = "";

    if (includeSignatureArea) {
      const boxTop = Math.max(y + 24, 355);
      const desiredBoxHeight = 445;
      const boxBottomLimit = 930;
      const boxHeight = Math.max(220, Math.min(desiredBoxHeight, boxBottomLimit - boxTop));
      const boxBottom = boxTop + boxHeight;

      middleBlock += `<text x="${marginX}" y="${boxTop - 14}" font-family="Times New Roman, serif" font-size="15" font-weight="700">Ek Özel Koşullar</text>`;
      middleBlock += `<rect x="${marginX}" y="${boxTop}" width="${usableWidth}" height="${boxHeight}" fill="none" stroke="#111" stroke-width="1" />`;

      const noteLines = extraText
        ? wrapSvgText(extraText, 95)
        : ["Taraflar arasında ayrıca kararlaştırılan özel koşullar bu alana yazılabilir."];
      middleBlock += svgTextBlock(
        noteLines,
        marginX + 12,
        boxTop + 28,
        22,
        'font-family="Times New Roman, serif" font-size="13.5"'
      );

      const guideStartY = boxTop + 68;
      const guideCount = Math.max(5, Math.floor((boxHeight - 82) / 28));
      for (let i = 0; i < guideCount; i += 1) {
        const lineY = guideStartY + i * 28;
        if (lineY < boxBottom - 14) {
          middleBlock += `<line x1="${marginX + 10}" y1="${lineY}" x2="${width - marginX - 10}" y2="${lineY}" stroke="#d7d7d7" stroke-width="0.8" />`;
        }
      }
    }

    if (includeBottomSignatures || includeSignatureArea) {
      const lineY = signatureLineY;
      const roleY = lineY - 10;
      const nameY = lineY + 20;
      const dateY = includeSignatureArea ? 976 : null;

      if (dateY) {
        signatureBlock += `<text x="${marginX}" y="${dateY}" font-family="Times New Roman, serif" font-size="14">Tarih: ${escapeSvgText(formatShortDate(rentAgreementForm.contractDate))}</text>`;
      }
      signatureBlock += `<line x1="90" y1="${lineY}" x2="210" y2="${lineY}" stroke="#111" stroke-width="1" />`;
      signatureBlock += `<line x1="337" y1="${lineY}" x2="457" y2="${lineY}" stroke="#111" stroke-width="1" />`;
      signatureBlock += `<line x1="584" y1="${lineY}" x2="704" y2="${lineY}" stroke="#111" stroke-width="1" />`;
      signatureBlock += `<text x="150" y="${roleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="14" font-weight="700">Kiraya Veren</text>`;
      signatureBlock += `<text x="${width / 2}" y="${roleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="14" font-weight="700">Kiracı</text>`;
      signatureBlock += `<text x="${width - 150}" y="${roleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="14" font-weight="700">Kefil</text>`;
      signatureBlock += `<text x="150" y="${nameY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="15">${escapeSvgText(rentAgreementForm.landlordName || "Kiraya Veren")}</text>`;
      signatureBlock += `<text x="${width / 2}" y="${nameY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="15">${escapeSvgText(rentAgreementForm.tenantName || "Kiracı")}</text>`;
      signatureBlock += `<text x="${width - 150}" y="${nameY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="15">${escapeSvgText(rentAgreementForm.guarantorName || "Kefil")}</text>`;
    }

    return `
      <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
        <rect width="100%" height="100%" fill="#ffffff" />
        <text x="${width / 2}" y="${titleY}" text-anchor="middle" font-family="Times New Roman, serif" font-size="${titleFontSize}" font-weight="700">${escapeSvgText(title)}</text>
        ${blocks.join("")}
        ${middleBlock}
        ${signatureBlock}
      </svg>
    `.trim();
  }

  async function exportTemplateRentAgreementPdf(customKiraTerms) {
    try {
      const pdf = new jsPDF({ unit: "mm", format: "a4" });
      const pageOneSvg = buildRentAgreementPageOneSvg();
      const pageTwoSvg = buildRentAgreementTextPageSvg("GENEL KOŞULLAR", RENT_GENERAL_CONDITIONS, "", false, true);
      const pageThreeSvg = buildRentAgreementTextPageSvg(
        "ÖZEL KOŞULLAR",
        RENT_SPECIAL_CONDITIONS,
        getRentAgreementFieldValues().ozelKosullar,
        true
      );

      await renderSvgPageToPdf(pdf, pageOneSvg);
      await renderSvgPageToPdf(pdf, pageTwoSvg, true);
      await renderSvgPageToPdf(pdf, pageThreeSvg, true);

      if (customKiraTerms && customKiraTerms.trim()) {
        const termsSvg = buildTermsPageSvg(customKiraTerms, "Kira Sözleşmesi — Ek Şartlar");
        pdf.addPage();
        await renderSvgPageToPdf(pdf, termsSvg);
      }

      pdf.save(`kira-sozlesmesi-${rentAgreementForm.contractNo || todayStr}.pdf`);
    } catch (e) {
      alert("PDF hatası: " + e.message);
    }
  }

  function buildRentAgreementWhatsappText() {
    return [
      `Merhaba, ${rentAgreementForm.contractDate || todayStr} tarihli kira sözleşmesi özeti aşağıdadır.`,
      `Ofis: ${getOfficeDisplayName()}`,
      `Kiraya Veren: ${rentAgreementForm.landlordName || "-"}`,
      `Kiracı: ${rentAgreementForm.tenantName || "-"}`,
      `Portföy: ${rentAgreementForm.propertyTitle || "-"}`,
      `Adres: ${rentAgreementForm.propertyAddress || "-"}`,
      `Aylık Kira: ${formatCurrency(rentAgreementForm.monthlyRent)}`,
      `Depozito: ${formatCurrency(rentAgreementForm.deposit)}`,
      `Dönem: ${formatShortDate(rentAgreementForm.startDate)} - ${formatShortDate(rentAgreementForm.endDate)}`,
      rentAgreementForm.note ? `Not: ${rentAgreementForm.note}` : "",
    ].filter(Boolean).join("\n");
  }

  async function exportSaleAgreementPdf() {
    await exportGenericContractPdf({
      fileName: `satis-sozlesmesi-${saleAgreementForm.contractNo || todayStr}.pdf`,
      title: "Satış Ön Protokolü",
      subtitle: "Taraflar arasında akdedilen gayrimenkul satış sözleşmesi",
      rows: [
        ["Sözleşme No", saleAgreementForm.contractNo || "-"],
        ["Sözleşme Tarihi", formatShortDate(saleAgreementForm.contractDate)],
        ["Satıcı", saleAgreementForm.sellerName || "-"],
        ["Satıcı Telefon", saleAgreementForm.sellerPhone || "-"],
        ["Alıcı", saleAgreementForm.buyerName || "-"],
        ["Alıcı Telefon", saleAgreementForm.buyerPhone || "-"],
        ["Portföy / Başlık", saleAgreementForm.propertyTitle || "-"],
        ["Adres", saleAgreementForm.propertyAddress || "-"],
        ["Tapu Bilgisi", saleAgreementForm.titleDeedInfo || "-"],
        ["Satış Bedeli", formatCurrency(saleAgreementForm.salePrice)],
        ["Kapora", formatCurrency(saleAgreementForm.deposit)],
        ["Komisyon Oranı", `%${saleAgreementForm.commissionRate || "-"}`],
        ["Komisyon Tutarı", formatCurrency(saleAgreementForm.commissionAmount)],
        ["Teslim Tarihi", formatShortDate(saleAgreementForm.deliveryDate)],
      ],
      footer: saleAgreementForm.note || "",
      signerLeftLabel: "Satıcı",
      signerLeft: saleAgreementForm.sellerName || "",
      signerMiddleLabel: "Alıcı",
      signerMiddle: saleAgreementForm.buyerName || "",
      signerRightLabel: "Emlakçı",
      signerRight: officeProfile.agentName || getOfficeDisplayName(),
    });
  }

  async function exportRentAgreementPdf() {
    await exportTemplateRentAgreementPdf(contractTemplates.kiraTerms);
  }

  function renderOfficeProfileSettings(title = "PDF / Ofis Bilgileri") {
    return (
      <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
        {sectionTitle(title, "Her emlakçı kendi ofis bilgisi, danışman adı ve logosuyla çıktı oluşturabilir.")}
        <div style={grid()}>
          {renderTextInput({ label: "Ofis Adı", value: officeProfile.officeName, onChange: (v) => setOfficeProfile({ ...officeProfile, officeName: v }) })}
          {renderTextInput({ label: "Danışman Adı", value: officeProfile.agentName, onChange: (v) => setOfficeProfile({ ...officeProfile, agentName: v }) })}
          {renderTextInput({ label: "Telefon", value: officeProfile.phone, onChange: (v) => setOfficeProfile({ ...officeProfile, phone: v }) })}
          {renderTextInput({ label: "E-posta", value: officeProfile.email, onChange: (v) => setOfficeProfile({ ...officeProfile, email: v }) })}
          {renderTextInput({ label: "Website", value: officeProfile.website, onChange: (v) => setOfficeProfile({ ...officeProfile, website: v }) })}
          {renderTextInput({ label: "Şehir", value: officeProfile.city, onChange: (v) => setOfficeProfile({ ...officeProfile, city: v }) })}
          {renderTextInput({ label: "Vergi Dairesi", value: officeProfile.taxOffice, onChange: (v) => setOfficeProfile({ ...officeProfile, taxOffice: v }) })}
          {renderTextInput({ label: "Vergi No", value: officeProfile.taxNo, onChange: (v) => setOfficeProfile({ ...officeProfile, taxNo: v }) })}
        </div>
        <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Adres", value: officeProfile.address, onChange: (v) => setOfficeProfile({ ...officeProfile, address: v }) })}</div>
        <div style={{ marginTop: 12, display: "grid", gap: 8 }}>
          <span style={{ fontWeight: 800, fontSize: 13, color: colors.sub }}>Logo</span>
          <input type="file" accept="image/*" onChange={(e) => handleOfficeLogoUpload(e.target.files?.[0])} style={{ color: colors.text }} />
          {officeProfile.logoDataUrl ? (
            <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
              <img src={officeProfile.logoDataUrl} alt="logo" style={{ width: 64, height: 64, objectFit: "contain", borderRadius: 12, border: `1px solid ${colors.border}`, background: colors.panel2 }} />
              <button onClick={() => setOfficeProfile({ ...officeProfile, logoDataUrl: "", logoName: "" })} style={actionButton(false)}>Logoyu Kaldır</button>
            </div>
          ) : <div style={{ color: colors.sub, fontSize: 13 }}>Yüklenen logo PDF çıktılarında kullanılacaktır.</div>}
        </div>
      </div>
    );
  }

  async function scanContractWithAI(file, contractType) {
    if (!file) return;
    setTemplateScanLoading(true);
    try {
      const base64Data = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("Dosya okunamadı"));
        r.readAsDataURL(file);
      });
      const isImage = file.type.startsWith("image/");
      const mediaType = isImage ? file.type : "application/pdf";
      const typeLabel = contractType === "yetki" ? "Emlak Yetki Sözleşmesi" : contractType === "satis" ? "Satış Ön Protokolü" : "Kira Sözleşmesi";
      const messages = [
        {
          role: "user",
          content: [
            {
              type: isImage ? "image" : "document",
              source: { type: "base64", media_type: mediaType, data: base64Data },
            },
            {
              type: "text",
              text: `Bu belge bir ${typeLabel} belgesidir. Belgedeki genel şartlar, özel koşullar ve maddeleri çıkar. Sadece numaralı maddeleri listele, her madde ayrı satırda olsun. Madde numarası ve içeriği şu formatta yaz: "1. [madde metni]". Belge başlığı, taraf bilgileri, fiyat gibi form alanlarını YAZMA, sadece genel hüküm ve şartları yaz. Yanıtı Türkçe ver.`,
            },
          ],
        },
      ];
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1000, messages }),
      });
      const data = await response.json();
      const text = (data.content || []).map((b) => b.text || "").join("\n").trim();
      if (!text) throw new Error("Maddeler çıkarılamadı");
      if (contractType === "yetki") setContractTemplates((t) => ({ ...t, yetkiTerms: text }));
      else if (contractType === "satis") setContractTemplates((t) => ({ ...t, satisTerms: text }));
      else setContractTemplates((t) => ({ ...t, kiraTerms: text }));
      alert("✅ Sözleşme maddeleri başarıyla tarandı! Aşağıdan inceleyip düzenleyebilirsiniz.");
    } catch (e) {
      alert("Tarama hatası: " + e.message);
    } finally {
      setTemplateScanLoading(false);
    }
  }

  function renderSettings() {
    const tabs = [
      { key: "ofis", label: "🏢 Ofis Profili" },
      { key: "yetki", label: "📄 Yetki Şablonu" },
      { key: "satis", label: "🧾 Satış Şablonu" },
      { key: "kira", label: "🏘️ Kira Şablonu" },
    ];
    const termKey = templateTab === "yetki" ? "yetkiTerms" : templateTab === "satis" ? "satisTerms" : "kiraTerms";
    const defaultMap = { yetki: DEFAULT_YETKI_TERMS, satis: DEFAULT_SATIS_TERMS, kira: DEFAULT_KIRA_TERMS };
    const typeLabel = templateTab === "yetki" ? "Yetki Sözleşmesi" : templateTab === "satis" ? "Satış Ön Protokolü" : "Kira Sözleşmesi";

    return (
      <div style={{ display: "grid", gap: 18 }}>
        <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 20, boxShadow: colors.shadow }}>
          {sectionTitle("⚙️ Ayarlar & Şablonlar", "Ofis bilgilerinizi ve sözleşme şablonlarınızı buradan yönetin.")}
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 4 }}>
            {tabs.map((t) => (
              <button key={t.key} onClick={() => setTemplateTab(t.key)} style={{
                ...actionButton(false),
                background: templateTab === t.key ? (dark ? "rgba(59,130,246,0.2)" : "rgba(37,99,235,0.1)") : colors.panel2,
                border: templateTab === t.key ? `1.5px solid ${colors.primary}` : `1px solid ${colors.border}`,
                color: templateTab === t.key ? colors.primary : colors.text,
                fontWeight: templateTab === t.key ? 900 : 700,
                fontSize: 13,
              }}>{t.label}</button>
            ))}
          </div>
        </div>

        {templateTab === "ofis" && (
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 20, boxShadow: colors.shadow }}>
            {sectionTitle("Ofis Profili", "Tüm PDF çıktılarında kullanılacak ofis ve danışman bilgileriniz.")}
            <div style={grid()}>
              {renderTextInput({ label: "Ofis / Şirket Adı", value: officeProfile.officeName, onChange: (v) => setOfficeProfile({ ...officeProfile, officeName: v }) })}
              {renderTextInput({ label: "Danışman Adı", value: officeProfile.agentName, onChange: (v) => setOfficeProfile({ ...officeProfile, agentName: v }) })}
              {renderTextInput({ label: "Telefon", value: officeProfile.phone, onChange: (v) => setOfficeProfile({ ...officeProfile, phone: v }) })}
              {renderTextInput({ label: "E-posta", value: officeProfile.email, onChange: (v) => setOfficeProfile({ ...officeProfile, email: v }) })}
              {renderTextInput({ label: "Website", value: officeProfile.website, onChange: (v) => setOfficeProfile({ ...officeProfile, website: v }) })}
              {renderTextInput({ label: "Şehir", value: officeProfile.city, onChange: (v) => setOfficeProfile({ ...officeProfile, city: v }) })}
              {renderTextInput({ label: "Vergi Dairesi", value: officeProfile.taxOffice, onChange: (v) => setOfficeProfile({ ...officeProfile, taxOffice: v }) })}
              {renderTextInput({ label: "Vergi No", value: officeProfile.taxNo, onChange: (v) => setOfficeProfile({ ...officeProfile, taxNo: v }) })}
            </div>
            <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Adres", value: officeProfile.address, onChange: (v) => setOfficeProfile({ ...officeProfile, address: v }) })}</div>
            <div style={{ marginTop: 16, display: "grid", gap: 10 }}>
              <span style={{ fontWeight: 800, fontSize: 13, color: colors.sub }}>Logo (PDF çıktılarında görünür)</span>
              <input type="file" accept="image/*" onChange={(e) => handleOfficeLogoUpload(e.target.files?.[0])} style={{ color: colors.text }} />
              {officeProfile.logoDataUrl ? (
                <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                  <img src={officeProfile.logoDataUrl} alt="logo" style={{ width: 72, height: 72, objectFit: "contain", borderRadius: 14, border: `1px solid ${colors.border}`, background: colors.panel2, padding: 6 }} />
                  <div>
                    <div style={{ fontWeight: 800, fontSize: 13 }}>{officeProfile.logoName || "Logo yüklendi"}</div>
                    <button onClick={() => setOfficeProfile({ ...officeProfile, logoDataUrl: "", logoName: "" })} style={{ ...actionButton(false, { fontSize: 12, marginTop: 6 }), color: colors.danger }}>Logoyu Kaldır</button>
                  </div>
                </div>
              ) : <div style={{ color: colors.sub, fontSize: 13, padding: "10px 14px", background: colors.panel2, borderRadius: 12 }}>Henüz logo yüklenmedi. Logo tüm PDF çıktılarında sol üstte görünür.</div>}
            </div>
            <div style={{ marginTop: 14, padding: "12px 16px", background: colors.primarySoft, borderRadius: 14, border: `1px solid ${dark ? "rgba(96,165,250,0.2)" : "rgba(37,99,235,0.15)"}`, fontSize: 13, color: colors.primary }}>
              ✅ Bilgileriniz otomatik kaydedilmektedir. Tüm sözleşme PDF çıktılarında bu bilgiler kullanılır.
            </div>
          </div>
        )}

        {templateTab !== "ofis" && (
          <div style={{ display: "grid", gap: 16 }}>
            <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 20, boxShadow: colors.shadow }}>
              {sectionTitle(`${typeLabel} — Mevcut Sözleşme Tara`, "Kullandığınız sözleşmeyi yükleyin, yapay zeka maddeleri otomatik çıkarsın.")}
              <div style={{ display: "grid", gap: 12 }}>
                <div style={{ padding: "14px 16px", background: colors.panel2, borderRadius: 16, border: `1px solid ${colors.border}`, fontSize: 13, lineHeight: 1.7, color: colors.sub }}>
                  📎 <strong>Desteklenen formatlar:</strong> PDF, JPG, PNG<br />
                  🤖 <strong>Yapay zeka</strong> sözleşmedeki genel şart ve hükümleri otomatik tanır<br />
                  ✏️ Taranan maddeler aşağıdaki editöre aktarılır, dilediğiniz gibi düzenleyebilirsiniz
                </div>
                <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
                  <input
                    type="file"
                    accept=".pdf,image/*"
                    id={`scan-input-${templateTab}`}
                    style={{ display: "none" }}
                    onChange={(e) => { const f = e.target.files?.[0]; if (f) scanContractWithAI(f, templateTab); e.target.value = ""; }}
                  />
                  <button
                    onClick={() => document.getElementById(`scan-input-${templateTab}`)?.click()}
                    disabled={templateScanLoading}
                    style={{ ...actionButton(true), opacity: templateScanLoading ? 0.65 : 1 }}
                  >
                    {templateScanLoading ? "⏳ Taranıyor..." : "📤 Sözleşme Yükle & Tara"}
                  </button>
                  {templateScanLoading && (
                    <span style={{ color: colors.sub, fontSize: 13 }}>Yapay zeka maddeleri çıkarıyor, lütfen bekleyin...</span>
                  )}
                </div>
              </div>
            </div>

            <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 20, boxShadow: colors.shadow }}>
              {sectionTitle(`${typeLabel} — Şablon Düzenleyici`, "Bu maddeler PDF çıktısının son sayfasında 'Genel Şartlar' olarak görünür.")}
              <textarea
                value={contractTemplates[termKey] || ""}
                onChange={(e) => setContractTemplates({ ...contractTemplates, [termKey]: e.target.value })}
                rows={18}
                style={{
                  width: "100%",
                  padding: "14px 16px",
                  borderRadius: 16,
                  border: `1.5px solid ${colors.border2}`,
                  background: colors.panel2,
                  color: colors.text,
                  outline: "none",
                  boxSizing: "border-box",
                  fontSize: 13,
                  lineHeight: 1.8,
                  resize: "vertical",
                  fontFamily: "Arial, sans-serif",
                }}
                placeholder={`${typeLabel} için genel şartları buraya yazın...\n\nÖrnek:\n1. Birinci madde metni\n2. İkinci madde metni`}
              />
              <div style={{ display: "flex", gap: 10, marginTop: 12, flexWrap: "wrap" }}>
                <button onClick={() => setContractTemplates({ ...contractTemplates, [termKey]: defaultMap[templateTab] })} style={actionButton(false, { fontSize: 13 })}>
                  🔄 Varsayılana Döndür
                </button>
                <button onClick={() => setContractTemplates({ ...contractTemplates, [termKey]: "" })} style={{ ...actionButton(false, { fontSize: 13 }), color: colors.danger }}>
                  🗑️ Temizle
                </button>
              </div>
              <div style={{ marginTop: 12, padding: "10px 14px", background: colors.primarySoft, borderRadius: 12, border: `1px solid ${dark ? "rgba(96,165,250,0.2)" : "rgba(37,99,235,0.15)"}`, fontSize: 12, color: colors.primary, lineHeight: 1.6 }}>
                ✅ Değişiklikler otomatik kaydedilir. PDF oluşturduğunuzda bu maddeler son sayfada görünür.
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }


    return (
      <div style={pageGrid(460)}>
        <div style={{ display: "grid", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Satış Sözleşme Paneli", "Satış sözleşmesini hızlı oluştur, PDF al ve WhatsApp ile paylaş.")}
            <div style={grid()}>
              {renderTextInput({ label: "Sözleşme No", value: saleAgreementForm.contractNo, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, contractNo: v }) })}
              {renderTextInput({ label: "Sözleşme Tarihi", value: saleAgreementForm.contractDate, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, contractDate: v }), type: "date" })}
              {renderTextInput({ label: "Satıcı", value: saleAgreementForm.sellerName, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, sellerName: v }) })}
              {renderTextInput({ label: "Satıcı Telefon", value: saleAgreementForm.sellerPhone, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, sellerPhone: v }) })}
              {renderTextInput({ label: "Alıcı", value: saleAgreementForm.buyerName, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, buyerName: v }) })}
              {renderTextInput({ label: "Alıcı Telefon", value: saleAgreementForm.buyerPhone, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, buyerPhone: v }) })}
              {renderTextInput({ label: "Portföy Başlığı", value: saleAgreementForm.propertyTitle, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, propertyTitle: v }) })}
              {renderTextInput({ label: "Tapu Bilgisi", value: saleAgreementForm.titleDeedInfo, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, titleDeedInfo: v }) })}
              {renderTextInput({ label: "Satış Bedeli", value: saleAgreementForm.salePrice, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, salePrice: v }) })}
              {renderTextInput({ label: "Kapora", value: saleAgreementForm.deposit, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, deposit: v }) })}
              {renderTextInput({ label: "Komisyon Oranı (%)", value: saleAgreementForm.commissionRate, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, commissionRate: v }) })}
              {renderTextInput({ label: "Komisyon Tutarı", value: saleAgreementForm.commissionAmount, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, commissionAmount: v }) })}
              {renderTextInput({ label: "Teslim Tarihi", value: saleAgreementForm.deliveryDate, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, deliveryDate: v }), type: "date" })}
            </div>
            <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Adres", value: saleAgreementForm.propertyAddress, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, propertyAddress: v }) })}</div>
            <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Not", value: saleAgreementForm.note, onChange: (v) => setSaleAgreementForm({ ...saleAgreementForm, note: v }) })}</div>
            <div style={{ display: "grid", gap: 10, marginTop: 16 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button onClick={exportSaleAgreementPdf} style={actionButton(true)}>📥 PDF İndir</button>
                <button onClick={() => shareOnWhatsApp(saleAgreementForm.sellerPhone || saleAgreementForm.buyerPhone, buildSaleAgreementWhatsappText())} style={actionButton(false, { background: "#16a34a", color: "#ffffff", border: "none" })}>💬 WhatsApp Aç</button>
                <button onClick={() => setSaleAgreementForm(emptySaleAgreement())} style={actionButton(false)}>Temizle</button>
              </div>
              <div style={{ fontSize: 12, color: colors.sub, lineHeight: 1.6, padding: "10px 14px", background: colors.panel2, borderRadius: 14, border: `1px solid ${colors.border}` }}>
                💡 Önce <strong>PDF İndir</strong> ile sözleşmeyi indirin, ardından <strong>WhatsApp Aç</strong> ile sohbeti başlatıp PDF'i paylaşın.
              </div>
            </div>
          </div>
        </div>
        <div style={{ display: "grid", gap: 16 }}>
          {renderOfficeProfileSettings("Satış PDF Ayarları")}
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Satış Önizleme", "PDF ve WhatsApp çıktısında kullanılacak hızlı özet.")}
            <div style={{ display: "grid", gap: 10 }}>
              <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 14 }}><strong>{saleAgreementForm.propertyTitle || "Portföy başlığı"}</strong><div style={{ color: colors.sub, marginTop: 6 }}>{saleAgreementForm.propertyAddress || "Adres bilgisi"}</div></div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Satıcı</div><div style={{ fontWeight: 800, marginTop: 5 }}>{saleAgreementForm.sellerName || "-"}</div></div>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Alıcı</div><div style={{ fontWeight: 800, marginTop: 5 }}>{saleAgreementForm.buyerName || "-"}</div></div>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Satış Bedeli</div><div style={{ fontWeight: 800, marginTop: 5 }}>{formatCurrency(saleAgreementForm.salePrice)}</div></div>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Komisyon</div><div style={{ fontWeight: 800, marginTop: 5 }}>%{saleAgreementForm.commissionRate || "-"}</div></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  function renderRentalContracts() {
    const values = getRentAgreementFieldValues();
    return (
      <div style={pageGrid(480)}>
        <div style={{ display: "grid", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Kira Sözleşme Paneli", "Yüklediğin örnek şablona yakın, 3 sayfalı kira sözleşmesi PDF'i üretir.")}
            <div style={{ display: "grid", gap: 18 }}>
              <div>
                <div style={{ fontWeight: 900, fontSize: 13, color: colors.sub, marginBottom: 10 }}>Sözleşme ve taşınmaz bilgileri</div>
                <div style={grid()}>
                  {renderTextInput({ label: "Sözleşme No", value: rentAgreementForm.contractNo, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, contractNo: v }) })}
                  {renderTextInput({ label: "Sözleşme Tarihi", value: rentAgreementForm.contractDate, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, contractDate: v }), type: "date" })}
                  {renderTextInput({ label: "Kiralananın Mahallesi", value: rentAgreementForm.propertyNeighborhood, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyNeighborhood: v }) })}
                  {renderTextInput({ label: "Kiralananın Cadde/Sokağı", value: rentAgreementForm.propertyStreet, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyStreet: v }) })}
                  {renderTextInput({ label: "Kiralananın Numarası", value: rentAgreementForm.propertyDoorNo, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyDoorNo: v }) })}
                  {renderTextInput({ label: "Kiralananın Cinsi", value: rentAgreementForm.propertyKind, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyKind: v }), placeholder: "Daire, iş yeri, depo..." })}
                  {renderTextInput({ label: "Portföy Başlığı", value: rentAgreementForm.propertyTitle, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyTitle: v }) })}
                  {renderTextInput({ label: "Kullanım Şekli", value: rentAgreementForm.usageType, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, usageType: v }), placeholder: "Mesken, iş yeri..." })}
                  {renderTextInput({ label: "Kiralananın Durumu", value: rentAgreementForm.propertyCondition, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyCondition: v }), placeholder: "Boyalı, boş, bakımlı..." })}
                  {renderTextInput({ label: "Kira Süresi", value: rentAgreementForm.rentDuration, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, rentDuration: v }), placeholder: "1 yıl / 12 ay" })}
                  {renderTextInput({ label: "Başlangıç Tarihi", value: rentAgreementForm.startDate, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, startDate: v }), type: "date" })}
                  {renderTextInput({ label: "Bitiş Tarihi", value: rentAgreementForm.endDate, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, endDate: v }), type: "date" })}
                  {renderTextInput({ label: "Aylık Kira", value: rentAgreementForm.monthlyRent, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, monthlyRent: v }) })}
                  {renderTextInput({ label: "Yıllık Kira", value: rentAgreementForm.annualRent, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, annualRent: v }), placeholder: "Boşsa aylıktan hesaplanır" })}
                  {renderTextInput({ label: "Depozito", value: rentAgreementForm.deposit, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, deposit: v }) })}
                  {renderTextInput({ label: "Ödeme Günü", value: rentAgreementForm.paymentDay, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, paymentDay: v }) })}
                  {renderTextInput({ label: "Ödeme Şekli", value: rentAgreementForm.paymentMethod, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, paymentMethod: v }), placeholder: "Banka havalesi, elden..." })}
                  {renderTextInput({ label: "Artış Oranı (%)", value: rentAgreementForm.increaseRate, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, increaseRate: v }) })}
                </div>
                <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Taşınmaz Açık Adresi", value: rentAgreementForm.propertyAddress, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, propertyAddress: v }) })}</div>
                <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Demirbaşlar", value: rentAgreementForm.fixtures, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, fixtures: v }), placeholder: "Klima, kombi, vitrin, masa..." })}</div>
              </div>

              <div>
                <div style={{ fontWeight: 900, fontSize: 13, color: colors.sub, marginBottom: 10 }}>Taraf bilgileri</div>
                <div style={grid()}>
                  {renderTextInput({ label: "Kiraya Veren", value: rentAgreementForm.landlordName, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, landlordName: v }) })}
                  {renderTextInput({ label: "Kiraya Veren Telefon", value: rentAgreementForm.landlordPhone, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, landlordPhone: v }) })}
                  {renderTextInput({ label: "Kiraya Veren T.C. Kimlik No", value: rentAgreementForm.landlordTcNo, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, landlordTcNo: v }) })}
                  {renderTextInput({ label: "Kiracı", value: rentAgreementForm.tenantName, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, tenantName: v }) })}
                  {renderTextInput({ label: "Kiracı Telefon", value: rentAgreementForm.tenantPhone, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, tenantPhone: v }) })}
                  {renderTextInput({ label: "Kiracı T.C. Kimlik No", value: rentAgreementForm.tenantTcNo, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, tenantTcNo: v }) })}
                  {renderTextInput({ label: "Kefil", value: rentAgreementForm.guarantorName, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, guarantorName: v }) })}
                  {renderTextInput({ label: "Kefil T.C. Kimlik No", value: rentAgreementForm.guarantorTcNo, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, guarantorTcNo: v }) })}
                </div>
                <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Kiraya Veren Adresi", value: rentAgreementForm.landlordAddress, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, landlordAddress: v }) })}</div>
                <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Kiracı Adresi", value: rentAgreementForm.tenantAddress, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, tenantAddress: v }) })}</div>
                <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Kefil Adresi", value: rentAgreementForm.guarantorAddress, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, guarantorAddress: v }) })}</div>
              </div>

              <div>
                <div style={{ fontWeight: 900, fontSize: 13, color: colors.sub, marginBottom: 10 }}>Ek özel koşullar</div>
                {renderTextarea({ label: "Özel Koşullar / Not", value: rentAgreementForm.note, onChange: (v) => setRentAgreementForm({ ...rentAgreementForm, note: v }), placeholder: "3. sayfadaki özel koşullara eklenecek notlar" })}
              </div>
            </div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
              <button onClick={exportRentAgreementPdf} style={actionButton(true)}>Şablona Göre PDF Oluştur</button>
              <button onClick={() => shareOnWhatsApp(rentAgreementForm.landlordPhone || rentAgreementForm.tenantPhone, buildRentAgreementWhatsappText())} style={actionButton(false, { background: "#16a34a", color: "#ffffff", border: "none" })}>WhatsApp Gönder</button>
              <button onClick={() => setRentAgreementForm(emptyRentAgreement())} style={actionButton(false)}>Temizle</button>
            </div>
          </div>
        </div>
        <div style={{ display: "grid", gap: 16 }}>
          {renderOfficeProfileSettings("Kira PDF Ayarları")}
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Kira Önizleme", "İlk sayfada görünecek temel alanların hızlı özeti.")}
            <div style={{ display: "grid", gap: 10 }}>
              <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 14 }}>
                <strong>{rentAgreementForm.propertyTitle || values.cins || "Kiralanan başlığı"}</strong>
                <div style={{ color: colors.sub, marginTop: 6 }}>{rentAgreementForm.propertyAddress || "Adres bilgisi"}</div>
                <div style={{ color: colors.sub, marginTop: 6 }}>Mahalle: {values.mahalle} • Cadde/Sokak: {values.cadde} • No: {values.numara}</div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 10 }}>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Kiraya Veren</div><div style={{ fontWeight: 800, marginTop: 5 }}>{rentAgreementForm.landlordName || "-"}</div></div>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Kiracı</div><div style={{ fontWeight: 800, marginTop: 5 }}>{rentAgreementForm.tenantName || "-"}</div></div>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Aylık / Yıllık Kira</div><div style={{ fontWeight: 800, marginTop: 5 }}>{formatCurrency(rentAgreementForm.monthlyRent)} • {values.yillikKira}</div></div>
                <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Kira Süresi</div><div style={{ fontWeight: 800, marginTop: 5 }}>{values.kiraSuresi}</div></div>
              </div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                <div style={{ fontSize: 12, color: colors.sub }}>Ödeme Şekli</div>
                <div style={{ fontWeight: 800, marginTop: 5 }}>{values.odemeSekli}</div>
              </div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                <div style={{ fontSize: 12, color: colors.sub }}>Demirbaşlar</div>
                <div style={{ marginTop: 5, lineHeight: 1.6 }}>{rentAgreementForm.fixtures || "-"}</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  function renderTextInput({ label, value, onChange, placeholder = "", type = "text" }) {
    return (
      <label style={{ display: "grid", gap: 8 }}>
        <span style={{ fontWeight: 800, fontSize: 13, color: colors.sub }}>{label}</span>
        <input
          type={type}
          value={value ?? ""}
          placeholder={placeholder}
          onChange={(e) => onChange(e.target.value)}
          style={{
            width: "100%",
            padding: "13px 15px",
            borderRadius: 16,
            border: `1px solid ${colors.border2}`,
            background: colors.panel2,
            color: colors.text,
            outline: "none",
            boxSizing: "border-box",
            fontSize: 14,
          }}
        />
      </label>
    );
  }

  function renderSelect({ label, value, onChange, options }) {
    return (
      <label style={{ display: "grid", gap: 8 }}>
        <span style={{ fontWeight: 800, fontSize: 13, color: colors.sub }}>{label}</span>
        <select
          value={value ?? ""}
          onChange={(e) => onChange(e.target.value)}
          style={{
            width: "100%",
            padding: "13px 15px",
            borderRadius: 16,
            border: `1px solid ${colors.border2}`,
            background: colors.panel2,
            color: colors.text,
            outline: "none",
            boxSizing: "border-box",
            fontSize: 14,
          }}
        >
          {options.map((opt) => {
            if (typeof opt === "string") {
              return (
                <option key={opt} value={opt}>
                  {opt === "" ? "Seçiniz" : opt}
                </option>
              );
            }
            return (
              <option key={opt.value} value={opt.value}>
                {opt.label}
              </option>
            );
          })}
        </select>
      </label>
    );
  }

  function renderTextarea({ label, value, onChange, placeholder = "" }) {
    return (
      <label style={{ display: "grid", gap: 8 }}>
        <span style={{ fontWeight: 800, fontSize: 13, color: colors.sub }}>{label}</span>
        <textarea
          value={value ?? ""}
          placeholder={placeholder}
          onChange={(e) => onChange(e.target.value)}
          style={{
            width: "100%",
            padding: "13px 15px",
            borderRadius: 16,
            border: `1px solid ${colors.border2}`,
            background: colors.panel2,
            color: colors.text,
            outline: "none",
            boxSizing: "border-box",
            fontSize: 14,
            minHeight: 110,
            resize: "vertical",
            fontFamily: "Arial, sans-serif",
          }}
        />
      </label>
    );
  }

  function renderEmpty(title, text = "") {
    return (
      <div
        style={{
          background: colors.panel2,
          border: `1px solid ${colors.border}`,
          borderRadius: 22,
          padding: 28,
          display: "grid",
          placeItems: "center",
          textAlign: "center",
          minHeight: 180,
          boxShadow: `${colors.shadowSoft}, ${colors.ring}`,
        }}
      >
        <div style={{ maxWidth: 360 }}>
          <div
            style={{
              width: 60,
              height: 60,
              borderRadius: 20,
              margin: "0 auto 14px",
              display: "grid",
              placeItems: "center",
              fontSize: 26,
              background: `linear-gradient(135deg, ${colors.primarySoft} 0%, ${colors.panel3} 100%)`,
              border: `1px solid ${colors.border}`,
            }}
          >
            ✨
          </div>
          <div style={{ fontWeight: 900, fontSize: 18 }}>{title}</div>
          <div style={{ color: colors.sub, marginTop: 8, lineHeight: 1.6 }}>
            {text || "Burada henüz görüntülenecek kayıt yok."}
          </div>
        </div>
      </div>
    );
  }

  function actionButton(primary = false, extra = {}) {
    return {
      background: primary
        ? `linear-gradient(135deg, ${colors.primary} 0%, ${dark ? "#3b82f6" : "#1d4ed8"} 100%)`
        : colors.panel2,
      color: primary ? colors.primaryText : colors.text,
      border: primary ? "none" : `1px solid ${colors.border2}`,
      borderRadius: 18,
      padding: "12px 16px",
      cursor: "pointer",
      fontWeight: 900,
      fontSize: 14,
      letterSpacing: "0.01em",
      boxSizing: "border-box",
      minWidth: 0,
      maxWidth: "100%",
      boxShadow: primary ? colors.shadowSoft : `${colors.shadowSoft}, ${colors.ring}`,
      ...extra,
    };
  }

  function sectionTitle(text, subText = "") {
    return (
      <div style={{ marginBottom: 18, display: "grid", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span
            style={{
              width: 34,
              height: 4,
              borderRadius: 999,
              background: `linear-gradient(90deg, ${colors.primary} 0%, ${dark ? "#38bdf8" : "#60a5fa"} 100%)`,
            }}
          />
          <span style={{ color: colors.sub, fontSize: 11, fontWeight: 900, textTransform: "uppercase", letterSpacing: "0.08em" }}>
            e-key paneli
          </span>
        </div>
        <h2 style={{ margin: 0, fontSize: 24, fontWeight: 900, lineHeight: 1.15 }}>{text}</h2>
        {subText ? <p style={{ margin: 0, color: colors.sub, lineHeight: 1.7 }}>{subText}</p> : null}
      </div>
    );
  }

  function grid(two = true) {
    return {
      display: "grid",
      gridTemplateColumns: isMobile ? "1fr" : two ? "1fr 1fr" : "1fr",
      gap: 12,
    };
  }

  function pageGrid(sideWidth = 400) {
    return {
      display: "grid",
      gridTemplateColumns: isMobile ? "1fr" : `minmax(${sideWidth}px, 460px) minmax(0, 1fr)`,
      gap: 16,
      alignItems: "start",
    };
  }

  function renderAuthScreen() {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
          padding: 20,
          background: dark
            ? "linear-gradient(180deg, #07111f 0%, #0d1728 100%)"
            : "linear-gradient(180deg, #eef4ff 0%, #f8fbff 100%)",
          color: colors.text,
          fontFamily: "Arial, sans-serif",
        }}
      >
        <div
          style={{
            width: "100%",
            maxWidth: 460,
            padding: 30,
            borderRadius: 28,
            background: colors.panel,
            border: `1px solid ${colors.border}`,
            boxShadow: colors.shadow,
          }}
        >
          <div
            style={{
              width: 58,
              height: 58,
              borderRadius: 18,
              background: dark ? "#60a5fa" : "#2563eb",
              color: "#ffffff",
              display: "grid",
              placeItems: "center",
              fontWeight: 900,
              marginBottom: 18,
              fontSize: 22,
            }}
          >
            E
          </div>

          <h1 style={{ margin: 0, fontSize: 32, fontWeight: 900 }}>e-key</h1>
          <p style={{ color: colors.sub, lineHeight: 1.6, marginTop: 10 }}>
            Müşteri, portföy, sözleşme ve randevularını modern bir panel üzerinden yönet.
          </p>

          <div style={{ display: "grid", gap: 10, marginTop: 18 }}>
            <input
              name="email"
              type="email"
              placeholder="E-posta"
              value={authForm.email}
              onChange={(e) => setAuthForm({ ...authForm, email: e.target.value })}
              style={{
                width: "100%",
                padding: "13px 15px",
                borderRadius: 16,
                border: `1px solid ${colors.border2}`,
                background: colors.panel2,
                color: colors.text,
                outline: "none",
                boxSizing: "border-box",
                fontSize: 14,
              }}
            />
            <input
              name="password"
              type="password"
              placeholder="Şifre"
              value={authForm.password}
              onChange={(e) => setAuthForm({ ...authForm, password: e.target.value })}
              style={{
                width: "100%",
                padding: "13px 15px",
                borderRadius: 16,
                border: `1px solid ${colors.border2}`,
                background: colors.panel2,
                color: colors.text,
                outline: "none",
                boxSizing: "border-box",
                fontSize: 14,
              }}
            />
          </div>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
            <button
              onClick={authMode === "login" ? handleLogin : handleRegister}
              style={actionButton(true)}
            >
              {authMode === "login" ? "Giriş Yap" : "Kayıt Ol"}
            </button>
            <button
              onClick={() => setAuthMode((m) => (m === "login" ? "register" : "login"))}
              style={actionButton(false)}
            >
              {authMode === "login" ? "Kayıt ekranına geç" : "Giriş ekranına geç"}
            </button>
          </div>
        </div>
      </div>
    );
  }

  function renderDashboard() {
    const dashboardCard = (title, subText, children, extra = {}) => (
      <div
        style={{
          background: colors.panel,
          border: `1px solid ${colors.border}`,
          borderRadius: 22,
          padding: 16,
          boxShadow: colors.shadowSoft,
          ...extra,
        }}
      >
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 16, fontWeight: 900 }}>{title}</div>
          {subText ? (
            <div style={{ color: colors.sub, marginTop: 4, lineHeight: 1.55, fontSize: 12.5 }}>
              {subText}
            </div>
          ) : null}
        </div>
        {children}
      </div>
    );

    const stat = (icon, title, value, helper, accent) => {
      const map = {
        primary: { bg: colors.primarySoft, color: colors.primary },
        success: { bg: colors.successSoft, color: colors.success },
        warning: { bg: colors.warningSoft, color: colors.warning },
        danger: { bg: colors.dangerSoft, color: colors.danger },
      };
      const style = map[accent] || map.primary;

      return (
        <div
          style={{
            background: colors.panel,
            border: `1px solid ${colors.border}`,
            borderRadius: 18,
            padding: 14,
            boxShadow: colors.shadowSoft,
            minWidth: 0,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div
              style={{
                width: 42,
                height: 42,
                borderRadius: 14,
                background: style.bg,
                color: style.color,
                display: "grid",
                placeItems: "center",
                fontSize: 20,
                fontWeight: 900,
                flexShrink: 0,
              }}
            >
              {icon}
            </div>
            <div style={{ minWidth: 0, flex: 1 }}>
              <div style={{ color: colors.sub, fontWeight: 700, fontSize: 12 }}>{title}</div>
              <div style={{ fontSize: 24, fontWeight: 900, marginTop: 3, lineHeight: 1.1 }}>{value}</div>
            </div>
          </div>
          <div style={{ color: colors.sub, marginTop: 10, lineHeight: 1.45, fontSize: 12 }}>{helper}</div>
        </div>
      );
    };

    const quickButton = (label, tab, emoji, primary = false) => (
      <button
        onClick={() => {
          setActiveTab(tab);
          window.scrollTo({ top: 0, behavior: "smooth" });
        }}
        style={actionButton(primary, {
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 10,
          width: "100%",
          textAlign: "left",
          padding: "12px 14px",
          minHeight: 50,
        })}
      >
        <span style={{ display: "flex", alignItems: "center", gap: 9, minWidth: 0 }}>
          <span style={{ fontSize: 17, flexShrink: 0 }}>{emoji}</span>
          <span style={{ fontSize: 13.5, fontWeight: 800 }}>{label}</span>
        </span>
        <span style={{ opacity: 0.8, flexShrink: 0 }}>→</span>
      </button>
    );

    const miniMetric = (label, value, helper) => (
      <div
        style={{
          background: colors.panel2,
          border: `1px solid ${colors.border}`,
          borderRadius: 16,
          padding: 12,
        }}
      >
        <div style={{ fontSize: 11.5, color: colors.sub }}>{label}</div>
        <div style={{ fontWeight: 900, fontSize: 22, marginTop: 4, lineHeight: 1.1 }}>{value}</div>
        <div style={{ color: colors.sub, fontSize: 12, marginTop: 4, lineHeight: 1.45 }}>{helper}</div>
      </div>
    );

    const recentCustomers = customers
      .slice()
      .sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")))
      .slice(0, 3);

    const highAttentionPortfolios = portfolios
      .map((item) => ({
        ...item,
        matchCount: getPortfolioMatchCount(item.id),
        appointmentCount: getPortfolioAppointmentCount(item.id),
        daysActive: getPortfolioDaysActive(item),
      }))
      .sort((a, b) => {
        const aScore = a.matchCount * 3 + a.appointmentCount * 2 + (a.status === "aktif" ? 1 : 0);
        const bScore = b.matchCount * 3 + b.appointmentCount * 2 + (b.status === "aktif" ? 1 : 0);
        if (bScore !== aScore) return bScore - aScore;
        return String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
      })
      .slice(0, 3);

    const weekAppointments = getCurrentWeekDates().map((date) => {
      const iso = date.toISOString().slice(0, 10);
      return {
        iso,
        label: date.toLocaleDateString("tr-TR", { weekday: "short" }),
        shortDay: date.toLocaleDateString("tr-TR", { day: "2-digit" }),
        count: appointments.filter((item) => item.date === iso).length,
        isToday: iso === todayStr,
      };
    });

    const portfolioStatusSummary = [
      { label: "Aktif", value: portfolios.filter((item) => item.status === "aktif").length },
      { label: "Satıldı", value: portfolios.filter((item) => item.status === "satıldı").length },
      { label: "Kiralandı", value: portfolios.filter((item) => item.status === "kiralandı").length },
      { label: "Pasif", value: portfolios.filter((item) => item.status === "pasif").length },
    ];

    const focusText = expiringContracts > 0
      ? `${expiringContracts} kritik sözleşme takip bekliyor.`
      : todayAppointments > 0
        ? `Bugün ${todayAppointments} randevu planlı görünüyor.`
        : totalCustomers > 0
          ? `Sistemde ${totalCustomers} müşteri ve ${totalPortfolios} portföy kayıtlı.`
          : "İlk kayıtlarını ekleyerek paneli doldurmaya başlayabilirsin.";

    const portfolioFillRate = totalPortfolios > 0
      ? Math.min(100, Math.round((closedPortfolioCount / totalPortfolios) * 100))
      : 0;

    return (
      <div style={{ display: "grid", gap: 12 }}>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: isMobile ? "1fr" : "minmax(0, 1.45fr) minmax(320px, 0.8fr)",
            gap: 12,
            alignItems: "stretch",
          }}
        >
          <div
            style={{
              background: dark
                ? "linear-gradient(135deg, rgba(37,99,235,0.20) 0%, rgba(15,23,42,0.92) 100%)"
                : "linear-gradient(135deg, rgba(37,99,235,0.12) 0%, rgba(255,255,255,0.96) 100%)",
              border: `1px solid ${colors.border}`,
              borderRadius: 24,
              padding: isMobile ? 18 : 20,
              boxShadow: colors.shadowSoft,
              overflow: "hidden",
              position: "relative",
            }}
          >
            <div
              style={{
                position: "absolute",
                right: -34,
                top: -42,
                width: 150,
                height: 150,
                borderRadius: "50%",
                background: dark ? "rgba(96,165,250,0.12)" : "rgba(96,165,250,0.16)",
                filter: "blur(4px)",
              }}
            />

            <div style={{ position: "relative", display: "grid", gap: 14 }}>
              <div style={{ maxWidth: 760 }}>
                <div
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 8,
                    padding: "7px 11px",
                    borderRadius: 999,
                    background: colors.panel2,
                    border: `1px solid ${colors.border}`,
                    color: colors.sub,
                    fontSize: 11.5,
                    fontWeight: 900,
                    letterSpacing: "0.03em",
                    textTransform: "uppercase",
                  }}
                >
                  e-key kontrol merkezi
                </div>

                <div style={{ marginTop: 12, fontSize: isMobile ? 25 : 30, fontWeight: 900, lineHeight: 1.1 }}>
                  {officeProfile.officeName || officeProfile.agentName || "e-key"} paneline hoş geldin
                </div>

                <div style={{ marginTop: 9, color: colors.sub, lineHeight: 1.6, fontSize: 14 }}>
                  {focusText}
                </div>
              </div>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4, minmax(0, 1fr))",
                  gap: 10,
                }}
              >
                {miniMetric("Bugün", todayAppointments, "planlanan randevu")}
                {miniMetric("Yaklaşan", upcomingAppointments, "takvim kaydı")}
                {miniMetric("Kapanış oranı", `%${portfolioFillRate}`, `${closedPortfolioCount}/${totalPortfolios} portföy`) }
                {miniMetric("Toplam değer", formatCurrency(totalPortfolioValue), "portföy toplamı")}
              </div>
            </div>
          </div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(6, minmax(0, 1fr))",
            gap: 10,
          }}
        >
          {stat("👥", "Toplam müşteri", totalCustomers, "Tüm müşteri kayıtların", "primary")}
          {stat("🏢", "Toplam portföy", totalPortfolios, "Aktif ve geçmiş portföyler", "success")}
          {stat("📄", "Aktif sözleşme", activeContracts, "Devam eden sözleşmeler", "warning")}
          {stat("🔗", "Eşleştirme", totalMatches, "Oluşturulan bağlantılar", "primary")}
          {stat("📅", "Aylık randevu", appointmentThisMonth, "Bu ay takvime eklenenler", "success")}
          {stat("⏰", "Kritik sözleşme", expiringContracts, "Takip bekleyen dosyalar", "danger")}
        </div>
      </div>
    );
  }

  function renderReports() {
    const topPortfolios = portfolios
      .map((item) => ({
        ...item,
        matchCount: getPortfolioMatchCount(item.id),
        appointmentCount: getPortfolioAppointmentCount(item.id),
      }))
      .sort((a, b) => b.matchCount + b.appointmentCount - (a.matchCount + a.appointmentCount))
      .slice(0, 6);

    const districtStats = Object.values(
      portfolios.reduce((acc, item) => {
        const key = `${item.city || "-"} / ${item.district || "-"}`;
        if (!acc[key]) acc[key] = { label: key, count: 0, totalValue: 0 };
        acc[key].count += 1;
        acc[key].totalValue += parsePrice(item.price);
        return acc;
      }, {})
    ).sort((a, b) => b.count - a.count).slice(0, 6);

    const maxDistrictCount = districtStats.length ? Math.max(...districtStats.map((item) => item.count)) : 1;
    const maxPortfolioScore = topPortfolios.length ? Math.max(...topPortfolios.map((item) => item.matchCount + item.appointmentCount)) : 1;

    const metricCard = (title, value, helper, accent = colors.primarySoft) => (
      <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 22, padding: 18, boxShadow: colors.shadow }}>
        <div style={{ color: colors.sub, fontSize: 12, fontWeight: 900, textTransform: "uppercase", letterSpacing: "0.08em" }}>{title}</div>
        <div style={{ fontSize: 28, fontWeight: 900, marginTop: 8 }}>{value}</div>
        <div style={{ color: colors.sub, marginTop: 8, lineHeight: 1.6 }}>{helper}</div>
        <div style={{ marginTop: 14, height: 6, borderRadius: 999, background: accent }} />
      </div>
    );

    return (
      <div style={{ display: "grid", gap: 16 }}>
        {sectionTitle("Raporlar", "Portföy, müşteri ve randevu verilerini yönetim bakışıyla takip et.")}
        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(4, minmax(0, 1fr))", gap: 12 }}>
          {metricCard("Toplam Portföy Değeri", formatCurrency(totalPortfolioValue), "Kayıtlı portföylerin ilan değer toplamı", colors.primarySoft)}
          {metricCard("Sonuçlanan İşlem", closedPortfolioCount, "Satılan veya kiralanan portföy sayısı", colors.successSoft)}
          {metricCard("Tamamlanan Görüşme", completedAppointments, "Tamamlandı durumuna alınan randevular", colors.warningSoft)}
          {metricCard("Tamamlanan Eşleşme", completedMatches, "Sonuçlanan müşteri-portföy eşleşmeleri", colors.dangerSoft)}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "minmax(0, 1.15fr) minmax(0, 0.85fr)", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("En Çok İlgi Gören Portföyler", "Eşleşme ve randevu toplamına göre en aktif portföyler.")}
            {topPortfolios.length === 0 ? renderEmpty("Portföy verisi yok", "Portföy ve eşleşme oluştuğunda burada performans görünür.") : (
              <div style={{ display: "grid", gap: 12 }}>
                {topPortfolios.map((item) => {
                  const score = item.matchCount + item.appointmentCount;
                  const percent = maxPortfolioScore ? Math.max(8, Math.round((score / maxPortfolioScore) * 100)) : 8;
                  return (
                    <button key={item.id} onClick={() => openPortfolioDetail(item)} style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 14, textAlign: "left", cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                        <div>
                          <div style={{ fontWeight: 900 }}>{item.title}</div>
                          <div style={{ color: colors.sub, fontSize: 13, marginTop: 4 }}>{item.city || "-"} / {item.district || "-"} • {formatCurrency(item.price)}</div>
                        </div>
                        <span style={statusBadge(item.status)}>{item.status}</span>
                      </div>
                      <div style={{ marginTop: 12, height: 8, borderRadius: 999, background: colors.panel3, overflow: "hidden", border: `1px solid ${colors.border}` }}>
                        <div style={{ width: `${percent}%`, height: "100%", background: `linear-gradient(90deg, ${colors.primary} 0%, ${dark ? "#38bdf8" : "#60a5fa"} 100%)` }} />
                      </div>
                      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginTop: 10, color: colors.sub, fontSize: 13 }}>
                        <span>{item.matchCount} eşleşme</span>
                        <span>{item.appointmentCount} randevu</span>
                        <span>{getPortfolioDaysActive(item)} gün aktif</span>
                      </div>
                    </button>
                  );
                })}
              </div>
            )}
          </div>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Yoğun Bölgeler", "Portföy yoğunluğunun en çok olduğu şehir / ilçe dağılımı.")}
            {districtStats.length === 0 ? renderEmpty("Bölge verisi yok", "Portföylerin şehir ve ilçe bilgisiyle burada analiz oluşur.") : (
              <div style={{ display: "grid", gap: 12 }}>
                {districtStats.map((item) => (
                  <div key={item.label} style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" }}>
                      <strong>{item.label}</strong>
                      <span style={statusBadge("aktif")}>{item.count} portföy</span>
                    </div>
                    <div style={{ marginTop: 12, height: 8, borderRadius: 999, background: colors.panel3, overflow: "hidden", border: `1px solid ${colors.border}` }}>
                      <div style={{ width: `${Math.max(8, Math.round((item.count / maxDistrictCount) * 100))}%`, height: "100%", background: `linear-gradient(90deg, ${colors.success} 0%, ${dark ? "#4ade80" : "#22c55e"} 100%)` }} />
                    </div>
                    <div style={{ color: colors.sub, fontSize: 13, marginTop: 8 }}>Toplam değer: {formatCurrency(item.totalValue)}</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(3, minmax(0, 1fr))", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Randevu Özeti", "Takvim performansını hızlıca gör.")}
            <div style={{ display: "grid", gap: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Bu ay planlanan</span><strong>{appointmentThisMonth}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Bugünkü randevu</span><strong>{todayAppointments}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Tamamlanan</span><strong>{completedAppointments}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Geciken</span><strong>{overdueAppointments}</strong></div>
            </div>
          </div>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Sözleşme Sağlığı", "Yetki sözleşmelerinin güncel dağılımı.")}
            <div style={{ display: "grid", gap: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Aktif</span><strong>{activeContracts}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Kritik</span><strong>{expiringContracts}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Komisyon potansiyeli</span><strong>{formatCurrency(contracts.reduce((sum, item) => sum + parsePrice(item.salePrice) * ((Number(item.commission || 0) || 0) / 100), 0))}</strong></div>
            </div>
          </div>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Müşteri ve Eşleşme", "CRM hareketliliğini hızlı ölç.")}
            <div style={{ display: "grid", gap: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Toplam müşteri</span><strong>{totalCustomers}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Toplam eşleşme</span><strong>{totalMatches}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Sonuçlanan eşleşme</span><strong>{completedMatches}</strong></div>
              <div style={{ display: "flex", justifyContent: "space-between" }}><span style={{ color: colors.sub }}>Portföy / müşteri oranı</span><strong>{totalCustomers ? (totalPortfolios / totalCustomers).toFixed(2) : "0.00"}</strong></div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  function renderCustomers() {
    const selectedCustomerMatches = selectedCustomer
      ? matches
          .filter((item) => item.customerId === selectedCustomer.id)
          .map((item) => ({
            ...item,
            portfolio: portfolios.find((portfolioItem) => portfolioItem.id === item.portfolioId) || null,
          }))
          .filter((item) => item.portfolio)
          .sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")))
      : [];

    const selectedCustomerAppointments = selectedCustomer
      ? appointments
          .filter((item) => item.customerId === selectedCustomer.id)
          .sort((a, b) => `${a.date || ""} ${a.time || ""}`.localeCompare(`${b.date || ""} ${b.time || ""}`))
      : [];

    const suggestedPortfolios = selectedCustomer ? getRecommendedPortfoliosForCustomer(selectedCustomer) : [];

    return (
      <div style={pageGrid()}>
        <div
          style={{
            background: colors.panel,
            border: `1px solid ${colors.border}`,
            borderRadius: 24,
            padding: 18,
            boxShadow: colors.shadow,
          }}
        >
          {sectionTitle(
            editingCustomerId ? "Müşteri Düzenle" : "Yeni Müşteri",
            "Müşteri iletişim, bütçe ve tercih bilgilerini tek yerden yönet."
          )}

          <div style={grid()}>
            {renderTextInput({
              label: "Ad Soyad",
              value: customerForm.fullName,
              onChange: (v) => setCustomerForm({ ...customerForm, fullName: v }),
            })}
            {renderTextInput({
              label: "Telefon",
              value: customerForm.phone,
              onChange: (v) => setCustomerForm({ ...customerForm, phone: v }),
            })}
            {renderTextInput({
              label: "E-posta",
              value: customerForm.email,
              onChange: (v) => setCustomerForm({ ...customerForm, email: v }),
              type: "email",
            })}
            {renderSelect({
              label: "Müşteri Tipi",
              value: customerForm.customerType,
              onChange: (v) => setCustomerForm({ ...customerForm, customerType: v }),
              options: ["alıcı", "satıcı", "kiracı", "kiraya veren"],
            })}
            {renderTextInput({
              label: "İlgilendiği Bölge",
              value: customerForm.interestedLocation,
              onChange: (v) => setCustomerForm({ ...customerForm, interestedLocation: v }),
            })}
            {renderTextInput({
              label: "İlgilendiği Mülk Tipi",
              value: customerForm.interestedPropertyType,
              onChange: (v) => setCustomerForm({ ...customerForm, interestedPropertyType: v }),
              placeholder: "Daire, villa, dükkan...",
            })}
            {renderTextInput({
              label: "Min. Bütçe",
              value: customerForm.budgetMin,
              onChange: (v) => setCustomerForm({ ...customerForm, budgetMin: v }),
            })}
            {renderTextInput({
              label: "Max. Bütçe",
              value: customerForm.budgetMax,
              onChange: (v) => setCustomerForm({ ...customerForm, budgetMax: v }),
            })}
          </div>

          <div style={{ marginTop: 12 }}>
            {renderTextarea({
              label: "Not",
              value: customerForm.note,
              onChange: (v) => setCustomerForm({ ...customerForm, note: v }),
            })}
          </div>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
            <button onClick={saveCustomer} style={actionButton(true)}>
              {editingCustomerId ? "Müşteriyi Güncelle" : "Müşteri Kaydet"}
            </button>
            <button
              onClick={() => {
                setCustomerForm(emptyCustomer());
                setEditingCustomerId(null);
              }}
              style={actionButton(false)}
            >
              Temizle
            </button>
          </div>
        </div>

        <div style={{ display: "grid", gap: 16 }}>
          {selectedCustomer ? (
            <div
              style={{
                background: colors.panel,
                border: `1px solid ${colors.border}`,
                borderRadius: 24,
                padding: 18,
                boxShadow: colors.shadow,
              }}
            >
              {sectionTitle("Seçili Müşteri Analizi", "Otomatik portföy önerilerini ve ilişki geçmişini tek bakışta gör.")}
              <div style={{ display: "grid", gap: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "flex-start" }}>
                  <div>
                    <div style={{ fontWeight: 900, fontSize: 20 }}>{selectedCustomer.fullName}</div>
                    <div style={{ color: colors.sub, marginTop: 6, lineHeight: 1.7 }}>
                      {selectedCustomer.phone || "Telefon yok"}
                      {selectedCustomer.email ? ` • ${selectedCustomer.email}` : ""}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <span style={statusBadge(selectedCustomer.customerType)}>{selectedCustomer.customerType}</span>
                    {selectedCustomer.interestedPropertyType ? <span style={statusBadge("önerildi")}>{selectedCustomer.interestedPropertyType}</span> : null}
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4, minmax(0, 1fr))", gap: 10 }}>
                  <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Bölge</div><div style={{ fontWeight: 800, marginTop: 5 }}>{selectedCustomer.interestedLocation || "-"}</div></div>
                  <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Bütçe</div><div style={{ fontWeight: 800, marginTop: 5 }}>{formatCurrency(selectedCustomer.budgetMin)} - {formatCurrency(selectedCustomer.budgetMax || selectedCustomer.budgetMin)}</div></div>
                  <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Mevcut Eşleşme</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getCustomerMatchCount(selectedCustomer.id)}</div></div>
                  <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Randevu</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getCustomerAppointmentCount(selectedCustomer.id)}</div></div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 16 }}>
                  <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 20, padding: 14 }}>
                    <div style={{ fontWeight: 900, marginBottom: 10 }}>Önerilen Portföyler</div>
                    {suggestedPortfolios.length ? (
                      <div style={{ display: "grid", gap: 10 }}>
                        {suggestedPortfolios.map((item) => {
                          const scoreAppearance = getMatchScoreAppearance(item._matchScore || 0);
                          return (
                            <div key={item.id} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12, display: "grid", gap: 8 }}>
                              <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                                <strong>{item.title}</strong>
                                <span style={{ padding: "7px 11px", borderRadius: 999, fontSize: 12, fontWeight: 900, background: scoreAppearance.background, color: scoreAppearance.color, border: `1px solid ${scoreAppearance.borderColor}` }}>{item._matchScore}/100 • {scoreAppearance.label}</span>
                              </div>
                              <div style={{ color: colors.sub, fontSize: 13 }}>
                                {item.city || "-"} / {item.district || "-"} • {item.propertyType || "Tip yok"} • {formatCurrency(item.price)}
                              </div>
                              {renderReasonChips(item._matchReasons || [], "positive")}
                              {renderReasonChips(item._matchMismatchReasons || [], "negative")}
                              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                                <button onClick={() => openPortfolioDetail(item)} style={actionButton(false, { padding: "9px 12px", fontSize: 13 })}>Portföy Aç</button>
                                <button onClick={() => saveSuggestedMatch(selectedCustomer.id, item.id, item._matchScore, item._matchReasons || [], item._matchMismatchReasons || [])} style={actionButton(false, { padding: "9px 12px", fontSize: 13 })}>Hızlı Eşleştir</button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : <div style={{ color: colors.sub }}>Bu müşteri için öne çıkan portföy önerisi bulunamadı.</div>}
                  </div>

                  <div style={{ display: "grid", gap: 16 }}>
                    <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 20, padding: 14 }}>
                      <div style={{ fontWeight: 900, marginBottom: 10 }}>Mevcut Eşleşmeler</div>
                      {selectedCustomerMatches.length ? (
                        <div style={{ display: "grid", gap: 10 }}>
                          {selectedCustomerMatches.slice(0, 4).map((item) => (
                            <div key={item.id} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                              <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                                <strong>{item.portfolio?.title || "-"}</strong>
                                <span style={statusBadge(item.status)}>{item.status}</span>
                              </div>
                              <div style={{ color: colors.sub, fontSize: 13, marginTop: 6 }}>{item.portfolio?.city || "-"} / {item.portfolio?.district || "-"} • {formatCurrency(item.portfolio?.price)}</div>
                            </div>
                          ))}
                        </div>
                      ) : <div style={{ color: colors.sub }}>Bu müşteri için henüz kayıtlı eşleşme yok.</div>}
                    </div>

                    <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 20, padding: 14 }}>
                      <div style={{ fontWeight: 900, marginBottom: 10 }}>Yaklaşan Randevular</div>
                      {selectedCustomerAppointments.length ? (
                        <div style={{ display: "grid", gap: 10 }}>
                          {selectedCustomerAppointments.slice(0, 4).map((item) => (
                            <div key={item.id} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                              <strong>{item.title}</strong>
                              <div style={{ color: colors.sub, fontSize: 13, marginTop: 6 }}>{formatLongDateTime(item.date, item.time)} • {getPortfolioTitleById(item.portfolioId)}</div>
                            </div>
                          ))}
                        </div>
                      ) : <div style={{ color: colors.sub }}>Bu müşteri için planlanmış randevu görünmüyor.</div>}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ) : null}

          <div
            style={{
              background: colors.panel,
              border: `1px solid ${colors.border}`,
              borderRadius: 24,
              padding: 18,
              boxShadow: colors.shadow,
            }}
          >
            {sectionTitle("Müşteri Listesi", "Arama, filtre ve akıllı önerilerle müşterilerini hızlıca yönet.")}
            <div style={grid()}>
              {renderTextInput({
                label: "Ara",
                value: customerSearch,
                onChange: setCustomerSearch,
                placeholder: "Ad, telefon, mail...",
              })}
              {renderSelect({
                label: "Tip Filtre",
                value: customerTypeFilter,
                onChange: setCustomerTypeFilter,
                options: ["tümü", "alıcı", "satıcı", "kiracı", "kiraya veren"],
              })}
            </div>

            <div style={{ display: "grid", gap: 12, marginTop: 16 }}>
              {filteredCustomers.length === 0 ? (
                renderEmpty("Müşteri bulunamadı", "Henüz bu kriterlere uygun müşteri kaydın yok.")
              ) : (
                filteredCustomers.map((item) => {
                  const recommendedCount = getRecommendedPortfoliosForCustomer(item).length;
                  return (
                    <div
                      key={item.id}
                      style={{
                        background: selectedCustomer?.id === item.id ? colors.primarySoft : colors.panel2,
                        border: `1px solid ${selectedCustomer?.id === item.id ? colors.primary : colors.border}`,
                        borderRadius: 18,
                        padding: 16,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          gap: 10,
                          flexWrap: "wrap",
                          alignItems: "flex-start",
                        }}
                      >
                        <div>
                          <div style={{ fontWeight: 900, fontSize: 19 }}>{item.fullName}</div>
                          <div style={{ color: colors.sub, marginTop: 6, fontSize: 14 }}>
                            {item.phone} {item.email ? `• ${item.email}` : ""}
                          </div>
                        </div>

                        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                          <span style={statusBadge(item.customerType)}>{item.customerType}</span>
                          {recommendedCount ? <span style={statusBadge("önerildi")}>{recommendedCount} öneri</span> : null}
                        </div>
                      </div>

                      <div
                        style={{
                          marginTop: 14,
                          display: "grid",
                          gridTemplateColumns: isMobile ? "1fr" : "repeat(3, minmax(0,1fr))",
                          gap: 10,
                        }}
                      >
                        <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                          <div style={{ fontSize: 12, color: colors.sub }}>Bölge</div>
                          <div style={{ fontWeight: 800, marginTop: 5 }}>{item.interestedLocation || "-"}</div>
                        </div>
                        <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                          <div style={{ fontSize: 12, color: colors.sub }}>Mülk Tipi</div>
                          <div style={{ fontWeight: 800, marginTop: 5 }}>{item.interestedPropertyType || "Serbest"}</div>
                        </div>
                        <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                          <div style={{ fontSize: 12, color: colors.sub }}>Bütçe</div>
                          <div style={{ fontWeight: 800, marginTop: 5 }}>{formatCurrency(item.budgetMin)} - {formatCurrency(item.budgetMax || item.budgetMin)}</div>
                        </div>
                      </div>

                      {item.note ? (
                        <div style={{ marginTop: 12, lineHeight: 1.6, color: colors.sub }}>{item.note}</div>
                      ) : null}

                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
                        <button onClick={() => startEditCustomer(item)} style={actionButton(false)}>
                          Düzenle
                        </button>
                        <button onClick={() => openCustomerDetail(item)} style={actionButton(false)}>
                          Detay
                        </button>
                        <button
                          onClick={() =>
                            shareOnWhatsApp(
                              item.phone,
                              `Merhaba ${item.fullName}, sizin için uygun portföyleri paylaşabilirim.`
                            )
                          }
                          style={actionButton(false, {
                            background: "#16a34a",
                            color: "#ffffff",
                            border: "none",
                          })}
                        >
                          WhatsApp
                        </button>
                        <button
                          onClick={() => removeItem("customers", item.id, setCustomers)}
                          style={actionButton(false, {
                            background: colors.dangerSoft,
                            color: colors.danger,
                            border: `1px solid ${dark ? "rgba(248,113,113,0.16)" : "#fecaca"}`,
                          })}
                        >
                          Sil
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
      </div>
    );
  }

  function renderPortfolios() {

    const selectedPortfolioMatches = selectedPortfolio
      ? matches
          .filter((item) => item.portfolioId === selectedPortfolio.id)
          .map((item) => ({
            ...item,
            customer: customers.find((customerItem) => customerItem.id === item.customerId) || null,
          }))
          .filter((item) => item.customer)
          .sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")))
      : [];

    const selectedPortfolioAppointments = selectedPortfolio
      ? appointments
          .filter((item) => item.portfolioId === selectedPortfolio.id)
          .sort((a, b) => `${a.date || ""} ${a.time || ""}`.localeCompare(`${b.date || ""} ${b.time || ""}`))
      : [];

    const suggestedCustomers = selectedPortfolio ? getRecommendedCustomersForPortfolio(selectedPortfolio) : [];

    const statCard = (label, value, helper) => (
      <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 14 }}>
        <div style={{ fontSize: 12, color: colors.sub, textTransform: "uppercase", letterSpacing: "0.08em", fontWeight: 900 }}>{label}</div>
        <div style={{ fontSize: 24, fontWeight: 900, marginTop: 6 }}>{value}</div>
        <div style={{ color: colors.sub, fontSize: 13, marginTop: 6, lineHeight: 1.6 }}>{helper}</div>
      </div>
    );

    return (
      <div style={pageGrid(440)}>
        <div
          style={{
            background: colors.panel,
            border: `1px solid ${colors.border}`,
            borderRadius: 24,
            padding: 18,
            boxShadow: colors.shadow,
          }}
        >
          {sectionTitle(editingPortfolioId ? "Portföy Düzenle" : "Yeni Portföy", "Portföyün tüm ticari ve görsel bilgilerini tek yerden yönet.")}
          <div style={grid()}>
            {renderTextInput({ label: "Portföy No", value: portfolioForm.portfolioNo, onChange: (v) => setPortfolioForm({ ...portfolioForm, portfolioNo: v }), placeholder: "Boş bırakılırsa otomatik oluşur" })}
            {renderTextInput({ label: "Başlık", value: portfolioForm.title, onChange: (v) => setPortfolioForm({ ...portfolioForm, title: v }) })}
            {renderSelect({ label: "Portföy Tipi", value: portfolioForm.portfolioType, onChange: (v) => setPortfolioForm({ ...portfolioForm, portfolioType: v }), options: ["satılık", "kiralık"] })}
            {renderTextInput({ label: "Mülk Tipi", value: portfolioForm.propertyType, onChange: (v) => setPortfolioForm({ ...portfolioForm, propertyType: v }), placeholder: "Daire, villa, arsa..." })}
            {renderTextInput({ label: "Fiyat", value: portfolioForm.price, onChange: (v) => setPortfolioForm({ ...portfolioForm, price: v }) })}
            {renderTextInput({ label: "Oda Sayısı", value: portfolioForm.roomCount, onChange: (v) => setPortfolioForm({ ...portfolioForm, roomCount: v }) })}
            {renderTextInput({ label: "Brüt Alan", value: portfolioForm.grossArea, onChange: (v) => setPortfolioForm({ ...portfolioForm, grossArea: v }) })}
            {renderTextInput({ label: "Şehir", value: portfolioForm.city, onChange: (v) => setPortfolioForm({ ...portfolioForm, city: v }) })}
            {renderTextInput({ label: "İlçe", value: portfolioForm.district, onChange: (v) => setPortfolioForm({ ...portfolioForm, district: v }) })}
            {renderTextInput({ label: "Adres", value: portfolioForm.address, onChange: (v) => setPortfolioForm({ ...portfolioForm, address: v }) })}
            {renderTextInput({ label: "Mülk Sahibi", value: portfolioForm.ownerName, onChange: (v) => setPortfolioForm({ ...portfolioForm, ownerName: v }) })}
            {renderTextInput({ label: "Sahip Telefon", value: portfolioForm.ownerPhone, onChange: (v) => setPortfolioForm({ ...portfolioForm, ownerPhone: v }) })}
            {renderTextInput({ label: "Görsel URL", value: portfolioForm.imageUrl, onChange: (v) => setPortfolioForm({ ...portfolioForm, imageUrl: v }) })}
            {renderSelect({ label: "Durum", value: portfolioForm.status, onChange: (v) => setPortfolioForm({ ...portfolioForm, status: v }), options: ["aktif", "satıldı", "kiralandı", "kapandı"] })}
          </div>
          <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Açıklama", value: portfolioForm.description, onChange: (v) => setPortfolioForm({ ...portfolioForm, description: v }) })}</div>
          <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Not", value: portfolioForm.note, onChange: (v) => setPortfolioForm({ ...portfolioForm, note: v }) })}</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
            <button onClick={savePortfolio} style={actionButton(true)}>{editingPortfolioId ? "Portföyü Güncelle" : "Portföy Kaydet"}</button>
            <button onClick={() => { setPortfolioForm(emptyPortfolio()); setEditingPortfolioId(null); }} style={actionButton(false)}>Temizle</button>
          </div>
        </div>

        <div style={{ display: "grid", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Portföy Detayı", "Seçili portföyün müşteri ilgisini, randevularını ve önerilerini tek panelde gör.")}
            {!selectedPortfolio ? (
              renderEmpty("Portföy seçilmedi", "Listeden bir portföy seçerek detay ekranını aç.")
            ) : (
              <div style={{ display: "grid", gap: 16 }}>
                {selectedPortfolio.imageUrl ? (
                  <img src={selectedPortfolio.imageUrl} alt={selectedPortfolio.title} style={{ width: "100%", maxHeight: 280, objectFit: "cover", borderRadius: 20, border: `1px solid ${colors.border}` }} onError={(e) => { e.currentTarget.style.display = "none"; }} />
                ) : null}
                <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "flex-start" }}>
                  <div>
                    <div style={{ fontWeight: 900, fontSize: 24 }}>{selectedPortfolio.title}</div>
                    <div style={{ color: colors.sub, marginTop: 6 }}>{selectedPortfolio.portfolioNo || "-"} • {selectedPortfolio.city || "-"} / {selectedPortfolio.district || "-"}</div>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <span style={statusBadge(selectedPortfolio.portfolioType)}>{selectedPortfolio.portfolioType}</span>
                    <span style={statusBadge(selectedPortfolio.status)}>{selectedPortfolio.status}</span>
                  </div>
                </div>
                <div style={{ fontSize: 28, fontWeight: 900 }}>{formatCurrency(selectedPortfolio.price)}</div>
                <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4, minmax(0, 1fr))", gap: 10 }}>
                  {statCard("Eşleşme", getPortfolioMatchCount(selectedPortfolio.id), "Bu portföyle kurulan toplam müşteri ilişkisi")}
                  {statCard("Randevu", getPortfolioAppointmentCount(selectedPortfolio.id), "Portföye bağlı planlanan görüşmeler")}
                  {statCard("Aktif Gün", getPortfolioDaysActive(selectedPortfolio), "İlanda kaldığı gün sayısı")}
                  {statCard("Uygun Müşteri", suggestedCustomers.length, "Bölge ve bütçeye göre öne çıkan müşteriler")}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 10 }}>
                  <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 14 }}>
                    <div style={{ color: colors.sub, fontSize: 12, fontWeight: 900, textTransform: "uppercase" }}>Sahip Bilgisi</div>
                    <div style={{ fontWeight: 900, marginTop: 8 }}>{selectedPortfolio.ownerName || "-"}</div>
                    <div style={{ color: colors.sub, marginTop: 6 }}>{selectedPortfolio.ownerPhone || "Telefon yok"}</div>
                  </div>
                  <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 14 }}>
                    <div style={{ color: colors.sub, fontSize: 12, fontWeight: 900, textTransform: "uppercase" }}>Temel Özellikler</div>
                    <div style={{ marginTop: 8, display: "grid", gap: 6, color: colors.muted }}>
                      <div>Mülk tipi: <strong style={{ color: colors.text }}>{selectedPortfolio.propertyType || "-"}</strong></div>
                      <div>Oda sayısı: <strong style={{ color: colors.text }}>{selectedPortfolio.roomCount || "-"}</strong></div>
                      <div>Brüt alan: <strong style={{ color: colors.text }}>{selectedPortfolio.grossArea || "-"}</strong></div>
                    </div>
                  </div>
                </div>
                {selectedPortfolio.address ? <div style={{ lineHeight: 1.7 }}>{selectedPortfolio.address}</div> : null}
                {selectedPortfolio.description ? <div style={{ lineHeight: 1.7, color: colors.muted }}>{selectedPortfolio.description}</div> : null}
                {selectedPortfolio.note ? <div style={{ color: colors.sub }}>Not: {selectedPortfolio.note}</div> : null}
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                  <button onClick={() => startEditPortfolio(selectedPortfolio)} style={actionButton(false)}>Düzenle</button>
                  <button onClick={() => shareOnWhatsApp(selectedPortfolio.ownerPhone, `Merhaba ${selectedPortfolio.ownerName || ""}, ${selectedPortfolio.title} portföyünüz hakkında bilgi vermek istiyorum.`)} style={actionButton(false, { background: "#16a34a", color: "#ffffff", border: "none" })}>WhatsApp</button>
                  <button onClick={() => { setMatchForm({ customerId: "", portfolioId: selectedPortfolio.id, status: "önerildi", note: `${selectedPortfolio.title} için yeni eşleşme` }); setActiveTab("matches"); window.scrollTo({ top: 0, behavior: "smooth" }); }} style={actionButton(false)}>Eşleştirme Aç</button>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 16 }}>
                  <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 20, padding: 14 }}>
                    <div style={{ fontWeight: 900, marginBottom: 10 }}>İlgilenen Müşteriler</div>
                    {selectedPortfolioMatches.length ? (
                      <div style={{ display: "grid", gap: 10 }}>
                        {selectedPortfolioMatches.slice(0, 4).map((item) => (
                          <div key={item.id} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                              <strong>{item.customer?.fullName || "-"}</strong>
                              <span style={statusBadge(item.status)}>{item.status}</span>
                            </div>
                            <div style={{ color: colors.sub, fontSize: 13, marginTop: 6 }}>{item.customer?.phone || "Telefon yok"}</div>
                          </div>
                        ))}
                      </div>
                    ) : <div style={{ color: colors.sub }}>Bu portföy için henüz eşleşme yok.</div>}
                  </div>
                  <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 20, padding: 14 }}>
                    <div style={{ fontWeight: 900, marginBottom: 10 }}>Yaklaşan Randevular</div>
                    {selectedPortfolioAppointments.length ? (
                      <div style={{ display: "grid", gap: 10 }}>
                        {selectedPortfolioAppointments.slice(0, 4).map((item) => (
                          <div key={item.id} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                            <strong>{item.title}</strong>
                            <div style={{ color: colors.sub, fontSize: 13, marginTop: 6 }}>{formatLongDateTime(item.date, item.time)} • {getCustomerNameById(item.customerId)}</div>
                          </div>
                        ))}
                      </div>
                    ) : <div style={{ color: colors.sub }}>Bu portföy için planlanmış randevu görünmüyor.</div>}
                  </div>
                </div>
                <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 20, padding: 14 }}>
                  <div style={{ fontWeight: 900, marginBottom: 10 }}>Önerilen Müşteriler</div>
                  {suggestedCustomers.length ? (
                    <div style={{ display: "grid", gap: 10 }}>
                      {suggestedCustomers.map((item) => {
                        const scoreAppearance = getMatchScoreAppearance(item._matchScore || 0);
                        return (
                          <div key={item.id} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12, display: "grid", gap: 10 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                              <div>
                                <strong>{item.fullName}</strong>
                                <div style={{ color: colors.sub, fontSize: 13, marginTop: 4 }}>{item.interestedLocation || "Bölge yok"} • {item.interestedPropertyType || "Tip serbest"} • {formatCurrency(item.budgetMax || item.budgetMin)}</div>
                              </div>
                              <span style={{ padding: "7px 11px", borderRadius: 999, fontSize: 12, fontWeight: 900, background: scoreAppearance.background, color: scoreAppearance.color, border: `1px solid ${scoreAppearance.borderColor}` }}>{item._matchScore}/100 • {scoreAppearance.label}</span>
                            </div>
                            {renderReasonChips(item._matchReasons || [], "positive")}
                            {renderReasonChips(item._matchMismatchReasons || [], "negative")}
                            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                              <button onClick={() => openCustomerDetail(item)} style={actionButton(false, { padding: "9px 12px", fontSize: 13 })}>Müşteri Aç</button>
                              <button onClick={() => { setMatchForm({ customerId: item.id, portfolioId: selectedPortfolio.id, status: "önerildi", note: `${selectedPortfolio.title} için önerildi • ${item._matchScore || 0}/100${item._matchReasons?.length ? ` • ${item._matchReasons.join(" • ")}` : ""}${item._matchMismatchReasons?.length ? ` • Dikkat: ${item._matchMismatchReasons.join(" • ")}` : ""}` }); setActiveTab("matches"); window.scrollTo({ top: 0, behavior: "smooth" }); }} style={actionButton(false, { padding: "9px 12px", fontSize: 13 })}>Eşleştir</button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : <div style={{ color: colors.sub }}>Bu portföy için öne çıkan müşteri önerisi bulunamadı.</div>}
                </div>
              </div>
            )}
          </div>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Portföy Listesi", "Portföylerini görsel, ilgi ve hareket özetleriyle yönet.")}
            <div style={grid()}>
              {renderTextInput({ label: "Ara", value: portfolioSearch, onChange: setPortfolioSearch, placeholder: "Başlık, ilçe, sahip, fiyat..." })}
              {renderSelect({ label: "Tip Filtre", value: portfolioTypeFilter, onChange: setPortfolioTypeFilter, options: ["tümü", "satılık", "kiralık"] })}
            </div>
            <div style={{ display: "grid", gap: 12, marginTop: 16 }}>
              {filteredPortfolios.length === 0 ? renderEmpty("Portföy bulunamadı", "Henüz bu kriterlere uygun portföy kaydın yok.") : filteredPortfolios.map((item) => (
                <div key={item.id} style={{ background: selectedPortfolio?.id === item.id ? colors.primarySoft : colors.panel2, border: `1px solid ${selectedPortfolio?.id === item.id ? colors.primary : colors.border}`, borderRadius: 18, padding: 16, display: "grid", gap: 12 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontWeight: 900, fontSize: 19 }}>{item.title}</div>
                      <div style={{ color: colors.sub, marginTop: 6, fontSize: 14 }}>{item.portfolioNo || "-"} • {item.city || "-"} / {item.district || "-"}</div>
                    </div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      <span style={statusBadge(item.portfolioType)}>{item.portfolioType}</span>
                      <span style={statusBadge(item.status)}>{item.status}</span>
                    </div>
                  </div>
                  <div style={{ fontWeight: 900, fontSize: 22 }}>{formatCurrency(item.price)}</div>
                  <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4, minmax(0, 1fr))", gap: 10 }}>
                    <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Mülk tipi</div><div style={{ fontWeight: 800, marginTop: 5 }}>{item.propertyType || "-"}</div></div>
                    <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Eşleşme</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getPortfolioMatchCount(item.id)}</div></div>
                    <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Randevu</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getPortfolioAppointmentCount(item.id)}</div></div>
                    <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Aktif gün</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getPortfolioDaysActive(item)}</div></div>
                  </div>
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <button onClick={() => startEditPortfolio(item)} style={actionButton(false)}>Düzenle</button>
                    <button onClick={() => openPortfolioDetail(item)} style={actionButton(false)}>Detay</button>
                    <button onClick={() => shareOnWhatsApp(item.ownerPhone, `Merhaba ${item.ownerName || ""}, ${item.title} portföyünüz hakkında bilgi vermek istiyorum.`)} style={actionButton(false, { background: "#16a34a", color: "#ffffff", border: "none" })}>WhatsApp</button>
                    <button onClick={() => removeItem("portfolios", item.id, setPortfolios)} style={actionButton(false, { background: colors.dangerSoft, color: colors.danger, border: `1px solid ${dark ? "rgba(248,113,113,0.16)" : "#fecaca"}` })}>Sil</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }


  function renderMatches() {
    const selectedCustomerForMatch = customers.find((item) => item.id === matchForm.customerId) || null;
    const selectedPortfolioForMatch = portfolios.find((item) => item.id === matchForm.portfolioId) || null;

    const suggestionRows = selectedCustomerForMatch && !selectedPortfolioForMatch
      ? getRecommendedPortfoliosForCustomer(selectedCustomerForMatch).map((portfolio) => ({
          customer: selectedCustomerForMatch,
          portfolio,
          score: portfolio._matchScore || 0,
          reasons: portfolio._matchReasons || [],
          mismatchReasons: portfolio._matchMismatchReasons || [],
        }))
      : selectedPortfolioForMatch && !selectedCustomerForMatch
      ? getRecommendedCustomersForPortfolio(selectedPortfolioForMatch).map((customer) => ({
          customer,
          portfolio: selectedPortfolioForMatch,
          score: customer._matchScore || 0,
          reasons: customer._matchReasons || [],
          mismatchReasons: customer._matchMismatchReasons || [],
        }))
      : getAutoMatchSuggestions(8);

    return (
      <div style={pageGrid()}>
        <div style={{ display: "grid", gap: 16 }}>
          <div
            style={{
              background: colors.panel,
              border: `1px solid ${colors.border}`,
              borderRadius: 24,
              padding: 18,
              boxShadow: colors.shadow,
            }}
          >
            {sectionTitle("Yeni Eşleştirme", "Müşteri ve uygun portföyleri birbirine bağla.")}

            <div style={grid()}>
              {renderSelect({
                label: "Müşteri",
                value: matchForm.customerId,
                onChange: (v) => setMatchForm({ ...matchForm, customerId: v }),
                options: customerOptions,
              })}
              {renderSelect({
                label: "Portföy",
                value: matchForm.portfolioId,
                onChange: (v) => setMatchForm({ ...matchForm, portfolioId: v }),
                options: portfolioOptions,
              })}
              {renderSelect({
                label: "Durum",
                value: matchForm.status,
                onChange: (v) => setMatchForm({ ...matchForm, status: v }),
                options: ["önerildi", "ilgileniyor", "gösterildi", "pazarlık", "tamamlandı"],
              })}
            </div>

            <div
              style={{
                marginTop: 12,
                background: colors.panel2,
                border: `1px solid ${colors.border}`,
                borderRadius: 18,
                padding: 14,
                color: colors.sub,
                lineHeight: 1.7,
                fontSize: 14,
              }}
            >
              <strong style={{ color: colors.text }}>Seçili müşteri:</strong>{" "}
              {matchForm.customerId ? getCustomerNameById(matchForm.customerId) : "-"}
              <br />
              <strong style={{ color: colors.text }}>Seçili portföy:</strong>{" "}
              {matchForm.portfolioId ? getPortfolioTitleById(matchForm.portfolioId) : "-"}
            </div>

            <div style={{ marginTop: 12 }}>
              {renderTextarea({
                label: "Not",
                value: matchForm.note,
                onChange: (v) => setMatchForm({ ...matchForm, note: v }),
              })}
            </div>

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
              <button onClick={saveMatch} style={actionButton(true)}>
                Eşleştirme Kaydet
              </button>
              <button onClick={() => setMatchForm(emptyMatch())} style={actionButton(false)}>
                Temizle
              </button>
            </div>
          </div>

          <div
            style={{
              background: colors.panel,
              border: `1px solid ${colors.border}`,
              borderRadius: 24,
              padding: 18,
              boxShadow: colors.shadow,
            }}
          >
            {sectionTitle("Akıllı Eşleştirme Önerileri", "Bütçe, bölge ve mülk tipi uyumuna göre otomatik öneriler.")}
            <div style={{ display: "grid", gap: 12 }}>
              {suggestionRows.length === 0 ? (
                renderEmpty("Öneri bulunamadı", "Yeni müşteri veya portföy ekleyince sistem burada öneriler gösterecek.")
              ) : (
                suggestionRows.map((item, index) => {
                  const scoreAppearance = getMatchScoreAppearance(item.score || 0);
                  return (
                    <div key={`${item.customer?.id || "c"}-${item.portfolio?.id || "p"}-${index}`} style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 16, display: "grid", gap: 10 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                        <div style={{ fontWeight: 900, fontSize: 18 }}>{item.customer?.fullName || "-"} ↔ {item.portfolio?.title || "-"}</div>
                        <span style={{ padding: "7px 11px", borderRadius: 999, fontSize: 12, fontWeight: 900, background: scoreAppearance.background, color: scoreAppearance.color, border: `1px solid ${scoreAppearance.borderColor}` }}>{item.score}/100 • {scoreAppearance.label}</span>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 10 }}>
                        <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                          <div style={{ fontSize: 12, color: colors.sub }}>Müşteri</div>
                          <div style={{ fontWeight: 800, marginTop: 5 }}>{item.customer?.customerType || "-"} • {item.customer?.interestedLocation || "Bölge yok"}</div>
                          <div style={{ color: colors.sub, fontSize: 13, marginTop: 5 }}>{item.customer?.interestedPropertyType || "Tip serbest"} • {formatCurrency(item.customer?.budgetMax || item.customer?.budgetMin)}</div>
                        </div>
                        <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}>
                          <div style={{ fontSize: 12, color: colors.sub }}>Portföy</div>
                          <div style={{ fontWeight: 800, marginTop: 5 }}>{item.portfolio?.portfolioType || "-"} • {item.portfolio?.city || "-"} / {item.portfolio?.district || "-"}</div>
                          <div style={{ color: colors.sub, fontSize: 13, marginTop: 5 }}>{item.portfolio?.propertyType || "Tip yok"} • {formatCurrency(item.portfolio?.price)}</div>
                        </div>
                      </div>
                      {renderReasonChips(item.reasons || [], "positive")}
                      {renderReasonChips(item.mismatchReasons || [], "negative")}
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                        <button onClick={() => setMatchForm({ customerId: item.customer?.id || "", portfolioId: item.portfolio?.id || "", status: "önerildi", note: `Akıllı öneri • ${item.score}/100${item.reasons?.length ? ` • ${item.reasons.join(" • ")}` : ""}${item.mismatchReasons?.length ? ` • Dikkat: ${item.mismatchReasons.join(" • ")}` : ""}` })} style={actionButton(false)}>Forma Doldur</button>
                        <button onClick={() => saveSuggestedMatch(item.customer?.id, item.portfolio?.id, item.score, item.reasons || [], item.mismatchReasons || [])} style={actionButton(false)}>Hızlı Eşleştir</button>
                        <button onClick={() => openCustomerDetail(item.customer)} style={actionButton(false)}>Müşteri Aç</button>
                        <button onClick={() => openPortfolioDetail(item.portfolio)} style={actionButton(false)}>Portföy Aç</button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>

        <div
          style={{
            background: colors.panel,
            border: `1px solid ${colors.border}`,
            borderRadius: 24,
            padding: 18,
            boxShadow: colors.shadow,
          }}
        >
          {sectionTitle("Eşleşme Listesi", "Portföy-müşteri ilgisini buradan takip et.")}

          <div style={{ display: "grid", gap: 12 }}>
            {matches.length === 0 ? (
              renderEmpty("Eşleşme bulunamadı", "İlk müşteri-portföy eşleşmeni oluştur.")
            ) : (
              matches.map((item) => (
                <div
                  key={item.id}
                  style={{
                    background: colors.panel2,
                    border: `1px solid ${colors.border}`,
                    borderRadius: 18,
                    padding: 16,
                  }}
                >
                  <div style={{ fontWeight: 900, fontSize: 18 }}>
                    {getCustomerNameById(item.customerId)} ↔ {getPortfolioTitleById(item.portfolioId)}
                  </div>

                  <div style={{ marginTop: 10 }}>
                    <span style={statusBadge(item.status)}>{item.status}</span>
                  </div>

                  {item.note ? <div style={{ marginTop: 12, lineHeight: 1.6 }}>{item.note}</div> : null}

                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
                    {["önerildi", "ilgileniyor", "gösterildi", "pazarlık", "tamamlandı"].map((st) => (
                      <button
                        key={st}
                        onClick={() => updateMatchStatus(item.id, st)}
                        style={actionButton(false)}
                      >
                        {st}
                      </button>
                    ))}

                    <button
                      onClick={() => removeItem("matches", item.id, setMatches)}
                      style={actionButton(false, {
                        background: colors.dangerSoft,
                        color: colors.danger,
                        border: `1px solid ${dark ? "rgba(248,113,113,0.16)" : "#fecaca"}`,
                      })}
                    >
                      Sil
                    </button>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    );
  }

  function renderAppointments() {

    const weekDates = getCurrentWeekDates();
    const weekLabel = `${weekDates[0].toLocaleDateString("tr-TR", { day: "2-digit", month: "short" })} - ${weekDates[6].toLocaleDateString("tr-TR", { day: "2-digit", month: "short" })}`;
    const weeklyColumns = weekDates.map((date) => {
      const key = date.toISOString().slice(0, 10);
      return {
        key,
        label: date.toLocaleDateString("tr-TR", { weekday: "short" }),
        day: date.toLocaleDateString("tr-TR", { day: "2-digit" }),
        items: appointmentRows.filter((item) => item.date === key),
      };
    });
    const overdueList = filteredAppointments.filter((item) => getAppointmentVisualStatus(item) === "gecikti");

    return (
      <div style={pageGrid(430)}>
        <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
          {sectionTitle(editingAppointmentId ? "Randevu Düzenle" : "Yeni Randevu", "Profesyonel takvim görünümüyle müşteri görüşmelerini planla.")}
          <div style={grid()}>
            {renderTextInput({ label: "Başlık", value: appointmentForm.title, onChange: (v) => setAppointmentForm({ ...appointmentForm, title: v }) })}
            {renderTextInput({ label: "Tarih", value: appointmentForm.date, onChange: (v) => setAppointmentForm({ ...appointmentForm, date: v }), type: "date" })}
            {renderTextInput({ label: "Saat", value: appointmentForm.time, onChange: (v) => setAppointmentForm({ ...appointmentForm, time: v }), type: "time" })}
            {renderTextInput({ label: "Konum", value: appointmentForm.location, onChange: (v) => setAppointmentForm({ ...appointmentForm, location: v }) })}
            {renderSelect({ label: "Müşteri", value: appointmentForm.customerId, onChange: (v) => setAppointmentForm({ ...appointmentForm, customerId: v }), options: customerOptions })}
            {renderSelect({ label: "Portföy", value: appointmentForm.portfolioId, onChange: (v) => setAppointmentForm({ ...appointmentForm, portfolioId: v }), options: portfolioOptions })}
            {renderSelect({ label: "Durum", value: appointmentForm.status, onChange: (v) => setAppointmentForm({ ...appointmentForm, status: v }), options: ["planlandı", "tamamlandı", "iptal"] })}
          </div>
          <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Not", value: appointmentForm.note, onChange: (v) => setAppointmentForm({ ...appointmentForm, note: v }) })}</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
            <button onClick={saveAppointment} style={actionButton(true)}>{editingAppointmentId ? "Randevuyu Güncelle" : "Randevu Kaydet"}</button>
            <button onClick={() => { setAppointmentForm(emptyAppointment()); setEditingAppointmentId(null); }} style={actionButton(false)}>Temizle</button>
          </div>
        </div>

        <div style={{ display: "grid", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Randevu Özeti", `Haftalık görünüm • ${weekLabel}`)}
            <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4, minmax(0, 1fr))", gap: 10, marginBottom: 14 }}>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Bugün</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{todayAppointments}</div></div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Yaklaşan</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{upcomingAppointments}</div></div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Tamamlanan</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{completedAppointments}</div></div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Geciken</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{overdueList.length}</div></div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(7, minmax(0, 1fr))", gap: 10 }}>
              {weeklyColumns.map((column) => (
                <div key={column.key} style={{ background: colors.panel2, border: `1px solid ${column.key === todayStr ? colors.primary : colors.border}`, borderRadius: 18, padding: 12, minHeight: 160 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", marginBottom: 10 }}>
                    <strong style={{ textTransform: "capitalize" }}>{column.label}</strong>
                    <span style={statusBadge(column.key === todayStr ? "bugün" : `${column.items.length} randevu`)}>{column.day}</span>
                  </div>
                  {column.items.length ? (
                    <div style={{ display: "grid", gap: 8 }}>
                      {column.items.slice(0, 3).map((item) => (
                        <button key={item.id} onClick={() => startEditAppointment(item)} style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 14, padding: 10, textAlign: "left", cursor: "pointer" }}>
                          <div style={{ fontWeight: 800, fontSize: 13 }}>{item.time || "Saat yok"}</div>
                          <div style={{ color: colors.sub, fontSize: 12, marginTop: 4 }}>{item.title}</div>
                        </button>
                      ))}
                      {column.items.length > 3 ? <div style={{ color: colors.sub, fontSize: 12 }}>+{column.items.length - 3} kayıt daha</div> : null}
                    </div>
                  ) : (
                    <div style={{ color: colors.sub, fontSize: 13 }}>Boş gün</div>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Randevu Listesi", "Yaklaşan, geciken ve tamamlanan görüşmeleri filtreleyerek yönet.")}
            <div style={grid()}>
              {renderTextInput({ label: "Ara", value: appointmentSearch, onChange: setAppointmentSearch, placeholder: "Başlık, konum, müşteri..." })}
              {renderSelect({ label: "Durum Filtre", value: appointmentStatusFilter, onChange: setAppointmentStatusFilter, options: ["tümü", "planlandı", "tamamlandı", "iptal"] })}
            </div>
            {overdueList.length ? (
              <div style={{ marginTop: 12, padding: 14, borderRadius: 18, background: colors.dangerSoft, color: colors.danger, border: `1px solid ${dark ? "rgba(248,113,113,0.18)" : "#fecaca"}` }}>
                {overdueList.length} adet gecikmiş randevu var. Öncelikli olarak müşteriye hatırlatma göndermen iyi olur.
              </div>
            ) : null}
            <div style={{ display: "grid", gap: 12, marginTop: 16 }}>
              {filteredAppointments.length === 0 ? renderEmpty("Randevu bulunamadı", "Henüz bu kriterlere uygun randevu yok.") : filteredAppointments.map((item) => (
                <div key={item.id} style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
                    <div>
                      <div style={{ fontWeight: 900, fontSize: 18 }}>{item.title}</div>
                      <div style={{ color: colors.sub, marginTop: 6, fontSize: 14 }}>{formatLongDateTime(item.date, item.time)} • {item.location || "Konum yok"}</div>
                    </div>
                    <span style={statusBadge(getAppointmentVisualStatus(item))}>{getAppointmentVisualStatus(item)}</span>
                  </div>
                  <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1fr 1fr", gap: 10 }}>
                    <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Müşteri</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getCustomerNameById(item.customerId)}</div></div>
                    <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Portföy</div><div style={{ fontWeight: 800, marginTop: 5 }}>{getPortfolioTitleById(item.portfolioId)}</div></div>
                  </div>
                  {item.note ? <div style={{ marginTop: 12, lineHeight: 1.6 }}>{item.note}</div> : null}
                  <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
                    <button onClick={() => startEditAppointment(item)} style={actionButton(false)}>Düzenle</button>
                    <button onClick={() => { const customer = customers.find((c) => c.id === item.customerId); if (customer) openCustomerDetail(customer); }} style={actionButton(false)}>Müşteri Aç</button>
                    <button onClick={() => { const portfolio = portfolios.find((p) => p.id === item.portfolioId); if (portfolio) openPortfolioDetail(portfolio); }} style={actionButton(false)}>Portföy Aç</button>
                    <button onClick={() => shareOnWhatsApp(customers.find((c) => c.id === item.customerId)?.phone, `Merhaba ${getCustomerNameById(item.customerId)}, ${formatLongDateTime(item.date, item.time)} tarihli randevumuzu hatırlatmak istedim.`)} style={actionButton(false, { background: "#16a34a", color: "#ffffff", border: "none" })}>Hatırlatma Gönder</button>
                    <button onClick={() => removeItem("appointments", item.id, setAppointments)} style={actionButton(false, { background: colors.dangerSoft, color: colors.danger, border: `1px solid ${dark ? "rgba(248,113,113,0.16)" : "#fecaca"}` })}>Sil</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }


  function renderContracts() {
    const handleExportContractsExcel = () => {
      try {
        const rows = filteredContracts.map((c) => ({
          "Sözleşme No": c.contractNo || "-",
          Şirket: c.companyName || "-",
          Emlakçı: c.agentName || "-",
          "Mülk Sahibi": c.owner || "-",
          Telefon: c.phone || "-",
          Adres: c.address || "-",
          "Mülk Tipi": c.propertyType || "-",
          Fiyat: c.salePrice || "-",
          Komisyon: c.commission || "-",
          Başlangıç: c.startDate || "-",
          Bitiş: c.endDate || "-",
          Durum: c.calculatedStatus || getContractStatus(c),
          Not: c.note || "-",
        }));
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Sozlesmeler");
        XLSX.writeFile(wb, `sozlesmeler-${new Date().toISOString().slice(0, 10)}.xlsx`);
      } catch (e) {
        alert("Excel hatası: " + e.message);
      }
    };

    const handleExportContractPdf = async (item) => {
      await exportGenericContractPdf({
        fileName: `yetki-sozlesmesi-${item.contractNo || "kayit"}.pdf`,
        title: "Emlak Yetki Sözleşmesi",
        subtitle: `${getOfficeDisplayName()} • ${formatShortDate(item.startDate)} - ${formatShortDate(item.endDate)}`,
        rows: [
          ["Sözleşme No", item.contractNo || "-"],
          ["Sözleşme Tarihi", formatShortDate(item.startDate)],
          ["Şirket", item.companyName || officeProfile.officeName || "-"],
          ["Emlakçı", item.agentName || officeProfile.agentName || "-"],
          ["Mülk Sahibi", item.owner || "-"],
          ["Telefon", item.phone || "-"],
          ["Adres", item.address || "-"],
          ["Mülk Tipi", item.propertyType || "-"],
          ["Satış Fiyatı", formatCurrency(item.salePrice)],
          ["Komisyon", `%${item.commission || "-"}`],
          ["Başlangıç Tarihi", formatShortDate(item.startDate)],
          ["Bitiş Tarihi", formatShortDate(item.endDate)],
          ["Durum", getContractStatus(item)],
        ],
        footer: item.note || "",
        signerLeftLabel: "Mülk Sahibi",
        signerLeft: item.owner || "",
        signerMiddleLabel: "Emlakçı",
        signerMiddle: item.agentName || officeProfile.agentName || "",
        signerRightLabel: "Tarih",
        signerRight: formatShortDate(item.endDate),
      });
    };

    const expiringTodayCount = contracts.filter((item) => getContractStatus(item) === "bugün bitiyor").length;
    const expiredCount = contracts.filter((item) => getContractStatus(item) === "süresi doldu").length;

    return (
      <div style={pageGrid(430)}>
        <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
          {sectionTitle(editingContractId ? "Sözleşme Düzenle" : "Yeni Sözleşme", "Yetki sözleşmelerini takip et, uyarıları kaçırma ve yenilemeleri hızlı hazırla.")}
          <div style={grid()}>
            {renderTextInput({ label: "Sözleşme No", value: contractForm.contractNo, onChange: (v) => setContractForm({ ...contractForm, contractNo: v }), placeholder: "Boş bırakılırsa otomatik oluşur" })}
            {renderTextInput({ label: "Şirket Adı", value: contractForm.companyName, onChange: (v) => setContractForm({ ...contractForm, companyName: v }) })}
            {renderTextInput({ label: "Emlakçı", value: contractForm.agentName, onChange: (v) => setContractForm({ ...contractForm, agentName: v }) })}
            {renderTextInput({ label: "Mülk Sahibi", value: contractForm.owner, onChange: (v) => setContractForm({ ...contractForm, owner: v }) })}
            {renderTextInput({ label: "Telefon", value: contractForm.phone, onChange: (v) => setContractForm({ ...contractForm, phone: v }) })}
            {renderTextInput({ label: "Mülk Tipi", value: contractForm.propertyType, onChange: (v) => setContractForm({ ...contractForm, propertyType: v }) })}
            {renderTextInput({ label: "Satış Fiyatı", value: contractForm.salePrice, onChange: (v) => setContractForm({ ...contractForm, salePrice: v }) })}
            {renderTextInput({ label: "Komisyon (%)", value: contractForm.commission, onChange: (v) => setContractForm({ ...contractForm, commission: v }) })}
            {renderTextInput({ label: "Başlangıç Tarihi", value: contractForm.startDate, onChange: (v) => setContractForm({ ...contractForm, startDate: v }), type: "date" })}
            {renderTextInput({ label: "Bitiş Tarihi", value: contractForm.endDate, onChange: (v) => setContractForm({ ...contractForm, endDate: v }), type: "date" })}
            {renderSelect({ label: "Durum", value: contractForm.status, onChange: (v) => setContractForm({ ...contractForm, status: v }), options: ["aktif", "iptal edildi"] })}
          </div>
          <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Adres", value: contractForm.address, onChange: (v) => setContractForm({ ...contractForm, address: v }) })}</div>
          <div style={{ marginTop: 12 }}>{renderTextarea({ label: "Not", value: contractForm.note, onChange: (v) => setContractForm({ ...contractForm, note: v }) })}</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
            <button onClick={saveContract} style={actionButton(true)}>{editingContractId ? "Sözleşmeyi Güncelle" : "Sözleşme Kaydet"}</button>
            <button onClick={() => { setContractForm(emptyContract()); setEditingContractId(null); }} style={actionButton(false)}>Temizle</button>
          </div>
        </div>

        <div style={{ display: "grid", gap: 16 }}>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            {sectionTitle("Sözleşme Sağlığı", "Bitiş, yenileme ve komisyon görünümünü tek ekranda takip et.")}
            <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr 1fr" : "repeat(4, minmax(0, 1fr))", gap: 10 }}>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Aktif</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{activeContracts}</div></div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Bugün Biten</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{expiringTodayCount}</div></div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Süresi Dolan</div><div style={{ fontWeight: 900, fontSize: 22, marginTop: 6 }}>{expiredCount}</div></div>
              <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Komisyon Potansiyeli</div><div style={{ fontWeight: 900, fontSize: 18, marginTop: 6 }}>{formatCurrency(contracts.reduce((sum, item) => sum + parsePrice(item.salePrice) * ((Number(item.commission || 0) || 0) / 100), 0))}</div></div>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: isMobile ? "1fr" : "1.15fr 0.85fr", gap: 12, marginTop: 14 }}>
              <div style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 14 }}>
                <div style={{ fontSize: 12, color: colors.sub }}>PDF / Ofis Bilgisi</div>
                <div style={{ fontWeight: 900, fontSize: 18, marginTop: 6 }}>{getOfficeDisplayName()}</div>
                <div style={{ color: colors.sub, marginTop: 8, lineHeight: 1.6 }}>{officeProfile.agentName || "Danışman bilgisi eklenmedi"} {officeProfile.phone ? `• ${officeProfile.phone}` : ""}</div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <button onClick={() => setActiveTab("saleContracts")} style={actionButton(false)}>Satış Paneli</button>
                <button onClick={() => setActiveTab("rentalContracts")} style={actionButton(false)}>Kira Paneli</button>
              </div>
            </div>
          </div>
          <div style={{ background: colors.panel, border: `1px solid ${colors.border}`, borderRadius: 24, padding: 18, boxShadow: colors.shadow }}>
            <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center", marginBottom: 16 }}>
              <div>
                <h2 style={{ margin: 0, fontSize: 24, fontWeight: 900 }}>Sözleşme Listesi</h2>
                <p style={{ margin: "8px 0 0", color: colors.sub }}>Dışa aktar, filtrele ve kritik tarihleri kolayca takip et.</p>
              </div>
              <button onClick={handleExportContractsExcel} style={actionButton(false)}>Excel Dışa Aktar</button>
            </div>
            <div style={grid()}>
              {renderTextInput({ label: "Ara", value: contractSearch, onChange: setContractSearch, placeholder: "No, sahip, telefon, adres..." })}
              {renderSelect({ label: "Durum Filtre", value: contractStatusFilter, onChange: setContractStatusFilter, options: ["tümü", "aktif", "yakında bitecek", "bugün bitiyor", "süresi doldu", "iptal edildi"] })}
            </div>
            <div style={{ display: "grid", gap: 12, marginTop: 16 }}>
              {filteredContracts.length === 0 ? renderEmpty("Sözleşme bulunamadı", "Henüz bu kriterlere uygun sözleşme yok.") : filteredContracts.map((item) => {
                const daysRemaining = getDaysRemaining(item.endDate);
                const totalDuration = item.startDate && item.endDate ? Math.max(1, Math.ceil((new Date(item.endDate) - new Date(item.startDate)) / (1000 * 60 * 60 * 24))) : 1;
                const consumedDays = item.startDate ? Math.max(0, totalDuration - Math.max(daysRemaining || 0, 0)) : 0;
                const progress = Math.max(0, Math.min(100, Math.round((consumedDays / totalDuration) * 100)));
                return (
                  <div key={item.id} style={{ background: colors.panel2, border: `1px solid ${colors.border}`, borderRadius: 18, padding: 16 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "flex-start" }}>
                      <div>
                        <div style={{ fontWeight: 900, fontSize: 19 }}>{item.contractNo || "-"}</div>
                        <div style={{ color: colors.sub, marginTop: 6, fontSize: 14 }}>{item.owner} • {item.phone}</div>
                      </div>
                      <span style={statusBadge(item.calculatedStatus)}>{item.calculatedStatus}</span>
                    </div>
                    <div style={{ marginTop: 14, display: "grid", gridTemplateColumns: isMobile ? "1fr" : "repeat(4, minmax(0,1fr))", gap: 10 }}>
                      <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Şirket</div><div style={{ fontWeight: 800, marginTop: 5 }}>{item.companyName || "-"}</div></div>
                      <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Emlakçı</div><div style={{ fontWeight: 800, marginTop: 5 }}>{item.agentName || "-"}</div></div>
                      <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Satış Fiyatı</div><div style={{ fontWeight: 800, marginTop: 5 }}>{formatCurrency(item.salePrice)}</div></div>
                      <div style={{ background: colors.panel3, border: `1px solid ${colors.border}`, borderRadius: 16, padding: 12 }}><div style={{ fontSize: 12, color: colors.sub }}>Komisyon</div><div style={{ fontWeight: 800, marginTop: 5 }}>%{item.commission || "-"}</div></div>
                    </div>
                    <div style={{ marginTop: 14 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 10, fontSize: 13, color: colors.sub, marginBottom: 6 }}>
                        <span>{formatShortDate(item.startDate)} → {formatShortDate(item.endDate)}</span>
                        <span>{daysRemaining == null ? "-" : `${daysRemaining} gün`}</span>
                      </div>
                      <div style={{ height: 8, borderRadius: 999, background: colors.panel3, border: `1px solid ${colors.border}`, overflow: "hidden" }}>
                        <div style={{ width: `${progress}%`, height: "100%", background: `linear-gradient(90deg, ${item.calculatedStatus === "aktif" ? colors.success : item.calculatedStatus === "yakında bitecek" || item.calculatedStatus === "bugün bitiyor" ? colors.warning : colors.danger} 0%, ${dark ? "#60a5fa" : "#93c5fd"} 100%)` }} />
                      </div>
                    </div>
                    {item.address ? <div style={{ marginTop: 12, lineHeight: 1.6 }}>{item.address}</div> : null}
                    {item.note ? <div style={{ marginTop: 10, color: colors.sub }}>Not: {item.note}</div> : null}
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
                      <button onClick={() => startEditContract(item)} style={actionButton(false)}>Düzenle</button>
                      <button onClick={() => renewContractFromItem(item)} style={actionButton(false)}>Yenileme Taslağı</button>
                      <button onClick={() => handleExportContractPdf(item)} style={actionButton(true)}>📥 PDF İndir</button>
                      <button onClick={() => shareOnWhatsApp(item.phone, `Merhaba, ${getOfficeDisplayName()} ofisinden ${item.contractNo || "sözleşme"} numaralı yetki sözleşmenizi PDF olarak iletiyorum.`)} style={actionButton(false, { background: "#16a34a", color: "#ffffff", border: "none" })}>💬 WhatsApp Aç</button>
                      <button onClick={() => removeItem("contracts", item.id, setContracts)} style={actionButton(false, { background: colors.dangerSoft, color: colors.danger, border: `1px solid ${dark ? "rgba(248,113,113,0.16)" : "#fecaca"}` })}>Sil</button>
                    </div>
                    <div style={{ marginTop: 8, fontSize: 12, color: colors.sub, padding: "8px 12px", background: colors.panel3, borderRadius: 12, border: `1px solid ${colors.border}` }}>
                      💡 Önce <strong>PDF İndir</strong> ile sözleşmeyi indirin, ardından <strong>WhatsApp Aç</strong> ile sohbeti başlatıp PDF'i paylaşın.
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    );
  }


  let tabContent = null;

  try {
    switch (activeTab) {
      case "dashboard":
        tabContent = renderDashboard();
        break;
      case "reports":
        tabContent = renderReports();
        break;
      case "customers":
        tabContent = renderCustomers();
        break;
      case "portfolios":
        tabContent = renderPortfolios();
        break;
      case "matches":
        tabContent = renderMatches();
        break;
      case "appointments":
        tabContent = renderAppointments();
        break;
      case "contracts":
        tabContent = renderContracts();
        break;
      case "saleContracts":
        tabContent = renderSaleContracts();
        break;
      case "rentalContracts":
        tabContent = renderRentalContracts();
        break;
      case "settings":
        tabContent = renderSettings();
        break;
      default:
        tabContent = renderDashboard();
    }
  } catch (err) {
    console.error("TAB ERROR:", activeTab, err);

    tabContent = (
      <div
        style={{
          background: colors.panel,
          color: colors.text,
          border: `1px solid ${colors.border}`,
          borderRadius: 16,
          padding: 20,
          boxShadow: colors.shadow,
        }}
      >
        <h2 style={{ marginTop: 0 }}>Sekme açılırken hata oluştu</h2>
        <p><strong>Sekme:</strong> {activeTab}</p>
        <p><strong>Mesaj:</strong> {String(err?.message || err)}</p>
      </div>
    );
  }

  if (authLoading) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
          background: colors.bg,
          color: colors.text,
          fontFamily: "Arial, sans-serif",
        }}
      >
        Yükleniyor...
      </div>
    );
  }

  if (user && dataLoading && !migrationDone) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
          background: colors.bg,
          color: colors.text,
          fontFamily: "Arial, sans-serif",
        }}
      >
        Veriler yükleniyor...
      </div>
    );
  }

  if (!user) {
    return renderAuthScreen();
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        background: dark
          ? "linear-gradient(180deg, #07111f 0%, #0d1728 100%)"
          : "linear-gradient(180deg, #eef4ff 0%, #f8fbff 100%)",
        color: colors.text,
        fontFamily: "Arial, sans-serif",
        display: "grid",
        gridTemplateColumns: isMobile ? "1fr" : "250px minmax(0, 1fr)",
        width: "100%",
        maxWidth: "100%",
        overflowX: "hidden",
        boxSizing: "border-box",
      }}
    >
      <aside
        style={{
          background: dark ? "rgba(16,26,43,0.9)" : "rgba(255,255,255,0.88)",
          borderRight: isMobile ? "none" : `1px solid ${colors.border}`,
          borderBottom: isMobile ? `1px solid ${colors.border}` : "none",
          padding: 16,
          display: "flex",
          flexDirection: "column",
          justifyContent: "space-between",
          gap: 18,
          position: isMobile ? "relative" : "sticky",
          top: 0,
          height: isMobile ? "auto" : "100vh",
          backdropFilter: "blur(14px)",
        }}
      >
        <div>
          <div
            style={{
              padding: 16,
              borderRadius: 22,
              background: colors.panel2,
              border: `1px solid ${colors.border}`,
              marginBottom: 16,
            }}
          >
            <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
              <div
                style={{
                  width: 54,
                  height: 54,
                  borderRadius: 18,
                  background: dark ? "#60a5fa" : "#2563eb",
                  color: "#ffffff",
                  display: "grid",
                  placeItems: "center",
                  fontWeight: 900,
                  fontSize: 22,
                }}
              >
                E
              </div>

              <div>
                <div style={{ fontSize: 22, fontWeight: 900 }}>e-key</div>
                <div style={{ fontSize: 13, color: colors.sub }}>Profesyonel Emlak CRM Paneli</div>
              </div>
            </div>

            <div style={{ marginTop: 14, color: colors.sub, lineHeight: 1.6, fontSize: 13 }}>
              Günlük takibini tek panelden yönet, müşterini ve portföyünü kontrol altında tut.
            </div>
          </div>

          <div
            style={{
              display: "flex",
              flexDirection: isMobile ? "row" : "column",
              gap: 10,
              overflowX: isMobile ? "auto" : "visible",
              paddingBottom: isMobile ? 6 : 0,
            }}
          >
            {navItems.map((item) => {
              const active = activeTab === item.key;
              return (
                <button
                  key={item.key}
                  onClick={() => setActiveTab(item.key)}
                  style={actionButton(false, {
                    minWidth: isMobile ? 170 : 0,
                    width: isMobile ? "auto" : "100%",
                    justifyContent: isMobile ? "center" : "flex-start",
                    display: "flex",
                    alignItems: "center",
                    gap: 12,
                    whiteSpace: "nowrap",
                    flexShrink: 0,
                    background: active ? colors.primary : colors.panel2,
                    color: active ? "#ffffff" : colors.text,
                    border: active ? "none" : `1px solid ${colors.border}`,
                  })}
                >
                  <span style={{ fontSize: 18 }}>{item.icon}</span>
                  <span>{item.label}</span>
                </button>
              );
            })}
          </div>
        </div>

        <div style={{ display: "grid", gap: 10 }}>
          <div
            style={{
              padding: "14px 16px",
              background: colors.panel2,
              borderRadius: 16,
              fontSize: 14,
              wordBreak: "break-word",
              border: `1px solid ${colors.border}`,
            }}
          >
            <div style={{ color: colors.sub, fontSize: 12, marginBottom: 6 }}>Oturum</div>
            <div style={{ fontWeight: 800 }}>{user.email}</div>
          </div>

          <button
            onClick={() => setTheme((t) => (t === "light" ? "dark" : "light"))}
            style={actionButton(false)}
          >
            {theme === "light" ? "🌙 Koyu Tema" : "☀️ Açık Tema"}
          </button>

          <button onClick={handleLogout} style={actionButton(false)}>
            Çıkış
          </button>
        </div>
      </aside>

      <main style={{ padding: isMobile ? 10 : 16, minWidth: 0, overflowX: "hidden", width: "100%", maxWidth: "100%", boxSizing: "border-box" }}>
        <div
          style={{
            background: colors.panel,
            border: `1px solid ${colors.border}`,
            borderRadius: 24,
            padding: 18,
            boxShadow: colors.shadow,
            marginBottom: 18,
          }}
        >
          <div
  style={{
    display: "grid",
    gridTemplateColumns: isMobile ? "1fr" : "minmax(0, 1.75fr) minmax(360px, 0.75fr)",
    gap: isMobile ? 16 : 18,
    alignItems: "start",
  }}
>
            <div>
              <div
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 8,
                  padding: "8px 12px",
                  borderRadius: 999,
                  background: colors.primarySoft,
                  color: colors.primary,
                  fontWeight: 800,
                  fontSize: 12,
                  marginBottom: 14,
                }}
              >
                ⚡ Profesyonel çalışma paneli
              </div>

             <h1
  style={{
    margin: 0,
    fontSize: isMobile ? 28 : 34,
    fontWeight: 900,
    lineHeight: 1.12,
    maxWidth: 780,
  }}
>
                e-key ile iş akışını tek yerden yönet
              </h1>

              <p style={{ margin: "10px 0 0", color: colors.sub, lineHeight: 1.7, maxWidth: "100%" }}>
                Müşteri, portföy, sözleşme, eşleştirme ve randevularını modern ve düzenli bir
                arayüzle takip et. Hızlı aksiyonlar, uyarılar ve canlı özetler tek ekranda.
              </p>

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 16 }}>
                <button onClick={() => setActiveTab("customers")} style={actionButton(false)}>
                  Müşteri Ekle
                </button>
                <button onClick={() => setActiveTab("portfolios")} style={actionButton(false)}>
                  Portföy Ekle
                </button>
                <button onClick={() => setActiveTab("contracts")} style={actionButton(false)}>
                 ⚡  Hızlı Sözleşme
                </button>
              </div>
            </div>

<div
  style={{
    background: colors.panel2,
    border: `1px solid ${colors.border}`,
    borderRadius: 18,
    padding: 16,
    display: "grid",
    gap: 10,
    minHeight: "100%",
    alignContent: "start",
  }}
>
              <div style={{ fontWeight: 900, fontSize: 16 }}>Bugün ve sıradaki işler</div>
              <div style={{ color: colors.sub, lineHeight: 1.6, fontSize: 13 }}>
                Yaklaşan randevuların ve sıradaki iş akışın burada görünür.
              </div>

              {dashboardAgendaItems.length === 0 ? (
                <div
                  style={{
                    background: colors.panel3,
                    border: `1px solid ${colors.border}`,
                    borderRadius: 16,
                    padding: 14,
                    color: colors.sub,
                    lineHeight: 1.6,
                    fontSize: 13,
                  }}
                >
                  Yeni bir randevu eklediğinde burada görünecek.
                </div>
              ) : (
                <div style={{ display: "grid", gap: 8 }}>
                  {dashboardAgendaItems.map((item) => (
                    <div
                      key={item.id}
                      style={{
                        background: colors.panel3,
                        border: `1px solid ${colors.border}`,
                        borderRadius: 16,
                        padding: 12,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          gap: 10,
                          alignItems: "flex-start",
                          flexWrap: "wrap",
                        }}
                      >
                        <div style={{ minWidth: 0, flex: 1 }}>
                          <div style={{ fontWeight: 900, fontSize: 13.5 }}>{item.title || "Randevu"}</div>
                          <div style={{ color: colors.sub, marginTop: 4, fontSize: 12, lineHeight: 1.45 }}>
                            {formatShortDate(item.date)} {item.time || "-"} • {getCustomerNameById(item.customerId)}
                          </div>
                          <div style={{ color: colors.sub, marginTop: 2, fontSize: 12 }}>
                            {getPortfolioTitleById(item.portfolioId)}
                          </div>
                          {item.location ? (
                            <div style={{ color: colors.sub, marginTop: 2, fontSize: 12 }}>
                              Konum: {item.location}
                            </div>
                          ) : null}
                        </div>
                        <span style={statusBadge(item.visualStatus)}>{item.visualStatus}</span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>

        {tabContent}
      </main>
    </div>
  );
}

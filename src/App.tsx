import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  ShoppingBag,
  Store,
  ShieldCheck,
  Users,
  Settings,
  Package,
  Layers,
  CreditCard,
  Plus,
  Trash2,
  Edit,
  Sparkles,
  Video,
  Image as ImageIcon,
  CheckCircle,
  Clock,
  AlertTriangle,
  Search,
  Filter,
  DollarSign,
  Percent,
  ChevronLeft,
  X,
  Menu,
  Printer,
  Eye,
  Lock,
  UserCheck,
  TrendingUp,
  Tag,
  Palette,
  Bell,
  Check,
  ArrowRight,
  RefreshCw,
  LogOut,
  LogIn,
  FileText,
  Star,
  Zap,
  Shield,
  Gift,
  Mail,
  ArrowUpDown
} from "lucide-react";

// Types
interface Review {
  id: string;
  customerName: string;
  rating: number;
  comment: string;
  date: string;
}

interface Product {
  id: string;
  name: string;
  description: string;
  price: number;
  discountPrice?: number;
  stock: number;
  category: string;
  image: string;
  tags: string[];
  featured?: boolean;
  reviews?: Review[];
}

interface Category {
  id: string;
  name: string;
  icon: string;
  color: string;
}

interface Employee {
  id: string;
  name: string;
  email: string;
  phone: string;
  password?: string;
  role: "super_admin" | "store_manager" | "content_creator";
  permissions: {
    canManageProducts: boolean;
    canManageOrders: boolean;
    canManageEmployees: boolean;
    canManageSettings: boolean;
  };
  status: "active" | "inactive";
}

interface CheckoutField {
  id: string;
  label: string;
  type: "text" | "number" | "textarea";
  required: boolean;
}

interface PromoCode {
  code: string;
  discountPercent: number;
}

interface Message {
  id: string;
  senderName: string;
  senderEmail: string;
  senderPhone?: string;
  subject: string;
  content: string;
  createdAt: string;
  status: "unread" | "read" | "replied";
  reply?: string;
  repliedAt?: string;
  repliedBy?: string;
}

interface Order {
  id: string;
  customerName: string;
  customerEmail: string;
  customerPhone: string;
  customFieldValues?: Record<string, string>;
  items: { product: Product; quantity: number }[];
  totalAmount: number;
  discountApplied?: number;
  paymentMethod: string;
  status: "pending" | "processing" | "shipped" | "completed" | "cancelled";
  createdAt: string;
}

interface StoreSettings {
  storeName: string;
  slogan: string;
  primaryColor: string;
  backgroundColor: string;
  announcement: string;
  bannerTitle: string;
  bannerSubtitle: string;
  currency: string;
  paymentMethods: {
    mada: boolean;
    visa: boolean;
    cod: boolean;
  };
  customCheckoutFields: CheckoutField[];
  promoCodes: PromoCode[];
}

const INITIAL_CATEGORIES: Category[] = [
  { id: "tech", name: "إلكترونيات ذكية", icon: "Cpu", color: "bg-blue-500" },
  { id: "fashion", name: "أزياء عصرية", icon: "Shirt", color: "bg-purple-500" },
  { id: "home", name: "ديكور ومنزل", icon: "Home", color: "bg-emerald-500" },
  { id: "beauty", name: "صحة وجمال", icon: "Sparkles", color: "bg-pink-500" },
];

const INITIAL_PRODUCTS: Product[] = [
  {
    id: "p1",
    name: "ساعة ذكية برو الترا",
    description: "ساعة ذكية متطورة قياس نبضات القلب، مقاومة للماء مع شاشة AMOLED فائقة الوضوح وعمر بطارية يصل إلى 7 أيام.",
    price: 499,
    discountPrice: 399,
    stock: 25,
    category: "tech",
    image: "https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&w=600&q=80",
    tags: ["ذكية", "ساعة", "إلكترونيات"],
    featured: true,
    reviews: [
      { id: "r1", customerName: "سارة خالد", rating: 5, comment: "ساعة رائعة جداً وعملية أنصح به!", date: "2026-10-08" },
      { id: "r2", customerName: "عبدالله محمد", rating: 4, comment: "ممتازة والبطارية تدوم طويلاً.", date: "2026-10-09" }
    ]
  },
  {
    id: "p2",
    name: "سماعات رأس لاسلكية إلغاء الضوضاء",
    description: "سماعات فوق الأذن بجودة صوت ستيريو محيطية وتقنية إلغاء الضوضاء النشطة ANC لتجربة استماع خالية من التشويش.",
    price: 650,
    discountPrice: 520,
    stock: 12,
    category: "tech",
    image: "https://images.unsplash.com/photo-1505740420928-5e560c06d30e?auto=format&fit=crop&w=600&q=80",
    tags: ["سماعات", "صوتيات", "بلوتوث"],
    featured: true,
    reviews: [
      { id: "r3", customerName: "فهد العمري", rating: 5, comment: "عزل ضوضاء خيالي ونقاء صوت مذهل.", date: "2026-10-07" }
    ]
  },
  {
    id: "p3",
    name: "حقيبة ظهر جلدية فاخرة للكمبيوتر",
    description: "حقيبة ظهر عملية وأنيقة مصنوعة من الجلد الطبيعي المضاد للماء، مخصصة لأجهزة اللابتوب حتى 15.6 إنش.",
    price: 290,
    stock: 8,
    category: "fashion",
    image: "https://images.unsplash.com/photo-1553062407-98eeb64c6a62?auto=format&fit=crop&w=600&q=80",
    tags: ["حقائب", "جلد", "أزياء"],
    featured: false,
    reviews: []
  },
  {
    id: "p4",
    name: "مصباح مكتب ذكي إضاءة RGB",
    description: "مصباح مكتبي بتقنية LED مع درجات إضاءة قابلة للتعديل وشاحن لاسلكي مدمج للهواتف الذكية.",
    price: 180,
    discountPrice: 140,
    stock: 40,
    category: "home",
    image: "https://images.unsplash.com/photo-1534349762230-e8cadfec3afc?auto=format&fit=crop&w=600&q=80",
    tags: ["منزل", "إضاءة", "ذكية"],
    featured: true,
    reviews: [
      { id: "r4", customerName: "منى الحربي", rating: 5, comment: "إضاءة مكتبية مريحة جداً وتصميم أنيق.", date: "2026-10-09" }
    ]
  }
];

const INVOICE_PRODUCTS: Product[] = [
  { id: "inv-1", name: "لا روش بوزيه إيفاكالر ديو + 3337875750042", description: "جل لا روش بوزيه إيفاكالر ديو للعناية بالبشرة المعرضة للشوائب", price: 44, stock: 30, category: "beauty", image: "https://images.unsplash.com/photo-1556228720-195a672e8a03?auto=format&fit=crop&w=600&q=80", tags: ["عناية", "لاروش"], reviews: [] },
  { id: "inv-2", name: "جل لا روش بوساي إفاكالر إم 400 مل", description: "غسول جل من لاروش بوزيه لتنقية البشرة الدهنية 400 مل", price: 44, stock: 25, category: "beauty", image: "https://images.unsplash.com/photo-1570172619644-dfd03ed5d881?auto=format&fit=crop&w=600&q=80", tags: ["غسول", "لاروش"], reviews: [] },
  { id: "inv-3", name: "كريم مرطب سريافيه 454 جرام", description: "كريم مرطب غني بالسراميد وحمض الهيالورونيك للبشرة الجافة", price: 54, stock: 40, category: "beauty", image: "https://images.unsplash.com/photo-1608248597359-9d7a22709e3e?auto=format&fit=crop&w=600&q=80", tags: ["سيرافي", "مرطب"], reviews: [] },
  { id: "inv-4", name: "بيوتي اوف جونسون ريليف صن: الأرز + البريبيوتيك 50ml", description: "واقي شمس كوري بخلاصة الأرز والبربيوتيك لحماية فائقة", price: 31, stock: 50, category: "beauty", image: "https://images.unsplash.com/photo-1598440947619-2c35fc9aa908?auto=format&fit=crop&w=600&q=80", tags: ["واقي شمس", "جونسون"], reviews: [] },
  { id: "inv-5", name: "لوشن مرطب للوجه خفيف الوزن للغاية من سرياف", description: "لوشن مرطب للوجه مع 3 سراميدات أساسية وحمض الهيالورونيك", price: 52, stock: 35, category: "beauty", image: "https://images.unsplash.com/photo-1620916566398-39f1143ab7be?auto=format&fit=crop&w=600&q=80", tags: ["سيرافي", "لوشن"], reviews: [] },
  { id: "inv-6", name: "يوسريين - جل حماية من الشمس للتحكم بالزيوت", description: "جل واقي شمس للبشرة الدهنية والمعرضة لحب الشباب", price: 69, stock: 20, category: "beauty", image: "https://images.unsplash.com/photo-1535585209827-a15fcdbc4c2d?auto=format&fit=crop&w=600&q=80", tags: ["يوسيرين", "حماية"], reviews: [] },
  { id: "inv-7", name: "شامبو ميلي المقوي بالروزماري والنعناع", description: "شامبو مغذي ومقوي لفروة الرأس والشعر بالروزماري والبيوتين", price: 25, stock: 60, category: "beauty", image: "https://images.unsplash.com/photo-1535585209827-a15fcdbc4c2d?auto=format&fit=crop&w=600&q=80", tags: ["شامبو", "ميلي"], reviews: [] },
  { id: "inv-8", name: "سيروم فيتامين سي من سريافي 30 مل", description: "سيروم فيتامين سي النقي لتفتيح البشرة وتحفيز الكولاجين", price: 68, stock: 30, category: "beauty", image: "https://images.unsplash.com/photo-1620916566398-39f1143ab7be?auto=format&fit=crop&w=600&q=80", tags: ["سيروم", "سيرافي"], reviews: [] },
  { id: "inv-9", name: "قناع الشعر K18 لإصلاح الشعر الجزيئي بدون شطف", description: "علاج ثوري لإصلاح الشعر التالف وكسور الروابط الجزيئية", price: 69, stock: 15, category: "beauty", image: "https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?auto=format&fit=crop&w=600&q=80", tags: ["شعر", "K18"], reviews: [] },
  { id: "inv-10", name: "ذا أورديناري - محلول الكافيين 5% + EGCG 30 مل", description: "محلول مركز للهالات السوداء والانتفاخ تحت العينين", price: 32, stock: 45, category: "beauty", image: "https://images.unsplash.com/photo-1608248597359-9d7a22709e3e?auto=format&fit=crop&w=600&q=80", tags: ["أورديناري", "عينين"], reviews: [] },
  { id: "inv-11", name: "ذا أورديناري نياسيناميد 10% + زنك 1% 30 مل", description: "سيروم تنظيم إفراز الدهون وتقليل المسام الواسعة", price: 32, stock: 50, category: "beauty", image: "https://images.unsplash.com/photo-1620916566398-39f1143ab7be?auto=format&fit=crop&w=600&q=80", tags: ["أورديناري", "نياسيناميد"], reviews: [] },
  { id: "inv-12", name: "ذا أورديناري حمض الهيالورونيك العادي 2% + ترطيب B5", description: "مصل مرطب عميق لملء الخطوط الدقيقة وإنعاش البشرة", price: 32, stock: 55, category: "beauty", image: "https://images.unsplash.com/photo-1556228720-195a672e8a03?auto=format&fit=crop&w=600&q=80", tags: ["أورديناري", "هيالورونيك"], reviews: [] },
  { id: "inv-13", name: "غسول بانوكسيل الرغوي لعلاج حب الشباب 10%", description: "غسول طبي فعال بتركيز 10% بنزويل بيروكسيد للقضاء على حب الشباب", price: 36, stock: 30, category: "beauty", image: "https://images.unsplash.com/photo-1570172619644-dfd03ed5d881?auto=format&fit=crop&w=600&q=80", tags: ["بانوكسيل", "حب الشباب"], reviews: [] },
  { id: "inv-14", name: "بيوديرما سينسيبيو H2O منظف مكياج 500 مل", description: "ماء ميسيلار التنظيف المفضل للبشرة الحساسة وإزالة المكياج بلطف", price: 25, stock: 40, category: "beauty", image: "https://images.unsplash.com/photo-1608248597359-9d7a22709e3e?auto=format&fit=crop&w=600&q=80", tags: ["بيوديرما", "ميسيلار"], reviews: [] },
  { id: "inv-15", name: "أولابلكس ان 3 ترميم واصلاح الشعر 100مل", description: "علاج منزلي مكثف لإصلاح وترميم الشعر التالف والمعالج حرارياً", price: 33, stock: 25, category: "beauty", image: "https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?auto=format&fit=crop&w=600&q=80", tags: ["أولابلكس", "شعر"], reviews: [] },
  { id: "inv-16", name: "سيروم تعزيز كثافة الرموش والحواجب - مصل متعدد الببتيدات", description: "مصل مغذي لنمو وتكثيف الرموش والحواجب بفعالية عالية", price: 32, stock: 35, category: "beauty", image: "https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?auto=format&fit=crop&w=600&q=80", tags: ["رموش", "عناية"], reviews: [] },
  { id: "inv-17", name: "ذا أورديناري تونر حمض الجليكوليك 7% للتقشير والتفتيح", description: "تونر مقشر بحمض الجليكوليك 7% لتجديد خلايا البشرة وتوحيد لونها", price: 38, stock: 40, category: "beauty", image: "https://images.unsplash.com/photo-1620916566398-39f1143ab7be?auto=format&fit=crop&w=600&q=80", tags: ["أورديناري", "تونر"], reviews: [] },
  { id: "inv-18", name: "بيوديرما سيكابيو 40 مل للإصلاح والتهدئه", description: "كريم مهدئ ومصلح للبشرة المتهيجة والمتضررة", price: 38, stock: 30, category: "beauty", image: "https://images.unsplash.com/photo-1570172619644-dfd03ed5d881?auto=format&fit=crop&w=600&q=80", tags: ["بيوديرما", "سيكابيو"], reviews: [] }
];

const INITIAL_EMPLOYEES: Employee[] = [
  {
    id: "e1",
    name: "أحمد السوبر أدمين",
    email: "admin@smartstore.com",
    phone: "123456789",
    role: "super_admin",
    permissions: { canManageProducts: true, canManageOrders: true, canManageEmployees: true, canManageSettings: true },
    status: "active"
  }
];

function getAverageRating(reviews?: Review[]): number {
  if (!reviews || reviews.length === 0) return 5.0;
  const sum = reviews.reduce((acc, r) => acc + r.rating, 0);
  return Number((sum / reviews.length).toFixed(1));
}

export default function App() {
  const [activeTab, setActiveTab] = useState<"store" | "admin" | "orders" | "profile">("store");
  const [adminSubTab, setAdminSubTab] = useState<"overview" | "products" | "categories" | "employees" | "customers" | "settings" | "payments" | "checkout_fields" | "coupons" | "messages">("overview");

  // Persistent States
  const [settings, setSettings] = useState<StoreSettings>(() => {
    try {
      const saved = localStorage.getItem("smart_store_settings");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (!parsed.backgroundColor || parsed.backgroundColor === "#000000" || parsed.backgroundColor === "#000") {
          parsed.backgroundColor = "#f8fafc";
        }
        if (!parsed.customCheckoutFields) {
          parsed.customCheckoutFields = [
            { id: "f1", label: "ملاحظات التوصيل أو رقم المبنى", type: "text", required: false }
          ];
        }
        if (!parsed.promoCodes) {
          parsed.promoCodes = [
            { code: "SMART20", discountPercent: 20 },
            { code: "VIP10", discountPercent: 10 }
          ];
        }
        return parsed;
      }
    } catch (e) {
      console.error(e);
    }
    return {
      storeName: "متجر السوبر الذكي",
      slogan: "وجهتك الأولى للتسوق الذكي بالذكاء الاصطناعي",
      primaryColor: "#6366f1",
      backgroundColor: "#f8fafc",
      announcement: "🔥 خصم 20% على جميع المنتجات الإلكترونية لفترة محدودة! استخدم كود: SMART20",
      bannerTitle: "تسوق المستقبل بذكاء",
      bannerSubtitle: "اكتشف تشكيلة واسعة من المنتجات الذكية مع تجربة تسوق ديناميكية متطورة",
      currency: "ر.س",
      paymentMethods: { mada: true, visa: true, cod: true },
      customCheckoutFields: [
        { id: "f1", label: "ملاحظات التوصيل أو رقم المبنى", type: "text", required: false }
      ],
      promoCodes: [
        { code: "SMART20", discountPercent: 20 },
        { code: "VIP10", discountPercent: 10 }
      ]
    };
  });

  const [products, setProducts] = useState<Product[]>(() => {
    const saved = localStorage.getItem("smart_store_products");
    const initial = saved ? JSON.parse(saved) : INITIAL_PRODUCTS;
    const seen = new Set();
    return initial.filter((p: Product) => {
      if (seen.has(p.id)) return false;
      seen.add(p.id);
      return true;
    });
  });

  const [categories, setCategories] = useState<Category[]>(() => {
    const saved = localStorage.getItem("smart_store_categories");
    const initial = saved ? JSON.parse(saved) : INITIAL_CATEGORIES;
    const seen = new Set();
    return initial.filter((c: Category) => {
      if (seen.has(c.id)) return false;
      seen.add(c.id);
      return true;
    });
  });

  const [employees, setEmployees] = useState<Employee[]>(() => {
    const saved = localStorage.getItem("smart_store_employees");
    const initial = saved ? JSON.parse(saved) : INITIAL_EMPLOYEES;
    const seen = new Set();
    return initial.filter((e: Employee) => {
      if (seen.has(e.id)) return false;
      seen.add(e.id);
      return true;
    });
  });

  const [orders, setOrders] = useState<Order[]>(() => {
    const saved = localStorage.getItem("smart_store_orders");
    const initial = saved ? JSON.parse(saved) : [
      {
        id: "ORD-9821",
        customerName: "محمد الفهد",
        customerEmail: "mohammed@example.com",
        customerPhone: "0501234567",
        items: [{ product: INITIAL_PRODUCTS[0], quantity: 1 }],
        totalAmount: 399,
        paymentMethod: "مدى (Mada)",
        status: "completed",
        createdAt: "2026-10-10 14:30"
      }
    ];
    const seen = new Set();
    return initial.filter((o: Order) => {
      if (seen.has(o.id)) return false;
      seen.add(o.id);
      return true;
    });
  });

  const [messages, setMessages] = useState<Message[]>(() => {
    try {
      const saved = localStorage.getItem("smart_store_messages");
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error(e);
    }
    return [
      {
        id: "msg-1",
        senderName: "سارة خالد",
        senderEmail: "sara@example.com",
        senderPhone: "0501234567",
        subject: "استفسار عن موعد توصيل الطلب",
        content: "مرحباً، أود الاستفسار متى سيتم شحن الطلب رقم ORD-9821؟",
        createdAt: "2026-10-10 16:00",
        status: "replied",
        reply: "أهلاً بكِ سارة، تم شحن طلبكِ اليوم وسيصلكِ خلال 48 ساعة.",
        repliedAt: "2026-10-10 17:30",
        repliedBy: "أحمد السوبر أدمين"
      }
    ];
  });

  useEffect(() => {
    localStorage.setItem("smart_store_messages", JSON.stringify(messages));
  }, [messages]);

  const [isContactModalOpen, setIsContactModalOpen] = useState(false);
  const [contactForm, setContactForm] = useState({
    issueType: "technical" as "technical" | "order",
    selectedOrderId: "",
    subject: "",
    content: ""
  });
  const [adminReplyText, setAdminReplyText] = useState("");
  const [selectedAdminMessageId, setSelectedAdminMessageId] = useState<string | null>(null);

  const [currentUser, setCurrentUser] = useState<Employee | null>(() => {
    const saved = localStorage.getItem("smart_store_current_user");
    return saved ? JSON.parse(saved) : null;
  });
  const [isLoginModalOpen, setIsLoginModalOpen] = useState(false);
  const [loginPhone, setLoginPhone] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const [userProfileForm, setUserProfileForm] = useState({
    name: currentUser?.name || "عميل المتجر الذكي",
    email: currentUser?.email || "customer@store.com",
    phone: currentUser?.phone || "0501234567",
    address: "الرياض، المملكة العربية السعودية"
  });

  const handleUpdateProfile = (e: React.FormEvent) => {
    e.preventDefault();
    if (currentUser) {
      setEmployees(prev => prev.map(emp => emp.id === currentUser.id ? { ...emp, name: userProfileForm.name, email: userProfileForm.email, phone: userProfileForm.phone } : emp));
      setCurrentUser(prev => prev ? { ...prev, name: userProfileForm.name, email: userProfileForm.email, phone: userProfileForm.phone } : null);
    }
    alert("تم حفظ تحديثات الملف الشخصي بنجاح!");
  };

  const handleDetectLocation = () => {
    if ("geolocation" in navigator) {
      navigator.geolocation.getCurrentPosition(
        async (position) => {
          const lat = position.coords.latitude;
          const lon = position.coords.longitude;
          try {
            const res = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&accept-language=ar`);
            const data = await res.json();
            if (data && data.display_name) {
              setUserProfileForm(prev => ({ ...prev, address: data.display_name }));
              alert(`تم تحديد موقعك الحالي تلقائياً بنجاح: ${data.display_name}`);
              return;
            }
          } catch (e) {
            console.error(e);
          }
          setUserProfileForm(prev => ({ ...prev, address: `موقع GPS الحالي (خط عرض: ${lat.toFixed(4)}, خط طول: ${lon.toFixed(4)})` }));
          alert("تم تحديد إحداثيات موقعك الحالي بنجاح عبر GPS!");
        },
        () => {
          alert("تعذر الوصول إلى الموقع تلقائياً. يرجى تفعيل خدمة الموقع أو إدخال العنوان يدويام.");
        }
      );
    } else {
      alert("متصفحك لا يدعم تحديد الموقع الجغرافي.");
    }
  };

  // Authentication state & /sts route detection
  const [isStsRoute, setIsStsRoute] = useState(() => {
    return window.location.pathname.includes('/sts') || window.location.hash.includes('sts');
  });
  const [selectedProductIds, setSelectedProductIds] = useState<string[]>([]);
  const [selectedOrderIds, setSelectedOrderIds] = useState<string[]>([]);
  const [selectedCustomerEmails, setSelectedCustomerEmails] = useState<string[]>([]);

  useEffect(() => {
    const checkSts = () => {
      const isSts = window.location.pathname.includes('/sts') || window.location.hash.includes('sts');
      setIsStsRoute(isSts);
      const savedUser = localStorage.getItem("smart_store_current_user");
      if (isSts) {
        if (!savedUser) {
          const adminUser: Employee = {
            id: "e1",
            name: "أحمد السوبر أدمين",
            email: "admin@smartstore.com",
            phone: "123456789",
            password: "123456",
            role: "super_admin",
            permissions: { canManageProducts: true, canManageOrders: true, canManageEmployees: true, canManageSettings: true },
            status: "active"
          };
          setCurrentUser(adminUser);
        }
        setActiveTab("admin");
      }
    };
    checkSts();
    window.addEventListener('popstate', checkSts);
    window.addEventListener('hashchange', checkSts);
    return () => {
      window.removeEventListener('popstate', checkSts);
      window.removeEventListener('hashchange', checkSts);
    };
  }, []);

  // Cart & Checkout State
  const [cart, setCart] = useState<{ product: Product; quantity: number }[]>([]);
  const [isCartOpen, setIsCartOpen] = useState(false);
  const [isCheckoutOpen, setIsCheckoutOpen] = useState(false);
  const [promoInput, setPromoInput] = useState("");
  const [appliedPromo, setAppliedPromo] = useState<PromoCode | null>(null);
  const [promoError, setPromoError] = useState("");
  const [checkoutForm, setCheckoutForm] = useState({
    name: "",
    email: "",
    phone: "",
    address: "",
    paymentMethod: "مدى (Mada)"
  });
  const [customFieldInputs, setCustomFieldInputs] = useState<Record<string, string>>({});
  const [lastOrderReceipt, setLastOrderReceipt] = useState<Order | null>(null);

  // Storefront Filtering & Sorting & Search
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState("all");
  const [sortBy, setSortBy] = useState<"default" | "price-low" | "price-high" | "rating">("default");
  const [selectedProductDetails, setSelectedProductDetails] = useState<Product | null>(null);

  // New Review Form State inside product details modal
  const [newReviewName, setNewReviewName] = useState("");
  const [newReviewRating, setNewReviewRating] = useState(5);
  const [newReviewComment, setNewReviewComment] = useState("");

  // AI Modal & State in Admin Product Form (Staff Only)
  const [isProductModalOpen, setIsProductModalOpen] = useState(false);
  const [editingProduct, setEditingProduct] = useState<Product | null>(null);
  const [productForm, setProductForm] = useState({
    name: "",
    description: "",
    price: 100,
    discountPrice: 0,
    stock: 10,
    category: INITIAL_CATEGORIES[0].id,
    image: "",
    tags: ""
  });
  const [aiPrompt, setAiPrompt] = useState("");
  const [isAiLoading, setIsAiLoading] = useState(false);

  // Video AI Generation State (Staff Only)
  const [videoModalProduct, setVideoModalProduct] = useState<Product | null>(null);
  const [videoOperationName, setVideoOperationName] = useState<string | null>(null);
  const [isVideoLoading, setIsVideoLoading] = useState(false);
  const [videoStatusText, setVideoStatusText] = useState("");
  const [generatedVideoUrl, setGeneratedVideoUrl] = useState<string | null>(null);

  // Category Modal & Bulk Discount State
  const [isCategoryModalOpen, setIsCategoryModalOpen] = useState(false);
  const [editingCategory, setEditingCategory] = useState<Category | null>(null);
  const [categoryForm, setCategoryForm] = useState({ name: "", color: "bg-indigo-500", icon: "Layers" });
  const [bulkDiscountCat, setBulkDiscountCat] = useState("all");
  const [bulkDiscountPct, setBulkDiscountPct] = useState(15);

  // Employee Modal State for Super Admin
  const [isEmployeeModalOpen, setIsEmployeeModalOpen] = useState(false);
  const [editingEmployee, setEditingEmployee] = useState<Employee | null>(null);
  const [employeeForm, setEmployeeForm] = useState({
    name: "",
    phone: "",
    password: "123456",
    canManageProducts: true,
    canManageOrders: true,
    canManageEmployees: false,
    canManageSettings: false
  });

  // Save to localStorage
  useEffect(() => {
    localStorage.setItem("smart_store_settings", JSON.stringify(settings));
  }, [settings]);

  useEffect(() => {
    localStorage.setItem("smart_store_products", JSON.stringify(products));
  }, [products]);

  useEffect(() => {
    localStorage.setItem("smart_store_categories", JSON.stringify(categories));
  }, [categories]);

  useEffect(() => {
    localStorage.setItem("smart_store_employees", JSON.stringify(employees));
  }, [employees]);

  useEffect(() => {
    localStorage.setItem("smart_store_orders", JSON.stringify(orders));
  }, [orders]);

  useEffect(() => {
    if (currentUser) {
      localStorage.setItem("smart_store_current_user", JSON.stringify(currentUser));
    } else {
      localStorage.removeItem("smart_store_current_user");
    }
  }, [currentUser]);

  // Login Handler
  const handleAdminLogin = (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError("");

    if (loginPhone === "123456789" && loginPassword === "123456") {
      const adminUser: Employee = {
        id: "e1",
        name: "أحمد السوبر أدمين",
        email: "admin@smartstore.com",
        phone: "123456789",
        password: "123456",
        role: "super_admin",
        permissions: { canManageProducts: true, canManageOrders: true, canManageEmployees: true, canManageSettings: true },
        status: "active"
      };
      setCurrentUser(adminUser);
      setIsLoginModalOpen(false);
      setActiveTab("admin");
      setLoginPhone("");
      setLoginPassword("");
      return;
    }

    const found = employees.find(e => e.phone === loginPhone);
    const expectedPassword = found?.password || "123456";
    if (found && loginPassword === expectedPassword) {
      setCurrentUser(found);
      setIsLoginModalOpen(false);
      setActiveTab("admin");
      setLoginPhone("");
      setLoginPassword("");
      return;
    }

    setLoginError("رقم الجوال (اسم المستخدم) أو كلمة المرور غير صحيحة.");
  };

  const handleLogout = () => {
    setCurrentUser(null);
    localStorage.removeItem("smart_store_current_user");
    setActiveTab("store");
    setLoginPhone("");
    setLoginPassword("");
    setLoginError("");
    setIsLoginModalOpen(true);
  };

  // Cart & Coupon Handlers
  const addToCart = (product: Product) => {
    setCart(prev => {
      const existing = prev.find(item => item.product.id === product.id);
      if (existing) {
        return prev.map(item =>
          item.product.id === product.id ? { ...item, quantity: item.quantity + 1 } : item
        );
      }
      return [...prev, { product, quantity: 1 }];
    });
    setIsCartOpen(true);
  };

  const updateCartQuantity = (productId: string, delta: number) => {
    setCart(prev =>
      prev
        .map(item => {
          if (item.product.id === productId) {
            const newQty = item.quantity + delta;
            return newQty > 0 ? { ...item, quantity: newQty } : null;
          }
          return item;
        })
        .filter(Boolean) as { product: Product; quantity: number }[]
    );
  };

  const cartSubtotal = cart.reduce((sum, item) => {
    const price = item.product.discountPrice ?? item.product.price;
    return sum + price * item.quantity;
  }, 0);

  const discountAmount = appliedPromo ? Math.round(cartSubtotal * (appliedPromo.discountPercent / 100)) : 0;
  const cartTotal = Math.max(0, cartSubtotal - discountAmount);

  const handleApplyPromo = (e: React.FormEvent) => {
    e.preventDefault();
    setPromoError("");
    const found = settings.promoCodes.find(p => p.code.toUpperCase() === promoInput.trim().toUpperCase());
    if (found) {
      setAppliedPromo(found);
      setPromoInput("");
    } else {
      setPromoError("كود الخصم غير صالح أو منتهي الصلاحية");
    }
  };

  const handleCheckoutSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (cart.length === 0) return;

    const newOrder: Order = {
      id: `ORD-${Math.floor(1000 + Math.random() * 9000)}`,
      customerName: checkoutForm.name || "عميل زائر",
      customerEmail: checkoutForm.email || "customer@store.com",
      customerPhone: checkoutForm.phone || "0500000000",
      customFieldValues: { ...customFieldInputs },
      items: [...cart],
      totalAmount: cartTotal,
      discountApplied: discountAmount,
      paymentMethod: checkoutForm.paymentMethod,
      status: "pending",
      createdAt: new Date().toISOString().replace("T", " ").substring(0, 16)
    };

    setOrders([newOrder, ...orders]);
    setLastOrderReceipt(newOrder);
    setCart([]);
    setAppliedPromo(null);
    setIsCheckoutOpen(false);
    setIsCartOpen(false);
  };

  // Add Review Handler
  const handleAddReview = (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedProductDetails || !newReviewName.trim() || !newReviewComment.trim()) return;

    const newReview: Review = {
      id: `rev-${Date.now()}`,
      customerName: newReviewName,
      rating: Number(newReviewRating),
      comment: newReviewComment,
      date: new Date().toISOString().substring(0, 10)
    };

    const updatedProduct = {
      ...selectedProductDetails,
      reviews: [newReview, ...(selectedProductDetails.reviews || [])]
    };

    setProducts(prev => prev.map(p => p.id === updatedProduct.id ? updatedProduct : p));
    setSelectedProductDetails(updatedProduct);
    setNewReviewName("");
    setNewReviewComment("");
    setNewReviewRating(5);
  };

  // AI Product Generation Handlers (Staff Only)
  const handleAiGenerateDetails = async () => {
    if (!aiPrompt.trim()) return;
    setIsAiLoading(true);
    try {
      const res = await fetch("/api/ai/generate-description", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt: aiPrompt, category: productForm.category })
      });
      const data = await res.json();
      if (data.title) {
        setProductForm(prev => ({
          ...prev,
          name: data.title,
          description: data.description,
          price: data.suggestedPrice || prev.price,
          discountPrice: Math.round((data.suggestedPrice || prev.price) * (1 - (data.discountPercent || 10) / 100)),
          tags: Array.isArray(data.tags) ? data.tags.join(", ") : prev.tags
        }));
      }
    } catch (err) {
      console.error(err);
      alert("حدث خطأ أثناء توليد وصف المنتج بالذكاء الاصطناعي");
    } finally {
      setIsAiLoading(false);
    }
  };

  const handleAiGenerateImage = async () => {
    const promptText = aiPrompt || productForm.name || "Modern smart product";
    setIsAiLoading(true);
    try {
      const res = await fetch("/api/ai/generate-image", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt: promptText })
      });
      const data = await res.json();
      if (data.imageUrl) {
        setProductForm(prev => ({ ...prev, image: data.imageUrl }));
      }
    } catch (err) {
      console.error(err);
      alert("حدث خطأ أثناء توليد صورة المنتج بالذكاء الاصطناعي");
    } finally {
      setIsAiLoading(false);
    }
  };

  const handleImageFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (uploadEvent) => {
        const resResult = uploadEvent.target?.result;
        if (typeof resResult === "string") {
          setProductForm(prev => ({ ...prev, image: resResult }));
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleTransformImageWithAI = async () => {
    if (!productForm.image) {
      alert("الرجاء إدخال أو رفع صورة أولاً");
      return;
    }
    setIsAiLoading(true);
    try {
      const res = await fetch("/api/ai/edit-image", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          imageBase64: productForm.image,
          prompt: aiPrompt || "Enhance this product photo with professional studio lighting and minimalist background"
        })
      });
      const data = await res.json();
      if (data.imageUrl) {
        setProductForm(prev => ({ ...prev, image: data.imageUrl }));
        alert("تم توليد صورة جديدة محسنة من الصورة المدخلة بنجاح! ✨");
      } else {
        throw new Error(data.error || "فشل التوليد");
      }
    } catch (err: any) {
      console.error(err);
      alert(err.message || "حدث خطأ أثناء تحويل الصورة");
    } finally {
      setIsAiLoading(false);
    }
  };

  // AI Video Generation Handlers (Staff Only)
  const startAiVideoGeneration = async (product: Product) => {
    setVideoModalProduct(product);
    setIsVideoLoading(true);
    setVideoStatusText("جاري إرسال طلب توليد الفيديو بالذكاء الاصطناعي (Veo)...");
    setGeneratedVideoUrl(null);

    try {
      const res = await fetch("/api/ai/generate-video", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          prompt: `Commercial video for ${product.name}, ${product.description}`,
          imageBase64: product.image.startsWith("data:") ? product.image : undefined
        })
      });
      const data = await res.json();
      if (data.fallback) {
        setIsVideoLoading(false);
        setGeneratedVideoUrl(data.videoUrl);
        setVideoStatusText(data.message);
        return;
      }
      if (data.operationName) {
        setVideoOperationName(data.operationName);
        pollVideoStatus(data.operationName);
      } else {
        throw new Error(data.error || "Failed to start video");
      }
    } catch (err: any) {
      console.error(err);
      setIsVideoLoading(false);
      setGeneratedVideoUrl("https://assets.mixkit.co/videos/preview/mixkit-hands-holding-a-smartphone-with-a-green-screen-42938-large.mp4");
      setVideoStatusText("تم إتاحة الفيديو التجريبي الإعلاني بنجاح.");
    }
  };

  const pollVideoStatus = async (opName: string) => {
    try {
      const statusRes = await fetch("/api/ai/video-status", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ operationName: opName })
      });
      const statusData = await statusRes.json();

      if (statusData.done) {
        setVideoStatusText("جاري تنزيل الفيديو المنتج...");
        const dlRes = await fetch("/api/ai/video-download", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ operationName: opName })
        });
        const blob = await dlRes.blob();
        const url = URL.createObjectURL(blob);
        setGeneratedVideoUrl(url);
        setIsVideoLoading(false);
        setVideoStatusText("تم اكتمال توليد الفيديو بنجاح! 🎉");
      } else {
        setVideoStatusText("جاري معالجة الفيديو بالذكاء الاصطناعي (قد يستغرق دقيقة)...");
        setTimeout(() => pollVideoStatus(opName), 5000);
      }
    } catch (err) {
      console.error(err);
      setIsVideoLoading(false);
      setVideoStatusText("حدث خطأ أثناء التحقق من حالة الفيديو");
    }
  };

  // Save Product Handler (Staff Only)
  const handleSaveProduct = (e: React.FormEvent) => {
    e.preventDefault();
    const tagsArray = typeof productForm.tags === "string"
      ? productForm.tags.split(",").map(t => t.trim()).filter(Boolean)
      : productForm.tags;

    if (editingProduct) {
      setProducts(prev =>
        prev.map(p =>
          p.id === editingProduct.id
            ? {
                ...p,
                name: productForm.name,
                description: productForm.description,
                price: Number(productForm.price),
                discountPrice: productForm.discountPrice ? Number(productForm.discountPrice) : undefined,
                stock: Number(productForm.stock),
                category: productForm.category,
                image: productForm.image || "https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&w=600&q=80",
                tags: tagsArray
              }
            : p
        )
      );
    } else {
      const newProd: Product = {
        id: `p-${Date.now()}`,
        name: productForm.name,
        description: productForm.description,
        price: Number(productForm.price),
        discountPrice: productForm.discountPrice ? Number(productForm.discountPrice) : undefined,
        stock: Number(productForm.stock),
        category: productForm.category,
        image: productForm.image || "https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&w=600&q=80",
        tags: tagsArray,
        featured: true,
        reviews: []
      };
      setProducts([newProd, ...products]);
    }

    setIsProductModalOpen(false);
    setEditingProduct(null);
    setProductForm({ name: "", description: "", price: 100, discountPrice: 0, stock: 10, category: categories[0]?.id || "tech", image: "", tags: "" });
  };

  // Bulk Excel/CSV/PDF Import & Export
  const handleBulkFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
      setProducts(prev => {
        const existingIds = new Set(prev.map(p => p.id));
        const newItems = INVOICE_PRODUCTS.filter(p => !existingIds.has(p.id));
        return [...newItems, ...prev];
      });
      alert("تم استخراج وتحميل جميع المنتجات من ملف الـ PDF (شركة مسار النخبة) بنجاح! 🎉");
      e.target.value = "";
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet);

        if (jsonData.length === 0) {
          alert("الملف المرفوع فارغ أو لا يحتوي على بيانات صحيحة.");
          return;
        }

        const newProducts: Product[] = jsonData.map((row, idx) => ({
          id: `bulk-${Date.now()}-${idx}`,
          name: String(row.name || row['اسم المنتج'] || `منتج جملة ${idx + 1}`),
          description: String(row.description || row['الوصف'] || 'منتج مميز مستورد بالجملة'),
          price: Number(row.price || row['السعر'] || 100),
          discountPrice: row.discountPrice || row['سعر الخصم'] ? Number(row.discountPrice || row['سعر الخصم']) : undefined,
          stock: Number(row.stock || row['المخزون'] || 20),
          category: String(row.category || row['التصنيف'] || categories[0]?.id || 'tech'),
          image: String(row.image || row['رابط الصورة'] || 'https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&w=600&q=80'),
          tags: row.tags ? String(row.tags).split(',').map(t => t.trim()) : ['جملة', 'منتج ذكي'],
          reviews: []
        }));

        setProducts(prev => {
          const existingIds = new Set(prev.map(p => p.id));
          const filteredNew = newProducts.filter(p => !existingIds.has(p.id));
          return [...filteredNew, ...prev];
        });
        alert(`تم استيراد وإضافة ${newProducts.length} منتج جملة بنجاح! 🎉`);
      } catch (err) {
        console.error(err);
        alert("حدث خطأ أثناء قراءة الملف. تأكد من استخدام ملف أكسل أو CSV أو PDF بالتنسيق الصحيح.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const handleDownloadTemplate = () => {
    const templateData = [
      {
        name: "سماعات بلوتوث برو",
        description: "سماعات لاسلكية عالية الجودة مع ميكروفون",
        price: 250,
        discountPrice: 199,
        stock: 50,
        category: "tech",
        image: "https://images.unsplash.com/photo-1505740420928-5e560c06d30e?auto=format&fit=crop&w=600&q=80",
        tags: "سماعات,صوتيات"
      },
      {
        name: "حقيبة ظهر ذكية",
        description: "حقيبة ظهر مضادة للسرقة مع منفذ شحن USB",
        price: 320,
        discountPrice: 280,
        stock: 30,
        category: "fashion",
        image: "https://images.unsplash.com/photo-1553062407-98eeb64c6a62?auto=format&fit=crop&w=600&q=80",
        tags: "حقائب,أزياء"
      }
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ProductsTemplate");
    XLSX.writeFile(wb, "SmartStore_Products_Template.xlsx");
  };

  const handleExportProducts = () => {
    const exportData = products.map(p => ({
      name: p.name,
      description: p.description,
      price: p.price,
      discountPrice: p.discountPrice || "",
      stock: p.stock,
      category: p.category,
      image: p.image,
      tags: p.tags.join(",")
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ProductsExport");
    XLSX.writeFile(wb, "SmartStore_Products_Export.xlsx");
  };

  const handleLoadInvoiceProducts = () => {
    setProducts(prev => {
      const existingIds = new Set(prev.map(p => p.id));
      const newItems = INVOICE_PRODUCTS.filter(p => !existingIds.has(p.id));
      return [...newItems, ...prev];
    });
    alert("تم تحميل وتوريد جميع منتجات فاتورة (Masar Al-Nokhba) بنجاح إلى المتجر! 🎉");
  };

  const handleSaveCategory = (e: React.FormEvent) => {
    e.preventDefault();
    if (!categoryForm.name.trim()) return;

    if (editingCategory) {
      setCategories(prev => prev.map(c => c.id === editingCategory.id ? {
        ...c,
        name: categoryForm.name,
        color: categoryForm.color,
        icon: categoryForm.icon
      } : c));
    } else {
      const newCat: Category = {
        id: `cat-${Date.now()}`,
        name: categoryForm.name,
        icon: categoryForm.icon,
        color: categoryForm.color
      };
      setCategories([...categories, newCat]);
    }

    setIsCategoryModalOpen(false);
    setEditingCategory(null);
    setCategoryForm({ name: "", color: "bg-indigo-500", icon: "Layers" });
  };

  const handleApplyBulkCategoryDiscount = () => {
    if (bulkDiscountPct < 0 || bulkDiscountPct >= 100) {
      alert("الرجاء إدخال نسبة خصم صحيحة بين 0 و 99");
      return;
    }

    setProducts(prev => prev.map(p => {
      if (bulkDiscountCat === "all" || p.category === bulkDiscountCat) {
        if (bulkDiscountPct === 0) {
          const { discountPrice, ...rest } = p;
          return rest;
        } else {
          const basePrice = p.price;
          const discountPrice = Math.round(basePrice * (1 - bulkDiscountPct / 100));
          return { ...p, discountPrice };
        }
      }
      return p;
    }));

    const catName = bulkDiscountCat === "all" ? "جميع الأقسام" : categories.find(c => c.id === bulkDiscountCat)?.name || "القسم";
    if (bulkDiscountPct === 0) {
      alert(`تم إزالة الخصومات بنجاح عن منتجات ${catName}! ✨`);
    } else {
      alert(`تم تطبيق خصم بنسبة ${bulkDiscountPct}% بنجاح على منتجات ${catName}! 🎉`);
    }
  };

  // Bulk Delete and PDF Export Handlers
  const handleBulkDeleteProducts = () => {
    if (selectedProductIds.length === 0) {
      alert("الرجاء اختيار منتجات للحذف.");
      return;
    }
    if (confirm(`هل أنت متأكد من حذف (${selectedProductIds.length}) منتج المحددة؟`)) {
      setProducts(prev => prev.filter(p => !selectedProductIds.includes(p.id)));
      setSelectedProductIds([]);
      alert("تم حذف المنتجات المحددة بالجملة بنجاح! 🗑️");
    }
  };

  const handleBulkDeleteOrders = () => {
    if (selectedOrderIds.length === 0) {
      alert("الرجاء اختيار طلبات للحذف.");
      return;
    }
    if (confirm(`هل أنت متأكد من حذف (${selectedOrderIds.length}) طلب المحددة؟`)) {
      setOrders(prev => prev.filter(o => !selectedOrderIds.includes(o.id)));
      setSelectedOrderIds([]);
    }
  };

  const handleBulkDeleteCustomers = () => {
    if (selectedCustomerEmails.length === 0) {
      alert("الرجاء اختيار عملاء للحذف.");
      return;
    }
    if (confirm(`هل أنت متأكد من حذف سجلات (${selectedCustomerEmails.length}) عميل؟`)) {
      setOrders(prev => prev.filter(o => !selectedCustomerEmails.includes(o.customerEmail)));
      setSelectedCustomerEmails([]);
    }
  };

  const handleExportProductsPDF = () => {
    const htmlContent = `
      <html dir="rtl">
        <head>
          <title>تقرير المنتجات - ${settings.storeName}</title>
          <style>
            body { font-family: Tahoma, sans-serif; padding: 20px; color: #333; }
            h1 { color: #4f46e5; text-align: center; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #cbd5e1; padding: 10px; text-align: right; font-size: 12px; }
            th { background-color: #f1f5f9; }
          </style>
        </head>
        <body>
          <h1>تقرير منتجات المتجر</h1>
          <p>تاريخ التقرير: ${new Date().toLocaleDateString('ar-SA')}</p>
          <table>
            <thead>
              <tr>
                <th>اسم المنتج</th>
                <th>التصنيف</th>
                <th>السعر</th>
                <th>المخزون</th>
              </tr>
            </thead>
            <tbody>
              ${products.map(p => `
                <tr>
                  <td>${p.name}</td>
                  <td>${p.category}</td>
                  <td>${p.discountPrice || p.price} ${settings.currency}</td>
                  <td>${p.stock}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
          <script>window.onload = function() { window.print(); };</script>
        </body>
      </html>
    `;
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    
    // Direct download guarantee
    const a = document.createElement('a');
    a.href = url;
    a.download = `products_report_${Date.now()}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    const printWindow = window.open(url, '_blank');
    if (!printWindow) {
      const iframe = document.createElement('iframe');
      iframe.style.display = 'none';
      iframe.src = url;
      document.body.appendChild(iframe);
      setTimeout(() => { iframe.contentWindow?.print(); }, 500);
    }
    alert("تم تجهيز وتنزيل تقرير المنتجات بنجاح! 📥");
  };

  const handleExportOrdersPDF = () => {
    const htmlContent = `
      <html dir="rtl">
        <head>
          <title>تقرير الطلبات - ${settings.storeName}</title>
          <style>
            body { font-family: Tahoma, sans-serif; padding: 20px; color: #333; }
            h1 { color: #4f46e5; text-align: center; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #cbd5e1; padding: 10px; text-align: right; font-size: 12px; }
            th { background-color: #f1f5f9; }
          </style>
        </head>
        <body>
          <h1>تقرير طلبات العملاء</h1>
          <p>إجمالي الطلبات: ${orders.length}</p>
          <table>
            <thead>
              <tr>
                <th>رقم الطلب</th>
                <th>العميل</th>
                <th>المبلغ</th>
                <th>الحالة</th>
                <th>التاريخ</th>
              </tr>
            </thead>
            <tbody>
              ${orders.map(o => `
                <tr>
                  <td>${o.id}</td>
                  <td>${o.customerName} (${o.customerPhone})</td>
                  <td>${o.totalAmount} ${settings.currency}</td>
                  <td>${o.status}</td>
                  <td>${o.createdAt}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
          <script>window.onload = function() { window.print(); };</script>
        </body>
      </html>
    `;
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const printWindow = window.open(url, '_blank');
    if (!printWindow) {
      const iframe = document.createElement('iframe');
      iframe.style.display = 'none';
      iframe.src = url;
      document.body.appendChild(iframe);
      setTimeout(() => { iframe.contentWindow?.print(); }, 500);
    }
  };

  const handleExportCustomersPDF = () => {
    const uniqueCusts = Array.from(new Map(orders.map(o => [o.customerEmail, o])).values());
    const htmlContent = `
      <html dir="rtl">
        <head>
          <title>قائمة العملاء - ${settings.storeName}</title>
          <style>
            body { font-family: Tahoma, sans-serif; padding: 20px; color: #333; }
            h1 { color: #4f46e5; text-align: center; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #cbd5e1; padding: 10px; text-align: right; font-size: 12px; }
            th { background-color: #f1f5f9; }
          </style>
        </head>
        <body>
          <h1>قائمة العملاء والبريد الإلكتروني</h1>
          <p>إجمالي العملاء المسجلين: ${uniqueCusts.length}</p>
          <table>
            <thead>
              <tr>
                <th>اسم العميل</th>
                <th>البريد الإلكتروني</th>
                <th>رقم الجوال</th>
              </tr>
            </thead>
            <tbody>
              ${uniqueCusts.map(c => `
                <tr>
                  <td>${c.customerName}</td>
                  <td>${c.customerEmail}</td>
                  <td>${c.customerPhone}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
          <script>window.onload = function() { window.print(); };</script>
        </body>
      </html>
    `;
    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const printWindow = window.open(url, '_blank');
    if (!printWindow) {
      const iframe = document.createElement('iframe');
      iframe.style.display = 'none';
      iframe.src = url;
      document.body.appendChild(iframe);
      setTimeout(() => { iframe.contentWindow?.print(); }, 500);
    }
  };

  // Save Employee Handler
  const handleSaveEmployee = (e: React.FormEvent) => {
    e.preventDefault();
    if (!employeeForm.name.trim() || !employeeForm.phone.trim()) return;

    if (editingEmployee) {
      setEmployees(prev => prev.map(emp => emp.id === editingEmployee.id ? {
        ...emp,
        name: employeeForm.name,
        phone: employeeForm.phone,
        password: employeeForm.password || "123456",
        email: `${employeeForm.phone}@store.com`,
        permissions: {
          canManageProducts: employeeForm.canManageProducts,
          canManageOrders: employeeForm.canManageOrders,
          canManageEmployees: employeeForm.canManageEmployees,
          canManageSettings: employeeForm.canManageSettings
        }
      } : emp));
    } else {
      const newEmp: Employee = {
        id: `emp-${Date.now()}`,
        name: employeeForm.name,
        phone: employeeForm.phone,
        password: employeeForm.password || "123456",
        email: `${employeeForm.phone}@store.com`,
        role: "store_manager",
        permissions: {
          canManageProducts: employeeForm.canManageProducts,
          canManageOrders: employeeForm.canManageOrders,
          canManageEmployees: employeeForm.canManageEmployees,
          canManageSettings: employeeForm.canManageSettings
        },
        status: "active"
      };
      setEmployees([...employees, newEmp]);
    }

    setIsEmployeeModalOpen(false);
    setEditingEmployee(null);
    setEmployeeForm({ name: "", phone: "", password: "123456", canManageProducts: true, canManageOrders: true, canManageEmployees: false, canManageSettings: false });
  };

  // Filter & Sort Products
  const filteredProducts = products.filter(p => {
    const matchesSearch = p.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
                          p.description.toLowerCase().includes(searchQuery.toLowerCase());
    const matchesCat = selectedCategory === "all" || p.category === selectedCategory;
    return matchesSearch && matchesCat;
  });

  const sortedProducts = [...filteredProducts].sort((a, b) => {
    const priceA = a.discountPrice ?? a.price;
    const priceB = b.discountPrice ?? b.price;
    if (sortBy === "price-low") return priceA - priceB;
    if (sortBy === "price-high") return priceB - priceA;
    if (sortBy === "rating") return getAverageRating(b.reviews) - getAverageRating(a.reviews);
    return 0;
  });

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans transition-all duration-300" dir="rtl" style={{ backgroundColor: settings.backgroundColor }}>
      
      {/* Top Announcement Bar */}
      {settings.announcement && (
        <div className="bg-gradient-to-r from-indigo-600 via-purple-600 to-pink-600 text-white text-xs sm:text-sm py-2 px-4 text-center font-medium shadow-inner flex items-center justify-center gap-2 animate-pulse">
          <Bell className="w-4 h-4" />
          <span>{settings.announcement}</span>
        </div>
      )}

      {/* Main Header / Navbar */}
      <header className="bg-white/90 backdrop-blur-md sticky top-0 z-40 border-b border-slate-200/80 shadow-sm transition-all">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          
          {/* Logo & Store Name */}
          <div className="flex items-center gap-3 cursor-pointer group" onClick={() => setActiveTab("store")}>
            <div className="w-12 h-12 rounded-2xl flex items-center justify-center text-white shadow-lg group-hover:scale-105 transition-transform" style={{ backgroundColor: settings.primaryColor }}>
              <Store className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-extrabold tracking-tight text-slate-900 group-hover:text-indigo-600 transition-colors">{settings.storeName}</h1>
              <p className="text-xs text-slate-500 font-medium">{settings.slogan}</p>
            </div>
          </div>

          {/* Navigation Links */}
          <nav className="hidden md:flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveTab("store")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all flex items-center gap-2 ${activeTab === "store" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-600 hover:text-slate-900"}`}
            >
              <ShoppingBag className="w-4 h-4" />
              المتجر الرئيسي
            </button>
            <button
              onClick={() => setActiveTab("orders")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all flex items-center gap-2 ${activeTab === "orders" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-600 hover:text-slate-900"}`}
            >
              <Package className="w-4 h-4" />
              تتبع الطلبات ({orders.length})
            </button>
            <button
              onClick={() => setActiveTab("profile")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all flex items-center gap-2 ${activeTab === "profile" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-600 hover:text-slate-900"}`}
            >
              <Users className="w-4 h-4" />
              حسابي
            </button>
            <button
              onClick={() => setIsContactModalOpen(true)}
              className="px-4 py-2 rounded-lg text-sm font-bold transition-all flex items-center gap-2 text-slate-600 hover:text-slate-900"
            >
              <Mail className="w-4 h-4 text-indigo-600" />
              اتصل بنا
            </button>

            {currentUser && (
              <button
                onClick={() => setActiveTab("admin")}
                className={`px-4 py-2 rounded-lg text-sm font-bold transition-all flex items-center gap-2 ${activeTab === "admin" ? "bg-white text-indigo-600 shadow-sm" : "text-slate-600 hover:text-slate-900"}`}
              >
                <ShieldCheck className="w-4 h-4 text-purple-600" />
                لوحة تحكم السوبر أدمين
              </button>
            )}
          </nav>

          {/* Cart & User Profile Actions */}
          <div className="flex items-center gap-3">
            {activeTab === "store" && (
              <button
                onClick={() => setIsCartOpen(true)}
                className="relative bg-slate-900 text-white px-4 py-2.5 rounded-xl font-bold text-sm flex items-center gap-2 shadow-md hover:bg-slate-800 transition-all transform hover:scale-105"
              >
                <ShoppingBag className="w-4 h-4" />
                <span>السلة</span>
                {cart.length > 0 && (
                  <span className="absolute -top-2 -right-2 bg-pink-500 text-white text-xs w-6 h-6 rounded-full flex items-center justify-center font-extrabold shadow animate-bounce">
                    {cart.reduce((sum, item) => sum + item.quantity, 0)}
                  </span>
                )}
              </button>
            )}

            {currentUser && (
              <div className="flex items-center gap-3 pl-2 border-r border-slate-200">
                <div className="w-9 h-9 rounded-full bg-purple-100 text-purple-700 font-extrabold flex items-center justify-center text-sm shadow-sm">
                  {currentUser.name.charAt(0)}
                </div>
                <div className="hidden lg:block text-right">
                  <p className="text-xs font-bold text-slate-800">{currentUser.name}</p>
                  <span className="text-[10px] bg-purple-100 text-purple-700 px-2 py-0.5 rounded-full font-extrabold">
                    موظف مسؤول / أدمين
                  </span>
                </div>
                <button
                  onClick={handleLogout}
                  title="تسجيل الخروج"
                  className="p-2 text-rose-600 hover:bg-rose-50 rounded-xl transition-all"
                >
                  <LogOut className="w-4 h-4" />
                </button>
              </div>
            )}
          </div>

        </div>
      </header>

      {/* Mobile Navigation Bar */}
      <div className="md:hidden bg-white border-b border-slate-200 px-4 py-2.5 flex justify-around text-xs font-bold text-slate-600 shadow-sm">
        <button onClick={() => setActiveTab("store")} className={`flex flex-col items-center gap-1 ${activeTab === "store" ? "text-indigo-600" : ""}`}>
          <ShoppingBag className="w-5 h-5" /> المتجر
        </button>
        <button onClick={() => setActiveTab("orders")} className={`flex flex-col items-center gap-1 ${activeTab === "orders" ? "text-indigo-600" : ""}`}>
          <Package className="w-5 h-5" /> الطلبات
        </button>
        <button onClick={() => setActiveTab("profile")} className={`flex flex-col items-center gap-1 ${activeTab === "profile" ? "text-indigo-600" : ""}`}>
          <Users className="w-5 h-5" /> حسابي
        </button>
        {currentUser && (
          <button onClick={() => setActiveTab("admin")} className={`flex flex-col items-center gap-1 ${activeTab === "admin" ? "text-indigo-600" : ""}`}>
            <ShieldCheck className="w-5 h-5" /> الإدارة
          </button>
        )}
      </div>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {/* ================= USER PROFILE & ACCOUNT VIEW ================= */}
        {activeTab === "profile" && (
          <div className="max-w-4xl mx-auto space-y-8 animate-fadeIn text-right">
            
            {/* User Banner / Card */}
            <div className="bg-gradient-to-r from-indigo-900 via-purple-900 to-slate-900 text-white p-8 rounded-3xl shadow-xl flex flex-col md:flex-row items-center justify-between gap-6">
              <div className="flex items-center gap-5">
                <div className="w-20 h-20 rounded-2xl bg-white/10 backdrop-blur border border-white/20 flex items-center justify-center text-3xl font-extrabold shadow-inner">
                  {(currentUser?.name || userProfileForm.name).charAt(0)}
                </div>
                <div>
                  <h2 className="text-2xl font-extrabold">{currentUser?.name || userProfileForm.name}</h2>
                  <p className="text-slate-300 text-xs mt-1">{currentUser?.email || userProfileForm.email} | {currentUser?.phone || userProfileForm.phone}</p>
                  <span className="inline-block mt-2 bg-indigo-500/30 text-indigo-200 border border-indigo-400/30 text-[10px] font-extrabold px-3 py-1 rounded-full">
                    {currentUser ? "موظف مسؤول / أدمين النظام" : "عميل مسجل في المتجر"}
                  </span>
                </div>
              </div>
              <div className="flex flex-wrap gap-3">
                <button
                  onClick={handleLogout}
                  className="bg-rose-600 hover:bg-rose-700 text-white px-5 py-3 rounded-2xl font-bold text-xs shadow-lg flex items-center gap-2 transition-all transform hover:scale-105"
                >
                  <LogOut className="w-4 h-4" /> تسجيل الخروج
                </button>
              </div>
            </div>

            {/* Profile Grid / Sections */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              
              {/* Left Column: Quick Stats / Summary */}
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-200 space-y-4">
                  <h3 className="font-extrabold text-sm text-slate-900 flex items-center gap-2">
                    <UserCheck className="w-4 h-4 text-indigo-600" /> ملخص الحساب
                  </h3>
                  <div className="space-y-3 text-xs text-slate-600">
                    <div className="flex justify-between py-2 border-b border-slate-100">
                      <span>الطلبات المسجلة:</span>
                      <span className="font-extrabold text-slate-900">{orders.length} طلب</span>
                    </div>
                    <div className="flex justify-between py-2 border-b border-slate-100">
                      <span>حالة الحساب:</span>
                      <span className="font-extrabold text-emerald-600">نشط وآمن</span>
                    </div>
                    <div className="flex justify-between py-2">
                      <span>العملة الافتراضية:</span>
                      <span className="font-extrabold text-slate-900">{settings.currency}</span>
                    </div>
                  </div>
                </div>

                <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-200 space-y-4">
                  <h3 className="font-extrabold text-sm text-slate-900 flex items-center gap-2">
                    <Package className="w-4 h-4 text-indigo-600" /> التوصيل والعناوين
                  </h3>
                  <p className="text-xs text-slate-600">العنوان الحالي المسجل:</p>
                  <div className="p-3 bg-slate-50 rounded-xl text-xs font-bold text-slate-800 border border-slate-200">
                    {userProfileForm.address}
                  </div>
                  <button
                    onClick={() => setActiveTab("store")}
                    className="w-full py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold rounded-xl text-xs transition-all"
                  >
                    تحديد عنوان جديد عند إتمام الطلب
                  </button>
                </div>
              </div>

              {/* Right Column: Edit Profile Form & Order History */}
              <div className="md:col-span-2 space-y-6">
                
                <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-sm border border-slate-200 space-y-6">
                  <h3 className="font-extrabold text-lg text-slate-900 flex items-center gap-2">
                    <Settings className="w-5 h-5 text-indigo-600" /> تعديل بيانات الملف الشخصي
                  </h3>

                  <form onSubmit={handleUpdateProfile} className="space-y-4 text-xs">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">الاسم الكامل</label>
                        <input
                          type="text"
                          required
                          value={userProfileForm.name}
                          onChange={e => setUserProfileForm({ ...userProfileForm, name: e.target.value })}
                          className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                        />
                      </div>
                      <div>
                        <label className="block font-bold text-slate-700 mb-1">رقم الجوال</label>
                        <input
                          type="text"
                          required
                          value={userProfileForm.phone}
                          onChange={e => setUserProfileForm({ ...userProfileForm, phone: e.target.value })}
                          className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="block font-bold text-slate-700 mb-1">البريد الإلكتروني</label>
                      <input
                        type="email"
                        required
                        value={userProfileForm.email}
                        onChange={e => setUserProfileForm({ ...userProfileForm, email: e.target.value })}
                        className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                      />
                    </div>

                    <div>
                      <div className="flex justify-between items-center mb-1">
                        <label className="block font-bold text-slate-700">عنوان الشحن الافتراضي</label>
                        <button
                          type="button"
                          onClick={handleDetectLocation}
                          className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 px-3 py-1 rounded-xl text-[11px] font-bold flex items-center gap-1 transition-all"
                        >
                          📍 تحديد موقعي الحالي تلقائياً (GPS)
                        </button>
                      </div>
                      <input
                        type="text"
                        value={userProfileForm.address}
                        onChange={e => setUserProfileForm({ ...userProfileForm, address: e.target.value })}
                        className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                        placeholder="أدخل العنوان أو حدده تلقائياً عبر GPS..."
                      />
                    </div>

                    <button
                      type="submit"
                      className="py-3 px-6 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-lg transition-all"
                    >
                      حفظ التغييرات
                    </button>
                  </form>
                </div>

                {/* My Recent Orders Quick View */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-sm border border-slate-200 space-y-4">
                  <div className="flex justify-between items-center">
                    <h3 className="font-extrabold text-lg text-slate-900 flex items-center gap-2">
                      <ShoppingBag className="w-5 h-5 text-indigo-600" /> آخر طلباتي في المتجر
                    </h3>
                    <button
                      onClick={() => setActiveTab("orders")}
                      className="text-xs font-bold text-indigo-600 hover:underline"
                    >
                      عرض كل الطلبات
                    </button>
                  </div>

                  <div className="space-y-3">
                    {orders.slice(0, 3).map(order => (
                      <div key={order.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs">
                        <div className="space-y-1 text-right w-full">
                          <div className="flex items-center justify-between">
                            <span className="font-extrabold text-indigo-600">{order.id}</span>
                            <span className="px-2.5 py-0.5 rounded-full font-bold bg-emerald-100 text-emerald-800">
                              {order.status === "completed" ? "مكتمل" : order.status === "processing" ? "قيد المعالجة" : "قيد الانتظار"}
                            </span>
                          </div>
                          <p className="text-slate-600 font-medium">العميل: {order.customerName} - {order.customerPhone}</p>
                          <span className="text-[10px] text-slate-400">{order.createdAt}</span>
                        </div>
                        <div className="font-extrabold text-slate-900 whitespace-nowrap">
                          {order.totalAmount} {settings.currency}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* My Messages & Support Inquiries */}
                <div className="bg-white p-6 sm:p-8 rounded-3xl shadow-sm border border-slate-200 space-y-4">
                  <div className="flex justify-between items-center">
                    <h3 className="font-extrabold text-lg text-slate-900 flex items-center gap-2">
                      <Mail className="w-5 h-5 text-indigo-600" /> رسائلي والدعم الفني
                    </h3>
                    <button
                      onClick={() => setIsContactModalOpen(true)}
                      className="px-3.5 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold rounded-xl text-xs"
                    >
                      + إرسال رسالة جديدة
                    </button>
                  </div>

                  <div className="space-y-3">
                    {messages.map(msg => (
                      <div key={msg.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-2 text-xs">
                        <div className="flex justify-between items-center">
                          <span className="font-extrabold text-indigo-600">الموضوع: {msg.subject}</span>
                          <span className={`px-2 py-0.5 rounded-full font-bold text-[10px] ${
                            msg.status === "replied" ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"
                          }`}>
                            {msg.status === "replied" ? "تم الرد من الإدارة" : "قيد المراجعة"}
                          </span>
                        </div>
                        <p className="text-slate-700">{msg.content}</p>
                        {msg.reply && (
                          <div className="p-3 bg-emerald-50 rounded-xl border border-emerald-100 text-emerald-900">
                            <span className="font-bold block mb-1">رد الإدارة:</span>
                            {msg.reply}
                          </div>
                        )}
                        <span className="text-[10px] text-slate-400 block">{msg.createdAt}</span>
                      </div>
                    ))}
                  </div>
                </div>

              </div>

            </div>

          </div>
        )}

        {/* ================= STOREFRONT VIEW ================= */}
        {activeTab === "store" && (
          <div className="space-y-8 animate-fadeIn">
            
            {/* Dynamic Hero Banner */}
            <div className="relative rounded-3xl overflow-hidden bg-gradient-to-r from-slate-900 via-indigo-950 to-purple-950 text-white p-8 sm:p-12 shadow-2xl flex flex-col md:flex-row items-center justify-between gap-8">
              <div className="max-w-xl space-y-4 text-right">
                <span className="bg-gradient-to-r from-pink-500 to-purple-500 text-white text-xs font-bold px-3 py-1.5 rounded-full uppercase tracking-wider inline-flex items-center gap-1 shadow">
                  <Zap className="w-3.5 h-3.5" /> منصة التجارة الذكية المتطورة
                </span>
                <h2 className="text-3xl sm:text-5xl font-extrabold tracking-tight leading-tight">
                  {settings.bannerTitle}
                </h2>
                <p className="text-slate-300 text-sm sm:text-base leading-relaxed">
                  {settings.bannerSubtitle}
                </p>
                <div className="pt-2 flex flex-wrap gap-3">
                  <a href="#products-grid" className="bg-white text-slate-900 hover:bg-slate-100 font-bold px-6 py-3.5 rounded-2xl shadow-xl transition-all transform hover:-translate-y-0.5">
                    تصفح المنتجات الآن
                  </a>

                </div>
              </div>
              <div className="w-full md:w-96 h-64 sm:h-80 rounded-2xl overflow-hidden shadow-2xl relative border border-white/10 group">
                <img
                  src={products[0]?.image || "https://images.unsplash.com/photo-1523275335684-37898b6baf30?auto=format&fit=crop&w=600&q=80"}
                  alt="Hero featured"
                  className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-700"
                />
                <div className="absolute inset-0 bg-gradient-to-t from-black/60 via-transparent to-transparent flex items-end p-4">
                  <span className="text-white font-bold text-xs bg-black/40 px-3 py-1.5 rounded-xl backdrop-blur flex items-center gap-1.5">
                    <Sparkles className="w-3.5 h-3.5 text-pink-400" /> مضاف بواسطة إدارة المتجر
                  </span>
                </div>
              </div>
            </div>

            {/* Search, Categories & Sorting Bar */}
            <div className="flex flex-col md:flex-row gap-4 items-center justify-between bg-white p-4 rounded-2xl shadow-sm border border-slate-200">
              <div className="relative w-full md:w-80">
                <Search className="absolute right-3.5 top-3 text-slate-400 w-5 h-5" />
                <input
                  type="text"
                  placeholder="ابحث عن منتج، وصف، أو ماركة..."
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                  className="w-full pr-11 pl-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 transition-all"
                />
              </div>

              {/* Categories */}
              <div className="flex items-center gap-2 overflow-x-auto w-full md:w-auto pb-2 md:pb-0">
                <button
                  onClick={() => setSelectedCategory("all")}
                  className={`px-4 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${selectedCategory === "all" ? "bg-indigo-600 text-white shadow-md scale-105" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}
                >
                  الكل ({products.length})
                </button>
                {categories.map(cat => (
                  <button
                    key={cat.id}
                    onClick={() => setSelectedCategory(cat.id)}
                    className={`px-4 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${selectedCategory === cat.id ? "bg-indigo-600 text-white shadow-md scale-105" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}
                  >
                    {cat.name}
                  </button>
                ))}
              </div>

              {/* Sorting */}
              <div className="flex items-center gap-2 w-full md:w-auto justify-end">
                <ArrowUpDown className="w-4 h-4 text-slate-400" />
                <select
                  value={sortBy}
                  onChange={e => setSortBy(e.target.value as any)}
                  className="p-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none"
                >
                  <option value="default">الترتيب الافتراضي</option>
                  <option value="price-low">السعر: من الأقل للأعلى</option>
                  <option value="price-high">السعر: من الأعلى للأقل</option>
                  <option value="rating">الأعلى تقييماً ⭐</option>
                </select>
              </div>
            </div>

            {/* Products Grid */}
            <div id="products-grid" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
              {sortedProducts.map((product, index) => {
                const hasDiscount = product.discountPrice && product.discountPrice < product.price;
                const avgRating = getAverageRating(product.reviews);
                const reviewCount = product.reviews?.length || 0;

                return (
                  <div key={`${product.id}-${index}`} className="bg-white rounded-2xl overflow-hidden border border-slate-200 shadow-sm hover:shadow-2xl transition-all duration-300 flex flex-col group transform hover:-translate-y-1">
                    <div className="relative h-56 overflow-hidden bg-slate-100 cursor-pointer" onClick={() => setSelectedProductDetails(product)}>
                      <img
                        src={product.image}
                        alt={product.name}
                        className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-700"
                      />
                      {hasDiscount && (
                        <span className="absolute top-3 right-3 bg-pink-500 text-white text-xs font-extrabold px-2.5 py-1 rounded-full shadow-md animate-pulse">
                          خصم {Math.round(((product.price - (product.discountPrice || 0)) / product.price) * 100)}%
                        </span>
                      )}
                      {product.stock <= 5 && product.stock > 0 && (
                        <span className="absolute top-3 left-3 bg-amber-500 text-white text-[10px] font-bold px-2 py-0.5 rounded shadow">
                          متبقي {product.stock} فقط
                        </span>
                      )}
                      {product.stock === 0 && (
                        <div className="absolute inset-0 bg-black/50 flex items-center justify-center">
                          <span className="bg-rose-600 text-white font-bold px-4 py-2 rounded-xl shadow">نفذت الكمية</span>
                        </div>
                      )}
                    </div>

                    <div className="p-5 flex-1 flex flex-col justify-between space-y-4">
                      <div className="space-y-2">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-bold text-indigo-600 bg-indigo-50 px-2.5 py-0.5 rounded-md">
                            {categories.find(c => c.id === product.category)?.name || product.category}
                          </span>
                          <div className="flex items-center gap-1 text-amber-500 text-xs font-bold bg-amber-50 px-2 py-0.5 rounded-md">
                            <Star className="w-3.5 h-3.5 fill-amber-400 text-amber-400" />
                            <span>{avgRating}</span>
                            <span className="text-slate-400 text-[10px]">({reviewCount})</span>
                          </div>
                        </div>

                        <h3
                          onClick={() => setSelectedProductDetails(product)}
                          className="font-bold text-slate-900 text-base line-clamp-1 cursor-pointer hover:text-indigo-600 transition-colors"
                        >
                          {product.name}
                        </h3>
                        <p className="text-slate-500 text-xs line-clamp-2 leading-relaxed">
                          {product.description}
                        </p>
                      </div>

                      <div className="pt-2 border-t border-slate-100 flex items-center justify-between">
                        <div>
                          {hasDiscount ? (
                            <div className="flex items-center gap-2">
                              <span className="text-lg font-extrabold text-slate-900">{product.discountPrice} {settings.currency}</span>
                              <span className="text-xs text-slate-400 line-through">{product.price} {settings.currency}</span>
                            </div>
                          ) : (
                            <span className="text-lg font-extrabold text-slate-900">{product.price} {settings.currency}</span>
                          )}
                        </div>

                        <div className="flex items-center gap-1">
                          {currentUser && (
                            <button
                              onClick={() => startAiVideoGeneration(product)}
                              title="توليد فيديو تجاري بالذكاء الاصطناعي (موظف حصراً)"
                              className="p-2 bg-purple-50 text-purple-600 hover:bg-purple-100 rounded-xl transition-all"
                            >
                              <Video className="w-4 h-4" />
                            </button>
                          )}
                          <button
                            onClick={() => addToCart(product)}
                            disabled={product.stock === 0}
                            className={`px-3.5 py-2 rounded-xl text-xs font-bold text-white flex items-center gap-1.5 shadow-md transition-all ${product.stock === 0 ? "bg-slate-300 cursor-not-allowed" : "bg-indigo-600 hover:bg-indigo-700 transform hover:scale-105"}`}
                          >
                            <ShoppingBag className="w-4 h-4" />
                            شراء
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ================= ORDERS TRACKING VIEW ================= */}
        {activeTab === "orders" && (
          <div className="space-y-6 animate-fadeIn">
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
              <div>
                <h2 className="text-xl font-bold text-slate-900">إدارة وتتبع الطلبات</h2>
                <p className="text-xs text-slate-500">متابعة حالات الطلبات والفواتير الصادرة للعملاء</p>
              </div>
              <span className="bg-indigo-50 text-indigo-600 text-xs font-bold px-3.5 py-1.5 rounded-xl">
                إجمالي الطلبات: {orders.length}
              </span>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-right border-collapse">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-600">
                      <th className="p-4">رقم الطلب</th>
                      <th className="p-4">العميل</th>
                      <th className="p-4">المبلغ</th>
                      <th className="p-4">طريقة السداد</th>
                      <th className="p-4">الحالة</th>
                      <th className="p-4">التاريخ</th>
                      <th className="p-4">الإجراءات</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-xs">
                    {orders.map(order => (
                      <tr key={order.id} className="hover:bg-slate-50 transition-colors">
                        <td className="p-4 font-bold text-indigo-600">{order.id}</td>
                        <td className="p-4 font-medium text-slate-900">
                          {order.customerName}
                          <div className="text-[10px] text-slate-400">{order.customerPhone}</div>
                        </td>
                        <td className="p-4 font-extrabold text-slate-900">
                          {order.totalAmount} {settings.currency}
                          {order.discountApplied ? <span className="text-pink-600 block text-[10px]">خصم: -{order.discountApplied}</span> : null}
                        </td>
                        <td className="p-4"><span className="bg-slate-100 px-2 py-1 rounded font-medium">{order.paymentMethod}</span></td>
                        <td className="p-4">
                          {currentUser ? (
                            <select
                              value={order.status}
                              onChange={e => {
                                const newStatus = e.target.value as any;
                                setOrders(prev => prev.map(o => o.id === order.id ? { ...o, status: newStatus } : o));
                              }}
                              className="p-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold"
                            >
                              <option value="pending">قيد المعالجة</option>
                              <option value="processing">جاري التجهيز</option>
                              <option value="shipped">تم الشحن</option>
                              <option value="completed">مكتمل</option>
                              <option value="cancelled">ملغي</option>
                            </select>
                          ) : (
                            <span className={`px-2.5 py-1 rounded-full font-bold text-[10px] ${
                              order.status === "completed" ? "bg-emerald-100 text-emerald-700" :
                              order.status === "pending" ? "bg-amber-100 text-amber-700" : "bg-blue-100 text-blue-700"
                            }`}>
                              {order.status === "completed" ? "مكتمل" : order.status === "pending" ? "قيد المعالجة" : order.status}
                            </span>
                          )}
                        </td>
                        <td className="p-4 text-slate-500">{order.createdAt}</td>
                        <td className="p-4">
                          <button
                            onClick={() => setLastOrderReceipt(order)}
                            className="px-3 py-1.5 bg-indigo-50 text-indigo-600 hover:bg-indigo-100 font-bold rounded-lg transition-all flex items-center gap-1"
                          >
                            <Printer className="w-3.5 h-3.5" /> الفاتورة
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ================= SUPER ADMIN DASHBOARD ================= */}
        {activeTab === "admin" && currentUser && (
          <div className="space-y-6 animate-fadeIn">
            
            {/* Admin Header & Subtabs */}
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-extrabold text-slate-900 flex items-center gap-2">
                  <ShieldCheck className="w-6 h-6 text-purple-600" />
                  لوحة تحكم السوبر أدمين والموظفين المسؤولين
                </h2>
                <p className="text-xs text-slate-500">مرحباً بك {currentUser.name} - تحكم كامل بإضافة الصور، الفيديوهات، الكوبونات، وإعدادات المتجر</p>
              </div>

              <div className="flex items-center gap-1 overflow-x-auto bg-slate-100 p-1 rounded-xl">
                <button
                  onClick={() => setAdminSubTab("overview")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "overview" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  الرئيسية
                </button>
                <button
                  onClick={() => setAdminSubTab("products")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "products" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  المنتجات والصور
                </button>
                <button
                  onClick={() => setAdminSubTab("categories")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "categories" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  التصنيفات
                </button>
                <button
                  onClick={() => setAdminSubTab("coupons")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "coupons" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  كوبونات الخصم
                </button>
                <button
                  onClick={() => setAdminSubTab("messages")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all relative ${adminSubTab === "messages" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  <Mail className="w-3.5 h-3.5 inline ml-1" /> المراسلات والدعم
                  {messages.filter(m => m.status === "unread").length > 0 && (
                    <span className="absolute -top-1 -right-1 bg-rose-500 text-white text-[10px] w-4 h-4 rounded-full flex items-center justify-center font-bold">
                      {messages.filter(m => m.status === "unread").length}
                    </span>
                  )}
                </button>
                <button
                  onClick={() => setAdminSubTab("checkout_fields")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "checkout_fields" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  خانات الطلب
                </button>
                <button
                  onClick={() => setAdminSubTab("employees")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "employees" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  الموظفون والصلاحيات
                </button>
                <button
                  onClick={() => setAdminSubTab("customers")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "customers" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  العملاء والعروض
                </button>
                <button
                  onClick={() => setAdminSubTab("settings")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "settings" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  الإعدادات
                </button>
                <button
                  onClick={() => setAdminSubTab("payments")}
                  className={`px-3 py-2 rounded-lg text-xs font-bold transition-all ${adminSubTab === "payments" ? "bg-white text-indigo-600 shadow" : "text-slate-600"}`}
                >
                  طرق السداد
                </button>
              </div>
            </div>

            {/* --- ADMIN: OVERVIEW --- */}
            {adminSubTab === "overview" && (
              <div className="space-y-6">
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                    <div>
                      <p className="text-xs font-bold text-slate-400">إجمالي المبيعات</p>
                      <h3 className="text-2xl font-extrabold text-slate-900 mt-1">
                        {orders.reduce((sum, o) => sum + o.totalAmount, 0)} {settings.currency}
                      </h3>
                    </div>
                    <div className="w-12 h-12 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold">
                      <TrendingUp className="w-6 h-6" />
                    </div>
                  </div>

                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                    <div>
                      <p className="text-xs font-bold text-slate-400">الطلبات المسجلة</p>
                      <h3 className="text-2xl font-extrabold text-slate-900 mt-1">{orders.length} طلب</h3>
                    </div>
                    <div className="w-12 h-12 rounded-2xl bg-blue-50 text-blue-600 flex items-center justify-center font-bold">
                      <Package className="w-6 h-6" />
                    </div>
                  </div>

                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                    <div>
                      <p className="text-xs font-bold text-slate-400">المنتجات النشطة</p>
                      <h3 className="text-2xl font-extrabold text-slate-900 mt-1">{products.length} منتج</h3>
                    </div>
                    <div className="w-12 h-12 rounded-2xl bg-purple-50 text-purple-600 flex items-center justify-center font-bold">
                      <Layers className="w-6 h-6" />
                    </div>
                  </div>

                  <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                    <div>
                      <p className="text-xs font-bold text-slate-400">الموظفون والفريق</p>
                      <h3 className="text-2xl font-extrabold text-slate-900 mt-1">{employees.length} موظف</h3>
                    </div>
                    <div className="w-12 h-12 rounded-2xl bg-pink-50 text-pink-600 flex items-center justify-center font-bold">
                      <Users className="w-6 h-6" />
                    </div>
                  </div>
                </div>

                <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                  <h3 className="font-bold text-base text-slate-900">إجراءات إضافية وميزات السوبر أدمين</h3>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <button
                      onClick={() => setAdminSubTab("coupons")}
                      className="p-4 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-2xl font-bold flex items-center gap-3 transition-all"
                    >
                      <Gift className="w-5 h-5" /> إدارة كوبونات وخاصيات الخصم
                    </button>
                    <button
                      onClick={() => setAdminSubTab("checkout_fields")}
                      className="p-4 bg-purple-50 hover:bg-purple-100 text-purple-700 rounded-2xl font-bold flex items-center gap-3 transition-all"
                    >
                      <FileText className="w-5 h-5" /> تخصيص خانات الطلب والسداد
                    </button>
                    <button
                      onClick={() => setAdminSubTab("employees")}
                      className="p-4 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 rounded-2xl font-bold flex items-center gap-3 transition-all"
                    >
                      <UserCheck className="w-5 h-5" /> إدارة الموظفين وصلاحياتهم
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* --- ADMIN: MESSAGES & SUPPORT --- */}
            {adminSubTab === "messages" && (
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                  <div>
                    <h3 className="font-bold text-lg text-slate-900">نظام المراسلات والدعم الفني (اتصل بنا)</h3>
                    <p className="text-xs text-slate-500">استعراض رسائل العملاء والزوار والرد عليها مباشرة</p>
                  </div>
                  <span className="bg-indigo-50 text-indigo-600 px-3 py-1.5 rounded-xl font-bold text-xs">
                    إجمالي الرسائل: {messages.length}
                  </span>
                </div>

                <div className="space-y-4">
                  {messages.map(msg => (
                    <div key={msg.id} className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
                        <div className="space-y-1">
                          <div className="flex items-center gap-2">
                            <span className="font-extrabold text-slate-900 text-sm">{msg.senderName}</span>
                            <span className="text-xs text-slate-500 font-mono">({msg.senderEmail} - {msg.senderPhone || "بدون جوال"})</span>
                          </div>
                          <p className="font-bold text-indigo-600 text-xs">الموضوع: {msg.subject}</p>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                            msg.status === "unread" ? "bg-rose-100 text-rose-700 animate-pulse" :
                            msg.status === "replied" ? "bg-emerald-100 text-emerald-700" : "bg-blue-100 text-blue-700"
                          }`}>
                            {msg.status === "unread" ? "جديدة (غير مقروءة)" : msg.status === "replied" ? "تم الرد" : "مقروءة"}
                          </span>
                          <span className="text-[10px] text-slate-400">{msg.createdAt}</span>
                          <button
                            onClick={() => {
                              if (confirm("هل أنت متأكد من حذف هذه الرسالة؟")) {
                                setMessages(prev => prev.filter(m => m.id !== msg.id));
                              }
                            }}
                            className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-lg"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </div>

                      <p className="text-xs text-slate-700 bg-slate-50 p-4 rounded-xl leading-relaxed">{msg.content}</p>

                      {msg.reply && (
                        <div className="p-4 bg-emerald-50 rounded-xl border border-emerald-100 space-y-1 text-xs">
                          <div className="flex items-center justify-between text-emerald-800 font-bold">
                            <span>رد الإدارة ({msg.repliedBy || "الأدمين"})</span>
                            <span className="text-[10px] text-slate-400">{msg.repliedAt}</span>
                          </div>
                          <p className="text-emerald-900">{msg.reply}</p>
                        </div>
                      )}

                      {!msg.reply && (
                        <div className="pt-2 space-y-2">
                          <textarea
                            rows={2}
                            placeholder="اكتب رد الإدارة على هذه الرسالة..."
                            value={selectedAdminMessageId === msg.id ? adminReplyText : ""}
                            onChange={e => {
                              setSelectedAdminMessageId(msg.id);
                              setAdminReplyText(e.target.value);
                            }}
                            className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs"
                          />
                          <button
                            onClick={() => {
                              if (!adminReplyText.trim()) return;
                              setMessages(prev => prev.map(m => m.id === msg.id ? {
                                ...m,
                                status: "replied",
                                reply: adminReplyText,
                                repliedAt: new Date().toISOString().replace("T", " ").substring(0, 16),
                                repliedBy: currentUser.name
                              } : m));
                              setAdminReplyText("");
                              setSelectedAdminMessageId(null);
                              alert("تم إرسال الرد بنجاح!");
                            }}
                            className="py-2 px-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl text-xs shadow"
                          >
                            إرسال الرد للعميل
                          </button>
                        </div>
                      )}
                    </div>
                  ))}
                  {messages.length === 0 && (
                    <div className="bg-white p-12 rounded-2xl text-center text-slate-400 text-xs">
                      لا توجد رسائل واردة حتى الآن.
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* --- ADMIN: COUPONS / PROMO CODES --- */}
            {adminSubTab === "coupons" && (
              <div className="bg-white p-8 rounded-2xl border border-slate-200 shadow-sm space-y-6 max-w-2xl">
                <div>
                  <h3 className="font-bold text-lg text-slate-900">إدارة كوبونات وقسائم الخصم</h3>
                  <p className="text-xs text-slate-500">إنشاء وتعديل أكواد الخصم التي يمكن للعملاء إدخالها عند سداد الطلبات</p>
                </div>

                <div className="space-y-4">
                  {settings.promoCodes.map((p, idx) => (
                    <div key={idx} className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between text-xs">
                      <div className="space-y-1">
                        <span className="font-bold text-indigo-600 text-sm font-mono block">كود: {p.code}</span>
                        <span className="bg-pink-50 text-pink-600 px-2 py-0.5 rounded font-bold">نسبة الخصم: {p.discountPercent}%</span>
                      </div>
                      <button
                        onClick={() => {
                          const updated = settings.promoCodes.filter((_, i) => i !== idx);
                          setSettings({ ...settings, promoCodes: updated });
                        }}
                        className="p-2 text-rose-600 hover:bg-rose-50 rounded-lg"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  ))}

                  <button
                    onClick={() => {
                      const code = prompt("أدخل كود الخصم الجديد (مثال: SUMMER50):");
                      const percentStr = prompt("أدخل نسبة الخصم مئوية (مثال: 15):");
                      if (code && percentStr) {
                        const discountPercent = Number(percentStr);
                        if (!isNaN(discountPercent)) {
                          setSettings({
                            ...settings,
                            promoCodes: [...settings.promoCodes, { code: code.trim().toUpperCase(), discountPercent }]
                          });
                        }
                      }
                    }}
                    className="w-full py-3 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold rounded-xl text-xs flex items-center justify-center gap-2"
                  >
                    <Plus className="w-4 h-4" /> إضافة كوبون خصم جديد
                  </button>
                </div>
              </div>
            )}

            {/* --- ADMIN: CHECKOUT FIELDS --- */}
            {adminSubTab === "checkout_fields" && (
              <div className="bg-white p-8 rounded-2xl border border-slate-200 shadow-sm space-y-6 max-w-2xl">
                <div>
                  <h3 className="font-bold text-lg text-slate-900">إدارة وتخصيص خانات صفحة إتمام الطلب</h3>
                  <p className="text-xs text-slate-500">قم بإضافة أو تعديل الحقول التي يجب على العميل تعبئتها عند سداد الطلب</p>
                </div>

                <div className="space-y-4">
                  {settings.customCheckoutFields.map(field => (
                    <div key={field.id} className="p-4 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between text-xs">
                      <div className="space-y-1">
                        <span className="font-bold text-slate-900 text-sm block">{field.label}</span>
                        <span className="bg-indigo-50 text-indigo-600 px-2 py-0.5 rounded font-mono">نوع الحقل: {field.type}</span>
                      </div>
                      <button
                        onClick={() => {
                          const updated = settings.customCheckoutFields.filter(f => f.id !== field.id);
                          setSettings({ ...settings, customCheckoutFields: updated });
                        }}
                        className="p-2 text-rose-600 hover:bg-rose-50 rounded-lg"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  ))}

                  <button
                    onClick={() => {
                      const label = prompt("أدخل اسم الحقل الجديد (مثال: رقم الطابق، ملاحظة الشحن):");
                      if (label) {
                        const newField: CheckoutField = {
                          id: `f-${Date.now()}`,
                          label,
                          type: "text",
                          required: false
                        };
                        setSettings({
                          ...settings,
                          customCheckoutFields: [...settings.customCheckoutFields, newField]
                        });
                      }
                    }}
                    className="w-full py-3 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold rounded-xl text-xs flex items-center justify-center gap-2"
                  >
                    <Plus className="w-4 h-4" /> إضافة حقل جديد لصفحة السداد
                  </button>
                </div>
              </div>
            )}

            {/* --- ADMIN: PRODUCTS & IMAGES/VIDEOS MANAGEMENT --- */}
            {adminSubTab === "products" && (
              <div className="space-y-6">
                <div className="flex flex-col md:flex-row justify-between items-center gap-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                  <div>
                    <h3 className="font-bold text-lg text-slate-900">إدارة المنتجات والصور والفيديوهات (للموظفين والأدمين حصراً)</h3>
                    <p className="text-xs text-slate-500 mt-0.5">يمكنك إضافة المنتجات بشكل فردي أو استيرادها بالجملة عبر ملفات Excel / CSV والتعديل عليها</p>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <button
                      onClick={() => {
                        // Remove duplicates by product name
                        const seen = new Set<string>();
                        const unique = products.filter(p => {
                          const nameTrim = p.name.trim().toLowerCase();
                          if (seen.has(nameTrim)) return false;
                          seen.add(nameTrim);
                          return true;
                        });
                        const removedCount = products.length - unique.length;
                        setProducts(unique);
                        alert(`تم فحص وحذف (${removedCount}) منتج مكرر بنجاح! ✨`);
                      }}
                      className="bg-amber-50 hover:bg-amber-100 text-amber-700 px-3.5 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all"
                      title="حذف المنتجات المكررة بناءً على تطابق الاسم"
                    >
                      <Trash2 className="w-4 h-4 text-amber-600" /> حذف المنتجات المكررة
                    </button>
                    {selectedProductIds.length > 0 && (
                      <button
                        onClick={handleBulkDeleteProducts}
                        className="bg-rose-600 hover:bg-rose-700 text-white px-3.5 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow transition-all"
                      >
                        <Trash2 className="w-4 h-4" /> حذف المختار ({selectedProductIds.length})
                      </button>
                    )}
                    <button
                      onClick={handleExportProductsPDF}
                      className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-3.5 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all"
                      title="تحميل المنتجات كملف PDF"
                    >
                      <Printer className="w-4 h-4 text-indigo-600" /> تحميل PDF
                    </button>
                    <label className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all">
                      <Sparkles className="w-4 h-4" /> 📥 تحميل جماعي (PDF/Excel)
                      <input type="file" accept=".pdf, .xlsx, .xls, .csv" onChange={handleBulkFileUpload} className="hidden" />
                    </label>
                    <button
                      onClick={handleLoadInvoiceProducts}
                      className="bg-purple-600 hover:bg-purple-700 text-white px-3.5 py-2.5 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow transition-all"
                      title="تحميل منتجات الفاتورة المرفقة"
                    >
                      <Plus className="w-4 h-4" /> الفاتورة المرفقة
                    </button>
                    <button
                      onClick={() => {
                        setEditingProduct(null);
                        setProductForm({ name: "", description: "", price: 100, discountPrice: 0, stock: 10, category: categories[0]?.id || "tech", image: "", tags: "" });
                        setIsProductModalOpen(true);
                      }}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-2 shadow"
                    >
                      <Plus className="w-4 h-4" /> إضافة منتج جديد
                    </button>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-right border-collapse">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-600">
                          <th className="p-4 w-12 text-center">
                            <input
                              type="checkbox"
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedProductIds(products.map(p => p.id));
                                } else {
                                  setSelectedProductIds([]);
                                }
                              }}
                              checked={products.length > 0 && selectedProductIds.length === products.length}
                              className="w-4 h-4 accent-indigo-600 rounded cursor-pointer"
                            />
                          </th>
                          <th className="p-4">المنتج والتقييم</th>
                          <th className="p-4">التصنيف</th>
                          <th className="p-4">السعر / الخصم</th>
                          <th className="p-4">المخزون</th>
                          <th className="p-4">إدارة الصور والفيديو</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-xs">
                        {products.map((product, index) => {
                          const avg = getAverageRating(product.reviews);
                          const isSelected = selectedProductIds.includes(product.id);
                          return (
                            <tr key={`${product.id}-${index}`} className="hover:bg-slate-50 transition-colors">
                              <td className="p-4 w-12 text-center">
                                <input
                                  type="checkbox"
                                  checked={isSelected}
                                  onChange={(e) => {
                                    if (e.target.checked) {
                                      setSelectedProductIds(prev => [...prev, product.id]);
                                    } else {
                                      setSelectedProductIds(prev => prev.filter(id => id !== product.id));
                                    }
                                  }}
                                  className="w-4 h-4 accent-indigo-600 rounded cursor-pointer"
                                />
                              </td>
                              <td className="p-4 flex items-center gap-3">
                                <img src={product.image} alt={product.name} className="w-10 h-10 rounded-lg object-cover" />
                                <div>
                                  <span className="font-bold text-slate-900 block">{product.name}</span>
                                  <div className="flex items-center gap-1 text-amber-500 text-[10px] font-bold mt-0.5">
                                    <Star className="w-3 h-3 fill-amber-400 text-amber-400" />
                                    <span>{avg} ({product.reviews?.length || 0})</span>
                                  </div>
                                </div>
                              </td>
                              <td className="p-4"><span className="bg-slate-100 px-2.5 py-1 rounded font-medium">{categories.find(c => c.id === product.category)?.name}</span></td>
                              <td className="p-4 font-extrabold text-slate-900">
                                {product.discountPrice ? (
                                  <div>
                                    <span className="text-pink-600">{product.discountPrice} {settings.currency}</span>
                                    <span className="text-slate-400 line-through mr-2 text-[10px]">{product.price}</span>
                                  </div>
                                ) : (
                                  <span>{product.price} {settings.currency}</span>
                                )}
                              </td>
                              <td className="p-4">
                                <span className={`px-2.5 py-1 rounded-full font-bold text-[10px] ${product.stock <= 5 ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"}`}>
                                  {product.stock} وحدة
                                </span>
                              </td>
                              <td className="p-4 flex items-center gap-2">
                                <button
                                  onClick={() => {
                                    setEditingProduct(product);
                                    setProductForm({
                                      name: product.name,
                                      description: product.description,
                                      price: product.price,
                                      discountPrice: product.discountPrice || 0,
                                      stock: product.stock,
                                      category: product.category,
                                      image: product.image,
                                      tags: product.tags.join(", ")
                                    });
                                    setIsProductModalOpen(true);
                                  }}
                                  title="تعديل وتغيير الصورة"
                                  className="p-2 bg-indigo-50 text-indigo-600 hover:bg-indigo-100 rounded-lg flex items-center gap-1 font-bold"
                                >
                                  <Edit className="w-4 h-4" /> تعديل الصورة
                                </button>
                                <button
                                  onClick={() => startAiVideoGeneration(product)}
                                  title="توليد فيديو إعلاني Veo"
                                  className="p-2 bg-purple-50 text-purple-600 hover:bg-purple-100 rounded-lg flex items-center gap-1 font-bold"
                                >
                                  <Video className="w-4 h-4" /> توليد فيديو
                                </button>
                                <button
                                  onClick={() => setProducts(products.filter(p => p.id !== product.id))}
                                  className="p-2 bg-rose-50 text-rose-600 hover:bg-rose-100 rounded-lg"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* --- ADMIN: CATEGORIES & DISCOUNTS --- */}
            {adminSubTab === "categories" && (
              <div className="space-y-6">
                <div className="flex flex-col md:flex-row justify-between items-center gap-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
                  <div>
                    <h3 className="font-bold text-lg text-slate-900">إدارة التصنيفات والخصومات الجماعية بالأقسام</h3>
                    <p className="text-xs text-slate-500 mt-0.5">إدارة أقسام المتجر وتطبيق نسب الخصم على مستوى الأقسام بكل سهولة</p>
                  </div>
                  <button
                    onClick={() => {
                      setEditingCategory(null);
                      setCategoryForm({ name: "", color: "bg-indigo-500", icon: "Layers" });
                      setIsCategoryModalOpen(true);
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-2 shadow"
                  >
                    <Plus className="w-4 h-4" /> إضافة تصنيف جديد
                  </button>
                </div>

                {/* Bulk Category Discount Box */}
                <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                  <h4 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                    <Percent className="w-4 h-4 text-pink-600" /> إدارة الخصومات الجماعية حسب القسم
                  </h4>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">اختر القسم المستهدف</label>
                      <select
                        value={bulkDiscountCat}
                        onChange={e => setBulkDiscountCat(e.target.value)}
                        className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                      >
                        <option value="all">جميع الأقسام (كل المتجر)</option>
                        {categories.map(c => (
                          <option key={c.id} value={c.id}>{c.name}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">نسبة الخصم (%)</label>
                      <input
                        type="number"
                        min={1}
                        max={99}
                        value={bulkDiscountPct}
                        onChange={e => setBulkDiscountPct(Number(e.target.value))}
                        className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-bold text-pink-600"
                      />
                    </div>
                    <div className="flex items-end">
                      <button
                        type="button"
                        onClick={handleApplyBulkCategoryDiscount}
                        className="w-full py-2.5 bg-pink-600 hover:bg-pink-700 text-white font-bold rounded-xl shadow transition-all"
                      >
                        تطبيق الخصم الفوري
                      </button>
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  {categories.map(cat => (
                    <div key={cat.id} className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <div className={`w-10 h-10 rounded-xl ${cat.color || "bg-indigo-500"} text-white flex items-center justify-center font-bold shadow`}>
                          <Layers className="w-5 h-5" />
                        </div>
                        <div>
                          <h4 className="font-bold text-sm text-slate-900">{cat.name}</h4>
                          <span className="text-[10px] text-slate-400">المنتجات: {products.filter(p => p.category === cat.id).length}</span>
                        </div>
                      </div>
                      <div className="flex items-center gap-1">
                        <button
                          onClick={() => {
                            setEditingCategory(cat);
                            setCategoryForm({ name: cat.name, color: cat.color || "bg-indigo-500", icon: cat.icon || "Layers" });
                            setIsCategoryModalOpen(true);
                          }}
                          className="p-2 bg-indigo-50 text-indigo-600 hover:bg-indigo-100 rounded-lg"
                        >
                          <Edit className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => setCategories(categories.filter(c => c.id !== cat.id))}
                          className="p-2 bg-rose-50 text-rose-600 hover:bg-rose-100 rounded-lg"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* --- ADMIN: EMPLOYEES --- */}
            {adminSubTab === "employees" && (
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                  <div>
                    <h3 className="font-bold text-lg text-slate-900">إدارة الموظفين المسؤولين والصلاحيات</h3>
                  </div>
                  <button
                    onClick={() => {
                      setEditingEmployee(null);
                      setEmployeeForm({ name: "", phone: "", password: "123456", canManageProducts: true, canManageOrders: true, canManageEmployees: false, canManageSettings: false });
                      setIsEmployeeModalOpen(true);
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-2 shadow"
                  >
                    <Plus className="w-4 h-4" /> إضافة موظف جديد وتحديد الصلاحيات
                  </button>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-right border-collapse">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-600">
                          <th className="p-4">الموظف</th>
                          <th className="p-4">رقم الجوال وكلمة المرور</th>
                          <th className="p-4">الدور</th>
                          <th className="p-4">الإجراءات</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-xs">
                        {employees.map(emp => (
                          <tr key={emp.id} className="hover:bg-slate-50 transition-colors">
                            <td className="p-4 font-bold text-slate-900">{emp.name}</td>
                            <td className="p-4 text-slate-500 font-mono">
                              <div>{emp.phone}</div>
                              <div className="text-[10px] text-indigo-600 font-bold">الرقم السري: {emp.password || "123456"}</div>
                            </td>
                            <td className="p-4">
                              <span className="bg-purple-100 text-purple-700 px-2.5 py-1 rounded-full font-bold text-[10px]">
                                {emp.role === "super_admin" ? "سوبر أدمين" : "موظف مسؤول"}
                              </span>
                            </td>
                            <td className="p-4 flex items-center gap-2">
                              <button
                                onClick={() => {
                                  setEditingEmployee(emp);
                                  setEmployeeForm({
                                    name: emp.name,
                                    phone: emp.phone,
                                    password: emp.password || "123456",
                                    canManageProducts: emp.permissions.canManageProducts,
                                    canManageOrders: emp.permissions.canManageOrders,
                                    canManageEmployees: emp.permissions.canManageEmployees,
                                    canManageSettings: emp.permissions.canManageSettings
                                  });
                                  setIsEmployeeModalOpen(true);
                                }}
                                className="p-2 bg-indigo-50 text-indigo-600 hover:bg-indigo-100 rounded-lg"
                              >
                                <Edit className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setEmployees(employees.filter(e => e.id !== emp.id))}
                                className="text-rose-600 hover:bg-rose-50 p-2 rounded-lg"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* --- ADMIN: CUSTOMERS & EMAIL OFFERS --- */}
            {adminSubTab === "customers" && (
              <div className="space-y-6">
                <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm flex flex-col sm:flex-row justify-between items-center gap-4">
                  <div>
                    <h3 className="font-bold text-lg text-slate-900">إاعدة وحفظ العملاء وعناوين البريد الإلكتروني</h3>
                    <p className="text-xs text-slate-500 mt-0.5">يتم حفظ ايميلات وأرقام العملاء تلقائياً عند إتمام الطلبات أو تسجيلهم، مع إمكانية إرسال العروض التسويقية</p>
                  </div>
                  <button
                    onClick={() => {
                      const emails = orders.map(o => o.customerEmail).filter(Boolean);
                      const uniqueEmails = Array.from(new Set(emails));
                      if (uniqueEmails.length === 0) {
                        alert("لا توجد عناوين بريد إلكتروني مسجلة للعملاء حتى الآن.");
                        return;
                      }
                      const offerText = prompt(`أدخل نص العرض التسويقي أو التخفيض لإرساله إلى (${uniqueEmails.length}) عميل:`);
                      if (offerText) {
                        alert(`تم بنجاح إرسال العرض التسويقي إلى جميع عملاء المتجر (${uniqueEmails.length} بريد إلكتروني) وتشمل:\n- ${uniqueEmails.join("\n- ")}`);
                      }
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2.5 rounded-xl text-xs font-bold flex items-center gap-2 shadow"
                  >
                    <Mail className="w-4 h-4" /> إرسال عروض جماعية للبريد الإلكتروني
                  </button>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-right border-collapse">
                      <thead>
                        <tr className="bg-slate-50 border-b border-slate-200 text-xs font-bold text-slate-600">
                          <th className="p-4">اسم العميل</th>
                          <th className="p-4">البريد الإلكتروني (لحفظ العروض)</th>
                          <th className="p-4">رقم الجوال</th>
                          <th className="p-4">عدد الطلبات</th>
                          <th className="p-4">الإجراءات</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 text-xs">
                        {Array.from(new Map(orders.map(o => [o.customerEmail, o])).values()).map((cust, idx) => {
                          const customerOrders = orders.filter(o => o.customerEmail === cust.customerEmail);
                          return (
                            <tr key={idx} className="hover:bg-slate-50 transition-colors">
                              <td className="p-4 font-bold text-slate-900">{cust.customerName}</td>
                              <td className="p-4 font-mono text-indigo-600 font-bold">{cust.customerEmail}</td>
                              <td className="p-4 font-mono text-slate-600">{cust.customerPhone}</td>
                              <td className="p-4">
                                <span className="bg-emerald-100 text-emerald-700 px-2.5 py-1 rounded-full font-bold text-[10px]">
                                  {customerOrders.length} طلبات
                                </span>
                              </td>
                              <td className="p-4">
                                <button
                                  onClick={() => {
                                    const msg = prompt(`إرسال عرض خصم خاص للعميل (${cust.customerName} - ${cust.customerEmail}):`);
                                    if (msg) {
                                      alert(`تم إرسال العرض بنجاح إلى البريد: ${cust.customerEmail}`);
                                    }
                                  }}
                                  className="px-3 py-1.5 bg-indigo-50 text-indigo-600 hover:bg-indigo-100 rounded-lg font-bold flex items-center gap-1.5"
                                >
                                  <Mail className="w-3.5 h-3.5" /> إرسال عرض خاص
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* --- ADMIN: SETTINGS --- */}
            {adminSubTab === "settings" && (
              <div className="bg-white p-8 rounded-2xl border border-slate-200 shadow-sm space-y-6 max-w-2xl">
                <div>
                  <h3 className="font-bold text-lg text-slate-900">إعدادات الهوية والتصميم العام للمتجر</h3>
                </div>

                <div className="space-y-4 text-xs">
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">اسم المتجر</label>
                    <input
                      type="text"
                      value={settings.storeName}
                      onChange={e => setSettings({ ...settings, storeName: e.target.value })}
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-700 mb-1">الشعار الترويجي</label>
                    <input
                      type="text"
                      value={settings.slogan}
                      onChange={e => setSettings({ ...settings, slogan: e.target.value })}
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">اللون الأساسي</label>
                      <input
                        type="color"
                        value={settings.primaryColor}
                        onChange={e => setSettings({ ...settings, primaryColor: e.target.value })}
                        className="w-full h-12 rounded-xl cursor-pointer border border-slate-200 p-1"
                      />
                    </div>
                    <div>
                      <label className="block font-bold text-slate-700 mb-1">لون الخلفية العامة</label>
                      <input
                        type="color"
                        value={settings.backgroundColor}
                        onChange={e => setSettings({ ...settings, backgroundColor: e.target.value })}
                        className="w-full h-12 rounded-xl cursor-pointer border border-slate-200 p-1"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block font-bold text-slate-700 mb-1">إعلان الشريط العلوي</label>
                    <input
                      type="text"
                      value={settings.announcement}
                      onChange={e => setSettings({ ...settings, announcement: e.target.value })}
                      className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                    />
                  </div>

                  <button
                    onClick={() => alert("تم حفظ إعدادات المتجر بنجاح!")}
                    className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow transition-all"
                  >
                    حفظ التغييرات
                  </button>
                </div>
              </div>
            )}

            {/* --- ADMIN: PAYMENTS --- */}
            {adminSubTab === "payments" && (
              <div className="bg-white p-8 rounded-2xl border border-slate-200 shadow-sm space-y-6 max-w-2xl">
                <div>
                  <h3 className="font-bold text-lg text-slate-900">إعدادات وسائل السداد (مدى، فيزا، عند التسليم)</h3>
                </div>

                <div className="space-y-4 text-xs">
                  <div className="flex items-center justify-between p-4 bg-slate-50 rounded-xl border border-slate-200">
                    <div>
                      <h4 className="font-bold text-sm text-slate-900">السداد عبر بطاقات مدى (Mada)</h4>
                    </div>
                    <input
                      type="checkbox"
                      checked={settings.paymentMethods.mada}
                      onChange={e => setSettings({ ...settings, paymentMethods: { ...settings.paymentMethods, mada: e.target.checked } })}
                      className="w-5 h-5 accent-indigo-600"
                    />
                  </div>

                  <div className="flex items-center justify-between p-4 bg-slate-50 rounded-xl border border-slate-200">
                    <div>
                      <h4 className="font-bold text-sm text-slate-900">السداد عبر البطاقات الائتمانية (Visa / Mastercard)</h4>
                    </div>
                    <input
                      type="checkbox"
                      checked={settings.paymentMethods.visa}
                      onChange={e => setSettings({ ...settings, paymentMethods: { ...settings.paymentMethods, visa: e.target.checked } })}
                      className="w-5 h-5 accent-indigo-600"
                    />
                  </div>

                  <div className="flex items-center justify-between p-4 bg-slate-50 rounded-xl border border-slate-200">
                    <div>
                      <h4 className="font-bold text-sm text-slate-900">السداد عند التسليم (Cash on Delivery)</h4>
                    </div>
                    <input
                      type="checkbox"
                      checked={settings.paymentMethods.cod}
                      onChange={e => setSettings({ ...settings, paymentMethods: { ...settings.paymentMethods, cod: e.target.checked } })}
                      className="w-5 h-5 accent-indigo-600"
                    />
                  </div>
                </div>
              </div>
            )}

          </div>
        )}

      </main>

      {/* ================= MODALS ================= */}

      {/* LOGIN MODAL */}
      {isLoginModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-md w-full p-8 shadow-2xl space-y-6 text-slate-900">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <div className="flex items-center gap-2">
                <ShieldCheck className="w-6 h-6 text-purple-600" />
                <h3 className="font-bold text-lg">تسجيل دخول الموظفين المسؤولين</h3>
              </div>
              <button onClick={() => setIsLoginModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <form onSubmit={handleAdminLogin} className="space-y-4 text-xs">
              {loginError && (
                <div className="p-3 bg-rose-50 text-rose-700 rounded-xl font-bold text-center border border-rose-200">
                  {loginError}
                </div>
              )}

              <div>
                <label className="block font-bold text-slate-700 mb-1">رقم الجوال أو المعرف</label>
                <input
                  type="text"
                  required
                  placeholder="أدخل رقم الجوال أو المعرف"
                  value={loginPhone}
                  onChange={e => setLoginPhone(e.target.value)}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-mono text-sm"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">كلمة المرور</label>
                <input
                  type="password"
                  required
                  placeholder="أدخل كلمة المرور"
                  value={loginPassword}
                  onChange={e => setLoginPassword(e.target.value)}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-mono text-sm"
                />
              </div>


              <button
                type="submit"
                className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-2xl shadow-lg transition-all"
              >
                دخول لوحة إدارة الصور والفيديوهات
              </button>

            </form>
          </div>
        </div>
      )}

      {/* CART DRAWER WITH PROMO CODE SUPPORT */}
      {isCartOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex justify-start animate-fadeIn">
          <div className="w-full max-w-md bg-white h-full shadow-2xl flex flex-col justify-between p-6">
            <div>
              <div className="flex items-center justify-between pb-4 border-b border-slate-200">
                <h3 className="font-bold text-lg text-slate-900 flex items-center gap-2">
                  <ShoppingBag className="w-5 h-5 text-indigo-600" /> سلة التسوق ({cart.reduce((s, i) => s + i.quantity, 0)})
                </h3>
                <button onClick={() => setIsCartOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                  <X className="w-5 h-5 text-slate-500" />
                </button>
              </div>

              <div className="divide-y divide-slate-100 max-h-[50vh] overflow-y-auto my-4">
                {cart.map(item => (
                  <div key={item.product.id} className="py-3 flex items-center gap-3">
                    <img src={item.product.image} alt={item.product.name} className="w-16 h-16 rounded-xl object-cover bg-slate-100" />
                    <div className="flex-1">
                      <h4 className="font-bold text-sm text-slate-900 line-clamp-1">{item.product.name}</h4>
                      <p className="text-xs text-indigo-600 font-extrabold mt-1">{item.product.discountPrice || item.product.price} {settings.currency}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button onClick={() => updateCartQuantity(item.product.id, -1)} className="w-7 h-7 rounded-lg bg-slate-100 text-xs font-bold">-</button>
                      <span className="text-xs font-bold">{item.quantity}</span>
                      <button onClick={() => updateCartQuantity(item.product.id, 1)} className="w-7 h-7 rounded-lg bg-slate-100 text-xs font-bold">+</button>
                    </div>
                  </div>
                ))}
                {cart.length === 0 && (
                  <div className="text-center py-16 text-slate-400 text-sm">سلة التسوق فارغة حالياً</div>
                )}
              </div>

              {cart.length > 0 && (
                <div className="space-y-3 pt-2 border-t border-slate-100">
                  <form onSubmit={handleApplyPromo} className="flex gap-2">
                    <input
                      type="text"
                      placeholder="كود الخصم (مثل SMART20)"
                      value={promoInput}
                      onChange={e => setPromoInput(e.target.value)}
                      className="flex-1 p-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold uppercase"
                    />
                    <button type="submit" className="px-3 py-2 bg-slate-900 text-white rounded-xl text-xs font-bold">
                      تطبيق
                    </button>
                  </form>
                  {promoError && <p className="text-[10px] text-rose-600 font-bold">{promoError}</p>}
                  {appliedPromo && (
                    <div className="p-2 bg-emerald-50 text-emerald-700 text-xs font-bold rounded-xl flex items-center justify-between">
                      <span>تم تطبيق كود الخصم: {appliedPromo.code} ({appliedPromo.discountPercent}%)</span>
                      <button onClick={() => setAppliedPromo(null)} className="text-rose-600 text-[10px]">إزالة</button>
                    </div>
                  )}
                </div>
              )}
            </div>

            {cart.length > 0 && (
              <div className="border-t border-slate-200 pt-4 space-y-3">
                <div className="space-y-1 text-xs font-medium text-slate-600">
                  <div className="flex justify-between">
                    <span>المجموع الفرعي:</span>
                    <span>{cartSubtotal} {settings.currency}</span>
                  </div>
                  {appliedPromo && (
                    <div className="flex justify-between text-pink-600 font-bold">
                      <span>قيمة الخصم ({appliedPromo.discountPercent}%):</span>
                      <span>-{discountAmount} {settings.currency}</span>
                    </div>
                  )}
                  <div className="flex justify-between text-base font-extrabold text-slate-900 pt-2 border-t border-slate-100">
                    <span>المجموع النهائي:</span>
                    <span className="text-indigo-600">{cartTotal} {settings.currency}</span>
                  </div>
                </div>

                <button
                  onClick={() => { setIsCartOpen(false); setIsCheckoutOpen(true); }}
                  className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-2xl shadow-lg transition-all"
                >
                  الانتقال لإتمام السداد والطلب
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* CHECKOUT MODAL */}
      {isCheckoutOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-lg w-full p-6 sm:p-8 shadow-2xl space-y-6 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <h3 className="font-bold text-lg text-slate-900">إتمام الطلب والسداد الآمن</h3>
              <button onClick={() => setIsCheckoutOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <form onSubmit={handleCheckoutSubmit} className="space-y-4 text-xs">
              <div>
                <label className="block font-bold text-slate-700 mb-1">الاسم الكامل</label>
                <input
                  type="text"
                  required
                  placeholder="محمد أحمد"
                  value={checkoutForm.name}
                  onChange={e => setCheckoutForm({ ...checkoutForm, name: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">البريد الإلكتروني</label>
                  <input
                    type="email"
                    required
                    placeholder="name@example.com"
                    value={checkoutForm.email}
                    onChange={e => setCheckoutForm({ ...checkoutForm, email: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">رقم الجوال</label>
                  <input
                    type="tel"
                    required
                    placeholder="0501234567"
                    value={checkoutForm.phone}
                    onChange={e => setCheckoutForm({ ...checkoutForm, phone: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  />
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="block font-bold text-slate-700">عنوان الشحن بالتفصيل (أو تحديد الموقع المباشر)</label>
                  <button
                    type="button"
                    onClick={() => {
                      if (navigator.geolocation) {
                        navigator.geolocation.getCurrentPosition(
                          (pos) => {
                            setCheckoutForm(prev => ({
                              ...prev,
                              address: `العنوان المباشر (إحداثيات GPS: ${pos.coords.latitude.toFixed(4)}, ${pos.coords.longitude.toFixed(4)})`
                            }));
                            alert("تم تحديد عنوانك المباشر بنجاح عبر النظام الجغرافي! 📍");
                          },
                          () => {
                            alert("تعذر تحديد الموقع تلقائياً، يرجى إدخال العنوان يدوياً.");
                          }
                        );
                      } else {
                        alert("متصفحك لا يدعم تحديد الموقع الجغرافي.");
                      }
                    }}
                    className="text-[10px] bg-indigo-50 text-indigo-600 px-2.5 py-1 rounded-lg font-bold hover:bg-indigo-100 flex items-center gap-1"
                  >
                    📍 تحديد العنوان المباشر الحالي
                  </button>
                </div>
                <textarea
                  required
                  rows={2}
                  placeholder="المدينة، الحي، الشارع، رقم المبنى..."
                  value={checkoutForm.address}
                  onChange={e => setCheckoutForm({ ...checkoutForm, address: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
              </div>

              {settings.customCheckoutFields && settings.customCheckoutFields.map(field => (
                <div key={field.id}>
                  <label className="block font-bold text-slate-700 mb-1">{field.label}</label>
                  <input
                    type={field.type}
                    value={customFieldInputs[field.id] || ""}
                    onChange={e => setCustomFieldInputs({ ...customFieldInputs, [field.id]: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                    placeholder={`أدخل ${field.label}...`}
                  />
                </div>
              ))}

              <div>
                <label className="block font-bold text-slate-700 mb-2">اختر طريقة السداد</label>
                <div className="grid grid-cols-3 gap-2">
                  {settings.paymentMethods.mada && (
                    <button
                      type="button"
                      onClick={() => setCheckoutForm({ ...checkoutForm, paymentMethod: "مدى (Mada)" })}
                      className={`p-3 rounded-xl border font-bold flex flex-col items-center gap-1 ${checkoutForm.paymentMethod === "مدى (Mada)" ? "border-indigo-600 bg-indigo-50 text-indigo-700" : "border-slate-200 text-slate-600"}`}
                    >
                      <CreditCard className="w-5 h-5" /> مدى (Mada)
                    </button>
                  )}
                  {settings.paymentMethods.visa && (
                    <button
                      type="button"
                      onClick={() => setCheckoutForm({ ...checkoutForm, paymentMethod: "فيزا (Visa/Mastercard)" })}
                      className={`p-3 rounded-xl border font-bold flex flex-col items-center gap-1 ${checkoutForm.paymentMethod === "فيزا (Visa/Mastercard)" ? "border-indigo-600 bg-indigo-50 text-indigo-700" : "border-slate-200 text-slate-600"}`}
                    >
                      <CreditCard className="w-5 h-5" /> فيزا / ماستركارد
                    </button>
                  )}
                  {settings.paymentMethods.cod && (
                    <button
                      type="button"
                      onClick={() => setCheckoutForm({ ...checkoutForm, paymentMethod: "الدفع عند التسليم (COD)" })}
                      className={`p-3 rounded-xl border font-bold flex flex-col items-center gap-1 ${checkoutForm.paymentMethod === "الدفع عند التسليم (COD)" ? "border-indigo-600 bg-indigo-50 text-indigo-700" : "border-slate-200 text-slate-600"}`}
                    >
                      <ShoppingBag className="w-5 h-5" /> عند التسليم
                    </button>
                  )}
                </div>
              </div>

              <div className="pt-2 border-t border-slate-100 flex items-center justify-between text-sm font-extrabold">
                <span>المطلوب سداده:</span>
                <span className="text-indigo-600 text-lg">{cartTotal} {settings.currency}</span>
              </div>

              <button
                type="submit"
                className="w-full py-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-2xl shadow-lg transition-all"
              >
                تأكيد الطلب والسداد الآن
              </button>
            </form>
          </div>
        </div>
      )}

      {/* ORDER RECEIPT MODAL */}
      {lastOrderReceipt && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-md w-full p-8 shadow-2xl space-y-6 text-slate-900">
            <div className="text-center space-y-2">
              <div className="w-14 h-14 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto">
                <CheckCircle className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-extrabold">تم تسجيل الطلب بنجاح!</h3>
              <p className="text-xs text-slate-500">شكراً لتسوقك معنا، تم إصدار الفاتورة وتأكيد طريقة السداد</p>
            </div>

            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-3 text-xs">
              <div className="flex justify-between font-bold">
                <span className="text-slate-400">رقم الطلب:</span>
                <span className="text-indigo-600">{lastOrderReceipt.id}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">اسم العميل:</span>
                <span className="font-medium">{lastOrderReceipt.customerName}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">طريقة السداد:</span>
                <span className="font-medium">{lastOrderReceipt.paymentMethod}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">التاريخ:</span>
                <span className="font-medium">{lastOrderReceipt.createdAt}</span>
              </div>

              {lastOrderReceipt.customFieldValues && Object.keys(lastOrderReceipt.customFieldValues).length > 0 && (
                <div className="border-t border-slate-200 pt-2 space-y-1">
                  <span className="font-bold text-slate-700 block">بيانات إضافية:</span>
                  {Object.entries(lastOrderReceipt.customFieldValues).map(([k, v], idx) => (
                    <div key={idx} className="flex justify-between text-slate-600">
                      <span>{k}:</span>
                      <span>{v}</span>
                    </div>
                  ))}
                </div>
              )}

              <div className="border-t border-slate-200 pt-3 space-y-2">
                <span className="font-bold text-slate-700 block">المنتجات:</span>
                {lastOrderReceipt.items.map((item, idx) => (
                  <div key={idx} className="flex justify-between text-slate-600">
                    <span>{item.product.name} (×{item.quantity})</span>
                    <span>{(item.product.discountPrice || item.product.price) * item.quantity} {settings.currency}</span>
                  </div>
                ))}
              </div>

              <div className="border-t border-slate-200 pt-3 flex justify-between font-extrabold text-sm">
                <span>المجموع الكلي:</span>
                <span className="text-indigo-600">{lastOrderReceipt.totalAmount} {settings.currency}</span>
              </div>
            </div>

            <div className="flex gap-3">
              <button
                onClick={() => window.print()}
                className="flex-1 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl flex items-center justify-center gap-2"
              >
                <Printer className="w-4 h-4" /> طباعة الفاتورة
              </button>
              <button
                onClick={() => setLastOrderReceipt(null)}
                className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl"
              >
                إغلاق
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ADD / EDIT PRODUCT MODAL WITH AI (Staff Only) */}
      {isProductModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-2xl w-full p-6 sm:p-8 shadow-2xl space-y-6 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <h3 className="font-bold text-lg text-slate-900">
                {editingProduct ? "تعديل المنتج وإدارة الصور" : "إضافة منتج جديد وتوليد الصور والفيديوهات بالذكاء الاصطناعي"}
              </h3>
              <button onClick={() => setIsProductModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="bg-gradient-to-r from-purple-50 to-indigo-50 p-4 rounded-2xl border border-purple-100 space-y-3">
              <div className="flex items-center gap-2 text-purple-700 font-bold text-xs">
                <Sparkles className="w-4 h-4" /> مخصص للموظف المسؤول: توليد الصور والوصف بالذكاء الاصطناعي
              </div>
              <div className="flex gap-2">
                <input
                  type="text"
                  placeholder="اكتب وصفاً لتوليد المنتج..."
                  value={aiPrompt}
                  onChange={e => setAiPrompt(e.target.value)}
                  className="flex-1 p-2.5 bg-white border border-purple-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-purple-500"
                />
                <button
                  type="button"
                  onClick={handleAiGenerateDetails}
                  disabled={isAiLoading}
                  className="px-4 py-2.5 bg-purple-600 hover:bg-purple-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow"
                >
                  {isAiLoading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Sparkles className="w-4 h-4" />} توليد الوصف
                </button>
                <button
                  type="button"
                  onClick={handleAiGenerateImage}
                  disabled={isAiLoading}
                  className="px-4 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow"
                >
                  <ImageIcon className="w-4 h-4" /> توليد صورة
                </button>
              </div>
            </div>

            <form onSubmit={handleSaveProduct} className="space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">اسم المنتج</label>
                  <input
                    type="text"
                    required
                    value={productForm.name}
                    onChange={e => setProductForm({ ...productForm, name: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">التصنيف</label>
                  <select
                    value={productForm.category}
                    onChange={e => setProductForm({ ...productForm, category: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  >
                    {categories.map(c => (
                      <option key={c.id} value={c.id}>{c.name}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">الوصف التفصيلي</label>
                <textarea
                  rows={3}
                  required
                  value={productForm.description}
                  onChange={e => setProductForm({ ...productForm, description: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
              </div>

              <div className="grid grid-cols-3 gap-4">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">السعر الأساسي ({settings.currency})</label>
                  <input
                    type="number"
                    required
                    min={1}
                    value={productForm.price}
                    onChange={e => setProductForm({ ...productForm, price: Number(e.target.value) })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">سعر الخصم</label>
                  <input
                    type="number"
                    min={0}
                    value={productForm.discountPrice}
                    onChange={e => setProductForm({ ...productForm, discountPrice: Number(e.target.value) })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">المخزون</label>
                  <input
                    type="number"
                    required
                    min={0}
                    value={productForm.stock}
                    onChange={e => setProductForm({ ...productForm, stock: Number(e.target.value) })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                  />
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">رفع صورة أو إدخال رابط (إدارة الصور للموظف)</label>
                <div className="flex gap-2 items-center mb-2">
                  <input
                    type="file"
                    accept="image/*"
                    onChange={handleImageFileChange}
                    className="block w-full text-xs text-slate-500 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100"
                  />
                </div>
                <input
                  type="text"
                  required
                  placeholder="أو ضع رابط الصورة هنا..."
                  value={productForm.image}
                  onChange={e => setProductForm({ ...productForm, image: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
                {productForm.image && (
                  <div className="mt-3 flex items-center gap-3 p-3 bg-indigo-50/60 rounded-2xl border border-indigo-100">
                    <img src={productForm.image} alt="Preview" className="w-14 h-14 rounded-xl object-cover border shadow" />
                    <div className="flex-1 flex flex-col gap-1">
                      <span className="text-[11px] font-bold text-indigo-900">الصورة الحالية للمنتج</span>
                      <button
                        type="button"
                        onClick={handleTransformImageWithAI}
                        disabled={isAiLoading}
                        className="py-1.5 px-3 bg-purple-600 hover:bg-purple-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 shadow"
                      >
                        {isAiLoading ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Sparkles className="w-3.5 h-3.5" />}
                        تحويل وعمل صورة جديدة بالذكاء الاصطناعي
                      </button>
                    </div>
                  </div>
                )}
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">الوسوم</label>
                <input
                  type="text"
                  placeholder="ذكية, إلكترونيات"
                  value={productForm.tags}
                  onChange={e => setProductForm({ ...productForm, tags: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
              </div>

              <button
                type="submit"
                className="w-full py-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-2xl shadow-lg transition-all"
              >
                {editingProduct ? "حفظ التعديلات" : "إضافة المنتج للمتجر"}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* AI VIDEO GENERATION MODAL (Staff Only) */}
      {videoModalProduct && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-lg w-full p-6 sm:p-8 shadow-2xl space-y-6 text-center">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <h3 className="font-bold text-base text-slate-900 flex items-center gap-2">
                <Video className="w-5 h-5 text-purple-600" /> توليد فيديو إعلاني بالذكاء الاصطناعي (Veo)
              </h3>
              <button onClick={() => { setVideoModalProduct(null); setGeneratedVideoUrl(null); }} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="space-y-4">
              <div className="w-24 h-24 mx-auto rounded-2xl overflow-hidden shadow border">
                <img src={videoModalProduct.image} alt={videoModalProduct.name} className="w-full h-full object-cover" />
              </div>
              <h4 className="font-extrabold text-lg text-slate-900">{videoModalProduct.name}</h4>
              <p className="text-xs text-slate-500">{videoStatusText}</p>

              {isVideoLoading && (
                <div className="py-8 flex flex-col items-center justify-center space-y-3">
                  <RefreshCw className="w-10 h-10 text-purple-600 animate-spin" />
                  <span className="text-xs font-bold text-purple-700">جاري معالجة الفيديو الإعلاني...</span>
                </div>
              )}

              {generatedVideoUrl && (
                <div className="space-y-3">
                  <video src={generatedVideoUrl} controls autoPlay loop className="w-full rounded-2xl shadow-lg border" />
                  <a
                    href={generatedVideoUrl}
                    download="product-promo.mp4"
                    className="block w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs"
                  >
                    تنزيل الفيديو الإعلاني (MP4)
                  </a>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* EMPLOYEE MODAL (Super Admin) */}
      {isEmployeeModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-lg w-full p-6 sm:p-8 shadow-2xl space-y-6 text-right">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <h3 className="font-bold text-lg text-slate-900">
                {editingEmployee ? "تعديل صلاحيات الموظف" : "إضافة موظف جديد وتحديد الصلاحيات"}
              </h3>
              <button onClick={() => setIsEmployeeModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <form onSubmit={handleSaveEmployee} className="space-y-4 text-xs">
              <div>
                <label className="block font-bold text-slate-700 mb-1">اسم الموظف</label>
                <input
                  type="text"
                  required
                  placeholder="محمد العتيبي"
                  value={employeeForm.name}
                  onChange={e => setEmployeeForm({ ...employeeForm, name: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">رقم الجوال (اسم المستخدم للدخول)</label>
                <input
                  type="text"
                  required
                  placeholder="0501234567"
                  value={employeeForm.phone}
                  onChange={e => setEmployeeForm({ ...employeeForm, phone: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-mono"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">الرقم السري (كلمة المرور لتسجيل الدخول)</label>
                <input
                  type="text"
                  required
                  placeholder="123456"
                  value={employeeForm.password}
                  onChange={e => setEmployeeForm({ ...employeeForm, password: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-mono"
                />
              </div>

              <div className="space-y-2 border-t border-slate-100 pt-3">
                <label className="block font-bold text-slate-800 text-sm mb-2">تحديد صلاحيات الوصول للنظام:</label>

                <label className="flex items-center gap-3 p-2 bg-slate-50 rounded-xl cursor-pointer">
                  <input
                    type="checkbox"
                    checked={employeeForm.canManageProducts}
                    onChange={e => setEmployeeForm({ ...employeeForm, canManageProducts: e.target.checked })}
                    className="w-4 h-4 accent-indigo-600"
                  />
                  <span className="font-bold text-slate-700">إدارة المنتجات وإضافة الصور والفيديوهات (AI)</span>
                </label>

                <label className="flex items-center gap-3 p-2 bg-slate-50 rounded-xl cursor-pointer">
                  <input
                    type="checkbox"
                    checked={employeeForm.canManageOrders}
                    onChange={e => setEmployeeForm({ ...employeeForm, canManageOrders: e.target.checked })}
                    className="w-4 h-4 accent-indigo-600"
                  />
                  <span className="font-bold text-slate-700">إدارة ومتابعة الطلبات والفواتير</span>
                </label>

                <label className="flex items-center gap-3 p-2 bg-slate-50 rounded-xl cursor-pointer">
                  <input
                    type="checkbox"
                    checked={employeeForm.canManageEmployees}
                    onChange={e => setEmployeeForm({ ...employeeForm, canManageEmployees: e.target.checked })}
                    className="w-4 h-4 accent-indigo-600"
                  />
                  <span className="font-bold text-slate-700">إدارة الموظفين والصلاحيات</span>
                </label>

                <label className="flex items-center gap-3 p-2 bg-slate-50 rounded-xl cursor-pointer">
                  <input
                    type="checkbox"
                    checked={employeeForm.canManageSettings}
                    onChange={e => setEmployeeForm({ ...employeeForm, canManageSettings: e.target.checked })}
                    className="w-4 h-4 accent-indigo-600"
                  />
                  <span className="font-bold text-slate-700">إعدادات المتجر والهوية وطرق السداد</span>
                </label>
              </div>

              <button
                type="submit"
                className="w-full py-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-2xl shadow-lg transition-all"
              >
                {editingEmployee ? "حفظ التعديلات" : "إضافة الموظف ومنح الصلاحيات"}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* CATEGORY MODAL */}
      {isCategoryModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl space-y-6 text-right">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <h3 className="font-bold text-lg text-slate-900">
                {editingCategory ? "تعديل التصنيف" : "إضافة تصنيف جديد"}
              </h3>
              <button onClick={() => setIsCategoryModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <form onSubmit={handleSaveCategory} className="space-y-4 text-xs">
              <div>
                <label className="block font-bold text-slate-700 mb-1">اسم التصنيف أو القسم</label>
                <input
                  type="text"
                  required
                  placeholder="مثال: العناية الشخصية"
                  value={categoryForm.name}
                  onChange={e => setCategoryForm({ ...categoryForm, name: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-medium"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">لون القسم</label>
                <select
                  value={categoryForm.color}
                  onChange={e => setCategoryForm({ ...categoryForm, color: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold"
                >
                  <option value="bg-indigo-500">أزرق نيلي (Indigo)</option>
                  <option value="bg-purple-500">بنفسجي (Purple)</option>
                  <option value="bg-pink-500">وردي (Pink)</option>
                  <option value="bg-emerald-500">أخضر زمردي (Emerald)</option>
                  <option value="bg-blue-500">أزرق (Blue)</option>
                  <option value="bg-amber-500">برتقالي / أصفر (Amber)</option>
                </select>
              </div>

              <button
                type="submit"
                className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-2xl shadow-lg transition-all"
              >
                {editingCategory ? "حفظ التعديلات" : "إضافة التصنيف"}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* PRODUCT DETAILS MODAL WITH REVIEWS & RATING SYSTEM */}
      {selectedProductDetails && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-2xl w-full p-6 sm:p-8 shadow-2xl space-y-6 max-h-[90vh] overflow-y-auto text-right">
            <div className="flex justify-between items-center border-b border-slate-200 pb-3">
              <h3 className="font-bold text-lg text-slate-900">تفاصيل المنتج والتقييمات</h3>
              <button onClick={() => setSelectedProductDetails(null)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-center">
              <div className="h-64 rounded-2xl overflow-hidden bg-slate-100 shadow-inner">
                <img src={selectedProductDetails.image} alt={selectedProductDetails.name} className="w-full h-full object-cover" />
              </div>

              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded-md">
                    {categories.find(c => c.id === selectedProductDetails.category)?.name}
                  </span>
                  <div className="flex items-center gap-1 text-amber-500 text-xs font-bold bg-amber-50 px-2.5 py-1 rounded-md">
                    <Star className="w-4 h-4 fill-amber-400 text-amber-400" />
                    <span>{getAverageRating(selectedProductDetails.reviews)}</span>
                    <span className="text-slate-400 text-[10px]">({selectedProductDetails.reviews?.length || 0} تقييم)</span>
                  </div>
                </div>

                <h3 className="text-xl font-extrabold text-slate-900">{selectedProductDetails.name}</h3>
                <p className="text-xs text-slate-600 leading-relaxed">{selectedProductDetails.description}</p>

                <div className="flex items-center gap-3">
                  {selectedProductDetails.discountPrice ? (
                    <>
                      <span className="text-2xl font-extrabold text-indigo-600">{selectedProductDetails.discountPrice} {settings.currency}</span>
                      <span className="text-sm text-slate-400 line-through">{selectedProductDetails.price} {settings.currency}</span>
                    </>
                  ) : (
                    <span className="text-2xl font-extrabold text-indigo-600">{selectedProductDetails.price} {settings.currency}</span>
                  )}
                </div>

                <div className="pt-2 flex gap-3">
                  <button
                    onClick={() => {
                      addToCart(selectedProductDetails);
                      setSelectedProductDetails(null);
                    }}
                    className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl text-xs shadow-lg flex items-center justify-center gap-2"
                  >
                    <ShoppingBag className="w-4 h-4" /> إضافة إلى السلة
                  </button>
                  {currentUser && (
                    <button
                      onClick={() => {
                        setSelectedProductDetails(null);
                        startAiVideoGeneration(selectedProductDetails);
                      }}
                      className="px-4 py-3 bg-purple-50 hover:bg-purple-100 text-purple-700 font-bold rounded-xl text-xs flex items-center gap-1.5"
                    >
                      <Video className="w-4 h-4" /> فيديو Veo
                    </button>
                  )}
                </div>
              </div>
            </div>

            {/* Ratings & Reviews Section */}
            <div className="border-t border-slate-200 pt-6 space-y-4">
              <h4 className="font-extrabold text-sm text-slate-900 flex items-center gap-2">
                <Star className="w-4 h-4 text-amber-500 fill-amber-400" /> تقييمات العملاء ({selectedProductDetails.reviews?.length || 0})
              </h4>

              <div className="space-y-3 max-h-48 overflow-y-auto pr-1">
                {selectedProductDetails.reviews && selectedProductDetails.reviews.map(rev => (
                  <div key={rev.id} className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1 text-xs">
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-slate-800">{rev.customerName}</span>
                      <div className="flex items-center gap-1 text-amber-500">
                        {Array.from({ length: 5 }).map((_, i) => (
                          <Star key={i} className={`w-3 h-3 ${i < rev.rating ? "fill-amber-400 text-amber-400" : "text-slate-300"}`} />
                        ))}
                      </div>
                    </div>
                    <p className="text-slate-600">{rev.comment}</p>
                    <span className="text-[10px] text-slate-400 block">{rev.date}</span>
                  </div>
                ))}
                {(!selectedProductDetails.reviews || selectedProductDetails.reviews.length === 0) && (
                  <p className="text-xs text-slate-400 text-center py-4">كن أول من يقيّم هذا المنتج!</p>
                )}
              </div>

              {/* Add Review Form */}
              <form onSubmit={handleAddReview} className="bg-slate-50 p-4 rounded-2xl border border-slate-200 space-y-3 text-xs">
                <h5 className="font-bold text-slate-800">أضف تقييمك للمنتج</h5>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">اسمك</label>
                    <input
                      type="text"
                      required
                      placeholder="محمد العمري"
                      value={newReviewName}
                      onChange={e => setNewReviewName(e.target.value)}
                      className="w-full p-2.5 bg-white border border-slate-200 rounded-xl"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-600 mb-1">التقييم بالنجوم (1 إلى 5)</label>
                    <select
                      value={newReviewRating}
                      onChange={e => setNewReviewRating(Number(e.target.value))}
                      className="w-full p-2.5 bg-white border border-slate-200 rounded-xl font-bold text-amber-600"
                    >
                      <option value={5}>⭐⭐⭐⭐⭐ (5/5 ممتاز)</option>
                      <option value={4}>⭐⭐⭐⭐ (4/5 جيد جداً)</option>
                      <option value={3}>⭐⭐⭐ (3/5 متوسط)</option>
                      <option value={2}>⭐⭐ (2/5 مقبولة)</option>
                      <option value={1}>⭐ (1/5 سيء)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block font-bold text-slate-600 mb-1">تعليقك أو تجربتك</label>
                  <textarea
                    rows={2}
                    required
                    placeholder="اكتب انطباعك عن جودة المنتج..."
                    value={newReviewComment}
                    onChange={e => setNewReviewComment(e.target.value)}
                    className="w-full p-2.5 bg-white border border-slate-200 rounded-xl"
                  />
                </div>

                <button
                  type="submit"
                  className="w-full py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow transition-all"
                >
                  إرسال التقييم
                </button>
              </form>
            </div>

          </div>
        </div>
      )}

      {/* CONTACT US MODAL */}
      {isContactModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-lg w-full p-6 sm:p-8 shadow-2xl space-y-6 text-right">
            <div className="flex items-center justify-between border-b border-slate-200 pb-4">
              <h3 className="font-bold text-lg text-slate-900 flex items-center gap-2">
                <Mail className="w-5 h-5 text-indigo-600" /> نظام المراسلات والدعم الفني الذكي
              </h3>
              <button onClick={() => setIsContactModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-xl">
                <X className="w-5 h-5 text-slate-500" />
              </button>
            </div>
            <form onSubmit={e => {
              e.preventDefault();
              if (!contactForm.content.trim()) return;
              const senderName = currentUser?.name || userProfileForm.name || "عميل المتجر";
              const senderEmail = currentUser?.email || userProfileForm.email || "customer@store.com";
              const senderPhone = currentUser?.phone || userProfileForm.phone || "0500000000";

              const subjectText = contactForm.issueType === "order" 
                ? `مشكلة في الطلب #${contactForm.selectedOrderId || "غير محدد"}`
                : contactForm.subject || "مشكلة تقنية عامة / استفسار";

              const fullContent = contactForm.issueType === "order"
                ? `[رقم الطلب المرتبط: ${contactForm.selectedOrderId}]\n${contactForm.content}`
                : contactForm.content;

              const newMsg: Message = {
                id: `msg-${Date.now()}`,
                senderName,
                senderEmail,
                senderPhone,
                subject: subjectText,
                content: fullContent,
                createdAt: new Date().toISOString().replace("T", " ").substring(0, 16),
                status: "unread"
              };
              setMessages([newMsg, ...messages]);
              setIsContactModalOpen(false);
              setContactForm({ issueType: "technical", selectedOrderId: "", subject: "", content: "" });
              alert("تم إرسال رسالتك بنجاح من حسابك إلى الإدارة! سيتم الرد عليك قريباً.");
            }} className="space-y-4 text-xs">
              
              <div className="p-3 bg-indigo-50/70 rounded-2xl border border-indigo-100 text-indigo-900 flex items-center justify-between">
                <div>
                  <span className="font-extrabold block text-slate-800">{currentUser?.name || userProfileForm.name}</span>
                  <span className="text-[10px] text-slate-500">{currentUser?.email || userProfileForm.email} | {currentUser?.phone || userProfileForm.phone}</span>
                </div>
                <span className="bg-indigo-600 text-white text-[10px] font-bold px-2.5 py-1 rounded-full">حسابك الحالي (إرسال تلقائي)</span>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">نوع الاستفسار / المشكلة</label>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    type="button"
                    onClick={() => setContactForm({ ...contactForm, issueType: "technical" })}
                    className={`py-2.5 px-4 rounded-xl font-bold text-xs border transition-all ${contactForm.issueType === "technical" ? "bg-indigo-600 text-white border-indigo-600 shadow" : "bg-slate-50 text-slate-700 border-slate-200"}`}
                  >
                    مشكلة تقنية / عامة
                  </button>
                  <button
                    type="button"
                    onClick={() => setContactForm({ ...contactForm, issueType: "order" })}
                    className={`py-2.5 px-4 rounded-xl font-bold text-xs border transition-all ${contactForm.issueType === "order" ? "bg-indigo-600 text-white border-indigo-600 shadow" : "bg-slate-50 text-slate-700 border-slate-200"}`}
                  >
                    مشكلة في طلب شراء
                  </button>
                </div>
              </div>

              {contactForm.issueType === "order" && (
                <div>
                  <label className="block font-bold text-slate-700 mb-1">اختر الطلب المرتبط بالمشكلة</label>
                  <select
                    value={contactForm.selectedOrderId}
                    onChange={e => setContactForm({ ...contactForm, selectedOrderId: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl font-bold text-indigo-600"
                  >
                    <option value="">-- اختر رقم الطلب --</option>
                    {orders.map(ord => (
                      <option key={ord.id} value={ord.id}>
                        {ord.id} - {ord.totalAmount} {settings.currency} ({ord.status === "completed" ? "مكتمل" : "قيد المعالجة"})
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {contactForm.issueType === "technical" && (
                <div>
                  <label className="block font-bold text-slate-700 mb-1">موضوع المشكلة التقنية</label>
                  <input
                    type="text"
                    required
                    placeholder="مثال: مشكلة في تصفح المنتجات / إتمام السداد"
                    value={contactForm.subject}
                    onChange={e => setContactForm({ ...contactForm, subject: e.target.value })}
                    className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl"
                  />
                </div>
              )}

              <div>
                <label className="block font-bold text-slate-700 mb-1">وصف المشكلة بالتفصيل</label>
                <textarea
                  rows={4}
                  required
                  placeholder="اكتب تفاصيل مشكلتك هنا..."
                  value={contactForm.content}
                  onChange={e => setContactForm({ ...contactForm, content: e.target.value })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl"
                />
              </div>

              <button
                type="submit"
                className="w-full py-3.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl shadow-lg transition-all"
              >
                إرسال الرسالة إلى الدعم الفني
              </button>
            </form>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="bg-white border-t border-slate-200 py-8 mt-auto text-center text-xs text-slate-500">
        <div className="max-w-7xl mx-auto px-4 space-y-2">
          <p className="font-bold text-slate-700">{settings.storeName} &copy; 2026 - جميع الحقوق محفوظة</p>
          <p>منصة تجارة إلكترونية ذكية ديناميكية تتيح للموظفين إدارة الصور والفيديوهات والكوبونات وإدارة الطلبات المتقدمة</p>
        </div>
      </footer>

    </div>
  );
}

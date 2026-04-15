import React, { useState, useMemo, useEffect } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  Calendar, Clock, CheckCircle, AlertCircle, User, Edit, Plus, X,
  MessageSquare, List, Layout, UploadCloud, DownloadCloud, FileText,
  FileSearch, Check, FileDown, CalendarDays, Printer, FileType2, Trash2,
  Lock, UserCircle, Settings, LogOut, ShieldCheck, ArrowRight, PenTool,
  ClipboardList, CheckSquare, PlayCircle, Milestone, Filter, Key,
  Percent
} from "lucide-react";

// ==========================================
// 🔴 Supabase 連線資訊
// ==========================================
const SUPABASE_URL = "https://mksmrupvgkehvfadynee.supabase.co";
const SUPABASE_KEY = "sb_publishable_0WCOlZOefS12mmupLA5YFg_fPv_8Xn8";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

const OFFICIAL_EVENTS = [
  { date: "04-03", name: "日月潭星光螢火季 (起跑)" },
  { date: "09-12", name: "花火音樂嘉年華 (開幕)" },
  { date: "09-20", name: "日月潭國際萬人泳渡" },
  { date: "10-03", name: "環法自行車挑戰賽" },
  { date: "10-10", name: "國慶花火音樂會" },
  { date: "12-31", name: "日月潭跨年晚會煙火秀" },
];

const PUBLIC_HOLIDAYS = [
  { date: "01-01", name: "元旦" }, { date: "02-16", name: "除夕" },
  { date: "02-17", name: "春節 (初一)" }, { date: "02-18", name: "春節 (初二)" },
  { date: "02-19", name: "春節 (初三)" }, { date: "02-20", name: "春節 (初四)" },
  { date: "02-28", name: "和平紀念日" }, { date: "04-04", name: "兒童節" },
  { date: "04-05", name: "清明節" }, { date: "06-19", name: "端午節" },
  { date: "09-25", name: "中秋節" }, { date: "10-10", name: "國慶日" },
];

const PRESET_BREAKDOWN_ITEMS = ["早餐", "午餐", "下午茶", "晚餐", "宵夜", "DIY", "備品"];

const DEPARTMENTS = ["客務部", "訂房組", "餐飲部", "休閒部", "業務部", "企劃部", "人資", "資訊", "總務", "採購", "財務部"];

const ROLES_INFO: any = {
  guest: { name: "訪客 (僅觀看)", color: "text-gray-500", bg: "bg-gray-100" },
  employee: { name: "部門員工", color: "text-blue-700", bg: "bg-blue-100" },
  gm: { name: "總經理", color: "text-green-700", bg: "bg-green-100" },
  admin: { name: "系統管理員", color: "text-purple-700", bg: "bg-purple-100" },
};

const YEARLY_DATA: any = {
  2026: {
    bigHolidays: ["02-14", "02-15", "02-16", "02-17", "02-18", "02-19", "02-20"],
    holidays: ["01-01", "01-02", "02-27", "02-28", "04-03", "04-04", "04-05", "05-01", "06-19", "06-20", "06-21", "09-25", "09-26", "09-27", "10-09", "10-10"],
    events: {
      "01-01": "元旦", "02-16": "除夕", "02-17": "春節 (初一)", "02-18": "春節 (初二)", "02-19": "春節 (初三)", "02-20": "春節 (初四)",
      "02-28": "和平紀念日", "04-04": "兒童節", "04-05": "清明節", "06-19": "端午節", "09-25": "中秋節", "10-10": "國慶日"
    }
  },
  2027: {
    bigHolidays: [
      "01-01", "01-02", "01-03", 
      "02-04", "02-05", "02-06", "02-07", "02-08", "02-09", "02-10", 
      "02-27", "02-28", "03-01", 
      "04-03", "04-04", "04-05", "04-06", 
      "04-30", "05-01", "05-02", 
      "06-09", "09-15", 
      "10-09", "10-10", "10-11", 
      "10-23", "10-24", "10-25", 
      "12-24", "12-25", "12-26"  
    ],
    holidays: ["09-28"], 
    events: {
      "01-01": "元旦", "02-04": "小年夜", "02-05": "除夕", "02-06": "春節 (初一)", "02-07": "春節 (初二)", "02-08": "春節 (初三)", "02-09": "春節補假", "02-10": "春節補假",
      "02-28": "和平紀念日", "04-04": "兒童節", "04-05": "清明節", "05-01": "勞動節",
      "06-09": "端午節", "09-15": "中秋節", "09-28": "教師節", "10-10": "國慶日",
      "10-25": "光復節", "12-25": "行憲紀念日"
    }
  }
};

const generateCalendar = (year: number, customEvents: any[]) => {
  const data: any = {};
  const yearData = YEARLY_DATA[year] || { bigHolidays: [], holidays: [], events: {} };

  for (let m = 1; m <= 12; m++) {
    const daysInMonth = new Date(year, m, 0).getDate();
    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${year}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
      const dateObj = new Date(dateStr);
      const day = dateObj.getDay();
      const md = `${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
      
      let type = (day === 6 || day === 0) ? "假日" : "平日";
      let events: string[] = [];
      let marketingEvents: string[] = [];
      
      if (yearData.bigHolidays.includes(md)) {
        type = "大假日";
      } else if (yearData.holidays.includes(md)) {
        type = "假日";
      } else if (type === "平日" && year === 2026) {
        const isWinterVacation = (m === 1 && d >= 21) || (m === 2 && d <= 13);
        const isSummerVacation = m === 7 || m === 8;
        if (day === 5 || isWinterVacation || isSummerVacation) {
          type = "旺日";
        }
      }

      if (yearData.events[md]) {
        events.push(`🧨 ${yearData.events[md]}`);
      }

      const officialEvent = OFFICIAL_EVENTS.find((e) => e.date === md);
      if (officialEvent) events.push(`✨ ${officialEvent.name}`);
      const mktEvent = OFFICIAL_EVENTS.find((e) => e.date === md); // 修正變數參考

      if (customEvents && customEvents.length > 0) {
        const dbMatch = customEvents.find(e => e.date === dateStr);
        if (dbMatch) {
          if (dbMatch.event_type) type = dbMatch.event_type;
          if (dbMatch.event_name) {
            const icon = dbMatch.is_public_holiday ? '🧨' : '✨';
            const customName = `${icon} ${dbMatch.event_name}`;
            if (!events.includes(customName)) events.push(customName);
          }
        }
      }

      data[dateStr] = { type, events, marketingEvents };
    }
  }
  return data;
};

const CAL_STYLES: any = {
  平日: { bg: "bg-white", border: "border-gray-200", text: "text-gray-800", tag: "bg-gray-100 text-gray-500" },
  旺日: { bg: "bg-blue-50/50", border: "border-blue-200", text: "text-blue-900", tag: "bg-blue-100 text-blue-600" },
  假日: { bg: "bg-pink-50/50", border: "border-pink-200", text: "text-pink-900", tag: "bg-pink-100 text-pink-600" },
  大假日: { bg: "bg-red-50", border: "border-red-300", text: "text-red-900", tag: "bg-red-500 text-white" },
};

const evaluateExpression = (expr: string) => {
  if (!expr) return 0;
  try { return new Function(`return ${expr.replace(/\s+/g, "").replace(/[^-()\d/*+.]/g, "")}`)() || 0; } catch (e) { return 0; }
};

const getCurrentTimeString = () => {
  const d = new Date();
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
};

export default function App() {
  const [view, setView] = useState("login");
  const [currentUser, setCurrentUser] = useState<any>(null);

  const [users, setUsers] = useState<any[]>([]);
  const [projects, setProjects] = useState<any[]>([]);
  const [dbEvents, setDbEvents] = useState<any[]>([]); 
  const [isLoading, setIsLoading] = useState(true);

  const currentYear = new Date().getFullYear();
  const [selectedYear, setSelectedYear] = useState(currentYear);

  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingProject, setEditingProject] = useState<any>(null);
  const [modalMode, setModalMode] = useState("view");

  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [importText, setImportText] = useState("");
  const [importFile, setImportFile] = useState<File | null>(null);

  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [exportConfig, setExportConfig] = useState({ year: currentYear, month: "all" });
  const [isPrintLayoutActive, setIsPrintLayoutActive] = useState(false);
  const [selectedDayInfo, setSelectedDayInfo] = useState<any>(null);

  const [editingUser, setEditingUser] = useState<any>(null);
  const [isUserModalOpen, setIsUserModalOpen] = useState(false);
  const [isChangePwdModalOpen, setIsChangePwdModalOpen] = useState(false);
  const [pwdForm, setPwdForm] = useState({ old: "", new: "", confirm: "" });

  const currentMonthStr = String(new Date().getMonth() + 1);
  const [dashboardDeptFilter, setDashboardDeptFilter] = useState("all");
  const [dashboardMonthFilter, setDashboardMonthFilter] = useState(currentMonthStr);
  const [ganttDeptFilter, setGanttDeptFilter] = useState("all");
  const [ganttMonthFilter, setGanttMonthFilter] = useState("all");

  const [rememberMe, setRememberMe] = useState(() => localStorage.getItem("mpr_remember") === "true");
  const [savedAccount, setSavedAccount] = useState(() => localStorage.getItem("mpr_account") || "");
  const [savedPassword, setSavedPassword] = useState(() => localStorage.getItem("mpr_password") || "");

  useEffect(() => { fetchData(); }, []);

  const fetchData = async () => {
    setIsLoading(true);
    try {
      const { data: usersData } = await supabase.from("users").select("*");
      if (usersData && usersData.length === 0) {
        const defaultAdmin = { id: "u1", account: "admin", password: "123", name: "管理員", role: "admin", dept: "系統" };
        await supabase.from("users").upsert(defaultAdmin);
        setUsers([defaultAdmin]);
      } else { setUsers(usersData || []); }

      const { data: projData } = await supabase.from("projects").select("*");
      if (projData) {
        const parsedProjects = projData.map((p: any) => {
          let bd;
          try { bd = typeof p.breakdown === "string" ? JSON.parse(p.breakdown || "[]") : (p.breakdown || []); } 
          catch (e) { bd = []; }
          if (!Array.isArray(bd)) {
            bd = [{ id: Date.now() + Math.random(), name: "主專案", price: bd.price || "", ota: "", items: bd.items || [], net: bd.net || "0" }];
          }
          return {
            ...p,
            projectType: p.projectType || 'leisure',
            breakdown: bd,
            countersign: typeof p.countersign === "string" ? JSON.parse(p.countersign || "[]") : (p.countersign || []),
          };
        });
        setProjects(parsedProjects);
      }
      const { data: calData } = await supabase.from("calendar_events").select("*");
      if (calData) setDbEvents(calData);
    } catch (e) { console.error("資料載入發生錯誤:", e); }
    setIsLoading(false);
  };

  const saveProjectToDb = async (proj: any) => {
    const dbProj = { ...proj, breakdown: JSON.stringify(proj.breakdown), countersign: JSON.stringify(proj.countersign), id: String(proj.id) };
    const { error } = await supabase.from("projects").upsert(dbProj);
    if (error) {
      alert(`⚠️ 儲存失敗！\n錯誤原因：${error.message}\n\n請放心，您的簽呈資料還在畫面上沒有遺失！`);
      return false; 
    } else { fetchData(); return true; }
  };

  const handleSaveDayRemark = async (dateStr: string, newRemark: string) => {
    const existing = dbEvents.find(e => e.date === dateStr);
    const payload = existing
      ? { ...existing, description: newRemark } 
      : { date: dateStr, event_name: "特殊備註", event_type: "平日", is_public_holiday: false, description: newRemark }; 
    const { error } = await supabase.from('calendar_events').upsert(payload, { onConflict: 'date' });
    if (error) alert("儲存備註失敗：" + error.message); else { alert("備註已成功儲存至資料庫！"); fetchData(); setSelectedDayInfo(null); }
  };

  const handleSaveUser = async (e: any) => {
    e.preventDefault();
    const userId = editingUser.id || "u_" + Date.now();
    const userToSave = { ...editingUser, id: userId };
    const { error } = await supabase.from("users").upsert(userToSave);
    if (error) alert("儲存帳號失敗！" + error.message); else { fetchData(); setIsUserModalOpen(false); }
  };

  const handleDeleteUser = async (id: string) => {
    if (window.confirm("確定要刪除此帳號嗎？")) {
      const { error } = await supabase.from("users").delete().eq("id", id);
      if (!error) fetchData();
    }
  };

  const handleChangePassword = async (e: any) => {
    e.preventDefault();
    if (pwdForm.old !== currentUser.password) { alert("原密碼錯誤！"); return; }
    if (pwdForm.new !== pwdForm.confirm) { alert("新密碼不一致！"); return; }
    const updatedUser = { ...currentUser, password: pwdForm.new };
    await supabase.from("users").upsert(updatedUser);
    setCurrentUser(updatedUser);
    alert("密碼變更成功！"); setIsChangePwdModalOpen(false); setPwdForm({ old: "", new: "", confirm: "" });
  };

  const handleProcessImport = async () => {
    let extractedContent = importText;
    if (importFile) {
      try {
        const data = await importFile.arrayBuffer();
        const workbook = XLSX.read(data);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        extractedContent = XLSX.utils.sheet_to_txt(sheet);
      } catch (error) { alert("讀取 Excel 失敗"); return; }
    }
    if (!extractedContent.trim()) { alert("未偵測到內容"); return; }
    setIsImportModalOpen(false);
    const today = new Date(); const twYear = today.getFullYear() - 1911; const mm = String(today.getMonth() + 1).padStart(2, "0");
    setEditingProject({
      id: Date.now(), title: "【從匯入自動建立】", refNo: `MPR-${twYear}-${mm}-${String(projects.length + 1).padStart(3, "0")}`, projectType: "leisure", applyDate: today.toISOString().split("T")[0], createTime: getCurrentTimeString(), purpose: extractedContent, startDate: `${selectedYear}-01-01`, endDate: `${selectedYear}-01-31`, content: "", breakdown: [{ id: Date.now(), name: "專案一", price: "", ota: "", items: [], net: "0" }], countersign: [], status: "countersigning", feedback: "", creator: `${currentUser?.dept || ""} - ${currentUser?.name || ""}`,
    });
    setModalMode("create"); setIsModalOpen(true); setImportText(""); setImportFile(null);
  };

  useEffect(() => {
    if (currentUser) {
      setDashboardDeptFilter(["admin", "gm"].includes(currentUser.role) ? "all" : currentUser.dept || "all");
      setDashboardMonthFilter(currentMonthStr);
    }
  }, [currentUser, currentMonthStr]);

  const yearOptions = useMemo(() => {
    let minYear = currentYear - 1; let maxYear = currentYear + 2;
    projects.forEach((p) => {
      const start = parseInt(p.startDate.split("-")[0], 10); const end = parseInt(p.endDate.split("-")[0], 10);
      if (!isNaN(start) && start < minYear) minYear = start; if (!isNaN(end) && end > maxYear) maxYear = end;
    });
    const options = []; for (let i = minYear; i <= maxYear; i++) options.push(i);
    return options;
  }, [projects, currentYear]);

  const calendarData = useMemo(() => generateCalendar(selectedYear, dbEvents), [selectedYear, dbEvents]);
  const [selectedMonthView, setSelectedMonthView] = useState<number | null>(null);

  const yearProjects = useMemo(() => projects.filter((p) => p.startDate.startsWith(String(selectedYear)) || p.endDate.startsWith(String(selectedYear))), [projects, selectedYear]);
  const scheduledProjects = useMemo(() => yearProjects.filter((p) => p.status === "scheduled"), [yearProjects]);
  
  const filteredScheduledProjects = useMemo(() => {
    return scheduledProjects
      .filter((p) => {
        let passDept = ganttDeptFilter === "all" || (p.creator && p.creator.includes(ganttDeptFilter));
        let passMonth = true;
        if (ganttMonthFilter !== "all") {
          const filterStart = new Date(selectedYear, parseInt(ganttMonthFilter) - 1, 1);
          const filterEnd = new Date(selectedYear, parseInt(ganttMonthFilter), 0);
          passMonth = new Date(p.startDate) <= filterEnd && new Date(p.endDate) >= filterStart;
        }
        return passDept && passMonth;
      })
      .sort((a, b) => {
        const isFullYear = (start: string, end: string) => {
          const d1 = new Date(start).getTime(); const d2 = new Date(end).getTime();
          return Math.ceil(Math.abs(d2 - d1) / (1000 * 60 * 60 * 24)) >= 360 || (start.endsWith("-01-01") && end.endsWith("-12-31"));
        };
        const fullA = isFullYear(a.startDate, a.endDate); const fullB = isFullYear(b.startDate, b.endDate);
        if (fullA && !fullB) return 1; if (!fullA && fullB) return -1;
        return new Date(a.startDate).getTime() - new Date(b.startDate).getTime();
      });
  }, [scheduledProjects, ganttDeptFilter, ganttMonthFilter, selectedYear]);

  const dashboardActiveProjects = useMemo(() => {
    return scheduledProjects
      .filter((p) => {
        let passDept = dashboardDeptFilter === "all" || (p.creator && p.creator.includes(dashboardDeptFilter)) || (p.countersign && p.countersign.some((c: any) => c.dept === dashboardDeptFilter));
        let passMonth = true;
        if (dashboardMonthFilter !== "all") {
          const filterStart = new Date(selectedYear, parseInt(dashboardMonthFilter) - 1, 1);
          const filterEnd = new Date(selectedYear, parseInt(dashboardMonthFilter), 0);
          passMonth = new Date(p.startDate) <= filterEnd && new Date(p.endDate) >= filterStart;
        }
        return passDept && passMonth;
      })
      .sort((a, b) => {
        const isFullYear = (start: string, end: string) => {
          const d1 = new Date(start).getTime(); const d2 = new Date(end).getTime();
          return Math.ceil(Math.abs(d2 - d1) / (1000 * 60 * 60 * 24)) >= 360 || (start.endsWith("-01-01") && end.endsWith("-12-31"));
        };
        const fullA = isFullYear(a.startDate, a.endDate); const fullB = isFullYear(b.startDate, b.endDate);
        if (fullA && !fullB) return 1; if (!fullA && fullB) return -1;
        return new Date(a.startDate).getTime() - new Date(b.startDate).getTime();
      });
  }, [scheduledProjects, dashboardDeptFilter, dashboardMonthFilter, selectedYear]);

  const myPendingCountersign = useMemo(() => {
    if (!currentUser) return [];
    return yearProjects.filter((p) => p.status === "countersigning" && p.countersign && p.countersign.some((c: any) => c.dept === currentUser.dept && c.status === "pending"));
  }, [yearProjects, currentUser]);
  
  const myOwnProposals = useMemo(() => {
    if (!currentUser) return [];
    return yearProjects.filter((p) => p.creator && p.creator.includes(currentUser.name) && ["countersigning", "revision", "unconfirmed"].includes(p.status));
  }, [yearProjects, currentUser]);
  
  const managerPendingApproval = useMemo(() => {
    if (!currentUser || !["admin", "gm"].includes(currentUser.role)) return [];
    return yearProjects.filter((p) => p.status === "unconfirmed");
  }, [yearProjects, currentUser]);

  const handleLoginSubmit = (e: any) => {
    e.preventDefault(); const formData = new FormData(e.target); const acc = formData.get("account"); const pass = formData.get("password");
    const user = users.find((u) => u.account === acc && u.password === pass);
    if (user) {
      if (rememberMe) { localStorage.setItem("mpr_account", acc as string); localStorage.setItem("mpr_password", pass as string); localStorage.setItem("mpr_remember", "true"); } 
      else { localStorage.removeItem("mpr_account"); localStorage.removeItem("mpr_password"); }
      setCurrentUser(user); setView("app");
    } else { alert("登入失敗"); }
  };
  const handleGuestLogin = () => { setCurrentUser({ id: "guest", name: "訪客", role: "guest", dept: "" }); setView("app"); };
  const handleLogout = () => { setCurrentUser(null); setView("login"); };

  const handleOpenCreate = () => {
    if (currentUser?.role === "guest") return;
    const today = new Date(); const twYear = today.getFullYear() - 1911; const mm = String(today.getMonth() + 1).padStart(2, "0");
    setEditingProject({
      id: Date.now(), title: "", refNo: `MPR-${twYear}-${mm}-${String(projects.length + 1).padStart(3, "0")}`, projectType: "leisure", applyDate: today.toISOString().split("T")[0], createTime: getCurrentTimeString(), purpose: "", startDate: `${selectedYear}-01-01`, endDate: `${selectedYear}-01-31`, content: "", breakdown: [{ id: Date.now(), name: "專案一", price: "", ota: "", items: [], net: "0" }], countersign: [], status: "countersigning", feedback: "", creator: `${currentUser.dept} - ${currentUser.name}`,
    });
    setModalMode("create"); setIsModalOpen(true);
  };

  const handleSave = async (e: any) => { 
    e.preventDefault(); 
    let updatedProj = { ...editingProject }; 
    if (modalMode === "create" && updatedProj.countersign.length === 0) updatedProj.status = "revision"; 
    const isSuccess = await saveProjectToDb(updatedProj); 
    if (isSuccess) setIsModalOpen(false); 
  };

  const updatePackage = (pIdx: number, field: string, value: any) => {
    const newBd = [...editingProject.breakdown]; newBd[pIdx][field] = value;
    if (['price', 'ota', 'items'].includes(field)) {
      const price = parseFloat(String(newBd[pIdx].price).replace(/,/g, "")) || 0;
      const ota = parseFloat(String(newBd[pIdx].ota)) || 0;
      const otaAmount = Math.round(price * (ota / 100));
      let totalDeductions = 0; (newBd[pIdx].items || []).forEach((item: any) => { totalDeductions += evaluateExpression(item.value); });
      const net = price - otaAmount - totalDeductions;
      newBd[pIdx].net = new Intl.NumberFormat("en-US").format(net);
    }
    setEditingProject({ ...editingProject, breakdown: newBd });
  };

  const handlePrintSingleProject = () => { setTimeout(() => { window.print(); }, 300); };
  const handleExportSystemPDF = () => { setIsPrintLayoutActive(true); setIsExportModalOpen(false); setTimeout(() => { window.print(); setIsPrintLayoutActive(false); }, 800); };

  const exportSingleProjectToWord = (project: any) => {
    const header = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>簽呈匯出</title><style>body{font-family:"Microsoft JhengHei",Arial,sans-serif;}table{border-collapse:collapse;width:100%;margin-bottom:20px;}th,td{border:1px solid black;padding:8px;text-align:left;vertical-align:top;}.center{text-align:center;}.no-border{border:none;}.no-border td{border:none;padding:4px 0;} .comments { color: black; font-weight: bold; background-color: #f9fafb; padding: 10px; border: 1px solid #ccc; }</style></head><body>`;
    const formatText = (t: string) => String(t || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br/>");
    let breakdownHtml = "";
    (project.breakdown || []).forEach((pkg: any) => {
      const otaVal = pkg.ota ? Math.round((parseFloat(String(pkg.price).replace(/,/g, ""))||0) * (parseFloat(pkg.ota)/100)) : 0;
      breakdownHtml += `<h4>${formatText(pkg.name)}</h4><table><tr><th>售價</th><th>OTA扣除</th><th>淨價</th></tr><tr><td>${formatText(pkg.price)}</td><td>${otaVal}</td><td>${formatText(pkg.net)}</td></tr></table>`;
    });
    const deptName = project.creator?.split('-')[0]?.trim() || "行銷公關部"; 
    const htmlContent = `<h2>${formatText(deptName)} 簽呈 Official application</h2><div>Date:${project.applyDate} Ref:${project.refNo}</div>${breakdownHtml}`; [cite: 1, 3]
    const blob = new Blob(["\ufeff", header + htmlContent + "</body></html>"], { type: "application/msword" });
    const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = `簽呈.doc`; link.click();
  };

  const getBarStyles = (start: string, end: string) => {
    const yearStart = new Date(`${selectedYear}-01-01`).getTime(); const yearEnd = new Date(`${selectedYear}-12-31`).getTime();
    const s = Math.max(yearStart, new Date(start).getTime()); const e = Math.min(yearEnd, new Date(end).getTime());
    let left = ((s - yearStart) / (yearEnd - yearStart)) * 100;
    let width = ((e - s) / (yearEnd - yearStart)) * 100;
    return { left: `${left}%`, width: `${Math.max(1, width)}%` };
  };

  return (
    <div className={`min-h-screen font-sans ${isPrintLayoutActive ? "bg-white" : "bg-slate-50"}`}>
      {/* 🌟 PDF 列印樣式優化 */}
      <style>{`
        .sign-table { width: 100%; border-collapse: collapse; border: 1px solid #d1d5db; margin-bottom: 1rem; }
        .sign-table th, .sign-table td { border: 1px solid #d1d5db; padding: 0.5rem; vertical-align: top; }
        .sign-table th { background-color: #f3f4f6; color: #374151; font-weight: 600; text-align: center; }
        
        @media print {
          body * { visibility: hidden; }
          body, html { height: auto !important; overflow: visible !important; }
          .print-modal, .print-modal * { visibility: visible; }
          .print-modal { position: absolute !important; left: 0 !important; top: 0 !important; width: 100% !important; border: none !important; box-shadow: none !important; }
          .no-print, .no-print * { display: none !important; }
          table { page-break-inside: auto; width: 100%; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          .sign-table th { background-color: #e5e7eb !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
      `}</style>

      {!isPrintLayoutActive && (
        <header className="bg-white shadow-sm sticky top-0 z-20 print:hidden">
          <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
            <div className="flex items-center gap-2 text-indigo-600"><Layout className="w-6 h-6" /><h1 className="font-bold text-xl">專案管理系統</h1></div>
            <div className="flex gap-2">
              {currentUser && <span className="text-sm font-medium">{currentUser.dept} - {currentUser.name}</span>}
              <button onClick={handleLogout} className="text-gray-400 hover:text-red-600 p-2"><LogOut className="w-5 h-5" /></button>
            </div>
          </div>
        </header>
      )}

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-6 print:hidden">
        <section className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
          <div className="flex justify-between items-center mb-6">
            <h2 className="text-2xl font-bold flex items-center gap-2"><PlayCircle className="w-6 h-6 text-indigo-600" /> 年度專案看板</h2>
            <button onClick={handleOpenCreate} className="bg-indigo-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 shadow-sm"><Plus className="w-4 h-4" /> 新增簽呈</button>
          </div>
          {/* 甘特圖簡略實現 */}
          <div className="overflow-x-auto border rounded-lg">
             <div className="min-w-[800px] bg-gray-50 flex border-b font-bold text-sm">
                <div className="w-48 p-2 border-r">專案名稱</div>
                <div className="flex-1 grid grid-cols-12 text-center">
                  {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <div key={m} className="p-2 border-r last:border-0">{m}月</div>)}
                </div>
             </div>
             {filteredScheduledProjects.map(p => (
               <div key={p.id} className="min-w-[800px] flex border-b last:border-0 group cursor-pointer hover:bg-slate-50" onClick={() => { setEditingProject(p); setModalMode("view"); setIsModalOpen(true); }}>
                  <div className="w-48 p-2 border-r text-sm truncate font-medium">{p.title}</div>
                  <div className="flex-1 relative h-10">
                    <div className="absolute top-2 bottom-2 bg-indigo-500 rounded text-[10px] text-white flex items-center px-2" style={getBarStyles(p.startDate, p.endDate)}>{p.title}</div>
                  </div>
               </div>
             ))}
          </div>
        </section>
      </main>

      {/* 💥 專案 Modal (PDF核心修改處) 💥 */}
      {isModalOpen && editingProject && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm print:static print:p-0">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-5xl flex flex-col max-h-[95vh] print-modal print:max-h-none overflow-hidden print:overflow-visible">
            
            <div className="px-6 py-4 border-b bg-gray-50 flex justify-between items-center no-print">
              <h2 className="font-bold text-gray-800">{modalMode === "view" ? "簽呈預覽" : "編輯簽呈"}</h2>
              <div className="flex items-center gap-2">
                {modalMode === "view" && <button onClick={handlePrintSingleProject} className="flex items-center gap-1 text-sm bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded"><Printer className="w-4 h-4" /> 列印 PDF</button>}
                <button onClick={() => setIsModalOpen(false)} className="text-gray-400 hover:text-gray-600"><X className="w-6 h-6" /></button>
              </div>
            </div>

            <div className="p-8 overflow-y-auto print:p-4 text-gray-900 bg-white">
              {/* 🌟 列印頁首：動態抓取部門 */}
              <h1 className="text-2xl font-bold text-center mb-6">
                {editingProject.creator?.split('-')[0]?.trim() || '飯店專案'} 簽呈 Official application 
              </h1>
              
              {/* 列印資料欄位 */}
              <div className="mb-4 font-medium flex justify-between text-sm">
                <span>Date 日期：{editingProject.applyDate}</span> [cite: 3, 8, 13]
                <span>Ref No 文檔號：{editingProject.refNo}</span> [cite: 3, 8, 13]
              </div>

              <table className="sign-table">
                <tbody>
                  <tr><th className="w-32">專案類型</th><td className="font-bold text-indigo-600">{editingProject.projectType === 'room' ? '🏨 住房專案' : '☕ 休閒/餐飲專案'}</td></tr> [cite: 4, 9, 14]
                  <tr><th>主旨</th><td className="font-bold">{editingProject.title}</td></tr> [cite: 4, 9, 14]
                  <tr><th>說明</th><td className="whitespace-pre-wrap">{editingProject.purpose}</td></tr> [cite: 4, 9, 14]
                  <tr><th>日期</th><td>{editingProject.startDate} ～ {editingProject.endDate}</td></tr> [cite: 4, 9, 14]
                </tbody>
              </table>

              <h3 className="font-bold mt-6 mb-2">財務內拆表</h3>
              {(editingProject.breakdown || []).map((pkg: any, idx: number) => {
                const otaVal = pkg.ota ? Math.round((parseFloat(String(pkg.price).replace(/,/g, "")) || 0) * (parseFloat(pkg.ota) / 100)) : 0;
                return (
                  <div key={idx} className="mb-4">
                    <p className="font-bold text-sm text-indigo-800 mb-1">{pkg.name}</p>
                    <table className="sign-table text-center text-sm">
                      <thead>
                        <tr><th width="20%">售價</th><th width="20%">OTA扣除 ({pkg.ota}%)</th><th>扣除項目</th><th width="20%">淨價</th></tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>{pkg.price}</td>
                          <td className="text-red-600 font-bold">-{otaVal}</td>
                          <td>{(pkg.items || []).map((i:any) => `${i.name}:${i.value}`).join(' | ')}</td>
                          <td className="font-bold text-indigo-700">{pkg.net}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                );
              })}

              <div className="mt-6">
                <h3 className="font-bold mb-2">會簽與意見備註</h3>
                <div className="space-y-3">
                  {(editingProject.countersign || []).map((c: any, i: number) => (
                    <div key={i} className="border p-2 rounded bg-gray-50 text-xs">
                      <div className="flex justify-between border-b pb-1 mb-1">
                        <span className="font-bold">{c.dept}</span>
                        <span className={c.status === "approved" ? "text-green-600" : "text-orange-500"}>
                          {c.status === "approved" ? `✓ 已確認 (${c.time})` : "待確認"}
                        </span>
                      </div>
                      <div className="italic">{c.comment || "無意見。"}</div>
                    </div>
                  ))}
                </div>
              </div>

            </div>
          </div>
        </div>
      )}
    </div>
  );
}

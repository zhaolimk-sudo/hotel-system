import React, { useState, useMemo, useEffect } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  Calendar, Clock, CheckCircle, AlertCircle, User, Edit, Plus, X,
  MessageSquare, List, Layout, UploadCloud, DownloadCloud, FileText,
  FileSearch, Check, FileDown, CalendarDays, Printer, FileType2, Trash2,
  Lock, UserCircle, Settings, LogOut, ShieldCheck, ArrowRight, PenTool,
  ClipboardList, CheckSquare, PlayCircle, Milestone, Filter, Key,
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

const PRESET_BREAKDOWN_ITEMS = ["早餐", "午餐", "下午茶", "晚餐", "宵夜", "DIY"];

const MARKETING_EVENTS = [
  { date: "02-14", name: "西洋情人節" }, { date: "05-10", name: "母親節檔期" },
  { date: "08-08", name: "父親節" }, { date: "10-31", name: "萬聖節" },
  { date: "12-25", name: "聖誕節" },
];

const ROLES_INFO: any = {
  guest: { name: "訪客 (僅觀看)", color: "text-gray-500", bg: "bg-gray-100" },
  employee: { name: "部門員工", color: "text-blue-700", bg: "bg-blue-100" },
  gm: { name: "總經理", color: "text-green-700", bg: "bg-green-100" },
  admin: { name: "系統管理員", color: "text-purple-700", bg: "bg-purple-100" },
};

const DEPARTMENTS = ["客務部", "訂房組", "餐飲部", "休閒部", "業務部", "企劃部", "人資", "資訊", "總務", "採購", "財務部"];

// 🌟 安全的日曆生成邏輯 (嚴格隔離 2026 與 2027)
const generateCalendar = (year: number, customEvents: any[]) => {
  const data: any = {};
  for (let m = 1; m <= 12; m++) {
    const daysInMonth = new Date(year, m, 0).getDate();
    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${year}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
      const dateObj = new Date(dateStr);
      const day = dateObj.getDay();
      const md = `${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
      
      // 1. 最基礎的白紙狀態：週末=假日，平日=平日
      let type = (day === 6 || day === 0) ? "假日" : "平日";
      let events: string[] = [];
      let marketingEvents: string[] = [];
      
      // 節日名稱標籤 (每年固定顯示，不影響平旺日背景色)
      const pubHoliday = PUBLIC_HOLIDAYS.find((e) => e.date === md);
      if (pubHoliday) events.push(`🧨 ${pubHoliday.name}`);

      const officialEvent = OFFICIAL_EVENTS.find((e) => e.date === md);
      if (officialEvent) events.push(`✨ ${officialEvent.name}`);
      
      const mktEvent = MARKETING_EVENTS.find((e) => e.date === md);
      if (mktEvent) marketingEvents.push(`🎯 ${mktEvent.name}`);

      // 2. 🛡️ 嚴格隔離：只有 2026 年才套用寫死的飯店特殊「平旺假日」規則
      if (year === 2026) {
        const bigHolidays = ["02-14", "02-15", "02-16", "02-17", "02-18", "02-19", "02-20"];
        const holidays = ["01-01", "01-02", "02-27", "02-28", "04-03", "04-04", "04-05", "05-01", "06-19", "06-20", "06-21", "09-25", "09-26", "09-27", "10-09", "10-10"];
        const isWinterVacation = (m === 1 && d >= 21) || (m === 2 && d <= 13);
        const isSummerVacation = m === 7 || m === 8;

        if (bigHolidays.includes(md) || md === "09-20") {
          type = "大假日";
        } else if (holidays.includes(md) || day === 6 || day === 0) {
          type = "假日";
        } else if (day === 5 || isWinterVacation || isSummerVacation) {
          type = "旺日";
        } else {
          type = "平日";
        }
      }

      // 3. 👑 最高指揮權：如果 Supabase 資料庫有設定，直接無條件覆蓋！
      if (customEvents && customEvents.length > 0) {
        const dbMatch = customEvents.find(e => e.date === dateStr);
        if (dbMatch) {
          if (dbMatch.event_type) type = dbMatch.event_type; // 覆蓋平旺日顏色
          if (dbMatch.event_name) {
            const icon = dbMatch.is_public_holiday ? '🧨' : '✨';
            const customName = `${icon} ${dbMatch.event_name}`;
            if (!events.includes(customName)) {
              events.push(customName);
            }
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
      } else { 
        setUsers(usersData || []); 
      }

      const { data: projData } = await supabase.from("projects").select("*");
      if (projData) {
        // 🛡️ 安全解析機制
        const parsedProjects = projData.map((p: any) => ({
          ...p,
          breakdown: typeof p.breakdown === "string" ? JSON.parse(p.breakdown || "{}") : (p.breakdown || {}),
          countersign: typeof p.countersign === "string" ? JSON.parse(p.countersign || "[]") : (p.countersign || []),
        }));
        setProjects(parsedProjects);
      }

      const { data: calData } = await supabase.from("calendar_events").select("*");
      if (calData) {
        setDbEvents(calData);
      }

    } catch (e) { 
      console.error("資料載入發生錯誤:", e); 
    }
    setIsLoading(false);
  };

  const saveProjectToDb = async (proj: any) => {
    const dbProj = { ...proj, breakdown: JSON.stringify(proj.breakdown), countersign: JSON.stringify(proj.countersign), id: String(proj.id) };
    const { error } = await supabase.from("projects").upsert(dbProj);
    if (error) alert("儲存失敗！" + error.message); else fetchData();
  };

  const handleSaveDayRemark = async (dateStr: string, newRemark: string) => {
    const existing = dbEvents.find(e => e.date === dateStr);
    const payload = existing
      ? { ...existing, description: newRemark } 
      : { date: dateStr, event_name: "特殊備註", event_type: "平日", is_public_holiday: false, description: newRemark }; 

    const { error } = await supabase.from('calendar_events').upsert(payload, { onConflict: 'date' });
    if (error) {
      alert("儲存備註失敗：" + error.message);
    } else {
      alert("備註已成功儲存至資料庫！");
      fetchData(); 
      setSelectedDayInfo(null);
    }
  };

  const handleSaveUser = async (e: any) => {
    e.preventDefault();
    const userId = editingUser.id || "u_" + Date.now();
    const userToSave = { ...editingUser, id: userId };
    const { error } = await supabase.from("users").upsert(userToSave);
    if (error) alert("儲存帳號失敗！" + error.message);
    else { fetchData(); setIsUserModalOpen(false); }
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
    if (localStorage.getItem("mpr_account") === currentUser.account) { localStorage.setItem("mpr_password", pwdForm.new); }
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
      } catch (error) {
        alert("讀取 Excel 失敗，請確認檔案格式是否正確。");
        return;
      }
    }
    if (!extractedContent.trim()) { alert("未偵測到任何內容，請手動輸入或選擇檔案。"); return; }

    setIsImportModalOpen(false);
    const today = new Date();
    const twYear = today.getFullYear() - 1911;
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    
    setEditingProject({
      id: Date.now(),
      title: "【從匯入自動建立】請修改標題",
      refNo: `MPR-${twYear}-${mm}-${String(projects.length + 1).padStart(3, "0")}`,
      applyDate: today.toISOString().split("T")[0],
      createTime: getCurrentTimeString(),
      purpose: "以下為系統匯入之原始資料，請人工整理：\n\n" + extractedContent,
      price: "",
      startDate: `${selectedYear}-01-01`,
      endDate: `${selectedYear}-01-31`,
      content: "",
      precautions: "",
      highlights: "",
      breakdown: { price: "", net: "", items: [] },
      countersign: [],
      status: "countersigning",
      feedback: "",
      creator: `${currentUser?.dept || ""} - ${currentUser?.name || ""}`,
    });
    setModalMode("create");
    setIsModalOpen(true);
    setImportText("");
    setImportFile(null);
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
      const start = parseInt(p.startDate.split("-")[0], 10);
      const end = parseInt(p.endDate.split("-")[0], 10);
      if (!isNaN(start) && start < minYear) minYear = start;
      if (!isNaN(end) && end > maxYear) maxYear = end;
    });
    const options = []; for (let i = minYear; i <= maxYear; i++) options.push(i);
    return options;
  }, [projects, currentYear]);

  // 將 dbEvents 作為生成月曆的依賴
  const calendarData = useMemo(() => generateCalendar(selectedYear, dbEvents), [selectedYear, dbEvents]);
  const [selectedMonthView, setSelectedMonthView] = useState<number | null>(null);

  const yearProjects = useMemo(() => projects.filter((p) => p.startDate.startsWith(String(selectedYear)) || p.endDate.startsWith(String(selectedYear))), [projects, selectedYear]);
  const scheduledProjects = useMemo(() => yearProjects.filter((p) => p.status === "scheduled"), [yearProjects]);
  
  const filteredScheduledProjects = useMemo(() => {
    return scheduledProjects.filter((p) => {
      // 🛡️ 安全部門過濾機制 (加回 p.creator && 保護)
      let passDept = ganttDeptFilter === "all" || (p.creator && p.creator.includes(ganttDeptFilter));
      let passMonth = true;
      if (ganttMonthFilter !== "all") {
        const filterStart = new Date(selectedYear, parseInt(ganttMonthFilter) - 1, 1);
        const filterEnd = new Date(selectedYear, parseInt(ganttMonthFilter), 0);
        passMonth = new Date(p.startDate) <= filterEnd && new Date(p.endDate) >= filterStart;
      }
      return passDept && passMonth;
    });
  }, [scheduledProjects, ganttDeptFilter, ganttMonthFilter, selectedYear]);

  const dashboardActiveProjects = useMemo(() => {
    return scheduledProjects.filter((p) => {
      // 🛡️ 保護機制
      let passDept = dashboardDeptFilter === "all" || (p.creator && p.creator.includes(dashboardDeptFilter)) || (p.countersign && p.countersign.some((c: any) => c.dept === dashboardDeptFilter));
      let passMonth = true;
      if (dashboardMonthFilter !== "all") {
        const filterStart = new Date(selectedYear, parseInt(dashboardMonthFilter) - 1, 1);
        const filterEnd = new Date(selectedYear, parseInt(dashboardMonthFilter), 0);
        passMonth = new Date(p.startDate) <= filterEnd && new Date(p.endDate) >= filterStart;
      }
      return passDept && passMonth;
    });
  }, [scheduledProjects, dashboardDeptFilter, dashboardMonthFilter, selectedYear]);

  const myPendingCountersign = useMemo(() => {
    if (!currentUser || ["admin", "gm"].includes(currentUser.role)) return [];
    return yearProjects.filter((p) => p.status === "countersigning" && p.countersign && p.countersign.some((c: any) => c.dept === currentUser.dept && c.status === "pending"));
  }, [yearProjects, currentUser]);
  
  const myOwnProposals = useMemo(() => {
    if (!currentUser || ["admin", "gm"].includes(currentUser.role)) return [];
    return yearProjects.filter((p) => p.creator && p.creator.includes(currentUser.name) && ["countersigning", "revision", "unconfirmed"].includes(p.status));
  }, [yearProjects, currentUser]);
  
  const managerPendingApproval = useMemo(() => {
    if (!currentUser || !["admin", "gm"].includes(currentUser.role)) return [];
    return yearProjects.filter((p) => p.status === "unconfirmed");
  }, [yearProjects, currentUser]);

  const getBarStyles = (startDate: string, endDate: string) => {
    const yearStart = new Date(`${selectedYear}-01-01`).getTime();
    const yearEnd = new Date(`${selectedYear}-12-31`).getTime();
    const totalYearTime = yearEnd - yearStart;
    const start = new Date(startDate).getTime();
    const end = new Date(endDate).getTime();
    const validStart = Math.max(yearStart, start);
    const validEnd = Math.min(yearEnd, end);
    let left = ((validStart - yearStart) / totalYearTime) * 100;
    let width = ((validEnd - validStart) / totalYearTime) * 100;
    if (width < 1) width = 1;
    return { left: `${left}%`, width: `${width}%` };
  };

  const handleLoginSubmit = (e: any) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const acc = formData.get("account"); const pass = formData.get("password");
    const user = users.find((u) => u.account === acc && u.password === pass);
    if (user) {
      if (rememberMe) { localStorage.setItem("mpr_account", acc as string); localStorage.setItem("mpr_password", pass as string); localStorage.setItem("mpr_remember", "true"); } 
      else { localStorage.removeItem("mpr_account"); localStorage.removeItem("mpr_password"); localStorage.setItem("mpr_remember", "false"); }
      setCurrentUser(user); setView("app");
    } else { 
      alert("登入失敗：帳號或密碼錯誤！(請先確認資料庫連線或資料庫內有無使用者)"); 
    }
  };
  const handleGuestLogin = () => { setCurrentUser({ id: "guest", name: "訪客", role: "guest", dept: "" }); setView("app"); };
  const handleLogout = () => { setCurrentUser(null); setView("login"); };

  const handleOpenCreate = () => {
    if (currentUser?.role === "guest") return;
    const today = new Date(); const twYear = today.getFullYear() - 1911; const mm = String(today.getMonth() + 1).padStart(2, "0");
    setEditingProject({
      id: Date.now(), title: "", refNo: `MPR-${twYear}-${mm}-${String(projects.length + 1).padStart(3, "0")}`, applyDate: today.toISOString().split("T")[0], createTime: getCurrentTimeString(), purpose: "", price: "", startDate: `${selectedYear}-01-01`, endDate: `${selectedYear}-01-31`, content: "", precautions: "", highlights: "", breakdown: { price: "", net: "", items: [] }, countersign: [], status: "countersigning", feedback: "", creator: `${currentUser.dept} - ${currentUser.name}`,
    });
    setModalMode("create"); setIsModalOpen(true);
  };

  const handleSave = async (e: any) => { 
    e.preventDefault(); 
    let updatedProj = { ...editingProject }; 
    if (modalMode === "create" && updatedProj.countersign.length === 0) updatedProj.status = "revision"; 
    await saveProjectToDb(updatedProj); 
    setIsModalOpen(false); 
  };

  const handleToggleDept = (dept: string) => {
    const current = editingProject.countersign || [];
    if (current.some((c: any) => c.dept === dept)) {
      setEditingProject({ ...editingProject, countersign: current.filter((c: any) => c.dept !== dept) });
    } else {
      setEditingProject({ ...editingProject, countersign: [...current, { dept, status: "pending", comment: "", time: "" }] });
    }
  };

  const submitDeptComment = async (deptName: string, comment: string) => {
    const updatedCountersign = editingProject.countersign.map((c: any) => {
      if (c.dept === deptName) return { ...c, status: "approved", comment, time: getCurrentTimeString() };
      return c;
    });
    let nextStatus = editingProject.status;
    if (updatedCountersign.every((c: any) => c.status === "approved")) nextStatus = "revision";
    const updatedProj = { ...editingProject, countersign: updatedCountersign, status: nextStatus };
    await saveProjectToDb(updatedProj);
  };

  const submitRevisionToManager = async () => { const updatedProj = { ...editingProject, status: "unconfirmed" }; await saveProjectToDb(updatedProj); setIsModalOpen(false); };
  const approveByManager = async () => { const updatedProj = { ...editingProject, status: "scheduled", feedback: "" }; await saveProjectToDb(updatedProj); setIsModalOpen(false); };
  const rejectByManager = async () => { const updatedProj = { ...editingProject, status: "revision" }; await saveProjectToDb(updatedProj); setIsModalOpen(false); };

  const calculateNetPrice = (currentBreakdown: any) => {
    const totalPriceStr = currentBreakdown.price || "0";
    const totalPrice = parseFloat(totalPriceStr.replace(/,/g, "")) || 0;
    let totalDeductions = 0;
    (currentBreakdown.items || []).forEach((item: any) => { totalDeductions += evaluateExpression(item.value); });
    const netPrice = totalPrice - totalDeductions;
    return new Intl.NumberFormat("en-US").format(netPrice);
  };

  const handleBreakdownPriceChange = (value: string) => {
    const newBreakdown = { ...editingProject.breakdown, price: value };
    newBreakdown.net = calculateNetPrice(newBreakdown);
    setEditingProject({ ...editingProject, breakdown: newBreakdown });
  };

  const handleTogglePreset = (presetName: string) => {
    let newItems = [...(editingProject.breakdown?.items || [])];
    if (newItems.some((i) => i.name === presetName)) newItems = newItems.filter((i) => i.name !== presetName);
    else newItems.push({ name: presetName, value: "" });
    const newBreakdown = { ...editingProject.breakdown, items: newItems };
    newBreakdown.net = calculateNetPrice(newBreakdown);
    setEditingProject({ ...editingProject, breakdown: newBreakdown });
  };

  const handleAddBreakdownItem = () => {
    const newBreakdown = { ...editingProject.breakdown, items: [...(editingProject.breakdown.items || []), { name: "新項目", value: "" }] };
    newBreakdown.net = calculateNetPrice(newBreakdown);
    setEditingProject({ ...editingProject, breakdown: newBreakdown });
  };

  const handleRemoveBreakdownItem = (index: number) => {
    const newItems = [...editingProject.breakdown.items]; newItems.splice(index, 1);
    const newBreakdown = { ...editingProject.breakdown, items: newItems };
    newBreakdown.net = calculateNetPrice(newBreakdown);
    setEditingProject({ ...editingProject, breakdown: newBreakdown });
  };

  const handleBreakdownItemChange = (index: number, field: string, value: string) => {
    const newItems = [...editingProject.breakdown.items]; newItems[index][field] = value;
    const newBreakdown = { ...editingProject.breakdown, items: newItems };
    newBreakdown.net = calculateNetPrice(newBreakdown);
    setEditingProject({ ...editingProject, breakdown: newBreakdown });
  };

  const handleExportSystemPDF = () => {
    setIsPrintLayoutActive(true); setIsExportModalOpen(false);
    setTimeout(() => { window.print(); setIsPrintLayoutActive(false); }, 800);
  };

  const exportSingleProjectToWord = (project: any) => {
    const header = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>簽呈匯出</title><style>body{font-family:"Microsoft JhengHei",Arial,sans-serif;}table{border-collapse:collapse;width:100%;margin-bottom:20px;}th,td{border:1px solid black;padding:8px;text-align:left;vertical-align:top;}.center{text-align:center;}.no-border{border:none;}.no-border td{border:none;padding:4px 0;} .comments { color: black; font-weight: bold; background-color: #f9fafb; padding: 10px; border: 1px solid #ccc; }</style></head><body>`;
    const formatText = (t: string) => String(t || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br/>");
    const itemsHeader = (project.breakdown?.items || []).map((i: any) => `<th class="center">${formatText(i.name)}</th>`).join("");
    const itemsData = (project.breakdown?.items || []).map((i: any) => `<td class="center">${formatText(i.value)}</td>`).join("");
    const commentsHtml = (project.countersign || []).map((c: any) => `<div>[${c.dept}] ${c.status === "approved" ? c.time + " - " + formatText(c.comment || "無意見") : "待確認..."}</div>`).join("");
    const htmlContent = `<h2 class="center">行銷公關部 簽呈 Official application</h2><table><tr><th class="center" width="20%">簽核</th><th class="center" width="20%">經辦人</th><th class="center" width="20%">部門主管</th><th class="center" width="20%">營運主管</th><th class="center" width="20%">總經理</th></tr><tr><td height="60"></td><td></td><td></td><td></td><td></td></tr></table><table class="no-border"><tr><td class="no-border">Date 日期：${formatText(project.applyDate)}</td></tr><tr><td class="no-border">Ref No 文檔號：${formatText(project.refNo)}</td></tr></table><table><tr><th width="15%">主旨</th><td>${formatText(project.title)}</td></tr><tr><th>說明</th><td>${formatText(project.purpose)}</td></tr><tr><th>活動售價</th><td>${formatText(project.price)}</td></tr><tr><th>活動日期</th><td>${formatText(project.startDate)} ～ ${formatText(project.endDate)}</td></tr><tr><th>內容說明</th><td>${formatText(project.content)}</td></tr><tr><th>注意事項</th><td>${formatText(project.precautions)}</td></tr></table><h3>內拆表</h3><table><tr><th class="center">售價</th>${itemsHeader}<th class="center">淨價</th></tr><tr><td class="center">${formatText(project.breakdown?.price)}</td>${itemsData}<td class="center"><b>${formatText(project.breakdown?.net)}</b></td></tr></table><table><tr><th width="15%">專案亮點</th><td>${formatText(project.highlights)}</td></tr><tr><th>會簽單位</th><td>${formatText((project.countersign || []).map((c: any) => c.dept).join("、 "))}</td></tr></table>${commentsHtml ? `<h3>會簽單位意見備註</h3><div class="comments">${commentsHtml}</div>` : ""}`;
    const blob = new Blob(["\ufeff", header + htmlContent + "</body></html>"], { type: "application/msword" });
    const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = `簽呈_${project.title.replace(/[\\/:*?"<>|]/g, "")}.doc`; document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(url);
  };

  const handleExportListWord = () => {
    let exportList = projects.filter((p) => {
      const startY = parseInt(p.startDate.split("-")[0]); const endY = parseInt(p.endDate.split("-")[0]);
      if (startY > exportConfig.year || endY < exportConfig.year) return false;
      if (exportConfig.month !== "all") {
        const startM = parseInt(p.startDate.split("-")[1]); const endM = parseInt(p.endDate.split("-")[1]); const targetM = parseInt(exportConfig.month);
        return targetM >= startM && targetM <= endM;
      }
      return true;
    });
    const header = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>專案清單匯出</title><style>body{font-family:"Microsoft JhengHei",Arial,sans-serif;}table{border-collapse:collapse;width:100%;margin-bottom:20px;}th,td{border:1px solid black;padding:8px;text-align:left;vertical-align:top;}th{background-color:#f2f2f2;}</style></head><body>`;
    const titleText = `${exportConfig.year} 年度 ${exportConfig.month === "all" ? "全年" : exportConfig.month + "月"} 專案清單`;
    let htmlContent = `<h2>${titleText}</h2><table><tr><th>專案名稱</th><th>活動日期</th><th>專案狀態</th><th>提案人</th></tr>`;
    if (exportList.length === 0) htmlContent += `<tr><td colspan="4">該區間尚無任何專案。</td></tr>`;
    else exportList.forEach((p) => { htmlContent += `<tr><td>${p.title}</td><td>${p.startDate} ~ ${p.endDate}</td><td>${p.status}</td><td>${p.creator}</td></tr>`; });
    htmlContent += `</table>`;
    const blob = new Blob(["\ufeff", header + htmlContent + "</body></html>"], { type: "application/msword" });
    const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = `${titleText}.doc`; document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(url); setIsExportModalOpen(false);
  };

  const WorkflowProgressBar = ({ project }: any) => {
    const isCountersignDone = project.status !== "countersigning" && project.status !== "start";
    const isRevisionDone = ["unconfirmed", "scheduled"].includes(project.status);
    const isApproved = project.status === "scheduled";
    const pendingDepts = (project.countersign || []).filter((c: any) => c.status === "pending").map((c: any) => c.dept);
    let countersignDesc = "無須會簽";
    if (project.countersign?.length > 0) {
      if (!isCountersignDone) countersignDesc = pendingDepts.length > 0 ? `待確認: ${pendingDepts.join("、")}` : "會簽完成，待提案人送交";
      else countersignDesc = "會簽已全數完成";
    }
    const steps = [
      { id: "start", label: "提案建立", desc: `${project.creator?.split("-")[0] || project.creator}\n${project.createTime || project.applyDate}`, done: true, active: false },
      { id: "countersigning", label: "會簽單位確認中", desc: countersignDesc, done: isCountersignDone, active: project.status === "countersigning" },
      { id: "revision", label: "版本確認", desc: isRevisionDone ? "已送交" : "提案人修改中", done: isRevisionDone, active: project.status === "revision" },
      { id: "unconfirmed", label: "審核決議中", desc: isApproved ? "已核准" : "等待主管核准", done: isApproved, active: project.status === "unconfirmed" },
      { id: "scheduled", label: "已排程", desc: "正式發佈", done: isApproved, active: isApproved },
    ];
    return (
      <div className="flex items-start justify-between w-full mb-10 mt-2 relative no-print px-4">
        <div className="absolute top-4 left-[10%] w-[80%] h-1.5 bg-gray-200 -z-10 rounded"></div>
        <div className="absolute top-4 left-[10%] h-1.5 bg-indigo-500 -z-10 transition-all duration-500 rounded" style={{ width: `${(steps.filter((s) => s.done).length / (steps.length - 1)) * 80}%` }}></div>
        {steps.map((step, idx) => (
          <div key={step.id} className="flex flex-col items-center flex-1 z-10 relative group">
            <div className={`w-9 h-9 rounded-full flex items-center justify-center font-bold text-sm shadow-md transition-all duration-300 ${step.done ? "bg-indigo-600 text-white" : step.active ? "bg-orange-500 text-white ring-4 ring-orange-100 scale-110" : "bg-gray-100 text-gray-400 border border-gray-300"}`}>
              {step.done ? <Check className="w-5 h-5" /> : idx + 1}
            </div>
            <div className={`mt-3 text-xs font-bold text-center ${step.done || step.active ? "text-gray-800" : "text-gray-400"}`}>{step.label}</div>
            <div className={`text-[10px] mt-1 whitespace-pre-wrap text-center ${step.done || step.active ? "text-indigo-600 font-medium" : "text-gray-400"}`}>{step.desc}</div>
          </div>
        ))}
      </div>
    );
  };

  const MonthCalendarView = ({ month }: any) => {
    const daysInMonth = new Date(selectedYear, month, 0).getDate();
    const firstDay = new Date(selectedYear, month - 1, 1).getDay();
    const days = Array.from({ length: daysInMonth }, (_, i) => i + 1);
    const blanks = Array.from({ length: firstDay }, (_, i) => i);

    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm print:hidden">
        <div className="bg-white rounded-xl shadow-xl w-full max-w-6xl overflow-hidden flex flex-col max-h-[95vh]">
          <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center bg-indigo-50">
            <h2 className="text-xl font-bold text-indigo-900 flex items-center gap-2"><CalendarDays className="w-6 h-6" /> {selectedYear} 年 {month} 月 行事曆與專案排程</h2>
            <button onClick={() => setSelectedMonthView(null)} className="text-gray-500 hover:text-gray-800"><X className="w-6 h-6" /></button>
          </div>
          <div className="p-6 overflow-y-auto bg-gray-50 flex-1">
            <div className="grid grid-cols-7 gap-2 mb-2">
              {["日", "一", "二", "三", "四", "五", "六"].map((d) => (<div key={d} className="text-center font-bold text-sm text-gray-500">{d}</div>))}
            </div>
            <div className="grid grid-cols-7 gap-2">
              {blanks.map((b) => (<div key={`blank-${b}`} className="bg-transparent p-2 rounded-lg"></div>))}
              {days.map((day) => {
                const dateStr = `${selectedYear}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
                const dailyData = calendarData[dateStr] || { type: "平日", events: [], marketingEvents: [] };
                const st = CAL_STYLES[dailyData.type] || CAL_STYLES["平日"];
                const dayProjects = yearProjects.filter((p) => p.startDate <= dateStr && p.endDate >= dateStr);
                const totalItems = (dailyData.events?.length || 0) + (dailyData.marketingEvents?.length || 0) + dayProjects.length;
                return (
                  <div key={day} className={`min-h-[70px] md:min-h-[110px] border rounded-lg p-1 md:p-2 flex flex-col ${st.bg} ${st.border} cursor-pointer hover:ring-2 hover:ring-indigo-400 transition-all`} onClick={() => setSelectedDayInfo({ date: dateStr, dailyData, dayProjects, st })}>
                    <div className="flex flex-col lg:flex-row justify-between items-center lg:items-start mb-1 gap-1 lg:gap-0">
                      <span className={`font-bold text-sm ${st.text}`}>{day}</span>
                      <span className={`text-[9px] md:text-[10px] px-1 rounded shadow-sm ${st.tag} whitespace-nowrap`}>{dailyData.type}</span>
                    </div>
                    {totalItems > 0 && (<div className="md:hidden mt-auto text-[10px] text-center font-bold text-indigo-600 bg-white border border-indigo-200 rounded py-0.5 shadow-sm">{totalItems} 備註</div>)}
                    <div className="hidden md:block flex-1 space-y-1 mt-1 overflow-y-auto">
                      {dailyData.events && dailyData.events.map((ev: any, idx: number) => (<div key={`ev-${idx}`} className="text-[10px] bg-yellow-100 text-yellow-800 border border-yellow-200 px-1 py-0.5 rounded mb-1 truncate" title={ev}>{ev}</div>))}
                      {dailyData.marketingEvents && dailyData.marketingEvents.map((ev: any, idx: number) => (<div key={`mk-ev-${idx}`} className="text-[10px] bg-green-50 text-green-700 border border-green-200 px-1 py-0.5 rounded mb-1 truncate" title={ev}>{ev}</div>))}
                      {dayProjects.map((p) => (<div key={p.id} className={`text-[10px] px-1.5 py-0.5 rounded truncate text-white shadow-sm ${p.status === "scheduled" ? "bg-indigo-500" : "bg-orange-400"}`} title={p.title}>{p.title}</div>))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const DashboardActiveProjectsCard = () => (
    <div className="bg-white rounded-xl shadow-sm border border-indigo-200 p-5 lg:col-span-2 flex flex-col h-full min-h-[350px]">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4 border-b border-gray-100 pb-3">
        <div className="flex items-center gap-3 text-indigo-600"><PlayCircle className="w-6 h-6" /><h3 className="font-bold text-lg">目前進行專案 (已排程)</h3></div>
        <div className="flex items-center gap-2">
          <div className="flex items-center bg-gray-50 rounded border border-gray-200 px-2 py-1">
            <Filter className="w-3 h-3 text-gray-500 mr-1" />
            <select className="bg-transparent text-sm outline-none text-gray-700 font-medium cursor-pointer" value={dashboardDeptFilter} onChange={(e) => setDashboardDeptFilter(e.target.value)}>
              <option value="all">全館所有部門</option>
              {DEPARTMENTS.map((d) => (<option key={d} value={d}>{d}</option>))}
            </select>
          </div>
          <div className="flex items-center bg-gray-50 rounded border border-gray-200 px-2 py-1">
            <Calendar className="w-3 h-3 text-gray-500 mr-1" />
            <select className="bg-transparent text-sm outline-none text-gray-700 font-medium cursor-pointer" value={dashboardMonthFilter} onChange={(e) => setDashboardMonthFilter(e.target.value)}>
              <option value="all">全年</option>
              {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((m) => (<option key={m} value={m}>{m}月</option>))}
            </select>
          </div>
        </div>
      </div>
      <p className="text-3xl font-black mb-4 text-indigo-700">{dashboardActiveProjects.length} <span className="text-sm text-gray-400 font-medium">個專案</span></p>
      <div className="space-y-2 flex-1 overflow-y-auto pr-2">
        {dashboardActiveProjects.length === 0 && (<div className="flex flex-col items-center justify-center h-full text-gray-400"><CheckCircle className="w-8 h-8 mb-2 opacity-20" /><p className="text-sm font-medium">該篩選條件下尚無進行中專案</p></div>)}
        {dashboardActiveProjects.map((p) => (
          <div key={p.id} onClick={() => { setEditingProject({ ...p }); setModalMode("view"); setIsModalOpen(true); }} className="text-sm p-3 bg-indigo-50 border border-indigo-100 rounded-lg cursor-pointer hover:bg-indigo-100 flex flex-col sm:flex-row sm:justify-between sm:items-center text-indigo-800 transition-colors shadow-sm gap-2 sm:gap-0">
            <span className="font-bold truncate">{p.title}</span>
            <span className="text-xs text-indigo-600 shrink-0 bg-white px-2 py-1 rounded border border-indigo-200 font-medium">{p.startDate.substring(5).replace("-", "/")} ~ {p.endDate.substring(5).replace("-", "/")}</span>
          </div>
        ))}
      </div>
    </div>
  );

  if (isLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <div className="w-8 h-8 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mb-4"></div>
        <div className="text-indigo-600 font-bold">系統與資料庫連線中...</div>
      </div>
    );
  }

  if (view === "login") {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
        <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden">
          <div className="bg-indigo-600 p-8 text-center text-white">
            <Layout className="w-12 h-12 mx-auto mb-4 opacity-90" />
            <h1 className="text-2xl font-bold tracking-wide">專案管理系統</h1>
          </div>
          <div className="p-8">
            <form onSubmit={handleLoginSubmit} className="space-y-5">
              <div><label className="block text-sm font-medium mb-1">登入帳號</label><div className="relative"><UserCircle className="w-5 h-5 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" /><input name="account" type="text" required defaultValue={savedAccount} className="w-full pl-10 p-2 border rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none" /></div></div>
              <div><label className="block text-sm font-medium mb-1">登入密碼</label><div className="relative"><Lock className="w-5 h-5 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" /><input name="password" type="password" required defaultValue={savedPassword} className="w-full pl-10 p-2 border rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none" /></div></div>
              <div className="flex items-center"><input id="rem" type="checkbox" className="mr-2 h-4 w-4 text-indigo-600" checked={rememberMe} onChange={(e) => setRememberMe(e.target.checked)} /><label htmlFor="rem" className="text-sm font-medium text-gray-700">記住密碼</label></div>
              <button type="submit" className="w-full bg-indigo-600 hover:bg-indigo-700 transition-colors text-white font-bold py-2.5 rounded-lg mt-4 shadow">登入系統</button>
            </form>
            <div className="mt-6 text-center border-t pt-4"><button onClick={handleGuestLogin} className="text-sm text-gray-500 hover:text-indigo-600">以訪客身分預覽</button></div>
          </div>
        </div>
      </div>
    );
  }

  if (view === "users") {
    return (
      <div className="min-h-screen bg-slate-50 text-slate-800">
        <header className="bg-white shadow-sm sticky top-0 z-20">
          <div className="max-w-5xl mx-auto px-4 h-16 flex items-center justify-between">
            <div className="flex items-center gap-2 text-indigo-600"><ShieldCheck className="w-6 h-6" /><h1 className="font-bold text-xl tracking-wide">系統後台 - 帳號密碼管理</h1></div>
            <div className="flex items-center gap-2">
              <span className="text-sm font-medium text-gray-700 max-w-[80px] sm:max-w-none truncate">{currentUser?.name}</span>
              <button onClick={() => setIsChangePwdModalOpen(true)} className="text-gray-400 hover:text-blue-600 hover:bg-blue-50 p-1.5 sm:p-2 rounded-lg transition-colors" title="變更密碼"><Key className="w-4 h-4 sm:w-5 sm:h-5" /></button>
              <button onClick={() => setView("app")} className="text-gray-600 hover:bg-gray-100 px-4 py-2 rounded-lg text-sm font-medium transition-colors ml-2 border border-gray-200">返回主系統</button>
            </div>
          </div>
        </header>
        <main className="max-w-5xl mx-auto px-4 py-8">
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            <div className="p-4 border-b border-gray-200 bg-gray-50 flex justify-between items-center">
              <h2 className="font-bold text-gray-800">使用者列表</h2>
              <button onClick={() => { setEditingUser({ account: "", password: "", name: "", role: "employee", dept: "企劃部" }); setIsUserModalOpen(true); }} className="bg-indigo-600 text-white px-4 py-2 rounded-lg text-sm flex items-center gap-1 shadow-sm hover:bg-indigo-700"><Plus className="w-4 h-4" /> 新增帳號</button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse min-w-[600px]">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200 text-gray-600 text-sm">
                    <th className="p-4">姓名</th><th className="p-4">所屬部門</th><th className="p-4">登入帳號</th><th className="p-4">登入密碼</th><th className="p-4">系統權限</th><th className="p-4 text-center">操作</th>
                  </tr>
                </thead>
                <tbody>
                  {users.map((u) => (
                    <tr key={u.id} className="border-b border-gray-100 hover:bg-gray-50">
                      <td className="p-4 font-medium">{u.name}</td><td className="p-4">{u.dept}</td><td className="p-4">{u.account}</td><td className="p-4 text-gray-400">••••</td>
                      <td className="p-4"><span className={`px-2 py-1 rounded text-xs font-bold ${ROLES_INFO[u.role]?.bg || "bg-gray-100"} ${ROLES_INFO[u.role]?.color || "text-gray-700"}`}>{ROLES_INFO[u.role]?.name || u.role}</span></td>
                      <td className="p-4 flex justify-center gap-2">
                        {u.account !== "admin" ? (
                          <><button onClick={() => { setEditingUser({ ...u }); setIsUserModalOpen(true); }} className="text-indigo-600 hover:bg-indigo-50 p-2 rounded"><Edit className="w-4 h-4" /></button>
                          <button onClick={() => handleDeleteUser(u.id)} className="text-red-600 hover:bg-red-50 p-2 rounded"><Trash2 className="w-4 h-4" /></button></>
                        ) : (<span className="text-xs font-bold text-gray-400 px-2 py-2">預設管理員</span>)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </main>

        {isUserModalOpen && editingUser && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <form onSubmit={handleSaveUser} className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden">
              <div className="px-6 py-4 border-b flex justify-between"><h2 className="font-bold">{editingUser.id ? "編輯帳號" : "新增帳號"}</h2><button type="button" onClick={() => setIsUserModalOpen(false)}><X className="w-5 h-5" /></button></div>
              <div className="p-6 space-y-4 text-sm">
                <div><label className="block mb-1 font-medium">姓名</label><input type="text" required className="w-full border rounded p-2 outline-none focus:ring-2 focus:ring-indigo-500" value={editingUser.name} onChange={(e) => setEditingUser({ ...editingUser, name: e.target.value })} /></div>
                <div><label className="block mb-1 font-medium">所屬部門</label><select className="w-full border rounded p-2 outline-none focus:ring-2 focus:ring-indigo-500" value={editingUser.dept} onChange={(e) => setEditingUser({ ...editingUser, dept: e.target.value })}>{DEPARTMENTS.map((d) => (<option key={d} value={d}>{d}</option>))}<option value="管理層">管理層</option><option value="系統管理">系統管理</option></select></div>
                <div><label className="block mb-1 font-medium">登入帳號</label><input type="text" required className="w-full border rounded p-2 outline-none focus:ring-2 focus:ring-indigo-500" value={editingUser.account} onChange={(e) => setEditingUser({ ...editingUser, account: e.target.value })} /></div>
                <div><label className="block mb-1 font-medium">登入密碼</label><input type="text" required className="w-full border rounded p-2 outline-none focus:ring-2 focus:ring-indigo-500" value={editingUser.password} onChange={(e) => setEditingUser({ ...editingUser, password: e.target.value })} /></div>
                <div><label className="block mb-1 font-medium">系統權限</label><select className="w-full border rounded p-2 outline-none focus:ring-2 focus:ring-indigo-500" value={editingUser.role} onChange={(e) => setEditingUser({ ...editingUser, role: e.target.value })}><option value="employee">一般員工/各部門 (可提案/會簽)</option><option value="gm">總經理 (觀看/審核簽呈)</option><option value="admin">系統管理員 (所有權限)</option></select></div>
              </div>
              <div className="px-6 py-4 bg-gray-50 flex justify-end gap-2"><button type="button" onClick={() => setIsUserModalOpen(false)} className="px-4 py-2 text-gray-600">取消</button><button type="submit" className="px-4 py-2 bg-indigo-600 text-white rounded">儲存帳號</button></div>
            </form>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className={`min-h-screen font-sans ${isPrintLayoutActive ? "bg-white" : "bg-slate-50 text-slate-800"}`}>
      
      {/* 列印預覽模式 */}
      {isPrintLayoutActive && (
        <div className="bg-white min-h-screen pb-12 font-sans" id="pdf-export-area">
          <div className="fixed top-0 left-0 w-full bg-slate-800 text-white p-3 flex justify-between items-center z-[100] print:hidden shadow-md">
            <div className="font-medium flex items-center gap-2"><Printer className="w-5 h-5" /> 列印預覽模式 </div>
            <div className="flex gap-3">
              <button onClick={() => window.print()} className="bg-indigo-500 hover:bg-indigo-600 px-4 py-1.5 rounded font-medium text-sm shadow">列印 (或 Ctrl+P)</button>
              <button onClick={() => setIsPrintLayoutActive(false)} className="bg-slate-600 hover:bg-slate-700 px-4 py-1.5 rounded font-medium text-sm">退出預覽</button>
            </div>
          </div>
          <div className="p-4 sm:p-8 max-w-5xl mx-auto mt-16 print:mt-0 print:p-0">
            <h1 className="text-2xl sm:text-3xl font-bold text-center mb-8">{exportConfig.year} 年度 {exportConfig.month === "all" ? "全年" : `${exportConfig.month}月`} 專案清單</h1>
            <table className="w-full border-collapse border border-gray-400 mb-8 text-sm">
              <thead><tr className="bg-gray-100"><th className="border border-gray-400 p-2 text-left">專案名稱</th><th className="border border-gray-400 p-2 text-center w-48">活動期間</th><th className="border border-gray-400 p-2 text-center w-32">狀態</th><th className="border border-gray-400 p-2 text-center w-32">提案人</th></tr></thead>
              <tbody>
                {projects.filter((p) => {
                  const startY = parseInt(p.startDate.split("-")[0]); const endY = parseInt(p.endDate.split("-")[0]);
                  if (startY > exportConfig.year || endY < exportConfig.year) return false;
                  if (exportConfig.month !== "all") { const startM = parseInt(p.startDate.split("-")[1]); const endM = parseInt(p.endDate.split("-")[1]); const targetM = parseInt(exportConfig.month); return targetM >= startM && targetM <= endM; }
                  return true;
                }).map((p) => (
                  <tr key={p.id}>
                    <td className="border border-gray-400 p-2 font-medium">{p.title}</td><td className="border border-gray-400 p-2 text-center">{p.startDate} ~ {p.endDate}</td>
                    <td className="border border-gray-400 p-2 text-center">{p.status === "scheduled" ? "已排程" : p.status === "unconfirmed" ? "主管審核中" : p.status === "revision" ? "提案人版本確認中" : "單位會簽中"}</td>
                    <td className="border border-gray-400 p-2 text-center">{p.creator?.split("-")[0] || ""}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* 一般主介面 */}
      {!isPrintLayoutActive && (
        <>
          <header className="bg-white shadow-sm sticky top-0 z-20 print:hidden">
            <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
              <div className="flex items-center gap-1 sm:gap-2 text-indigo-600">
                <Layout className="w-5 h-5 sm:w-6 sm:h-6" />
                <h1 className="font-bold text-lg sm:text-xl tracking-wide hidden md:block">專案管理</h1>
                <select className="ml-1 sm:ml-2 bg-indigo-50 border border-indigo-200 text-indigo-800 text-xs sm:text-sm font-bold rounded-lg focus:ring-indigo-500 p-1.5 cursor-pointer outline-none" value={selectedYear} onChange={(e) => setSelectedYear(Number(e.target.value))}>
                  {yearOptions.map((y) => (<option key={y} value={y}>{y} 年</option>))}
                </select>
              </div>
              <div className="flex items-center gap-1 sm:gap-3">
                <button onClick={() => setIsImportModalOpen(true)} className="text-white bg-green-600 hover:bg-green-700 px-2 sm:px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-1 transition-colors shadow-sm">
                  <UploadCloud className="w-4 h-4" /> <span className="hidden lg:inline">匯入資料</span>
                </button>
                <div className="h-4 sm:h-6 w-px bg-gray-200 mx-1 sm:mx-1"></div>
                <button onClick={() => setIsExportModalOpen(true)} className="text-gray-600 hover:text-indigo-600 hover:bg-indigo-50 px-2 sm:px-3 py-2 rounded-lg text-sm font-medium flex items-center gap-1">
                  <DownloadCloud className="w-4 h-4" /> <span className="hidden lg:inline">匯出清單</span>
                </button>
                {currentUser?.role === "admin" && (
                  <button onClick={() => setView("users")} className="text-indigo-600 bg-indigo-50 hover:bg-indigo-100 px-2 sm:px-3 py-2 rounded-lg text-sm font-medium flex items-center gap-1 border border-indigo-100 transition-colors">
                    <Settings className="w-4 h-4" /> <span className="hidden lg:inline">後台管理</span>
                  </button>
                )}
                <div className="h-4 sm:h-6 w-px bg-gray-200 mx-1 sm:mx-2"></div>
                <div className="flex items-center gap-1 sm:gap-2 px-1">
                  <span className={`hidden md:inline px-2 py-0.5 rounded text-xs font-bold ${ROLES_INFO[currentUser?.role || "guest"]?.bg} ${ROLES_INFO[currentUser?.role || "guest"]?.color}`}>{currentUser?.dept || "訪客"}</span>
                  <span className="text-sm font-medium text-gray-700 max-w-[80px] sm:max-w-none truncate">{currentUser?.name}</span>
                </div>
                {currentUser && currentUser.role !== "guest" && (<button onClick={() => setIsChangePwdModalOpen(true)} className="text-gray-400 hover:text-blue-600 hover:bg-blue-50 p-1.5 sm:p-2 rounded-lg transition-colors" title="變更密碼"><Key className="w-4 h-4 sm:w-5 sm:h-5" /></button>)}
                <button onClick={handleLogout} className="text-gray-400 hover:text-red-600 hover:bg-red-50 p-1.5 sm:p-2 rounded-lg transition-colors" title="登出"><LogOut className="w-4 h-4 sm:w-5 sm:h-5" /></button>
                {currentUser?.role !== "guest" && (
                  <button onClick={() => handleOpenCreate()} className="bg-indigo-600 hover:bg-indigo-700 text-white px-2 sm:px-4 py-1.5 sm:py-2 ml-1 sm:ml-2 rounded-lg text-sm font-medium flex items-center gap-1 sm:gap-2 shadow-sm transition-colors">
                    <Plus className="w-4 h-4" /> <span className="hidden md:inline">新增簽呈</span>
                  </button>
                )}
              </div>
            </div>
          </header>

          <main className="max-w-7xl mx-auto px-4 py-8 space-y-6">
            {currentUser && currentUser.role !== "guest" && (
              <section className="mb-8">
                <div className="flex items-center justify-between mb-4"><h2 className="text-2xl font-bold text-gray-800 flex items-center gap-2">👋 歡迎回來，{currentUser.name} <span className="text-sm text-gray-500 font-medium ml-2">({currentUser.dept} 專屬主控板)</span></h2></div>
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  {currentUser.role === "employee" && (
                    <><div className="lg:col-span-1 flex flex-col gap-6">
                        <div className="bg-white rounded-xl shadow-sm border border-orange-200 p-5 flex-1 flex flex-col min-h-[200px]">
                          <div className="flex items-center gap-3 mb-4 text-orange-600"><AlertCircle className="w-6 h-6" /><h3 className="font-bold text-lg">待我會辦的專案</h3></div>
                          <p className="text-3xl font-black mb-4">{myPendingCountersign.length}</p>
                          <div className="space-y-2 flex-1 overflow-y-auto pr-2">
                            {myPendingCountersign.length === 0 && (<p className="text-sm text-gray-400">目前無待會辦事項</p>)}
                            {myPendingCountersign.map((p) => (<div key={p.id} onClick={() => { setEditingProject({ ...p }); setModalMode("view"); setIsModalOpen(true); }} className="text-sm p-3 bg-orange-50 border border-orange-100 rounded cursor-pointer hover:bg-orange-100 truncate text-orange-800 shadow-sm font-medium">{p.title}</div>))}
                          </div>
                        </div>
                        <div className="bg-white rounded-xl shadow-sm border border-blue-200 p-5 flex-1 flex flex-col min-h-[200px]">
                          <div className="flex items-center gap-3 mb-4 text-blue-600"><PenTool className="w-6 h-6" /><h3 className="font-bold text-lg">我的提案 (待處理)</h3></div>
                          <p className="text-3xl font-black mb-4">{myOwnProposals.length}</p>
                          <div className="space-y-2 flex-1 overflow-y-auto pr-2">
                            {myOwnProposals.length === 0 && (<p className="text-sm text-gray-400">尚無進行中的提案</p>)}
                            {myOwnProposals.map((p) => (
                              <div key={p.id} onClick={() => { setEditingProject({ ...p }); setModalMode("view"); setIsModalOpen(true); }} className="text-sm p-2 bg-blue-50 border border-blue-100 rounded cursor-pointer hover:bg-blue-100 text-blue-800 shadow-sm flex flex-col gap-1">
                                <span className="font-bold truncate">{p.title}</span>
                                <span className={`self-start font-bold shrink-0 px-2 py-0.5 rounded text-[10px] ${p.status === "revision" ? "bg-red-500 text-white" : "bg-white text-blue-600 border border-blue-200"}`}>{p.status === "revision" ? "等待我確認修改" : p.status === "unconfirmed" ? "送交主管審核中" : "單位會簽中"}</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                      <DashboardActiveProjectsCard />
                    </>
                  )}
                  {["admin", "gm"].includes(currentUser.role) && (
                    <><div className="lg:col-span-1 flex flex-col gap-6">
                        <div className="bg-white rounded-xl shadow-sm border border-orange-200 p-5 relative overflow-hidden flex-1 flex flex-col min-h-[300px]">
                          <div className="absolute top-0 right-0 w-2 h-full bg-orange-500"></div>
                          <div className="flex items-center gap-3 mb-4 text-orange-600"><ClipboardList className="w-6 h-6" /><h3 className="font-bold text-lg">待審核簽呈 (決議中)</h3></div>
                          <p className="text-3xl font-black mb-4 text-orange-600">{managerPendingApproval.length} <span className="text-sm text-gray-400 font-medium">件</span></p>
                          <div className="space-y-2 flex-1 overflow-y-auto pr-2">
                            {managerPendingApproval.length === 0 && (<p className="text-sm text-gray-400">目前無待審核事項</p>)}
                            {managerPendingApproval.map((p) => (<div key={p.id} onClick={() => { setEditingProject({ ...p }); setModalMode("view"); setIsModalOpen(true); }} className="text-sm p-3 bg-orange-50 border border-orange-100 rounded cursor-pointer hover:bg-orange-100 truncate text-orange-800 font-bold shadow-sm">{p.title}</div>))}
                          </div>
                        </div>
                      </div>
                      <DashboardActiveProjectsCard />
                    </>
                  )}
                </div>
              </section>
            )}

            <div className="bg-white p-3 rounded-lg shadow-sm border border-gray-200 flex flex-wrap gap-4 text-sm items-center">
              <span className="font-bold text-gray-600">行事曆圖例：</span>
              <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-white border border-gray-300"></span>平日</span>
              <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-blue-100 border border-blue-300"></span>旺日</span>
              <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-pink-100 border border-pink-300"></span>假日</span>
              <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-red-100 border border-red-300"></span>大假日 (除夕、春節、連假)</span>
            </div>

            {/* 甘特圖區塊 */}
            <section className={`bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden`}>
              <div className="p-4 border-b border-gray-200 bg-gray-50 flex flex-col md:flex-row justify-between md:items-center gap-4">
                <h2 className="text-lg font-bold flex items-center gap-2 text-gray-800"><Calendar className="w-5 h-5 text-indigo-600" /> {selectedYear} 年度專案甘特圖 (全館)</h2>
                <div className="flex items-center gap-3">
                  <div className="flex items-center gap-2"><Filter className="w-4 h-4 text-gray-500" /><span className="text-sm font-bold text-gray-700">部門:</span><select className="border border-gray-300 rounded p-1 text-sm outline-none focus:border-indigo-500 bg-white" value={ganttDeptFilter} onChange={(e) => setGanttDeptFilter(e.target.value)}><option value="all">全部</option>{DEPARTMENTS.map((d) => (<option key={d} value={d}>{d}</option>))}</select></div>
                  <div className="flex items-center gap-2"><span className="text-sm font-bold text-gray-700">月份:</span><select className="border border-gray-300 rounded p-1 text-sm outline-none focus:border-indigo-500 bg-white" value={ganttMonthFilter} onChange={(e) => setGanttMonthFilter(e.target.value)}><option value="all">全年</option>{[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((m) => (<option key={m} value={m}>{m}月</option>))}</select></div>
                </div>
              </div>
              <div className="overflow-x-auto">
                <div className="min-w-[800px]">
                  <div className="flex border-b border-gray-200 bg-gray-50">
                    <div className="w-48 flex-shrink-0 p-3 border-r border-gray-200 font-semibold text-gray-600 text-sm">專案名稱</div>
                    <div className="flex-1 grid grid-cols-12">
                      {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((m) => (<div key={m} onClick={() => setSelectedMonthView(m)} className={`border-r border-gray-200 p-2 text-center text-xs font-medium cursor-pointer hover:bg-indigo-100 transition-colors text-gray-700`} title="點擊查看本月詳細平旺日與活動">{m}月</div>))}
                    </div>
                  </div>
                  <div className="relative">
                    <div className="absolute inset-0 flex ml-48 pointer-events-none">
                      {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((m) => (<div key={`bg-${m}`} className="flex-1 border-r border-gray-100"></div>))}
                    </div>
                    <div className="relative z-10">
                      {filteredScheduledProjects.length === 0 && (<div className="py-12 text-center text-gray-400 font-medium bg-gray-50/50">此條件下尚未有符合的專案排程</div>)}
                      {filteredScheduledProjects.map((p) => (
                        <div key={p.id} className="flex items-center group cursor-pointer border-b border-gray-100 hover:bg-gray-50 transition-colors" onClick={() => { setEditingProject({ ...p }); setModalMode("view"); setIsModalOpen(true); }}>
                          <div className="w-48 flex-shrink-0 p-3 pr-4 border-r border-gray-200"><div className="font-medium text-sm text-gray-800 truncate">{p.title}</div><div className="text-xs text-gray-400 mt-1">{p.startDate.substring(5)} ~ {p.endDate.substring(5)}</div></div>
                          <div className="flex-1 relative h-12"><div className={`absolute top-3 bottom-3 rounded-md shadow-sm transition-all bg-indigo-500 hover:bg-indigo-600`} style={getBarStyles(p.startDate, p.endDate)}><div className="px-2 text-xs text-white leading-6 truncate drop-shadow-sm">{p.title}</div></div></div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </section>
          </main>
        </>
      )}

      {/* 匯入資料 Modal */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-lg overflow-hidden">
            <div className="px-6 py-4 border-b flex justify-between bg-green-50">
              <h2 className="text-lg font-bold text-green-900 flex items-center gap-2"><UploadCloud className="w-5 h-5" /> 匯入外部資料</h2>
              <button onClick={() => setIsImportModalOpen(false)}><X className="w-5 h-5 text-gray-500 hover:text-gray-800" /></button>
            </div>
            <div className="p-6 space-y-4">
              <p className="text-sm text-gray-600 mb-4">系統可讀取 <b className="text-green-600">Excel (.xlsx, .csv)</b> 檔案，或請直接貼上文件文字。<br/>讀取後會自動幫您填入「新增簽呈」的企劃目的欄位中。</p>
              <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center bg-gray-50 hover:bg-gray-100 transition-colors">
                <input type="file" accept=".xlsx, .xls, .csv" className="hidden" id="excel-upload" onChange={(e) => setImportFile(e.target.files?.[0] || null)} />
                <label htmlFor="excel-upload" className="cursor-pointer flex flex-col items-center justify-center">
                  <FileText className={`w-10 h-10 mb-2 ${importFile ? 'text-green-500' : 'text-gray-400'}`} />
                  <span className="font-medium text-indigo-600 hover:text-indigo-800">{importFile ? `已選擇檔案：${importFile.name}` : "點擊上傳 Excel 檔案"}</span>
                </label>
              </div>
              <div className="flex items-center gap-4 my-2"><div className="h-px bg-gray-200 flex-1"></div><span className="text-xs text-gray-400 font-bold">或者貼上純文字</span><div className="h-px bg-gray-200 flex-1"></div></div>
              <textarea className="w-full border rounded-lg p-3 text-sm focus:ring-2 focus:ring-green-500 outline-none h-32" placeholder="請在此貼上您的專案內容文字..." value={importText} onChange={(e) => setImportText(e.target.value)} disabled={importFile !== null}></textarea>
              {importFile && <p className="text-xs text-orange-500 font-bold">* 已選擇檔案，文字框暫時停用。欲輸入文字請先清除檔案。</p>}
            </div>
            <div className="px-6 py-4 bg-gray-50 flex justify-between items-center border-t border-gray-100">
              <button onClick={() => { setImportFile(null); setImportText(""); }} className="text-gray-500 hover:text-red-500 text-sm font-medium">清除重設</button>
              <div className="flex gap-2">
                <button onClick={() => setIsImportModalOpen(false)} className="px-4 py-2 text-gray-600 bg-white border rounded-lg font-medium text-sm">取消</button>
                <button onClick={handleProcessImport} className="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-lg font-bold text-sm shadow-sm transition-colors">解析並建立簽呈</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 單日詳細資訊與備註彈出視窗 */}
      {selectedDayInfo && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm print:hidden">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden border">
            <div className={`px-4 py-3 border-b flex justify-between items-center ${selectedDayInfo.st.bg} ${selectedDayInfo.st.text}`}>
              <h3 className="font-bold flex items-center gap-2"><CalendarDays className="w-5 h-5" />{selectedDayInfo.date}</h3>
              <button onClick={() => setSelectedDayInfo(null)} className="hover:opacity-70"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-5 space-y-4">
              <div><span className={`text-xs px-2 py-1 rounded shadow-sm ${selectedDayInfo.st.tag} font-bold`}>{selectedDayInfo.dailyData.type}</span></div>
              
              {selectedDayInfo.dailyData.events?.length > 0 && (
                <div>
                  <h4 className="text-xs font-bold text-gray-500 mb-1 border-b pb-1">節慶與官方活動</h4>
                  <ul className="space-y-1">
                    {selectedDayInfo.dailyData.events.map((ev: string, idx: number) => (<li key={idx} className="text-sm bg-yellow-50 text-yellow-800 border border-yellow-200 px-2 py-1 rounded">{ev}</li>))}
                  </ul>
                </div>
              )}
              {selectedDayInfo.dailyData.marketingEvents?.length > 0 && (
                <div>
                  <h4 className="text-xs font-bold text-gray-500 mb-1 border-b pb-1">行銷節慶</h4>
                  <ul className="space-y-1">
                    {selectedDayInfo.dailyData.marketingEvents.map((ev: string, idx: number) => (<li key={idx} className="text-sm bg-green-50 text-green-700 border border-green-200 px-2 py-1 rounded">{ev}</li>))}
                  </ul>
                </div>
              )}
              {selectedDayInfo.dayProjects?.length > 0 && (
                <div>
                  <h4 className="text-xs font-bold text-gray-500 mb-1 border-b pb-1">館內專案排程</h4>
                  <ul className="space-y-1">
                    {selectedDayInfo.dayProjects.map((p: any) => (<li key={p.id} className={`text-sm px-2 py-1 rounded text-white shadow-sm cursor-pointer hover:opacity-80 transition-opacity ${p.status === "scheduled" ? "bg-indigo-500" : "bg-orange-400"}`} onClick={() => { setEditingProject({ ...p }); setModalMode("view"); setSelectedDayInfo(null); setIsModalOpen(true); }}>{p.title}</li>))}
                  </ul>
                </div>
              )}
              
              {(!selectedDayInfo.dailyData.events?.length) && (!selectedDayInfo.dailyData.marketingEvents?.length) && (!selectedDayInfo.dayProjects?.length) && (
                 <div className="text-center text-gray-400 text-sm py-2">本日無任何活動或專案排程</div>
              )}

              {/* 日程特殊備註與權限管理 */}
              <div className="mt-4 pt-4 border-t border-gray-200">
                <h4 className="text-sm font-bold text-indigo-800 flex items-center gap-1 mb-2">
                  <MessageSquare className="w-4 h-4" /> 日程特殊備註
                </h4>
                {currentUser?.role !== "admin" && (
                  <div className="bg-gray-50 p-3 rounded-lg border border-gray-200 text-sm text-gray-700 whitespace-pre-wrap min-h-[60px]">
                    {dbEvents.find(e => e.date === selectedDayInfo.date)?.description || "本日無特殊備註。"}
                  </div>
                )}
                {currentUser?.role === "admin" && (
                  <div className="flex flex-col gap-2">
                    <textarea
                      id="day-remark-input"
                      className="w-full border border-indigo-200 bg-indigo-50/30 rounded-lg p-3 text-sm focus:ring-2 focus:ring-indigo-500 outline-none transition-all"
                      rows={3}
                      placeholder="管理員專用：請輸入本日備註 (如交通管制、滿房提醒)..."
                      defaultValue={dbEvents.find(e => e.date === selectedDayInfo.date)?.description || ""}
                    ></textarea>
                    <button
                      onClick={() => handleSaveDayRemark(selectedDayInfo.date, (document.getElementById("day-remark-input") as HTMLTextAreaElement).value)}
                      className="self-end bg-indigo-600 text-white px-4 py-1.5 rounded text-sm font-bold hover:bg-indigo-700 shadow-sm"
                    >
                      儲存至資料庫
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {selectedMonthView && <MonthCalendarView month={selectedMonthView} />}

      {/* 匯出 Modal */}
      {isExportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm print:hidden">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden">
            <div className="px-6 py-4 border-b flex justify-between bg-indigo-50">
              <h2 className="text-lg font-bold text-indigo-900 flex items-center gap-2"><FileDown className="w-5 h-5" /> 匯出專案清單</h2>
              <button onClick={() => setIsExportModalOpen(false)}><X className="w-5 h-5" /></button>
            </div>
            <div className="p-6 space-y-4 text-sm">
              <p className="text-gray-500 mb-2">請選擇匯出範圍：</p>
              <div><label className="block font-bold mb-1">匯出年度</label><select className="w-full border rounded-lg p-2 outline-none focus:border-indigo-500" value={exportConfig.year} onChange={(e) => setExportConfig({ ...exportConfig, year: Number(e.target.value) })}>{yearOptions.map((y) => (<option key={y} value={y}>{y} 年</option>))}</select></div>
              <div><label className="block font-bold mb-1">匯出月份</label><select className="w-full border rounded-lg p-2 outline-none focus:border-indigo-500" value={exportConfig.month} onChange={(e) => setExportConfig({ ...exportConfig, month: e.target.value })}><option value="all">全年度</option>{[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12].map((m) => (<option key={m} value={m}>{m} 月</option>))}</select></div>
            </div>
            <div className="px-6 py-4 bg-gray-50 flex justify-between items-center border-t border-gray-100">
              <button onClick={handleExportListWord} className="flex items-center gap-1 px-4 py-2 bg-blue-50 text-blue-700 hover:bg-blue-100 rounded-lg font-medium text-sm transition-colors border border-blue-200"><FileType2 className="w-4 h-4" /> 下載 Word</button>
              <button onClick={handleExportSystemPDF} className="flex items-center gap-1 px-4 py-2 bg-indigo-600 text-white hover:bg-indigo-700 rounded-lg font-medium text-sm transition-colors"><Printer className="w-4 h-4" /> 列印 (PDF)</button>
            </div>
          </div>
        </div>
      )}

      {/* 變更密碼 Modal */}
      {isChangePwdModalOpen && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm print:hidden">
          <form onSubmit={handleChangePassword} className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden">
            <div className="px-6 py-4 border-b flex justify-between bg-indigo-50">
              <h2 className="font-bold text-indigo-900 flex items-center gap-2"><Key className="w-4 h-4" /> 變更密碼</h2>
              <button type="button" onClick={() => { setIsChangePwdModalOpen(false); setPwdForm({ old: "", new: "", confirm: "" }); }}><X className="w-5 h-5" /></button>
            </div>
            <div className="p-6 space-y-4 text-sm">
              <div><label className="block mb-1 font-medium">原密碼</label><input type="password" required className="w-full border rounded p-2 focus:ring-2 focus:ring-indigo-500" value={pwdForm.old} onChange={(e) => setPwdForm({ ...pwdForm, old: e.target.value })} /></div>
              <div><label className="block mb-1 font-medium">新密碼</label><input type="password" required className="w-full border rounded p-2 focus:ring-2 focus:ring-indigo-500" value={pwdForm.new} onChange={(e) => setPwdForm({ ...pwdForm, new: e.target.value })} /></div>
              <div><label className="block mb-1 font-medium">確認新密碼</label><input type="password" required className="w-full border rounded p-2 focus:ring-2 focus:ring-indigo-500" value={pwdForm.confirm} onChange={(e) => setPwdForm({ ...pwdForm, confirm: e.target.value })} /></div>
            </div>
            <div className="px-6 py-4 bg-gray-50 flex justify-end gap-2"><button type="button" onClick={() => { setIsChangePwdModalOpen(false); setPwdForm({ old: "", new: "", confirm: "" }); }} className="px-4 py-2 text-gray-600">取消</button><button type="submit" className="px-4 py-2 bg-indigo-600 text-white rounded">確認變更</button></div>
          </form>
        </div>
      )}

      {/* 專案 Modal (檢視 / 編輯 / 新增) */}
      <style>{`
        .sign-table { width: 100%; border-collapse: collapse; margin-bottom: 1.5rem; border: 1px solid #d1d5db; }
        .sign-table th, .sign-table td { border: 1px solid #d1d5db; padding: 0.75rem; text-align: left; vertical-align: top; }
        .sign-table th { background-color: #f3f4f6; color: #374151; font-weight: 600; }
        .sign-table .center { text-align: center; }
        @media print {
          body * { visibility: hidden; }
          .print-modal, .print-modal * { visibility: visible; }
          .print-modal { position: absolute; left: 0; top: 0; width: 100%; border: none !important; box-shadow: none !important; }
          .no-print { display: none !important; }
          .sign-table th, .sign-table td { border: 1px solid #000; color: #000; }
          .sign-table th { background-color: #e5e7eb !important; -webkit-print-color-adjust: exact; }
          .print-system-only { visibility: visible !important; }
        }
      `}</style>

      {isModalOpen && editingProject && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm print:bg-white print:p-0">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-4xl flex flex-col max-h-[95vh] print-modal print:max-h-none print:h-auto overflow-hidden print:overflow-visible relative" id="pdf-export-area">
            <div className="px-6 py-4 border-b border-gray-100 flex justify-between items-center bg-gray-50 no-print">
              <h2 className="text-xl font-bold text-gray-800">
                {modalMode === "create" ? "新增簽呈" : modalMode === "edit" ? "編輯簽呈" : "專案簽呈內容"}
              </h2>
              <div className="flex items-center gap-2">
                {modalMode === "view" && (
                  <>
                    <button onClick={handleExportSystemPDF} className="flex items-center gap-1 text-sm bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded"><Printer className="w-4 h-4" /> 列印 PDF</button>
                    <button onClick={() => exportSingleProjectToWord(editingProject)} className="flex items-center gap-1 text-sm bg-blue-50 text-blue-700 px-3 py-1.5 rounded"><FileType2 className="w-4 h-4" /> 下載 Word</button>
                  </>
                )}
                <button onClick={() => setIsModalOpen(false)} className="text-gray-400 hover:text-gray-600 ml-2"><X className="w-6 h-6" /></button>
              </div>
            </div>

            <div className="p-8 overflow-y-auto print:p-4 text-sm print:text-base text-gray-900 bg-white">
              {modalMode === "view" ? (
                <div className="max-w-3xl mx-auto">
                  <WorkflowProgressBar project={editingProject} />
                  <h1 className="text-2xl font-bold text-center mb-6">行銷公關部 簽呈 Official application</h1>
                  <table className="sign-table center">
                    <thead><tr><th className="center w-1/5">簽核</th><th className="center w-1/5">經辦人</th><th className="center w-1/5">部門主管</th><th className="center w-1/5">營運主管</th><th className="center w-1/5">總經理</th></tr></thead>
                    <tbody><tr><td className="h-16"></td><td></td><td></td><td></td><td></td></tr></tbody>
                  </table>
                  <div className="mb-4 font-medium text-gray-700 print:text-black flex justify-between">
                    <span>Date：{editingProject.applyDate}</span>
                    <span>Ref No：{editingProject.refNo}</span>
                  </div>
                  <table className="sign-table">
                    <tbody>
                      <tr><th className="w-32 center">主旨</th><td className="font-bold text-lg">{editingProject.title}</td></tr>
                      <tr><th className="center">說明</th><td className="whitespace-pre-wrap">{editingProject.purpose}</td></tr>
                      <tr><th className="center">活動售價</th><td>{editingProject.price}</td></tr>
                      <tr><th className="center">活動日期</th><td>{editingProject.startDate} ～ {editingProject.endDate}</td></tr>
                      <tr><th className="center">內容說明</th><td className="whitespace-pre-wrap">{editingProject.content}</td></tr>
                      <tr><th className="center">注意事項</th><td className="whitespace-pre-wrap">{editingProject.precautions}</td></tr>
                    </tbody>
                  </table>
                  <h3 className="font-bold mb-2 text-lg">內拆表</h3>
                  <table className="sign-table center">
                    <thead>
                      <tr>
                        <th className="center">售價</th>
                        {(editingProject.breakdown?.items || []).map((item: any, idx: number) => (<th key={idx} className="center">{item.name}</th>))}
                        <th className="center">淨價</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="font-medium">{editingProject.breakdown?.price}</td>
                        {(editingProject.breakdown?.items || []).map((item: any, idx: number) => (<td key={idx}>{item.value}</td>))}
                        <td className="font-bold text-indigo-700 print:text-black">{editingProject.breakdown?.net}</td>
                      </tr>
                    </tbody>
                  </table>
                  <div className="mt-6 mb-8">
                    <h3 className="font-bold text-lg text-red-600 mb-2">專案亮點</h3>
                    <p className="bg-red-50 p-4 rounded border font-medium whitespace-pre-wrap">{editingProject.highlights}</p>
                  </div>
                  <div className="mt-8">
                    <p className="font-bold mb-2 text-lg">擬辦：奉 核後，函知各相關部門後續作業</p>
                    <table className="sign-table">
                      <tbody><tr><th className="w-32 center">會簽單位</th><td className="font-medium">{(editingProject.countersign || []).map((c: any) => c.dept).join("、 ") || "無須會簽"}</td></tr></tbody>
                    </table>
                  </div>

                  {editingProject.countersign?.length > 0 && (
                    <div className="mt-8 no-print border border-gray-300 rounded-lg overflow-hidden shadow-sm">
                      <div className="bg-gray-100 px-4 py-2 font-bold text-gray-800 border-b flex items-center gap-2"><PenTool className="w-4 h-4" /> 會簽意見</div>
                      <div className="p-4 bg-gray-50 space-y-4">
                        {editingProject.countersign.map((c: any) => (
                          <div key={c.dept} className="flex flex-col border-b pb-3 last:border-0 last:pb-0">
                            <div className="flex items-center gap-2 mb-1">
                              <span className="font-bold text-gray-900 bg-white px-2 py-0.5 rounded border">{c.dept}</span>
                              {c.status === "approved" ? (<span className="text-xs text-green-600 font-bold flex items-center gap-1"><CheckCircle className="w-3 h-3" /> 已確認 ({c.time})</span>) : (<span className="text-xs text-orange-600 font-bold">待確認...</span>)}
                            </div>
                            {c.status === "approved" ? (
                              <div className="text-black font-bold whitespace-pre-wrap pl-1 border-l-4 border-gray-400 ml-1 mt-1 p-2 bg-white rounded">{c.comment || "無意見。"}</div>
                            ) : (
                              editingProject.status === "countersigning" && currentUser?.dept === c.dept && (
                                <div className="flex gap-2 mt-2">
                                  <input type="text" id={`comment-${c.dept}`} className="flex-1 border rounded px-3 py-1.5 text-sm" placeholder="請填寫會簽意見" />
                                  <button onClick={() => submitDeptComment(c.dept, (document.getElementById(`comment-${c.dept}`) as HTMLInputElement).value)} className="bg-indigo-600 text-white px-4 py-1.5 rounded text-sm font-bold hover:bg-indigo-700">送出確認</button>
                                </div>
                              )
                            )}
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {editingProject.status === "unconfirmed" && editingProject.feedback && (
                    <div className="bg-orange-50 p-4 rounded-lg border border-orange-200 mt-6 no-print">
                      <div className="flex gap-2 text-orange-800 font-bold mb-1"><MessageSquare className="w-5 h-5" /> 主管退回意見：</div>
                      <p className="text-orange-700 font-bold border-l-4 border-orange-400 pl-2 ml-1">{editingProject.feedback}</p>
                    </div>
                  )}
                  {editingProject.status === "revision" && editingProject.feedback && (
                    <div className="bg-orange-50 p-4 rounded-lg border border-orange-200 mt-6 no-print">
                      <div className="flex gap-2 text-orange-800 font-bold mb-1"><MessageSquare className="w-5 h-5" /> 主管退回意見：</div>
                      <p className="text-orange-700 font-bold border-l-4 border-orange-400 pl-2 ml-1">{editingProject.feedback}</p>
                    </div>
                  )}

                  <div className="flex gap-3 pt-6 border-t mt-6 no-print justify-end">
                    {editingProject.status === "revision" && editingProject.creator?.includes(currentUser?.name) && (
                      <button onClick={submitRevisionToManager} className="bg-indigo-600 text-white px-6 py-2 rounded-lg font-bold flex items-center gap-2">修改完成，送交審核 <ArrowRight className="w-4 h-4" /></button>
                    )}
                    {["admin", "gm"].includes(currentUser?.role) && editingProject.status === "unconfirmed" && (
                      <><button onClick={approveByManager} className="bg-green-600 text-white px-6 py-2 rounded-lg font-bold">核准排程</button>
                      <button onClick={rejectByManager} className="bg-orange-500 text-white px-6 py-2 rounded-lg font-bold">退回修改</button></>
                    )}
                    {(["admin", "gm"].includes(currentUser?.role) || (currentUser?.role === "employee" && editingProject.creator?.includes(currentUser?.name))) && editingProject.status !== "scheduled" && (
                      <button onClick={() => setModalMode("edit")} className="flex items-center gap-2 bg-indigo-50 text-indigo-700 px-6 py-2 rounded-lg font-medium"><Edit className="w-4 h-4" /> 編輯</button>
                    )}
                  </div>
                </div>
              ) : (
                <form id="project-form" onSubmit={handleSave} className="space-y-8 max-w-4xl mx-auto bg-white">
                  <div className="space-y-4">
                    <h3 className="font-bold text-lg border-b pb-2 text-indigo-800">一、基本資料</h3>
                    <div className="grid grid-cols-2 gap-4">
                      <div><label className="block font-medium mb-1">文檔號</label><input type="text" className="w-full border rounded-lg p-2.5 bg-gray-50" value={editingProject.refNo} onChange={(e) => setEditingProject({ ...editingProject, refNo: e.target.value })} /></div>
                      <div><label className="block font-medium mb-1">申請日期</label><input type="date" className="w-full border rounded-lg p-2.5 bg-gray-50" value={editingProject.applyDate} onChange={(e) => setEditingProject({ ...editingProject, applyDate: e.target.value })} /></div>
                    </div>
                  </div>
                  <div className="space-y-4">
                    <h3 className="font-bold text-lg border-b pb-2 text-indigo-800">二、活動內容</h3>
                    <div><label className="block font-medium mb-1">專案名稱</label><input type="text" required className="w-full border rounded-lg p-2.5 focus:ring-2 focus:ring-indigo-500" value={editingProject.title} onChange={(e) => setEditingProject({ ...editingProject, title: e.target.value })} /></div>
                    <div><label className="block font-medium mb-1">企劃目的</label><textarea rows={2} className="w-full border rounded-lg p-2.5 focus:ring-2 focus:ring-indigo-500" value={editingProject.purpose} onChange={(e) => setEditingProject({ ...editingProject, purpose: e.target.value })}></textarea></div>
                    <div className="grid grid-cols-3 gap-4">
                      <div><label className="block font-medium mb-1">活動售價</label><input type="text" className="w-full border rounded-lg p-2.5" placeholder="例: 每房 NT$5,188" value={editingProject.price} onChange={(e) => setEditingProject({ ...editingProject, price: e.target.value })} /></div>
                      <div><label className="block font-medium mb-1">開始日期</label><input type="date" required className="w-full border rounded-lg p-2.5" value={editingProject.startDate} onChange={(e) => setEditingProject({ ...editingProject, startDate: e.target.value })} /></div>
                      <div><label className="block font-medium mb-1">結束日期</label><input type="date" required className="w-full border rounded-lg p-2.5" value={editingProject.endDate} onChange={(e) => setEditingProject({ ...editingProject, endDate: e.target.value })} /></div>
                    </div>
                    <div><label className="block font-medium mb-1">內容說明</label><textarea rows={4} className="w-full border rounded-lg p-2.5" value={editingProject.content} onChange={(e) => setEditingProject({ ...editingProject, content: e.target.value })}></textarea></div>
                    <div><label className="block font-medium mb-1">注意事項</label><textarea rows={3} className="w-full border rounded-lg p-2.5" value={editingProject.precautions} onChange={(e) => setEditingProject({ ...editingProject, precautions: e.target.value })}></textarea></div>
                    <div><label className="block font-medium text-red-600 mb-1">專案亮點</label><textarea rows={2} className="w-full border-red-200 rounded-lg p-2.5 bg-red-50 focus:ring-red-500" value={editingProject.highlights} onChange={(e) => setEditingProject({ ...editingProject, highlights: e.target.value })}></textarea></div>
                  </div>
                  <div className="space-y-4">
                    <div className="flex justify-between items-center border-b pb-2">
                      <h3 className="font-bold text-lg text-indigo-800">三、財務內拆表</h3>
                      <button type="button" onClick={handleAddBreakdownItem} className="text-sm flex items-center gap-1 bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-lg"><Plus className="w-4 h-4" /> 自訂項目</button>
                    </div>
                    <div className="bg-slate-50 p-3 rounded-lg border">
                      <label className="block font-bold mb-2">快速勾選：</label>
                      <div className="flex flex-wrap gap-3">
                        {PRESET_BREAKDOWN_ITEMS.map((preset) => (
                          <label key={preset} className="flex items-center gap-2 bg-white border px-3 py-1.5 rounded cursor-pointer">
                            <input type="checkbox" checked={(editingProject.breakdown?.items || []).some((i: any) => i.name === preset)} onChange={() => handleTogglePreset(preset)} className="w-4 h-4 text-indigo-600" />
                            <span className="text-sm font-medium">{preset}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                    <div className="overflow-x-auto pb-2">
                      <table className="w-full border-collapse border min-w-[600px] text-sm">
                        <thead>
                          <tr className="bg-indigo-50">
                            <th className="border p-3 w-32 font-bold">總售價</th>
                            {(editingProject.breakdown?.items || []).map((item: any, idx: number) => (
                              <th key={idx} className="border p-2 relative group min-w-[120px]">
                                <input type="text" className="w-full bg-transparent border-b border-indigo-300 focus:border-indigo-600 outline-none text-center font-bold text-indigo-800" value={item.name} onChange={(e) => handleBreakdownItemChange(idx, "name", e.target.value)} placeholder="名稱" />
                                <button type="button" onClick={() => handleRemoveBreakdownItem(idx)} className="absolute top-1 right-1 bg-red-100 text-red-600 rounded-full p-1 opacity-0 group-hover:opacity-100"><Trash2 className="w-3 h-3" /></button>
                              </th>
                            ))}
                            <th className="border p-3 text-indigo-700 w-32 font-bold">客房淨價</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td className="border p-2"><input type="text" className="w-full border rounded p-2 text-center" placeholder="總金額" value={editingProject.breakdown?.price || ""} onChange={(e) => handleBreakdownPriceChange(e.target.value)} /></td>
                            {(editingProject.breakdown?.items || []).map((item: any, idx: number) => (
                              <td key={idx} className="border p-2"><input type="text" className="w-full border rounded p-2 text-center" value={item.value} onChange={(e) => handleBreakdownItemChange(idx, "value", e.target.value)} placeholder="金額或算式" /></td>
                            ))}
                            <td className="border p-2 bg-indigo-50/50"><div className="w-full border-2 border-indigo-400 bg-white font-bold text-indigo-800 rounded p-2 text-center min-h-[36px] flex items-center justify-center">{editingProject.breakdown?.net || "0"}</div></td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                  <div className="space-y-4">
                    <h3 className="font-bold text-lg border-b pb-2 text-indigo-800">四、會簽與設定</h3>
                    <div>
                      <label className="block font-medium mb-2">需會簽之部門</label>
                      <div className="flex flex-wrap gap-3">
                        {DEPARTMENTS.map((dept) => (
                          <label key={dept} className="flex items-center gap-2 bg-gray-50 border px-3 py-2 rounded-lg cursor-pointer">
                            <input type="checkbox" checked={(editingProject.countersign || []).some((c: any) => c.dept === dept)} onChange={() => handleToggleDept(dept)} className="w-4 h-4 text-indigo-600" />
                            <span className="text-sm font-medium">{dept}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                  </div>
                </form>
              )}
            </div>
            {modalMode !== "view" && (
              <div className="px-6 py-4 border-t bg-gray-50 flex justify-end gap-3 no-print">
                <button type="button" onClick={() => setIsModalOpen(false)} className="px-6 py-2 text-gray-600 bg-white border rounded-lg font-medium">取消</button>
                <button type="submit" form="project-form" className="bg-indigo-600 text-white px-8 py-2 rounded-lg font-medium">儲存</button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

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
    bigHolidays: ["01-01", "01-02", "01-03", "02-04", "02-05", "02-06", "02-07", "02-08", "02-09", "02-10", "02-27", "02-28", "03-01", "04-03", "04-04", "04-05", "04-06", "04-30", "05-01", "05-02", "06-09", "09-15", "10-09", "10-10", "10-11", "10-23", "10-24", "10-25", "12-24", "12-25", "12-26"],
    holidays: ["09-28"], 
    events: {
      "01-01": "元旦", "02-04": "小年夜", "02-05": "除夕", "02-06": "春節 (初一)", "02-07": "春節 (初二)", "02-08": "春節 (初三)", "02-09": "春節補假", "02-10": "春節補假",
      "02-28": "和平紀念日", "04-04": "兒童節", "04-05": "清明節", "05-01": "勞動節", "06-09": "端午節", "09-15": "中秋節", "09-28": "教師節", "10-10": "國慶日", "10-25": "光復節", "12-25": "行憲紀念日"
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
      if (yearData.bigHolidays.includes(md)) { type = "大假日"; } 
      else if (yearData.holidays.includes(md)) { type = "假日"; } 
      else if (type === "平日" && year === 2026) {
        const isWinter = (m === 1 && d >= 21) || (m === 2 && d <= 13);
        const isSummer = m === 7 || m === 8;
        if (day === 5 || isWinter || isSummer) type = "旺日";
      }
      if (yearData.events[md]) events.push(`🧨 ${yearData.events[md]}`);
      const offEv = OFFICIAL_EVENTS.find(e => e.date === md); if (offEv) events.push(`✨ ${offEv.name}`);
      if (customEvents && customEvents.length > 0) {
        const match = customEvents.find(e => e.date === dateStr);
        if (match) {
          if (match.event_type) type = match.event_type;
          if (match.event_name) events.push(`${match.is_public_holiday ? '🧨' : '✨'} ${match.event_name}`);
        }
      }
      data[dateStr] = { type, events, marketingEvents };
    }
  }
  return data;
};

const evaluateExpression = (expr: string) => {
  if (!expr) return 0;
  try { return new Function(`return ${String(expr).replace(/\s+/g, "").replace(/[^-()\d/*+.]/g, "")}`)() || 0; } catch (e) { return 0; }
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

  useEffect(() => { fetchData(); }, []);

  const fetchData = async () => {
    setIsLoading(true);
    try {
      const { data: uData } = await supabase.from("users").select("*");
      if (uData && uData.length === 0) {
        const admin = { id: "u1", account: "admin", password: "123", name: "管理員", role: "admin", dept: "系統" };
        await supabase.from("users").upsert(admin); setUsers([admin]);
      } else { setUsers(uData || []); }
      const { data: pData } = await supabase.from("projects").select("*");
      if (pData) {
        const parsed = pData.map((p: any) => {
          let bd; try { bd = typeof p.breakdown === "string" ? JSON.parse(p.breakdown || "[]") : (p.breakdown || []); } catch(e){bd=[];}
          if (!Array.isArray(bd)) bd = [{ id: Date.now(), name: "主專案", price: bd.price||"", ota: "", items: bd.items||[], net: bd.net||"0" }];
          return { ...p, projectType: p.projectType || 'leisure', breakdown: bd, countersign: typeof p.countersign === "string" ? JSON.parse(p.countersign || "[]") : (p.countersign || []) };
        });
        setProjects(parsed);
      }
      const { data: cData } = await supabase.from("calendar_events").select("*");
      if (cData) setDbEvents(cData);
    } catch (e) { console.error(e); }
    setIsLoading(false);
  };

  const saveProjectToDb = async (proj: any) => {
    const payload = { ...proj, breakdown: JSON.stringify(proj.breakdown), countersign: JSON.stringify(proj.countersign), id: String(proj.id) };
    const { error } = await supabase.from("projects").upsert(payload);
    if (error) {
      alert(`⚠️ 儲存失敗！\n錯誤原因：${error.message}\n\n資料還在畫面上，請放心複製文字備份。`);
      return false;
    } else { fetchData(); return true; }
  };

  const handleLoginSubmit = (e: any) => {
    e.preventDefault(); const fd = new FormData(e.target);
    const u = users.find(x => x.account === fd.get("account") && x.password === fd.get("password"));
    if (u) {
      if (rememberMe) { localStorage.setItem("mpr_account", u.account); localStorage.setItem("mpr_password", u.password); localStorage.setItem("mpr_remember", "true"); }
      setCurrentUser(u); setView("app");
    } else { alert("帳號或密碼錯誤"); }
  };

  const handleOpenCreate = () => {
    const today = new Date(); const twY = today.getFullYear() - 1911; const mm = String(today.getMonth()+1).padStart(2,"0");
    setEditingProject({
      id: Date.now(), title: "", refNo: `MPR-${twY}-${mm}-${String(projects.length+1).padStart(3,"0")}`, projectType: "leisure", applyDate: today.toISOString().split("T")[0], createTime: getCurrentTimeString(), purpose: "", startDate: `${selectedYear}-01-01`, endDate: `${selectedYear}-01-31`, content: "", precautions: "", highlights: "", 
      breakdown: [{ id: Date.now(), name: "專案一", price: "", ota: "", items: [], net: "0" }], countersign: [], status: "countersigning", creator: `${currentUser.dept} - ${currentUser.name}`,
    });
    setModalMode("create"); setIsModalOpen(true);
  };

  const updatePackage = (idx: number, field: string, val: any) => {
    const newBd = [...editingProject.breakdown]; newBd[idx][field] = val;
    if (['price', 'ota', 'items'].includes(field)) {
      const price = parseFloat(String(newBd[idx].price).replace(/,/g, "")) || 0;
      const ota = parseFloat(String(newBd[idx].ota)) || 0;
      const otaAmt = Math.round(price * (ota / 100));
      let deduct = 0; (newBd[idx].items || []).forEach((i: any) => { deduct += evaluateExpression(i.value); });
      newBd[idx].net = new Intl.NumberFormat("en-US").format(price - otaAmt - deduct);
    }
    setEditingProject({ ...editingProject, breakdown: newBd });
  };

  const handlePrintSingle = () => { setTimeout(() => window.print(), 300); };

  const exportSingleToWord = (project: any) => {
    const formatText = (t: string) => String(t || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br/>");
    let bdHtml = "";
    (project.breakdown || []).forEach((p: any) => {
      const oVal = Math.round((parseFloat(String(p.price).replace(/,/g,""))||0) * (parseFloat(p.ota)/100));
      bdHtml += `<h4>${formatText(p.name)}</h4><table border="1" style="width:100%; border-collapse:collapse;"><tr><th>售價</th><th>OTA扣除</th><th>淨價</th></tr><tr><td align="center">${p.price}</td><td align="center">-${oVal}</td><td align="center"><b>${p.net}</b></td></tr></table>`;
    });
    const dept = project.creator?.split('-')[0]?.trim() || "飯店";
    const content = `<h2>${dept} 簽呈</h2><div>日期:${project.applyDate} 案號:${project.refNo}</div>${bdHtml}`;
    const blob = new Blob(["\ufeff", content], { type: "application/msword" });
    const url = URL.createObjectURL(blob); const a = document.createElement("a"); a.href = url; a.download = `簽呈.doc`; a.click();
  };

  return (
    <div className={`min-h-screen font-sans ${isPrintLayoutActive ? "bg-white" : "bg-slate-50"}`}>
      <style>{`
        .sign-table { width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: 10px; }
        .sign-table th, .sign-table td { border: 1px solid #000; padding: 6px; font-size: 13px; vertical-align: top; }
        .sign-table th { background-color: #f2f2f2; text-align: center; }
        @media print {
          body * { visibility: hidden; }
          .print-modal, .print-modal * { visibility: visible; }
          .print-modal { position: absolute; left: 0; top: 0; width: 100%; border: none !important; box-shadow: none !important; overflow: visible !important; }
          .no-print { display: none !important; }
          table { page-break-inside: auto; } tr { page-break-inside: avoid; }
        }
      `}</style>

      {view === "login" ? (
        <div className="flex items-center justify-center min-h-screen"><form onSubmit={handleLoginSubmit} className="p-8 bg-white rounded-xl shadow-lg w-80">
          <h1 className="mb-6 text-xl font-bold text-center">專案管理系統</h1>
          <input name="account" placeholder="帳號" required className="w-full p-2 mb-3 border rounded"/>
          <input name="password" type="password" placeholder="密碼" required className="w-full p-2 mb-4 border rounded"/>
          <button className="w-full p-2 text-white bg-indigo-600 rounded">登入</button>
        </form></div>
      ) : (
        <>
          <header className="flex items-center justify-between h-16 px-6 bg-white shadow-sm no-print">
            <h1 className="text-lg font-bold text-indigo-700">飯店專案管理系統</h1>
            <div className="flex gap-4 items-center">
               <span className="text-sm">{currentUser.dept} - {currentUser.name}</span>
               <button onClick={handleOpenCreate} className="px-4 py-1.5 text-white bg-indigo-600 rounded-lg text-sm">+ 新增簽呈</button>
               <button onClick={() => setView("login")} className="text-gray-400"><LogOut size={20}/></button>
            </div>
          </header>
          
          <main className="p-8 no-print">
             <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="p-6 bg-white rounded-xl border border-indigo-100">
                   <h2 className="mb-4 font-bold flex items-center gap-2"><ClipboardList className="text-indigo-500"/> 進行中的專案</h2>
                   <div className="space-y-2">
                      {projects.filter(p => p.status === 'scheduled').map(p => (
                        <div key={p.id} onClick={() => {setEditingProject(p); setModalMode("view"); setIsModalOpen(true);}} className="p-3 text-sm bg-indigo-50 rounded-lg cursor-pointer hover:bg-indigo-100 flex justify-between">
                           <span>{p.title}</span><span className="text-xs text-indigo-400">{p.startDate}</span>
                        </div>
                      ))}
                   </div>
                </div>
                <div className="p-6 bg-white rounded-xl border border-orange-100">
                   <h2 className="mb-4 font-bold flex items-center gap-2"><AlertCircle className="text-orange-500"/> 待核簽呈</h2>
                   <div className="space-y-2">
                      {projects.filter(p => p.status !== 'scheduled').map(p => (
                        <div key={p.id} onClick={() => {setEditingProject(p); setModalMode("view"); setIsModalOpen(true);}} className="p-3 text-sm bg-orange-50 rounded-lg cursor-pointer hover:bg-orange-100">
                           {p.title}
                        </div>
                      ))}
                   </div>
                </div>
             </div>
          </main>
        </>
      )}

      {isModalOpen && editingProject && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 print:p-0">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[95vh] overflow-y-auto print-modal">
            <div className="flex items-center justify-between p-4 border-b bg-gray-50 no-print">
               <h2 className="font-bold">{modalMode === 'view' ? '簽呈預覽' : '編輯簽呈'}</h2>
               <div className="flex gap-2">
                  {modalMode === 'view' && <button onClick={handlePrintSingle} className="p-1.5 bg-indigo-100 text-indigo-700 rounded"><Printer size={18}/></button>}
                  <button onClick={() => setIsModalOpen(false)} className="text-gray-400"><X/></button>
               </div>
            </div>

            <div className="p-8">
               <h1 className="text-2xl font-bold text-center mb-6">
                 {editingProject.creator?.split('-')[0]?.trim() || '飯店'} 簽呈 Official application
               </h1>
               
               <div className="flex justify-between mb-4 text-sm font-medium">
                  <span>Date 日期：{editingProject.applyDate}</span>
                  <span>Ref No 文檔號：{editingProject.refNo}</span>
               </div>

               <table className="sign-table">
                  <tbody>
                    <tr><th className="w-32">專案類型</th><td className="font-bold text-indigo-600">{editingProject.projectType === 'room' ? '🏨 住房專案' : '☕ 休閒/餐飲專案'}</td></tr>
                    <tr><th>主旨</th><td className="font-bold text-lg">{editingProject.title || (modalMode !== 'view' && <input className="w-full p-1 border" onChange={e => setEditingProject({...editingProject, title: e.target.value})}/>)}</td></tr>
                    {modalMode === 'view' ? (
                       <>
                         <tr><th>說明</th><td className="whitespace-pre-wrap">{editingProject.purpose}</td></tr>
                         <tr><th>日期</th><td>{editingProject.startDate} ~ {editingProject.endDate}</td></tr>
                         <tr><th>內容注意事項</th><td className="whitespace-pre-wrap">{editingProject.content} {editingProject.precautions}</td></tr>
                       </>
                    ) : (
                       <tr><th>編輯內容</th><td>請在編輯欄位填寫主旨與目的</td></tr>
                    )}
                  </tbody>
               </table>

               <h3 className="mt-6 mb-2 font-bold border-l-4 border-indigo-500 pl-2">財務內拆表</h3>
               {(editingProject.breakdown || []).map((pkg: any, idx: number) => {
                 const otaAmt = pkg.ota ? Math.round((parseFloat(String(pkg.price).replace(/,/g,""))||0) * (parseFloat(pkg.ota)/100)) : 0;
                 return (
                   <div key={idx} className="mb-6 p-4 bg-slate-50 rounded-lg border">
                      <div className="flex justify-between items-center mb-2">
                         <span className="font-bold text-indigo-800">{pkg.name}</span>
                         {modalMode !== 'view' && <button onClick={() => {const newBd = [...editingProject.breakdown]; newBd.splice(idx,1); setEditingProject({...editingProject, breakdown: newBd})}} className="text-red-500"><Trash2 size={16}/></button>}
                      </div>
                      <table className="w-full text-center border-collapse bg-white">
                         <thead><tr className="bg-gray-100 text-xs"><th>售價</th><th>OTA扣除({pkg.ota}%)</th><th>扣除項目</th><th>淨價</th></tr></thead>
                         <tbody>
                            <tr className="text-sm">
                               <td className="border p-2">{modalMode === 'view' ? pkg.price : <input className="w-full text-center" value={pkg.price} onChange={e => updatePackage(idx, 'price', e.target.value)}/>}</td>
                               <td className="border p-2 text-red-500">-{otaAmt} {modalMode !== 'view' && <input placeholder="%" className="w-10 ml-2 border" value={pkg.ota} onChange={e => updatePackage(idx, 'ota', e.target.value)}/>}</td>
                               <td className="border p-2 text-xs">{(pkg.items || []).map((i:any) => `${i.name}:${i.value}`).join(', ')}</td>
                               <td className="border p-2 font-bold text-indigo-700">{pkg.net}</td>
                            </tr>
                         </tbody>
                      </table>
                   </div>
                 );
               })}

               <div className="mt-8">
                  <h3 className="mb-2 font-bold border-l-4 border-green-500 pl-2">會簽單位意見</h3>
                  <div className="space-y-3">
                     {(editingProject.countersign || []).map((c: any, i: number) => (
                        <div key={i} className="p-3 border rounded bg-gray-50 text-sm">
                           <div className="flex justify-between mb-1 border-b pb-1 font-bold">
                              <span>{c.dept}</span>
                              <span className={c.status === 'approved' ? 'text-green-600' : 'text-orange-500'}>
                                 {c.status === 'approved' ? `✓ 已確認 (${c.time})` : '待確認'}
                              </span>
                           </div>
                           <div className="italic text-gray-600">{c.comment || "無意見。"}</div>
                        </div>
                     ))}
                  </div>
               </div>
            </div>

            {modalMode !== 'view' && (
               <div className="p-4 border-t bg-gray-50 flex justify-end gap-3 no-print">
                  <button onClick={() => setIsModalOpen(false)} className="px-6 py-2 border rounded">取消</button>
                  <button onClick={async () => { if(await saveProjectToDb(editingProject)) setIsModalOpen(false); }} className="px-6 py-2 bg-indigo-600 text-white rounded">儲存簽呈</button>
               </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

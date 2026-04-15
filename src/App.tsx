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

const DEPARTMENTS = ["客務部", "訂房組", "餐飲部", "休閒部", "業務部", "企劃部", "人資", "資訊", "總務", "採購", "財務部"];

const ROLES_INFO: any = {
  guest: { name: "訪客 (僅觀看)", color: "text-gray-500", bg: "bg-gray-100" },
  employee: { name: "部門員工", color: "text-blue-700", bg: "bg-blue-100" },
  gm: { name: "總經理", color: "text-green-700", bg: "bg-green-100" },
  admin: { name: "系統管理員", color: "text-purple-700", bg: "bg-purple-100" },
};

// 🌟 核心：日曆生成引擎 (完全依賴資料庫)
const generateCalendar = (year: number, dbEvents: any[]) => {
  const data: any = {};
  for (let m = 1; m <= 12; m++) {
    const daysInMonth = new Date(year, m, 0).getDate();
    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${year}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
      const dateObj = new Date(dateStr);
      const day = dateObj.getDay();
      
      // 1. 基礎判斷：週末為假日，其餘為平日
      let type = (day === 6 || day === 0) ? "假日" : "平日";
      let events: string[] = [];

      // 2. 寒暑假預設旺日邏輯 (1/21~2/13, 7~8月)
      const isWinter = (m === 1 && d >= 21) || (m === 2 && d <= 13);
      const isSummer = m === 7 || m === 8;
      if (type === "平日" && (day === 5 || isWinter || isSummer)) {
        type = "旺日";
      }

      // 3. 🔍 匹配 Supabase 資料庫中的設定 (這會決定大假日與節慶名稱)
      const dbMatch = dbEvents.find(e => e.date === dateStr);
      if (dbMatch) {
        if (dbMatch.event_type) type = dbMatch.event_type;
        if (dbMatch.event_name) {
          const icon = dbMatch.is_public_holiday ? '🧨' : '🎯';
          events.push(`${icon} ${dbMatch.event_name}`);
        }
      }

      data[dateStr] = { type, events, marketingEvents: [] };
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

  const [isChangePwdModalOpen, setIsChangePwdModalOpen] = useState(false);
  const [pwdForm, setPwdForm] = useState({ old: "", new: "", confirm: "" });
  const [isUserModalOpen, setIsUserModalOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<any>(null);

  const [rememberMe, setRememberMe] = useState(() => localStorage.getItem("mpr_remember") === "true");
  const [savedAccount, setSavedAccount] = useState(() => localStorage.getItem("mpr_account") || "");
  const [savedPassword, setSavedPassword] = useState(() => localStorage.getItem("mpr_password") || "");

  const [ganttDeptFilter, setGanttDeptFilter] = useState("all");
  const [ganttMonthFilter, setGanttMonthFilter] = useState("all");
  const [dashboardDeptFilter, setDashboardDeptFilter] = useState("all");
  const [dashboardMonthFilter, setDashboardMonthFilter] = useState(String(new Date().getMonth() + 1));

  useEffect(() => { fetchData(); }, []);

  const fetchData = async () => {
    setIsLoading(true);
    try {
      const { data: u } = await supabase.from("users").select("*"); if (u) setUsers(u);
      const { data: p } = await supabase.from("projects").select("*"); 
      if (p) setProjects(p.map(x => ({ ...x, breakdown: JSON.parse(x.breakdown || "{}"), countersign: JSON.parse(x.countersign || "[]") })));
      const { data: c } = await supabase.from("calendar_events").select("*"); if (c) setDbEvents(c);
    } catch (e) { console.log(e); }
    setIsLoading(false);
  };

  const handleSaveDayRemark = async (dateStr: string, newRemark: string) => {
    const existing = dbEvents.find(e => e.date === dateStr);
    const payload = existing ? { ...existing, description: newRemark } : { date: dateStr, event_name: "自訂備註", event_type: "平日", description: newRemark };
    const { error } = await supabase.from('calendar_events').upsert(payload, { onConflict: 'date' });
    if (!error) { fetchData(); setSelectedDayInfo(null); alert("儲存成功"); }
  };

  const yearOptions = useMemo(() => {
    const years = [currentYear - 1, currentYear, currentYear + 1, currentYear + 2];
    return Array.from(new Set(years)).sort();
  }, [currentYear]);

  const calendarData = useMemo(() => generateCalendar(selectedYear, dbEvents), [selectedYear, dbEvents]);

  const yearProjects = useMemo(() => projects.filter(p => p.startDate.startsWith(String(selectedYear)) || p.endDate.startsWith(String(selectedYear))), [projects, selectedYear]);
  const filteredScheduledProjects = useMemo(() => {
    return yearProjects.filter(p => p.status === "scheduled").filter(p => {
      let passDept = ganttDeptFilter === "all" || p.creator.includes(ganttDeptFilter);
      let passMonth = ganttMonthFilter === "all" || (new Date(p.startDate).getMonth() + 1 === Number(ganttMonthFilter) || new Date(p.endDate).getMonth() + 1 === Number(ganttMonthFilter));
      return passDept && passMonth;
    });
  }, [yearProjects, ganttDeptFilter, ganttMonthFilter]);

  const getBarStyles = (start: string, end: string) => {
    const yStart = new Date(`${selectedYear}-01-01`).getTime();
    const yEnd = new Date(`${selectedYear}-12-31`).getTime();
    const s = Math.max(yStart, new Date(start).getTime());
    const e = Math.min(yEnd, new Date(end).getTime());
    return { left: `${((s - yStart) / (yEnd - yStart)) * 100}%`, width: `${((e - s) / (yEnd - yStart)) * 100}%` };
  };

  if (isLoading) return <div className="min-h-screen flex items-center justify-center font-bold text-indigo-600 animate-pulse">系統資料同步中...</div>;

  if (view === "login") {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
        <form onSubmit={(e: any) => {
          e.preventDefault();
          const acc = e.target.account.value; const pass = e.target.password.value;
          const u = users.find(x => x.account === acc && x.password === pass);
          if (u) { setCurrentUser(u); setView("app"); } else alert("登入失敗");
        }} className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden">
          <div className="bg-indigo-600 p-8 text-center text-white"><Layout className="w-12 h-12 mx-auto mb-4" /><h1 className="text-2xl font-bold">專案管理系統</h1></div>
          <div className="p-8 space-y-4">
            <input name="account" placeholder="帳號" required className="w-full p-3 border rounded-lg" />
            <input name="password" type="password" placeholder="密碼" required className="w-full p-3 border rounded-lg" />
            <button type="submit" className="w-full bg-indigo-600 text-white p-3 rounded-lg font-bold">登入系統</button>
          </div>
        </form>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans">
      <header className="bg-white shadow-sm sticky top-0 z-30 px-4 h-16 flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 text-indigo-600"><Layout className="w-6 h-6" /><h1 className="font-bold text-xl">專案管理</h1></div>
          <select className="bg-indigo-50 border border-indigo-200 text-indigo-800 font-bold rounded-lg p-1" value={selectedYear} onChange={(e) => setSelectedYear(Number(e.target.value))}>
            {yearOptions.map(y => <option key={y} value={y}>{y} 年</option>)}
          </select>
        </div>
        <div className="flex items-center gap-3">
          <button onClick={() => setIsImportModalOpen(true)} className="bg-green-600 text-white px-3 py-2 rounded-lg flex items-center gap-1 font-bold text-sm shadow-sm"><UploadCloud className="w-4 h-4" /> 匯入資料</button>
          <button onClick={() => { setEditingProject({ id: Date.now(), title: "", refNo: `MPR-${selectedYear}-01`, applyDate: new Date().toISOString().split('T')[0], purpose: "", price: "", startDate: `${selectedYear}-01-01`, endDate: `${selectedYear}-01-31`, content: "", precautions: "", highlights: "", breakdown: { price: "", net: "", items: [] }, countersign: [], status: "countersigning", creator: `${currentUser.dept}-${currentUser.name}` }); setModalMode("create"); setIsModalOpen(true); }} className="bg-indigo-600 text-white px-3 py-2 rounded-lg flex items-center gap-1 font-bold text-sm shadow-sm"><Plus className="w-4 h-4" /> 新增簽呈</button>
          <div className="h-6 w-px bg-gray-200 mx-2" />
          <span className="text-sm font-bold bg-gray-100 px-2 py-1 rounded">{currentUser.name}</span>
          <button onClick={handleLogout} className="text-gray-400 hover:text-red-500"><LogOut className="w-5 h-5" /></button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-6 space-y-6">
        {/* 甘特圖區塊 */}
        <section className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
          <div className="p-4 border-b bg-gray-50 flex justify-between items-center">
            <h2 className="font-bold flex items-center gap-2"><Calendar className="w-5 h-5 text-indigo-600" /> 年度專案甘特圖</h2>
            <div className="flex gap-4">
               <select className="border rounded p-1 text-sm" value={ganttDeptFilter} onChange={e => setGanttDeptFilter(e.target.value)}>
                 <option value="all">所有部門</option>
                 {DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}
               </select>
            </div>
          </div>
          <div className="overflow-x-auto">
            <div className="min-w-[1000px]">
              <div className="flex border-b bg-gray-50 text-xs font-bold text-gray-500">
                <div className="w-48 p-3 border-r">專案名稱</div>
                <div className="flex-1 grid grid-cols-12">
                  {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <div key={m} onClick={() => setSelectedMonthView(m)} className="border-r p-2 text-center cursor-pointer hover:bg-indigo-100">{m}月</div>)}
                </div>
              </div>
              <div className="relative">
                <div className="absolute inset-0 flex ml-48 pointer-events-none">
                  {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <div key={m} className="flex-1 border-r border-gray-100" />)}
                </div>
                <div className="relative z-10 min-h-[200px]">
                  {filteredScheduledProjects.map(p => (
                    <div key={p.id} className="flex items-center border-b hover:bg-gray-50 cursor-pointer" onClick={() => { setEditingProject(p); setModalMode("view"); setIsModalOpen(true); }}>
                      <div className="w-48 p-3 text-sm font-medium border-r truncate">{p.title}</div>
                      <div className="flex-1 h-12 relative">
                        <div className="absolute top-3 bottom-3 rounded-md bg-indigo-500 shadow-sm text-[10px] text-white flex items-center px-2 truncate" style={getBarStyles(p.startDate, p.endDate)}>{p.title}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* 月曆圖例 */}
        <div className="bg-white p-3 rounded-lg shadow-sm border flex flex-wrap gap-4 text-xs font-bold">
          <span className="text-gray-500">行事曆圖例：</span>
          <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-white border border-gray-300" />平日</span>
          <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-blue-100 border border-blue-300" />旺日</span>
          <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-pink-100 border border-pink-300" />假日</span>
          <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-red-100 border border-red-300" />大假日 (除夕/春節/連假)</span>
        </div>
      </main>

      {/* --- 單日彈出視窗 --- */}
      {selectedDayInfo && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-sm overflow-hidden border">
            <div className={`px-4 py-3 border-b flex justify-between items-center ${selectedDayInfo.st.bg} ${selectedDayInfo.st.text}`}>
              <h3 className="font-bold flex items-center gap-2"><CalendarDays className="w-5 h-5" />{selectedDayInfo.date}</h3>
              <button onClick={() => setSelectedDayInfo(null)} className="hover:opacity-70"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-5 space-y-4">
              <span className={`text-xs px-2 py-1 rounded shadow-sm ${selectedDayInfo.st.tag} font-bold`}>{selectedDayInfo.dailyData.type}</span>
              {selectedDayInfo.dailyData.events?.length > 0 && (
                <div className="space-y-1">
                  <h4 className="text-xs font-bold text-gray-500 border-b pb-1">節慶與活動</h4>
                  {selectedDayInfo.dailyData.events.map((ev: string, i: number) => <div key={i} className="text-sm bg-yellow-50 text-yellow-800 border border-yellow-200 px-2 py-1 rounded">{ev}</div>)}
                </div>
              )}
              {selectedDayInfo.dayProjects?.length > 0 && (
                <div className="space-y-1">
                  <h4 className="text-xs font-bold text-gray-500 border-b pb-1">館內專案</h4>
                  {selectedDayInfo.dayProjects.map((p: any) => <div key={p.id} className="text-sm bg-indigo-500 text-white px-2 py-1 rounded shadow-sm">{p.title}</div>)}
                </div>
              )}
              <div className="mt-4 pt-4 border-t">
                <h4 className="text-sm font-bold text-indigo-800 flex items-center gap-1 mb-2"><MessageSquare className="w-4 h-4" /> 日程備註</h4>
                {currentUser?.role === "admin" ? (
                  <div className="space-y-2">
                    <textarea id="day-remark" className="w-full border rounded-lg p-2 text-sm" rows={3} defaultValue={dbEvents.find(e => e.date === selectedDayInfo.date)?.description || ""} placeholder="管理員輸入備註..." />
                    <button onClick={() => handleSaveDayRemark(selectedDayInfo.date, (document.getElementById("day-remark") as HTMLTextAreaElement).value)} className="w-full bg-indigo-600 text-white py-1.5 rounded font-bold text-sm">儲存備註</button>
                  </div>
                ) : (
                  <div className="bg-gray-50 p-3 rounded-lg border text-sm text-gray-600 whitespace-pre-wrap">{dbEvents.find(e => e.date === selectedDayInfo.date)?.description || "本日無備註。"}</div>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* --- 月份詳細日曆彈窗 --- */}
      {selectedMonthView && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="p-4 border-b bg-indigo-50 flex justify-between items-center">
              <h2 className="font-bold text-indigo-900 text-lg">{selectedYear} 年 {selectedMonthView} 月 行事曆</h2>
              <button onClick={() => setSelectedMonthView(null)}><X className="w-6 h-6" /></button>
            </div>
            <div className="p-4 overflow-y-auto grid grid-cols-7 gap-2">
              {["日", "一", "二", "三", "四", "五", "六"].map(w => <div key={w} className="text-center font-bold text-gray-400 text-xs">{w}</div>)}
              {Array.from({ length: new Date(selectedYear, selectedMonthView - 1, 1).getDay() }).map((_, i) => <div key={i} />)}
              {Array.from({ length: new Date(selectedYear, selectedMonthView, 0).getDate() }).map((_, i) => {
                const d = i + 1;
                const ds = `${selectedYear}-${String(selectedMonthView).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                const data = calendarData[ds] || { type: "平日", events: [] };
                const st = CAL_STYLES[data.type];
                return (
                  <div key={d} onClick={() => setSelectedDayInfo({ date: ds, dailyData: data, dayProjects: yearProjects.filter(p => p.startDate <= ds && p.endDate >= ds), st })} className={`min-h-[80px] border rounded-lg p-2 cursor-pointer hover:ring-2 hover:ring-indigo-400 ${st.bg} ${st.border}`}>
                    <div className="flex justify-between items-start">
                      <span className={`font-bold ${st.text}`}>{d}</span>
                      <span className="text-[8px] px-1 rounded bg-white/50">{data.type}</span>
                    </div>
                    {data.events.map((ev: any, idx: number) => <div key={idx} className="text-[9px] mt-1 truncate font-bold text-red-600">{ev}</div>)}
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* --- 匯入資料彈窗 --- */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-lg p-6 space-y-4">
            <h2 className="text-xl font-bold flex items-center gap-2 text-green-700"><UploadCloud /> 匯入專案資料</h2>
            <p className="text-sm text-gray-500">上傳 Excel 或貼上文字，系統會自動預填簽呈內容。</p>
            <input type="file" accept=".xlsx,.csv" onChange={e => setImportFile(e.target.files?.[0] || null)} className="w-full border p-2 rounded" />
            <textarea className="w-full border rounded p-3 h-32 text-sm" placeholder="或是貼上專案說明文字..." value={importText} onChange={e => setImportText(e.target.value)} />
            <div className="flex justify-end gap-2">
              <button onClick={() => setIsImportModalOpen(false)} className="px-4 py-2 text-gray-500">取消</button>
              <button onClick={handleProcessImport} className="bg-green-600 text-white px-6 py-2 rounded-lg font-bold">開始解析</button>
            </div>
          </div>
        </div>
      )}

      {/* --- 簽呈編輯/檢視彈窗 (完整保留) --- */}
      {isModalOpen && editingProject && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm overflow-y-auto">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-4xl my-8">
            <div className="p-6 border-b flex justify-between items-center bg-gray-50 sticky top-0 z-10">
              <h2 className="text-xl font-bold">{modalMode === "view" ? "專案簽呈檢視" : "編輯簽呈"}</h2>
              <button onClick={() => setIsModalOpen(false)}><X className="w-6 h-6 text-gray-400" /></button>
            </div>
            <div className="p-8">
              {modalMode === "view" ? (
                <div className="space-y-6">
                  <WorkflowProgressBar project={editingProject} />
                  <div className="border p-6 rounded-lg space-y-4">
                    <div className="flex justify-between border-b pb-4">
                      <div className="font-bold text-gray-500">Ref No: {editingProject.refNo}</div>
                      <div className="font-bold text-gray-500">日期: {editingProject.applyDate}</div>
                    </div>
                    <div className="grid grid-cols-4 gap-4">
                       <div className="bg-gray-100 p-2 font-bold text-center border">主旨</div>
                       <div className="col-span-3 p-2 border font-bold text-lg">{editingProject.title}</div>
                    </div>
                    <div className="grid grid-cols-4 gap-4">
                       <div className="bg-gray-100 p-2 font-bold text-center border">說明</div>
                       <div className="col-span-3 p-2 border whitespace-pre-wrap min-h-[100px]">{editingProject.purpose}</div>
                    </div>
                    <div className="grid grid-cols-4 gap-4">
                       <div className="bg-gray-100 p-2 font-bold text-center border">活動日期</div>
                       <div className="col-span-3 p-2 border font-bold text-indigo-600">{editingProject.startDate} ~ {editingProject.endDate}</div>
                    </div>
                  </div>
                </div>
              ) : (
                <form onSubmit={handleSave} className="space-y-4">
                  <div><label className="block font-bold mb-1">專案主旨</label><input className="w-full border p-2 rounded" value={editingProject.title} onChange={e => setEditingProject({...editingProject, title: e.target.value})} required /></div>
                  <div className="grid grid-cols-2 gap-4">
                    <div><label className="block font-bold mb-1">開始日期</label><input type="date" className="w-full border p-2 rounded" value={editingProject.startDate} onChange={e => setEditingProject({...editingProject, startDate: e.target.value})} /></div>
                    <div><label className="block font-bold mb-1">結束日期</label><input type="date" className="w-full border p-2 rounded" value={editingProject.endDate} onChange={e => setEditingProject({...editingProject, endDate: e.target.value})} /></div>
                  </div>
                  <div><label className="block font-bold mb-1">說明內容</label><textarea className="w-full border p-2 rounded h-40" value={editingProject.purpose} onChange={e => setEditingProject({...editingProject, purpose: e.target.value})} /></div>
                  <div className="flex justify-end gap-2 pt-4"><button type="button" onClick={() => setIsModalOpen(false)} className="px-4 py-2 border rounded">取消</button><button type="submit" className="bg-indigo-600 text-white px-8 py-2 rounded font-bold">儲存簽呈</button></div>
                </form>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// 輔助登出與格式化
const handleLogout = () => { window.location.reload(); };

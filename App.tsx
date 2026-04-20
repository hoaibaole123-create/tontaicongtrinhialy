
import React, { useState, useEffect, useCallback, useRef } from 'react';
import { 
  PlusCircle, 
  Moon, 
  Sun,
  Signal,
  Wifi, 
  Battery, 
  LayoutDashboard,
  ClipboardList,
  Settings,
  CheckCircle2,
  Loader2,
  TableProperties,
  Search,
  FileSpreadsheet,
  RefreshCw,
  Download,
  AlertCircle,
  ExternalLink,
  X,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  UploadCloud,
  Trash2,
  Save,
  Edit3,
  Camera,
  CheckSquare,
  Filter,
  Clock,
  Activity as ActivityIcon,
  Users
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  Legend,
  Radar,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Area,
  AreaChart,
  PieChart,
  Pie,
  Cell,
  Sector
} from 'recharts';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// --- Consolidated Types ---
export enum DefectStatus {
  PENDING = 'PENDING',
  PROCESSING = 'PROCESSING',
  COMPLETED = 'COMPLETED',
  URGENT = 'URGENT'
}

export interface ChartData {
  name: string;
  detected: number;
  processed: number;
  nvvh: number;
}

// --- Configuration ---
const SHEET_ID = '1EVA37o8kSgi3Z86hwUQN5uyBtVwERDo3REO0xMtMqE0';

// AppScript URL
const REPORT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbytJsnBmwEMosm1dLK8VZTLYTt2CvR0E-ApUHFMDgWV6B0T1GEBnkk400Q4v0XBrRVO/exec';
const PROCESS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbz6EOtoLlEu4qUDZPllqs2eET8VOQ14WwJbM0drY-sVWKWVL1nKJcAFqo7nsnGdZ6jl/exec';
const EDIT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwpCbRAAzIKjQhBq96OqYoAHgDaaahzvFkKo2NczHntmkwZOeGnSSFecvg44ZXZUhs/exec';

const CATEGORIES = [
  'Quản lý hành chính',
  'Thiết bị công trình',
  'An toàn vệ sinh lao động',
  'TPM, Kaizen'
];

// --- Sub-components ---

const FormLabel: React.FC<{ icon?: string, children: React.ReactNode, required?: boolean }> = ({ icon, children, required }) => (
  <label className="block text-[14px] font-semibold text-slate-700 dark:text-slate-300 mb-2 flex items-center gap-2">
    {icon && <span>{icon}</span>} {children} {required && <span className="text-red-500">*</span>}
  </label>
);

const CustomRadio: React.FC<{ label: string, description?: string, name: string, value: string, checked?: boolean, onChange?: (e: any) => void }> = ({ label, description, name, value, checked, onChange }) => (
  <label className="flex items-start gap-2 mb-2 cursor-pointer group">
    <div className="mt-1 relative flex items-center justify-center">
      <input 
        type="radio" 
        name={name} 
        value={value} 
        checked={checked} 
        onChange={onChange}
        className="w-4 h-4 text-blue-600 focus:ring-blue-500 border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 cursor-pointer" 
      />
    </div>
    <div className="flex-1">
      <div className={`text-[13px] font-medium ${checked ? 'text-blue-700 dark:text-blue-400' : 'text-slate-700 dark:text-slate-300'}`}>{label}</div>
      {description && <div className="text-[11px] text-slate-400 dark:text-slate-500 leading-tight italic">{description}</div>}
    </div>
  </label>
);

// --- Dashboard Component ---

const Dashboard: React.FC<{ isDarkMode: boolean, onActivityClick: (sheet: string, row: number) => void, onStatClick: (sheet: string, status: 'all' | 'processed' | 'pending' | 'nvvh') => void }> = ({ isDarkMode, onActivityClick, onStatClick }) => {
  const [isLoading, setIsLoading] = useState(true);
  const [stats, setStats] = useState({ total: 0, pending: 0, completed: 0, notStarted: 0 });
  const [chartData, setChartData] = useState<any[]>([]);
  const [recentActivities, setRecentActivities] = useState<any[]>([]);
  const [topContributors, setTopContributors] = useState<{name: string, total: number, monthly: {[key: string]: number}}[]>([]);
  const [isStatsExpanded, setIsStatsExpanded] = useState(false);

  const fetchDashboardData = useCallback(async () => {
    setIsLoading(true);
    try {
      let combinedActivities: any[] = [];
      let totalCount = 0;
      let completedCount = 0;
      let processingCount = 0;
      let notStartedCount = 0;
      const contributorMap: {[key: string]: { total: number, monthly: {[key: string]: number} }} = {};

      const parseDateObj = (val: any) => {
        if (!val) return null;
        const str = String(val);
        if (str.startsWith('Date(')) {
          const p = str.match(/\d+/g);
          if (p) return new Date(Number(p[0]), Number(p[1]), Number(p[2]), Number(p[3]||0), Number(p[4]||0), Number(p[5]||0));
        }
        const ts = Date.parse(str);
        return isNaN(ts) ? null : new Date(ts);
      };

      const promises = CATEGORIES.map(async (cat) => {
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(cat)}`;
        const response = await fetch(url);
        const text = await response.text();
        const match = text.match(/google\.visualization\.Query\.setResponse\((.*)\);/);
        
        if (match) {
          const json = JSON.parse(match[1]);
          if (json.table && json.table.rows) {
            const rows = json.table.rows;
            const detected = rows.length;
            
            const processed = rows.filter((r: any) => r.c[11] && r.c[11].v).length;
            const processing = rows.filter((r: any) => (r.c[12] && r.c[12].v) && !(r.c[11] && r.c[11].v)).length;
            const pending = rows.filter((r: any) => !(r.c[12] && r.c[12].v) && !(r.c[11] && r.c[11].v)).length;

            rows.forEach((r: any, idx: number) => {
              const physicalRow = idx + 2; // Physical row number in Google Sheets
              
              // Aggregate contributors (case-insensitive) - Using robust cell value extraction
              const reporterCell = r.c[2];
              const rawName = reporterCell ? (reporterCell.f || (reporterCell.v != null ? String(reporterCell.v) : '')) : '';
              const trimmedName = rawName.trim();
              
              if (trimmedName) {
                const normalizedName = trimmedName.toUpperCase();
                if (!contributorMap[normalizedName]) {
                  contributorMap[normalizedName] = { total: 0, monthly: {} };
                }
                contributorMap[normalizedName].total += 1;
                
                const dateObj = parseDateObj(r.c[1]?.v);
                if (dateObj) {
                  const monthKey = `${String(dateObj.getMonth() + 1).padStart(2, '0')}/${dateObj.getFullYear()}`;
                  contributorMap[normalizedName].monthly[monthKey] = (contributorMap[normalizedName].monthly[monthKey] || 0) + 1;
                }
              }

              combinedActivities.push({
                time: r.c[1]?.f || 'N/A',
                rawTime: r.c[1]?.v || '', 
                title: String(r.c[5]?.v || 'Không rõ'),
                location: String(r.c[6]?.v || 'N/A'),
                category: cat,
                row: physicalRow,
                isDone: !!(r.c[11] && r.c[11].v)
              });
            });

            totalCount += detected;
            completedCount += processed;
            processingCount += processing;
            notStartedCount += pending;

            return {
              name: cat,
              detected, 
              processed, 
              processing,
              pending      
            };
          }
        }
        return { name: cat, detected: 0, processed: 0, processing: 0, pending: 0 };
      });

      const results = await Promise.all(promises);
      setChartData(results);
      setStats({ total: totalCount, completed: completedCount, pending: processingCount, notStarted: notStartedCount });
      
      // Process contributors
      const sortedContributors = Object.entries(contributorMap)
        .map(([name, data]) => ({ name, total: data.total, monthly: data.monthly }))
        .sort((a, b) => b.total - a.total);
      setTopContributors(sortedContributors);

      // Improved sorting: parse Google Date strings and fallback to row index
      const sortedActivities = combinedActivities.sort((a, b) => {
        const parseDate = (val: any) => {
          if (!val) return 0;
          const str = String(val);
          if (str.startsWith('Date(')) {
            const p = str.match(/\d+/g);
            if (p) return new Date(Number(p[0]), Number(p[1]), Number(p[2]), Number(p[3]||0), Number(p[4]||0), Number(p[5]||0)).getTime();
          }
          const ts = Date.parse(str);
          return isNaN(ts) ? 0 : ts;
        };

        const timeA = parseDate(a.rawTime);
        const timeB = parseDate(b.rawTime);

        if (timeB !== timeA) return timeB - timeA;
        return b.row - a.row; // If same time, higher row index is newer
      });
      
      setRecentActivities(sortedActivities.slice(0, 6));

    } catch (err) {
      console.error("Dashboard error:", err);
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchDashboardData();
  }, [fetchDashboardData]);

  return (
    <div className="animate-in fade-in duration-500 pb-10">
      <div className="relative w-full h-64 overflow-hidden">
        <img alt="Factory" className="w-full h-full object-cover brightness-[0.25]" src="https://i.ibb.co/zWPTxZvg/123.png" />
        <div className="absolute inset-0 flex flex-col justify-center items-center px-10 text-center bg-gradient-to-b from-blue-900/30 via-transparent to-slate-900/90">
          <h1 className="text-white text-2xl font-black uppercase tracking-tight drop-shadow-2xl mb-2">
            KIỂM TRA VÀ CẬP NHẬT CÁC HƯ HỎNG, TỒN TẠI VÀ CÁC ĐIỂM KHÔNG PHÙ HỢP
          </h1>
          <div className="flex items-center gap-2 text-blue-200 text-[10px] font-black uppercase tracking-widest bg-white/10 backdrop-blur-md px-5 py-2 rounded-full border border-white/20">
            <ActivityIcon size={14} className="animate-pulse" />
            
          </div>
        </div>
      </div>

      <main className="px-4 -mt-12 relative z-10 flex-1 w-full">
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6">
          <div 
            onClick={() => onStatClick('all', 'all')}
            className="bg-white dark:bg-slate-800 p-4 rounded-3xl shadow-xl border border-white/50 dark:border-slate-800 text-center group transition-all hover:bg-blue-50/10 cursor-pointer active:scale-95"
          >
            <p className="text-[16px] text-slate-400 font-black uppercase mb-1 tracking-widest">Tổng</p>
            <p className="text-2xl font-black text-blue-600 dark:text-blue-400">{isLoading ? <Loader2 className="animate-spin inline" size={16} /> : stats.total}</p>
          </div>
          <div 
            onClick={() => onStatClick('all', 'processed')}
            className="bg-white dark:bg-slate-800 p-4 rounded-3xl shadow-xl border border-white/50 dark:border-slate-800 text-center group transition-all hover:bg-emerald-50/10 cursor-pointer active:scale-95"
          >
            <p className="text-[16px] text-slate-400 font-black uppercase mb-1 tracking-widest">Hoàn thành</p>
            <p className="text-2xl font-black text-emerald-500">{isLoading ? <Loader2 className="animate-spin inline" size={16} /> : stats.completed}</p>
          </div>
          <div 
            onClick={() => onStatClick('all', 'nvvh')}
            className="bg-white dark:bg-slate-800 p-4 rounded-3xl shadow-xl border border-white/50 dark:border-slate-800 text-center group transition-all hover:bg-amber-50/10 cursor-pointer active:scale-95"
          >
            <p className="text-[16px] text-slate-400 font-black uppercase mb-1 tracking-widest">Đang xử lý</p>
            <p className="text-2xl font-black text-amber-500">{isLoading ? <Loader2 className="animate-spin inline" size={16} /> : stats.pending}</p>
          </div>
          <div 
            onClick={() => onStatClick('all', 'pending')}
            className="bg-white dark:bg-slate-800 p-4 rounded-3xl shadow-xl border border-white/50 dark:border-slate-800 text-center group transition-all hover:bg-rose-50/10 cursor-pointer active:scale-95"
          >
            <p className="text-[20px] text-slate-400 font-black uppercase mb-1 tracking-widest">Chưa xử lý</p>
            <p className="text-2xl font-black text-rose-500">{isLoading ? <Loader2 className="animate-spin inline" size={16} /> : stats.notStarted}</p>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 p-8 rounded-[2.5rem] shadow-2xl border border-slate-100 dark:border-slate-800 mb-6">
          <div className="flex items-center justify-between mb-8">
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-blue-100 dark:bg-blue-900/30 rounded-2xl text-blue-600"><ActivityIcon size={18} /></div>
              <h2 className="text-sm font-black uppercase tracking-widest text-slate-700 dark:text-slate-300">Biểu đồ hoạt động</h2>
            </div>
            <button onClick={fetchDashboardData} className="p-2 text-slate-400 hover:text-blue-500 transition-colors">
              <RefreshCw size={16} className={isLoading ? 'animate-spin' : ''} />
            </button>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-8">
            {isLoading ? (
              Array(4).fill(0).map((_, i) => (
                <div key={i} className="h-48 bg-slate-50 dark:bg-slate-900/50 rounded-3xl animate-pulse" />
              ))
            ) : (
              chartData.map((entry, idx) => {
                const completed = entry.processed;
                const totalStarted = entry.processing;
                const pending = entry.pending;
                const inProgressOnly = Math.max(0, entry.detected - completed - pending);
                
                const data = [
                  { name: 'Hoàn thành', value: completed, color: '#10b981', grad: `gradGreen-${idx}` },
                  { name: 'Đang xử lý', value: inProgressOnly, color: '#f59e0b', grad: `gradAmber-${idx}` },
                  { name: 'Chưa xử lý', value: pending, color: '#ef4444', grad: `gradRed-${idx}` }
                ].filter(d => d.value > 0);

                return (
                  <div key={idx} className="flex flex-col items-center group">
                    <div className="h-60 w-full relative bg-slate-900/5 dark:bg-slate-900/20 rounded-[2.5rem] overflow-hidden shadow-inner border border-slate-200/50 dark:border-slate-700/30 transition-all hover:shadow-2xl hover:scale-[1.02] duration-500">
                      <div className="absolute inset-0 bg-gradient-to-br from-white/40 to-transparent dark:from-white/5 pointer-events-none" />
                      
                      {/* Corner Buttons */}
                      <button 
                        onClick={() => onStatClick(entry.name, 'processed')}
                        className="absolute top-4 left-4 z-20 flex flex-col items-center p-3.5 bg-emerald-50/90 dark:bg-emerald-900/50 backdrop-blur-md rounded-[1.5rem] border border-emerald-100 dark:border-emerald-800/50 hover:scale-110 active:scale-95 transition-all shadow-xl min-w-[70px]"
                      >
                        <span className="text-[9px] font-black text-emerald-600 dark:text-emerald-400 uppercase tracking-widest mb-1">Hoàn thành</span>
                        <span className="text-lg font-black text-emerald-700 dark:text-emerald-300 leading-none">{completed}</span>
                      </button>

                      <button 
                        onClick={() => onStatClick(entry.name, 'nvvh')}
                        className="absolute top-4 right-4 z-20 flex flex-col items-center p-3.5 bg-amber-50/90 dark:bg-amber-900/50 backdrop-blur-md rounded-[1.5rem] border border-amber-100 dark:border-amber-800/50 hover:scale-110 active:scale-95 transition-all shadow-xl min-w-[70px]"
                      >
                        <span className="text-[9px] font-black text-amber-600 dark:text-amber-400 uppercase tracking-widest mb-1">Đang xử lý</span>
                        <span className="text-lg font-black text-amber-700 dark:text-amber-300 leading-none">{totalStarted}</span>
                      </button>

                      <button 
                        onClick={() => onStatClick(entry.name, 'pending')}
                        className="absolute bottom-4 left-4 z-20 flex flex-col items-center p-3.5 bg-rose-50/90 dark:bg-rose-900/50 backdrop-blur-md rounded-[1.5rem] border border-rose-100 dark:border-rose-800/50 hover:scale-110 active:scale-95 transition-all shadow-xl min-w-[70px]"
                      >
                        <span className="text-[9px] font-black text-rose-600 dark:text-rose-400 uppercase tracking-widest mb-1">Chưa xử lý</span>
                        <span className="text-lg font-black text-rose-700 dark:text-rose-300 leading-none">{pending}</span>
                      </button>

                      <ResponsiveContainer width="100%" height="100%" minWidth={0} minHeight={0}>
                        <PieChart>
                          <defs>
                            <linearGradient id={`gradGreen-${idx}`} x1="0" y1="0" x2="0" y2="1">
                              <stop offset="0%" stopColor="#10b981" />
                              <stop offset="100%" stopColor="#059669" />
                            </linearGradient>
                            <linearGradient id={`gradAmber-${idx}`} x1="0" y1="0" x2="0" y2="1">
                              <stop offset="0%" stopColor="#f59e0b" />
                              <stop offset="100%" stopColor="#d97706" />
                            </linearGradient>
                            <linearGradient id={`gradRed-${idx}`} x1="0" y1="0" x2="0" y2="1">
                              <stop offset="0%" stopColor="#ef4444" />
                              <stop offset="100%" stopColor="#b91c1c" />
                            </linearGradient>
                            <filter id="vividShadow" height="150%">
                              <feGaussianBlur in="SourceAlpha" stdDeviation="3" />
                              <feOffset dx="0" dy="10" result="offsetblur" />
                              <feComponentTransfer>
                                <feFuncA type="linear" slope="0.5" />
                              </feComponentTransfer>
                              <feMerge>
                                <feMergeNode />
                                <feMergeNode in="SourceGraphic" />
                              </feMerge>
                            </filter>
                          </defs>
                          <Pie
                            data={data}
                            cx="50%"
                            cy="50%"
                            startAngle={90}
                            endAngle={-270}
                            innerRadius={35}
                            outerRadius={75}
                            paddingAngle={4}
                            dataKey="value"
                            stroke="none"
                            filter="url(#vividShadow)"
                            animationDuration={1800}
                            labelLine={false}
                            label={({ cx, cy, midAngle, innerRadius, outerRadius, percent }) => {
                              const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
                              const x = cx + radius * Math.cos(-midAngle * (Math.PI / 180));
                              const y = cy + radius * Math.sin(-midAngle * (Math.PI / 180));
                              if (percent < 0.1) return null;
                              return (
                                <text 
                                  x={x} 
                                  y={y} 
                                  fill="white" 
                                  textAnchor="middle" 
                                  dominantBaseline="central" 
                                  className="text-[11px] font-black drop-shadow-lg"
                                >
                                  {(percent * 100).toFixed(0)}%
                                </text>
                              );
                            }}
                          >
                            {data.map((entry, index) => (
                              <Cell 
                                key={`cell-${index}`} 
                                fill={`url(#${entry.grad})`}
                                className="transition-all duration-500 hover:opacity-80 cursor-pointer"
                              />
                            ))}
                          </Pie>
                          <Tooltip 
                            content={({ active, payload }) => {
                              if (active && payload && payload.length) {
                                const d = payload[0].payload;
                                return (
                                  <div className="bg-white dark:bg-slate-900 p-3 rounded-2xl shadow-2xl border border-slate-100 dark:border-slate-800 animate-in fade-in zoom-in duration-200">
                                    <p className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">{d.name}</p>
                                    <p className="text-lg font-black text-slate-900 dark:text-white">{d.value}</p>
                                  </div>
                                );
                              }
                              return null;
                            }}
                          />
                        </PieChart>
                      </ResponsiveContainer>
                      <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-center pointer-events-none">
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-tighter leading-none">Tổng</p>
                        <p className="text-lg font-black text-slate-700 dark:text-slate-300">{entry.detected}</p>
                      </div>
                    </div>
                    <div className="mt-5 text-center w-full">
                      <p className="text-xs font-black uppercase tracking-[0.2em] text-slate-800 dark:text-slate-200 mb-1">{entry.name}</p>
                    </div>
                  </div>
                );
              })
            )}
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 p-6 rounded-[2.5rem] shadow-2xl border border-slate-100 dark:border-slate-800 mb-6">
          <div className="flex items-center gap-3 mb-6">
            <div className="p-2.5 bg-blue-100 dark:bg-blue-900/30 rounded-2xl text-blue-600"><Clock size={18} /></div>
            <h2 className="text-sm font-black uppercase tracking-widest text-slate-700 dark:text-slate-300">Hoạt động mới nhất</h2>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {isLoading ? (
              Array(4).fill(0).map((_, i) => (
                <div key={i} className="h-20 bg-slate-50 dark:bg-slate-900/50 rounded-3xl animate-pulse" />
              ))
            ) : recentActivities.length > 0 ? (
              recentActivities.map((act, idx) => (
                <div 
                  key={idx} 
                  onClick={() => onActivityClick(act.category, act.row)}
                  className="flex items-center gap-4 p-4 bg-slate-50 dark:bg-slate-900/40 rounded-3xl border border-slate-100 dark:border-slate-800/50 group transition-all hover:shadow-lg cursor-pointer active:scale-[0.98]"
                >
                  <div className={`shrink-0 w-12 h-12 rounded-[1.2rem] flex items-center justify-center shadow-sm ${act.isDone ? 'bg-emerald-100 text-emerald-600' : 'bg-amber-100 text-amber-600'}`}>
                    {act.isDone ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
                  </div>
                  <div className="flex-1 min-w-0">
                    <h3 className="text-[16px] font-black text-slate-800 dark:text-slate-200 truncate leading-tight mb-1 uppercase tracking-tighter">
                      {act.title}
                    </h3>
                    <p className="text-[15px] text-slate-500 dark:text-slate-400 truncate mb-1">
                      📍 {act.location}
                    </p>
                    <div className="flex items-center gap-2">
                      <span className="text-[15px] font-black text-blue-500 uppercase tracking-widest bg-blue-50 dark:bg-blue-900/30 px-2 py-0.5 rounded-md">
                        {act.category.split(',')[0]}
                      </span>
                      <span className="text-[15px] font-bold text-slate-400 flex items-center gap-1">
                        <Clock size={10} /> {act.time.split(' ')[0]}
                      </span>
                    </div>
                  </div>
                  <ChevronRight size={14} className="text-slate-300 group-hover:translate-x-1 transition-transform" />
                </div>
              ))
            ) : (
              <div className="col-span-full py-12 text-center opacity-30 flex flex-col items-center">
                <ActivityIcon size={40} strokeWidth={1} />
                <p className="text-[10px] uppercase font-black tracking-widest mt-3">Hiện chưa có hoạt động</p>
              </div>
            )}
          </div>
        </div>

        <div className="mb-6">
          <div className="bg-white dark:bg-slate-800 p-8 rounded-[2.5rem] shadow-2xl border border-slate-100 dark:border-slate-800">
            <button 
              onClick={() => setIsStatsExpanded(!isStatsExpanded)}
              className="w-full flex items-center justify-between mb-0 group"
            >
              <div className="flex items-center gap-3">
                <div className="p-3 bg-emerald-100 dark:bg-emerald-900/30 rounded-2xl text-emerald-600 shadow-inner"><Users size={20} /></div>
                <div className="text-left">
                  <h2 className="text-sm font-black uppercase tracking-widest text-slate-700 dark:text-slate-300">Thống kê hoạt động</h2>
                  <p className="text-[9px] text-slate-400 font-bold uppercase tracking-tighter">Bảng thông kê phát hiện tồn tại của từng chức danh</p>
                </div>
              </div>
              <div className="p-2 rounded-full bg-slate-50 dark:bg-slate-900 group-hover:bg-slate-100 dark:group-hover:bg-slate-700 transition-colors">
                {isStatsExpanded ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
              </div>
            </button>
            
            {isStatsExpanded && (
              <div className="mt-8 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 max-h-[400px] overflow-y-auto pr-2 pb-4 custom-scrollbar animate-in slide-in-from-top-4 duration-300">
                {isLoading ? (
                  Array(6).fill(0).map((_, i) => (
                    <div key={i} className="h-24 bg-slate-50 dark:bg-slate-900/50 rounded-3xl animate-pulse" />
                  ))
                ) : topContributors.length > 0 ? (
                  topContributors.map((person, idx) => (
                    <div key={idx} className="relative group p-5 bg-slate-50 dark:bg-slate-900/40 rounded-[2rem] border border-slate-100 dark:border-slate-800/50 transition-all hover:shadow-xl hover:bg-white dark:hover:bg-slate-800 hover:-translate-y-1">
                      {/* Badge for Top 3 */}
                      {idx < 3 && (
                        <div className={`absolute -top-2 -right-2 w-8 h-8 rounded-full flex items-center justify-center text-[10px] font-black shadow-lg z-10 ${
                          idx === 0 ? 'bg-amber-400 text-white' : 
                          idx === 1 ? 'bg-slate-300 text-slate-700' : 
                          'bg-orange-400 text-white'
                        }`}>
                          {idx === 0 ? '🏆' : idx + 1}
                        </div>
                      )}
                      
                      <div className="flex items-center gap-4 mb-4">
                        <div className={`w-14 h-14 rounded-2xl flex items-center justify-center text-lg font-black shadow-inner shrink-0 ${
                          idx === 0 ? 'bg-gradient-to-br from-amber-100 to-amber-200 text-amber-600' :
                          idx === 1 ? 'bg-gradient-to-br from-slate-100 to-slate-200 text-slate-600' :
                          'bg-gradient-to-br from-blue-100 to-blue-200 text-blue-600'
                        }`}>
                          {person.name.charAt(0)}
                        </div>
                        <div className="min-w-0 flex-1">
                          <h3 className="text-[12px] font-black text-slate-800 dark:text-slate-100 uppercase tracking-tight truncate">
                            {person.name}
                          </h3>
                          <div className="flex items-center gap-1.5 mt-1">
                            <span className="text-[10px] font-black text-blue-600 dark:text-blue-400">
                              {person.total}
                            </span>
                            <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">Đóng góp</span>
                          </div>
                        </div>
                      </div>
                      
                      <div className="flex gap-1.5 overflow-x-auto pb-1 no-scrollbar">
                        {Object.entries(person.monthly)
                          .sort((a, b) => {
                            const [m1, y1] = a[0].split('/').map(Number);
                            const [m2, y2] = b[0].split('/').map(Number);
                            return y2 !== y1 ? y2 - y1 : m2 - m1;
                          })
                          .map(([month, count]) => (
                            <div key={month} className="flex flex-col items-center min-w-[45px] py-1.5 bg-white dark:bg-slate-800 rounded-xl border border-slate-100 dark:border-slate-700 shadow-sm transition-all group-hover:border-blue-200 dark:group-hover:border-blue-900/50">
                              <span className="text-[7px] font-black text-slate-400 uppercase tracking-tighter mb-0.5">T{month}</span>
                              <span className="text-[10px] font-black text-slate-700 dark:text-slate-200">{count}</span>
                            </div>
                          ))}
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="col-span-full py-12 text-center opacity-30 flex flex-col items-center">
                    <Users size={40} strokeWidth={1} />
                    <p className="text-[10px] uppercase font-black tracking-widest mt-3">Chưa có dữ liệu nhân sự</p>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
};

// --- Form Tab Component ---

const DefectForm: React.FC = () => {
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);
  const [images, setImages] = useState<{file: File, preview: string}[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const [formData, setFormData] = useState({
    reporterName: '',
    category: '', 
    area: '',     
    equipmentName: '',
    location: '',
    description: ''
  });

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files).map((file: File) => ({
        file,
        preview: URL.createObjectURL(file)
      }));
      setImages(prev => [...prev, ...newFiles]);
    }
  };

  const removeImage = (index: number) => {
    setImages(prev => {
      const newImages = [...prev];
      URL.revokeObjectURL(newImages[index].preview);
      newImages.splice(index, 1);
      return newImages;
    });
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.category || !formData.area) {
      alert("Vui lòng chọn đầy đủ Phân loại và Khu vực!");
      return;
    }

    setIsSubmitting(true);
    try {
      const filesPayload = await Promise.all(images.map(img => {
        return new Promise<any>((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve({
            dataURL: reader.result as string,
            type: img.file.type,
            name: img.file.name
          });
          reader.readAsDataURL(img.file);
        });
      }));

      const payload = { ...formData, files: filesPayload };
      await fetch(REPORT_WEB_APP_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });
      
      setIsSubmitting(false);
      setShowSuccess(true);
      setTimeout(() => {
        setShowSuccess(false);
        setFormData({ reporterName: '', category: '', area: '', equipmentName: '', location: '', description: '' });
        setImages([]);
      }, 3000);

    } catch (err) {
      alert("Có lỗi xảy ra khi kết nối với máy chủ.");
      setIsSubmitting(false);
    }
  };

  if (showSuccess) return (
    <div className="flex flex-col items-center justify-center min-h-[60vh] px-8 text-center animate-in zoom-in duration-300">
      <div className="w-20 h-20 bg-emerald-100 dark:bg-emerald-900/30 rounded-full flex items-center justify-center text-emerald-600 mb-6 shadow-inner"><CheckCircle2 size={48} /></div>
      <h2 className="text-xl font-bold mb-2">Gửi báo cáo thành công!</h2>
    </div>
  );

  return (
    <div className="animate-in fade-in duration-500">
      <div className="bg-blue-800 p-4 shadow-md text-center"><h2 className="text-white text-[12px] font-black uppercase tracking-widest">CẬP NHẬT TỒN TẠI & HƯ HỎNG</h2></div>
      <div className="w-full px-4 py-8">
        <form onSubmit={handleSubmit} className="space-y-6 pb-20">
          <section><FormLabel required>Họ và tên người phát hiện</FormLabel><input type="text" className="w-full p-2.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-[13px] outline-none" required value={formData.reporterName} onChange={(e) => setFormData({...formData, reporterName: e.target.value})} /></section>
          <section><FormLabel required>Phân loại</FormLabel>
            <div className="mt-2 space-y-1">
              <CustomRadio name="category" value="administrative" label="Quản lý hành chính" checked={formData.category === 'administrative'} onChange={(e) => setFormData({...formData, category: e.target.value})} />
              <CustomRadio name="category" value="construction-equipment" label="Hư hỏng thiết bị công trình" checked={formData.category === 'construction-equipment'} onChange={(e) => setFormData({...formData, category: e.target.value})} />
              <CustomRadio name="category" value="safety" label="An toàn vệ sinh lao động" checked={formData.category === 'safety'} onChange={(e) => setFormData({...formData, category: e.target.value})} />
              <CustomRadio name="category" value="iso-kaizen" label="ISO, KAIZEN 5S, TPM" checked={formData.category === 'iso-kaizen'} onChange={(e) => setFormData({...formData, category: e.target.value})} />
            </div>
          </section>
          <section><FormLabel required>Khu vực</FormLabel>
            <div className="mt-2 flex flex-col gap-1">
              <CustomRadio name="area" value="ialy-hien-huu" label="Ialy hiện hữu" checked={formData.area === 'ialy-hien-huu'} onChange={(e) => setFormData({...formData, area: e.target.value})} />
              <CustomRadio name="area" value="ialy-mo-rong" label="Ialy mở rộng" checked={formData.area === 'ialy-mo-rong'} onChange={(e) => setFormData({...formData, area: e.target.value})} />
              <CustomRadio name="area" value="cua-nhan-nuoc" label="Cửa nhận nước" checked={formData.area === 'cua-nhan-nuoc'} onChange={(e) => setFormData({...formData, area: e.target.value})} />
              <CustomRadio name="area" value="opy-500" label="OPY 500kV" checked={formData.area === 'opy-500'} onChange={(e) => setFormData({...formData, area: e.target.value})} />
            </div>
          </section>
          <section><FormLabel required>Tên thiết bị</FormLabel><input type="text" className="w-full p-2.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-[13px] outline-none" required value={formData.equipmentName} onChange={(e) => setFormData({...formData, equipmentName: e.target.value})} /></section>
          <section><FormLabel required>Địa điểm</FormLabel><input type="text" className="w-full p-2.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-[13px] outline-none" required value={formData.location} onChange={(e) => setFormData({...formData, location: e.target.value})} /></section>
          <section><FormLabel required>Mô tả</FormLabel><textarea rows={3} className="w-full p-2.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-[13px] outline-none resize-none" required value={formData.description} onChange={(e) => setFormData({...formData, description: e.target.value})} /></section>
          <section><FormLabel>Hình ảnh</FormLabel>
            <div onClick={() => fileInputRef.current?.click()} className="border-2 border-dashed border-slate-300 dark:border-slate-700 rounded-2xl p-8 flex flex-col items-center justify-center text-center bg-slate-50 dark:bg-slate-900 cursor-pointer"><input type="file" ref={fileInputRef} multiple accept="image/*" className="hidden" onChange={handleImageChange} /><UploadCloud className="text-blue-500 mb-4" size={32} /><span className="px-6 py-2 bg-blue-600 text-white text-[11px] font-bold rounded-lg uppercase shadow-md">Chọn hình ảnh</span></div>
            <div className="flex flex-wrap gap-2 mt-4">{images.map((img, i) => (<div key={i} className="relative w-20 h-20 rounded-xl overflow-hidden border"><img src={img.preview} className="w-full h-full object-cover" /><button type="button" onClick={() => removeImage(i)} className="absolute top-1 right-1 bg-red-600 text-white p-1 rounded-full"><Trash2 size={12}/></button></div>))}</div>
          </section>
          <div className="flex justify-center pt-6"><button type="submit" disabled={isSubmitting} className="px-12 py-3.5 bg-blue-600 text-white rounded-xl font-black uppercase tracking-widest text-[13px] shadow-xl flex items-center gap-2">{isSubmitting ? <Loader2 className="animate-spin" size={18} /> : null} GỬI BÁO CÁO</button></div>
        </form>
      </div>
    </div>
  );
};

// --- Summary Tab Component ---

const TableCellContent: React.FC<{ value: any, header: string }> = ({ value, header }) => {
  const [previewIndex, setPreviewIndex] = useState<number | null>(null);
  const valStr = value != null ? String(value).trim() : '';
  const lowerHeader = String(header || '').toLowerCase();
  
  const isImageColumn = lowerHeader.includes('hình') || lowerHeader.includes('minh chứng') || lowerHeader.includes('ảnh');
  
  if (!isImageColumn) {
    const isSTT = lowerHeader === 'stt';
    const isTime = lowerHeader.includes('thời gian');
    const isReporter = lowerHeader.includes('người phát hiện');
    const isCategory = lowerHeader.includes('phân loại');
    const isArea = lowerHeader.includes('khu vực');
    
    return (
      <div className={`block break-words leading-normal ${
        isSTT ? 'min-w-[30px] text-center' : 
        isTime ? 'min-w-[80px]' : 
        isReporter ? 'min-w-[100px]' : 
        isCategory || isArea ? 'min-w-[100px]' :
        'min-w-[160px]'
      }`}>
        {value ?? ''}
      </div>
    );
  }

  const potentialUrls = valStr.split(/[,\n\s]+/).map(s => s.trim()).filter(s => s.length > 5);
  const images = potentialUrls.filter(url => url.startsWith('http')).map(url => {
    let displayUrl = url;
    if (url.includes('drive.google.com')) {
      const driveMatch = url.match(/\/d\/(.+?)\/(view|edit|usp)/) || url.match(/id=(.+?)(&|$)/);
      if (driveMatch && driveMatch[1]) {
        displayUrl = `https://drive.google.com/thumbnail?id=${driveMatch[1]}&sz=w400`;
      }
    }
    return { original: url, display: displayUrl };
  });

  if (images.length === 0) return <div className="block break-words min-w-[150px] leading-normal">{valStr}</div>;

  return (
    <>
      <div className="flex flex-wrap gap-2 items-center justify-start p-1 min-w-[200px]">
        {images.map((img, idx) => (
          <div key={idx} onClick={() => setPreviewIndex(idx)} className="w-20 h-20 bg-slate-100 rounded-lg border border-slate-200 overflow-hidden shrink-0 cursor-pointer shadow-sm hover:scale-105 transition-transform"><img src={img.display} className="w-full h-full object-cover" /></div>
        ))}
      </div>
      {previewIndex !== null && (
        <div className="fixed inset-0 z-[100] bg-black/95 flex flex-col items-center justify-center p-4 animate-in fade-in duration-200" onClick={() => setPreviewIndex(null)}>
          <button className="absolute top-6 right-6 text-white p-2 hover:bg-white/10 rounded-full z-[110]"><X size={32} /></button>
          {images.length > 1 && (
            <>
              <button className="absolute left-4 top-1/2 -translate-y-1/2 text-white p-3 bg-white/10 hover:bg-white/20 rounded-full z-[110]" onClick={(e) => { e.stopPropagation(); setPreviewIndex((previewIndex - 1 + images.length) % images.length); }}><ChevronLeft size={32} /></button>
              <button className="absolute right-4 top-1/2 -translate-y-1/2 text-white p-3 bg-white/10 hover:bg-white/20 rounded-full z-[110]" onClick={(e) => { e.stopPropagation(); setPreviewIndex((previewIndex + 1) % images.length); }}><ChevronRight size={32} /></button>
            </>
          )}
          <div className="relative max-w-full max-h-[75vh] flex items-center justify-center" onClick={(e) => e.stopPropagation()}>
            <img src={images[previewIndex].display.includes('thumbnail') ? images[previewIndex].display.replace('w400', 'w1000') : images[previewIndex].original} alt="Full Preview" className="max-w-full max-h-[75vh] object-contain rounded-lg shadow-2xl" />
          </div>
          <div className="mt-8 flex flex-col items-center gap-4 text-white">
            <p className="text-xs font-bold uppercase tracking-widest opacity-60">Hình ảnh {previewIndex + 1} / {images.length}</p>
            <a href={images[previewIndex].original} target="_blank" rel="noreferrer" className="px-8 py-3 bg-blue-600 text-white rounded-full font-bold text-xs uppercase tracking-widest flex items-center gap-2" onClick={(e) => e.stopPropagation()}>Xem link gốc <ExternalLink size={14} /></a>
          </div>
        </div>
      )}
    </>
  );
};

// --- Modal Chỉnh sửa ---

interface EditModalProps {
  sheet: string;
  row: number;
  headers: string[];
  rowData: any[];
  onClose: () => void;
  onSave: () => void;
}

const EditModal: React.FC<EditModalProps> = ({ sheet, row, headers, rowData, onClose, onSave }) => {
  const [isSaving, setIsSaving] = useState(false);
  const [editedData, setEditedData] = useState<any[]>(rowData);

  const handleSave = async () => {
    setIsSaving(true);
    try {
      // Làm sạch dữ liệu trước khi gửi
      const cleanData = editedData.map(v => (v === null || v === undefined) ? '' : v);
      
      const payload = {
        action: 'updateRowData',
        sheetName: sheet,
        row: row,
        rowData: cleanData,
        sheetId: SHEET_ID
      };

      console.log("Sending update via proxy:", payload);

      // Gửi qua server proxy để tránh lỗi CORS và nhận được phản hồi thực tế
      const response = await fetch('/api/proxy-apps-script', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: EDIT_WEB_APP_URL,
          payload: payload
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        console.error("EditModal Proxy error:", errorData);
        throw new Error(errorData.details || "Lỗi server proxy (Edit)");
      }

      const resultText = await response.text();
      console.log("Update result raw:", resultText);
      
      let result;
      try {
        result = JSON.parse(resultText);
      } catch (e) {
        result = resultText;
      }

      alert("Dữ liệu đã được cập nhật thành công! Hệ thống đang làm mới (vui lòng đợi 3 giây)...");
      onSave();
    } catch (err: any) {
      console.error("Save error:", err);
      alert("Có lỗi xảy ra khi lưu: " + err.message);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[100] bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-300">
      <div className="bg-white dark:bg-slate-900 w-full max-w-lg rounded-[2rem] shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        <div className="p-6 bg-blue-600 flex items-center justify-between shrink-0">
          <h2 className="text-white font-black text-xs uppercase tracking-widest">Chỉnh sửa hàng #{row}</h2>
          <button onClick={onClose} className="text-white/80 hover:text-white transition-colors"><X size={20}/></button>
        </div>
        <div className="p-6 overflow-y-auto space-y-4 flex-1 custom-scrollbar">
          {headers.map((h, idx) => {
            const isReadOnly = idx === 0 || idx === 1; // ID và Timestamp thường không nên sửa
            return (
              <div key={idx} className="space-y-1">
                <label className="text-[10px] font-black uppercase text-slate-400 tracking-tighter ml-1">{h}</label>
                <input 
                  type="text" 
                  disabled={isReadOnly}
                  className={`w-full p-3 rounded-xl border text-xs font-semibold outline-none transition-all ${isReadOnly ? 'bg-slate-50 text-slate-400 border-slate-100' : 'bg-white border-slate-200 focus:border-blue-500 focus:ring-2 focus:ring-blue-100 dark:bg-slate-800 dark:border-slate-700 dark:text-white'}`}
                  value={editedData[idx] || ''}
                  onChange={(e) => {
                    const newData = [...editedData];
                    newData[idx] = e.target.value;
                    setEditedData(newData);
                  }}
                />
              </div>
            );
          })}
        </div>
        <div className="p-6 border-t border-slate-100 dark:border-slate-800 bg-slate-50 dark:bg-slate-900/50 flex gap-3 shrink-0">
          <button onClick={onClose} className="flex-1 py-3 text-xs font-bold text-slate-500 bg-white border border-slate-200 rounded-xl hover:bg-slate-50 transition-colors uppercase tracking-widest dark:bg-slate-800 dark:border-slate-700 dark:text-slate-400">Hủy</button>
          <button 
            disabled={isSaving}
            onClick={handleSave} 
            className="flex-1 py-3 text-xs font-black text-white bg-blue-600 rounded-xl shadow-lg shadow-blue-200 hover:bg-blue-700 active:scale-95 transition-all uppercase tracking-widest flex items-center justify-center gap-2"
          >
            {isSaving ? <Loader2 size={16} className="animate-spin"/> : <Save size={16}/>}
            {isSaving ? 'Đang gửi...' : 'Lưu thay đổi'}
          </button>
        </div>
      </div>
    </div>
  );
};

const DefectSummary: React.FC<{ jumpTo?: { sheet: string, row?: number, status?: 'all' | 'processed' | 'pending' | 'nvvh' } | null }> = ({ jumpTo }) => {
  const MAX_COLS = 14; 
  const [isLoading, setIsLoading] = useState(false);
  const [data, setData] = useState<any[]>([]);
  const [activeSheetName, setActiveSheetName] = useState(jumpTo?.sheet || 'Quản lý hành chính');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('all');
  const [selectedStatus, setSelectedStatus] = useState<'all' | 'processed' | 'pending' | 'nvvh'>(jumpTo?.status || 'all');
  const [editTarget, setEditTarget] = useState<{ row: number, data: any[], sheet: string } | null>(null);
  const [selectedRowIndex, setSelectedRowIndex] = useState<number | null>(null);
  const lastScrolledRef = useRef<string | null>(null);
  const fetchIdRef = useRef(0);

  useEffect(() => {
    if (jumpTo) {
      setActiveSheetName(jumpTo.sheet);
      setSelectedStatus(jumpTo.status || 'all');
      setSearchTerm(''); 
    }
  }, [jumpTo]);

  const categories = [
    { name: 'Tất cả', value: 'all' },
    { name: 'Quản lý hành chính', value: 'Quản lý hành chính' }, 
    { name: 'Thiết bị công trình', value: 'Thiết bị công trình' }, 
    { name: 'An toàn vệ sinh lao động', value: 'An toàn vệ sinh lao động' }, 
    { name: 'TPM, Kaizen', value: 'TPM, Kaizen' }
  ];

  const fetchSheetData = useCallback(async () => {
    const currentFetchId = ++fetchIdRef.current;
    setIsLoading(true);
    setData([]); 
    
    try {
      if (activeSheetName === 'all') {
        const promises = CATEGORIES.map(async (cat) => {
          const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(cat)}&t=${Date.now()}`;
          const response = await fetch(url);
          const text = await response.text();
          const match = text.match(/google\.visualization\.Query\.setResponse\((.*)\);/);
          if (match) {
            const json = JSON.parse(match[1]);
            if (json.table && json.table.rows) {
              const rows = json.table.rows.map((row: any) => {
                const cells = row.c.map((cell: any) => {
                  if (!cell) return '';
                  return cell.f != null ? cell.f : (cell.v != null ? String(cell.v) : '');
                });
                // Add sheet name to the row data for reference
                return [...cells, cat];
              });
              return { rows, headers: json.table.cols.map((col: any) => col.label || '') };
            }
          }
          return { rows: [], headers: [] };
        });

        const results = await Promise.all(promises);
        if (currentFetchId === fetchIdRef.current) {
          const allRows = results.flatMap(r => r.rows);
          const baseHeaders = results.find(r => r.headers.length > 0)?.headers || [];
          const headersWithSheet = [...baseHeaders, 'Bộ phận'];
          setData([headersWithSheet, ...allRows]);
        }
      } else {
        const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(activeSheetName)}&t=${Date.now()}`;
        const response = await fetch(url);
        const text = await response.text();
        const match = text.match(/google\.visualization\.Query\.setResponse\((.*)\);/);
        
        if (match && currentFetchId === fetchIdRef.current) {
          const json = JSON.parse(match[1]);
          if (json.table && json.table.rows) {
            const rows = json.table.rows.map((row: any) => 
              row.c.map((cell: any) => {
                if (!cell) return '';
                return cell.f != null ? cell.f : (cell.v != null ? String(cell.v) : '');
              })
            );
            const headers = json.table.cols.map((col: any) => col.label || '');
            setData([headers, ...rows]);
          }
        }
      }
    } catch (err) { 
      if (currentFetchId === fetchIdRef.current) console.error(err); 
    } finally { 
      if (currentFetchId === fetchIdRef.current) setIsLoading(false); 
    }
  }, [activeSheetName]);

  useEffect(() => { fetchSheetData(); }, [fetchSheetData]);

  useEffect(() => {
    const jumpKey = jumpTo ? `${jumpTo.sheet}-${jumpTo.row}` : null;
    
    if (!jumpTo || isLoading || data.length === 0 || lastScrolledRef.current === jumpKey) {
      if (!jumpTo) lastScrolledRef.current = null;
      return;
    }

    const targetId = `row-${jumpTo.row}`;
    let attempts = 0;
    const maxAttempts = 30;
    
    const tryScroll = () => {
      const el = document.getElementById(targetId);
      if (el) {
        lastScrolledRef.current = jumpKey;
        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      } else if (attempts < maxAttempts) {
        attempts++;
        setTimeout(tryScroll, 150);
      }
    };

    const timer = setTimeout(tryScroll, 200);
    return () => clearTimeout(timer);
  }, [jumpTo, isLoading, data, activeSheetName]);

  const safeSearchTerm = String(searchTerm || '').trim().toLowerCase();
  const headers = data[0] || [];
  const rowsWithIndex = data.slice(1).map((row, idx) => {
    const physicalRow = idx + 2;
    return { values: row, index: physicalRow };
  });

  const availableMonths = React.useMemo(() => {
    const months = new Set<string>();
    rowsWithIndex.forEach(item => {
      const dateStr = String(item.values[1] || '');
      const match = dateStr.match(/(\d{2})\/(\d{4})/);
      if (match) months.add(`${match[1]}/${match[2]}`);
    });
    return Array.from(months).sort((a, b) => {
      const [m1, y1] = a.split('/').map(Number);
      const [m2, y2] = b.split('/').map(Number);
      return y2 !== y1 ? y2 - y1 : m2 - m1;
    });
  }, [rowsWithIndex]);
  
  const filteredRows = rowsWithIndex.filter(item => {
    const matchesSearch = safeSearchTerm === '' || 
      item.values.some((cell: any) => String(cell).toLowerCase().includes(safeSearchTerm));
    
    const isProcessed = !!(item.values[11] && String(item.values[11]).trim() !== '');
    const isNVVH = !!(item.values[12] && String(item.values[12]).trim() !== '');
    const matchesStatus = selectedStatus === 'all' || 
      (selectedStatus === 'processed' && isProcessed) || 
      (selectedStatus === 'nvvh' && isNVVH && !isProcessed) ||
      (selectedStatus === 'pending' && !isNVVH && !isProcessed);
      
    const dateStr = String(item.values[1] || '');
    const matchesMonth = selectedMonth === 'all' || dateStr.includes(selectedMonth);

    return matchesSearch && matchesStatus && matchesMonth;
  });

  const exportToExcel = async () => {
    setIsLoading(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(activeSheetName);

      // Add headers
      const headerRow = worksheet.addRow(headers.slice(0, MAX_COLS));
      headerRow.font = { bold: true };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }
      };

      // Set column widths
      worksheet.columns = headers.slice(0, MAX_COLS).map((h: string) => ({
        header: h,
        key: h,
        width: h.toLowerCase().includes('hình') || h.toLowerCase().includes('ảnh') ? 25 : 20
      }));

      // Add data rows and collect image tasks
      const imageTasks: { 
        url: string, 
        fetchUrl: string, 
        rowNumber: number, 
        colIndex: number 
      }[] = [];

      for (const item of filteredRows) {
        const rowData = item.values.slice(0, MAX_COLS);
        const row = worksheet.addRow(rowData);
        row.height = 80; // Set height for images
        row.alignment = { vertical: 'middle', wrapText: true };

        // Handle images
        for (let i = 0; i < rowData.length; i++) {
          const header = headers[i].toLowerCase();
          const isImageCol = header.includes('hình') || header.includes('minh chứng') || header.includes('ảnh');
          const valStr = String(rowData[i] || '').trim();

          if (isImageCol && valStr) {
            const potentialUrls = valStr.split(/[,\n\s]+/).map(s => s.trim()).filter(s => s.length > 5);
            const imageUrls = potentialUrls.filter(url => url.startsWith('http'));

            if (imageUrls.length > 0) {
              const url = imageUrls[0];
              let fetchUrl = url;
              
              // Enhanced Drive detection
              if (url.includes('drive.google.com') || url.includes('docs.google.com')) {
                const driveMatch = url.match(/\/d\/(.+?)\/(view|edit|usp|copy)/) || 
                                  url.match(/[?&]id=(.+?)(&|$)/) ||
                                  url.match(/\/file\/d\/(.+?)\//);
                if (driveMatch && driveMatch[1]) {
                  fetchUrl = `https://drive.google.com/thumbnail?id=${driveMatch[1]}&sz=w800`;
                }
              }
              
              imageTasks.push({ url, fetchUrl, rowNumber: row.number, colIndex: i });
            }
          }
        }
      }

      // Process images in parallel chunks to avoid overwhelming the server
      const CHUNK_SIZE = 5;
      for (let i = 0; i < imageTasks.length; i += CHUNK_SIZE) {
        const chunk = imageTasks.slice(i, i + CHUNK_SIZE);
        await Promise.all(chunk.map(async (task) => {
          let objectUrl = null;
          try {
            let response;
            let blob;
            
            // Helper for fetch with timeout
            const fetchWithTimeout = async (url: string, timeout = 10000) => {
              const controller = new AbortController();
              const id = setTimeout(() => controller.abort(), timeout);
              try {
                const res = await fetch(url, { signal: controller.signal });
                clearTimeout(id);
                return res;
              } catch (err) {
                clearTimeout(id);
                throw err;
              }
            };

            // Try local proxy first
            try {
              const proxyUrl = `/api/proxy-image?url=${encodeURIComponent(task.fetchUrl)}`;
              response = await fetchWithTimeout(proxyUrl);
              if (!response.ok) throw new Error(`Local proxy failed with ${response.status}`);
              blob = await response.blob();
            } catch (localProxyErr) {
              console.warn("Local proxy failed, trying public fallback...", localProxyErr);
              // Fallback to a public Google proxy
              const fallbackUrl = `https://images1-focus-opensocial.googleusercontent.com/gadgets/proxy?container=focus&refresh=2592000&url=${encodeURIComponent(task.fetchUrl)}`;
              response = await fetchWithTimeout(fallbackUrl);
              if (!response.ok) throw new Error(`Public fallback failed with ${response.status}`);
              blob = await response.blob();
            }
            
            if (blob.size < 100) throw new Error("Invalid image size");

            const arrayBuffer = await blob.arrayBuffer();
            
            // Determine valid extension for ExcelJS
            let extension: 'png' | 'jpeg' | 'gif' = 'png';
            const mimeType = blob.type.toLowerCase();
            if (mimeType.includes('png')) extension = 'png';
            else if (mimeType.includes('gif')) extension = 'gif';
            else if (mimeType.includes('jpg') || mimeType.includes('jpeg')) extension = 'jpeg';
            else extension = 'jpeg';

            // Get image dimensions with timeout
            const imgObj = new Image();
            objectUrl = URL.createObjectURL(blob);
            imgObj.src = objectUrl;
            
            await Promise.race([
              new Promise((resolve) => {
                imgObj.onload = resolve;
                imgObj.onerror = resolve;
              }),
              new Promise((_, reject) => setTimeout(() => reject(new Error("Image load timeout")), 5000))
            ]);
            
            const naturalWidth = imgObj.width || 100;
            const naturalHeight = imgObj.height || 100;

            const maxWidth = 180; 
            const maxHeight = 100;
            const ratio = Math.min(maxWidth / naturalWidth, maxHeight / naturalHeight, 1);
            const finalWidth = naturalWidth * ratio;
            const finalHeight = naturalHeight * ratio;

            const imageId = workbook.addImage({
              buffer: arrayBuffer,
              extension: extension,
            });

            worksheet.addImage(imageId, {
              tl: { col: task.colIndex + 0.05, row: task.rowNumber - 0.95 } as any,
              ext: { width: finalWidth, height: finalHeight },
              editAs: 'oneCell'
            });
            
            worksheet.getRow(task.rowNumber).getCell(task.colIndex + 1).value = '';
          } catch (e) {
            console.error("Failed to fetch image for excel:", task.fetchUrl, e);
            worksheet.getRow(task.rowNumber).getCell(task.colIndex + 1).value = 'Lỗi ảnh: ' + task.url;
          } finally {
            if (objectUrl) URL.revokeObjectURL(objectUrl);
          }
        }));
      }

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `${activeSheetName}_${new Date().toLocaleDateString()}.xlsx`);
    } catch (error) {
      console.error("Excel export error:", error);
      alert("Có lỗi khi xuất Excel. Vui lòng thử lại.");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="flex flex-col h-full animate-in fade-in duration-500 overflow-hidden">
      <div className="bg-blue-800 p-4 shadow-xl flex items-center justify-between shrink-0">
        <div className="flex items-center gap-3">
          <TableProperties className="text-white" size={20} />
          <h1 className="text-white font-black text-[12px] uppercase">TỔNG HỢP DỮ LIỆU</h1>
        </div>
        <div className="flex items-center gap-2">
          <button 
            onClick={exportToExcel} 
            disabled={isLoading || data.length === 0}
            className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-[10px] font-bold uppercase tracking-widest hover:bg-emerald-700 transition-all disabled:opacity-50"
          >
            <Download size={14} />
            Xuất Excel
          </button>
          <button onClick={fetchSheetData} className="p-2 text-white">
            <RefreshCw size={18} className={isLoading ? 'animate-spin' : ''} />
          </button>
        </div>
      </div>
      <div className="p-4 bg-white dark:bg-slate-900 border-b space-y-4 shrink-0">
        <div className="flex gap-2 overflow-x-auto pb-2 no-scrollbar">
          {categories.map((cat) => (
            <button 
              key={cat.value} 
              onClick={() => {
                setActiveSheetName(cat.value);
                setSelectedMonth('all');
                setSelectedStatus('all');
              }} 
              className={`px-4 py-2 rounded-xl whitespace-nowrap text-[10px] font-black uppercase tracking-widest transition-all ${activeSheetName === cat.value ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-500'}`}
            >
              {cat.name}
            </button>
          ))}
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
          <div className="relative">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input 
              type="text" 
              placeholder="Tìm kiếm nhanh..." 
              className="w-full pl-9 pr-4 py-2.5 bg-slate-100 dark:bg-slate-800 rounded-xl text-xs outline-none" 
              value={searchTerm} 
              onChange={(e) => setSearchTerm(e.target.value)} 
            />
          </div>
          
          <div className="flex gap-2 md:col-span-2">
            <select 
              className="flex-1 p-2.5 bg-slate-100 dark:bg-slate-800 rounded-xl text-[10px] font-bold uppercase outline-none border-none"
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
            >
              <option value="all">TẤT CẢ THÁNG</option>
              {availableMonths.map(m => <option key={m} value={m}>THÁNG {m}</option>)}
            </select>
            
            <select 
              className="flex-1 p-2.5 bg-slate-100 dark:bg-slate-800 rounded-xl text-[10px] font-bold uppercase outline-none border-none"
              value={selectedStatus}
              onChange={(e) => setSelectedStatus(e.target.value as any)}
            >
              <option value="all">TẤT CẢ TRẠNG THÁI</option>
              <option value="nvvh"> ĐANG XỬ LÝ</option>
              <option value="pending">CHƯA XỬ LÝ</option>
              <option value="processed">HOÀN THÀNH</option>
            </select>

            <button 
              onClick={() => {
                if (selectedRowIndex !== null) {
                  const item = filteredRows.find(r => r.index === selectedRowIndex);
                  if (item) {
                    setEditTarget({ 
                      row: item.index, 
                      data: item.values, 
                      sheet: activeSheetName === 'all' ? item.values[item.values.length - 1] : activeSheetName 
                    });
                  }
                }
              }}
              disabled={selectedRowIndex === null}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-blue-700 transition-all disabled:opacity-30 disabled:grayscale disabled:cursor-not-allowed shadow-lg shadow-blue-200 dark:shadow-none"
            >
              <Edit3 size={14} />
              <span className="hidden sm:inline">Sửa dòng chọn</span>
            </button>
          </div>
        </div>
      </div>
      <div className="flex-1 min-h-0 p-4 flex flex-col">
        {isLoading ? (
          <div className="h-full flex flex-col items-center justify-center space-y-3">
            <Loader2 size={48} className="animate-spin text-blue-500" />
          </div>
        ) : data.length > 0 ? (
          <div className="flex-1 flex flex-col shadow-2xl rounded-3xl bg-white dark:bg-slate-900 border border-slate-100 dark:border-slate-800 overflow-hidden relative">
            <div className="flex-1 overflow-auto custom-scrollbar pb-2">
              <table className="border-separate border-spacing-0 table-fixed" style={{ width: 'max-content', minWidth: '100%' }}>
                <thead className="bg-slate-50 dark:bg-slate-800 sticky top-0 z-20 shadow-sm">
                  <tr>
                    <th style={{ width: '50px' }} className="px-4 py-4 text-center text-[9px] font-black text-slate-400 uppercase tracking-widest border-r border-b border-slate-100 dark:border-slate-700 bg-slate-50 dark:bg-slate-800">Chọn</th>
                    {headers.slice(0, MAX_COLS).map((h: string, idx: number) => {
                      const lowerH = h.toLowerCase();
                      const isSTT = lowerH === 'stt';
                      const isTime = lowerH.includes('thời gian');
                      const isReporter = lowerH.includes('người phát hiện');
                      const isCategory = lowerH.includes('phân loại');
                      const isArea = lowerH.includes('khu vực');
                      const isImage = lowerH.includes('hình') || lowerH.includes('ảnh') || lowerH.includes('minh chứng');
                      
                      let width = '180px';
                      if (isSTT) width = '50px';
                      else if (isTime) width = '100px';
                      else if (isReporter) width = '120px';
                      else if (isCategory || isArea) width = '120px';
                      else if (isImage) width = '240px';

                      return (
                        <th key={idx} style={{ width }} className="px-4 py-4 text-left text-[9px] font-black text-slate-400 uppercase tracking-widest border-r border-b border-slate-100 dark:border-slate-700 last:border-r-0 bg-slate-50 dark:bg-slate-800 whitespace-nowrap">
                          {h}
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-800/50">
                  {filteredRows.map((item) => {
                    const isJumped = jumpTo?.row === item.index && jumpTo?.sheet === activeSheetName;
                    return (
                      <tr
                        key={item.index}
                        id={`row-${item.index}`}
                        onClick={() => setSelectedRowIndex(item.index)}
                        className={`${isJumped ? 'bg-yellow-100 dark:bg-yellow-900/40 ring-2 ring-yellow-400 ring-inset' : ''} ${selectedRowIndex === item.index ? 'bg-blue-50 dark:bg-blue-900/20' : ''} transition-all duration-500 cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/50`}
                      >
                        <td className="px-4 py-4 text-center border-r border-slate-100 dark:border-slate-800/50">
                          <div className={`w-4 h-4 rounded-full border-2 flex items-center justify-center transition-all ${selectedRowIndex === item.index ? 'border-blue-600 bg-blue-600' : 'border-slate-300'}`}>
                            {selectedRowIndex === item.index && <div className="w-1.5 h-1.5 bg-white rounded-full" />}
                          </div>
                        </td>
                        {item.values.slice(0, MAX_COLS).map((cell: any, cIdx: number) => (
                          <td key={cIdx} className="px-4 py-4 text-xs text-slate-700 dark:text-slate-300 border-r border-slate-100 dark:border-slate-800/50 last:border-r-0 align-top overflow-hidden">
                            <TableCellContent value={cell} header={headers[cIdx]} />
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          <div className="h-full flex flex-col items-center justify-center opacity-40">
            <FileSpreadsheet size={80} />
          </div>
        )}
      </div>

      {editTarget && (
        <EditModal 
          sheet={editTarget.sheet} 
          row={editTarget.row} 
          headers={data[0].slice(0, MAX_COLS)} 
          rowData={editTarget.data.slice(0, MAX_COLS)} 
          onClose={() => setEditTarget(null)}
          onSave={() => {
            setEditTarget(null);
            setTimeout(fetchSheetData, 3500); // Tăng lên 3.5 giây để chắc chắn server đã cập nhật
          }}
        />
      )}
    </div>
  );
};

// --- Tab: Xử lý tồn tại ---

const ProcessingForm: React.FC = () => {
  const [showSuccess1, setShowSuccess1] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isLoadingList, setIsLoadingList] = useState(false);
  const [defectList, setDefectList] = useState<any[]>([]);
  const [images, setImages] = useState<{file: File, preview: string}[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [formData, setFormData] = useState({ sheet: '', row: '', tinhTrang: '', ghiChu: '', NVVH: '' });

  const categories = [
    { label: 'Quản lý hành chính', value: 'Quản lý hành chính' },
    { label: 'Thiết bị công trình', value: 'Thiết bị công trình' },
    { label: 'An toàn vệ sinh lao động', value: 'An toàn vệ sinh lao động' },
    { label: 'TPM, Kaizen', value: 'TPM, Kaizen' }
  ];

  useEffect(() => {
    if (formData.sheet) fetchDefects();
    else setDefectList([]);
  }, [formData.sheet]);

  const fetchDefects = async () => {
    setIsLoadingList(true);
    try {
      const payload = { action: 'getPendingList', sheetName: formData.sheet, sheetId: SHEET_ID };
      // Sử dụng mode no-cors cho POST request nếu cần, nhưng GET/POST JSON đơn giản thường ok với Apps Script nếu thiết lập đúng.
      const response = await fetch(PROCESS_WEB_APP_URL, { 
        method: 'POST', 
        body: JSON.stringify(payload) 
      });
      const data = await response.json();
      setDefectList(data);
    } catch (err) { console.error(err); } finally { setIsLoadingList(false); }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.sheet || !formData.row) return alert("Vui lòng chọn đầy đủ thông tin!");
    setIsSubmitting(true);
    try {
      const filesPayload = await Promise.all(images.map(img => {
        return new Promise<any>((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve({ data: (reader.result as string).split(',')[1], type: img.file.type, name: img.file.name });
          reader.readAsDataURL(img.file);
        });
      }));
      const payload = { action: 'uploadFiles', form: { sheetId: SHEET_ID, sheet: formData.sheet, row: parseInt(formData.row), tinhTrang: formData.tinhTrang, ghiChu: formData.ghiChu, NVVH: formData.NVVH, files: filesPayload } };
      
      // Gửi qua server proxy
      const response = await fetch('/api/proxy-apps-script', { 
        method: 'POST', 
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          url: PROCESS_WEB_APP_URL,
          payload: payload
        }) 
      });

      if (!response.ok) {
        const errorData = await response.json();
        console.error("ProcessingForm Proxy error:", errorData);
        throw new Error(errorData.details || "Lỗi server proxy (Processing)");
      }

      setShowSuccess1(true);
      setFormData({ sheet: '', row: '', tinhTrang: '', ghiChu: '', NVVH: '' });
      setImages([]);
      setDefectList([]);
    } catch (err: any) { alert("Lỗi: " + err.message); } finally { setIsSubmitting(false); }
  };
if (showSuccess1) {
  return (
    <div className="flex flex-col items-center justify-center min-h-[60vh] px-8 text-center animate-in zoom-in duration-300">
      <div className="w-20 h-20 bg-emerald-100 rounded-full flex items-center justify-center text-emerald-600 mb-6 shadow-inner">
        <CheckCircle2 size={48} />
      </div>
      <h2 className="text-xl font-bold mb-2">
        Cập nhật tồn tại đã xử lý thành công!
      </h2>
      <button
        onClick={() => setShowSuccess1(false)}
        className="mt-4 px-6 py-2 bg-blue-600 text-white rounded-lg"
      >
        Quay lại
      </button>
    </div>
  );
}
  return (
    <div className="animate-in slide-in-from-right duration-500 w-full px-4 py-8">
      <div className="flex flex-col items-center mb-8"><h1 className="text-xl font-bold text-blue-600 flex items-center gap-2">📸 Cập nhật xử lý tồn tại</h1></div>
      <form onSubmit={handleSubmit} className="space-y-6 pb-24">
        <section><FormLabel icon="📄">Chọn loại</FormLabel><select className="w-full p-3 rounded-lg border border-slate-300 bg-white" value={formData.sheet} onChange={(e) => setFormData({...formData, sheet: e.target.value, row: ''})}><option value="">-- Chọn loại --</option>{categories.map(c => <option key={c.value} value={c.value}>{c.label}</option>)}</select></section>
        <section><FormLabel icon="📁">Chọn tồn tại</FormLabel><div className="relative"><select className="w-full p-3 rounded-lg border border-slate-300 bg-white disabled:bg-slate-50" value={formData.row} onChange={(e) => setFormData({...formData, row: e.target.value})} disabled={!formData.sheet || isLoadingList}><option value="">-- Chọn tồn tại --</option>{defectList.map((item, idx) => (<option key={idx} value={item.row}>{`[${item.colE}] - ${item.colF} - ${item.colG}`}</option>))}</select>{isLoadingList && <Loader2 className="absolute right-3 top-3.5 animate-spin text-blue-500" size={18} />}</div></section>
        <section><FormLabel icon="⚠️">Tình trạng</FormLabel><input type="text" className="w-full p-3 rounded-lg border border-slate-300 
             bg-white text-slate-900 
             placeholder:text-slate-400 
             focus:outline-none focus:ring-2 focus:ring-blue-500" value={formData.tinhTrang} onChange={(e) => setFormData({...formData, tinhTrang: e.target.value})} /></section>
        <section><FormLabel icon="📝">Ghi chú</FormLabel><textarea rows={3} className="w-full p-3 rounded-lg border border-slate-300 
             bg-white text-slate-900 
             placeholder:text-slate-400 
             resize-none
             focus:outline-none focus:ring-2 focus:ring-blue-500" value={formData.ghiChu} onChange={(e) => setFormData({...formData, ghiChu: e.target.value})} /></section>
        <section><FormLabel icon="📝">NVVH xử lý ( khi xác nhận kết thúc tồn tại thì không điền ô này)</FormLabel><textarea rows={3} className="w-full p-3 rounded-lg border border-slate-300 
             bg-white text-slate-900 
             placeholder:text-slate-400 
             resize-none
             focus:outline-none focus:ring-2 focus:ring-blue-500" value={formData.NVVH} onChange={(e) => setFormData({...formData, NVVH: e.target.value})} /></section>
        <section><FormLabel icon="🖼️">Hình ảnh minh chứng ( bắt buộc phải có hình ảnh để kết thúc tồn tại)</FormLabel><div className="flex items-center gap-4 p-3 border border-slate-300 rounded-lg bg-white"><button type="button" onClick={() => fileInputRef.current?.click()} className="px-4 py-1.5 bg-slate-100 border border-slate-300 rounded text-sm font-medium">Chọn tệp</button><span className="text-sm text-slate-500">{images.length > 0 ? `${images.length} tệp` : "Chưa chọn"}</span><input type="file" ref={fileInputRef} multiple className="hidden" onChange={(e) => { if(e.target.files) setImages([...images, ...Array.from(e.target.files).map((f: File) => ({file: f, preview: URL.createObjectURL(f)}))]); }} /></div><div className="flex flex-wrap gap-2 mt-4">{images.map((img, i) => (<div key={i} className="relative w-24 h-24 rounded-lg border-2 overflow-hidden shadow-sm hover:scale-105 transition-transform"><img src={img.preview} className="w-full h-full object-cover" /><button type="button" onClick={() => setImages(images.filter((_, idx) => idx !== i))} className="absolute top-1 right-1 bg-red-500 text-white rounded-full p-1"><Trash2 size={12}/></button></div>))}</div></section>
        <button type="submit" disabled={isSubmitting} className="w-full bg-blue-600 text-white py-4 rounded-xl font-black uppercase tracking-widest text-sm shadow-xl shadow-blue-200 active:scale-[0.98] transition-all">{isSubmitting ? <Loader2 size={24} className="animate-spin inline" /> : "Gửi dữ liệu"}</button>
      </form>
    </div>
  );
};

// --- App Shell ---

const App: React.FC = () => {
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'report' | 'processing' | 'summary'>('dashboard');
  const [currentTime, setCurrentTime] = useState(new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }));
  const [summaryJump, setSummaryJump] = useState<{ sheet: string, row?: number, status?: 'all' | 'processed' | 'pending' | 'nvvh' } | null>(null);

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })), 10000);
    return () => clearInterval(timer);
  }, []);

  const toggleDarkMode = () => {
    setIsDarkMode(!isDarkMode);
    document.documentElement.classList.toggle('dark');
  };

  const handleActivityClick = (sheet: string, row: number) => {
    setSummaryJump({ sheet, row, status: 'all' });
    setActiveTab('summary');
  };

  const handleStatClick = (sheet: string, status: 'all' | 'processed' | 'pending' | 'nvvh') => {
    setSummaryJump({ sheet, status });
    setActiveTab('summary');
  };

  return (
    <div 
      className={`h-screen relative flex flex-col pb-24 shadow-2xl overflow-hidden w-full bg-cover bg-center bg-no-repeat transition-all duration-500`}
      style={{ 
        backgroundImage: isDarkMode 
          ? `linear-gradient(rgba(15, 23, 42, 0.85), rgba(15, 23, 42, 0.85)), url('https://images.unsplash.com/photo-1553095066-5014bc7b7f2d?auto=format&fit=crop&q=80&w=2000')`
          : `linear-gradient(rgba(248, 250, 252, 0.75), rgba(248, 250, 252, 0.75)), url('https://images.unsplash.com/photo-1553095066-5014bc7b7f2d?auto=format&fit=crop&q=80&w=2000')`
      }}
    >
      <button onClick={toggleDarkMode} className="fixed top-6 right-6 w-10 h-10 bg-white/80 dark:bg-slate-800/80 backdrop-blur-md rounded-full shadow-lg flex items-center justify-center border border-white dark:border-slate-700 z-50 transition-all active:scale-90">
        {isDarkMode ? <Sun className="text-amber-400" size={18} /> : <Moon className="text-slate-600" size={18} />}
      </button>

      <div className="flex-1 min-h-0 overflow-hidden w-full flex flex-col">
        {activeTab === 'dashboard' && (
          <div className="flex-1 overflow-y-auto custom-scrollbar">
            <Dashboard isDarkMode={isDarkMode} onActivityClick={handleActivityClick} onStatClick={handleStatClick} />
          </div>
        )}
        {activeTab === 'report' && (
          <div className="flex-1 overflow-y-auto custom-scrollbar">
            <DefectForm />
          </div>
        )}
        {activeTab === 'processing' && (
          <div className="flex-1 overflow-y-auto custom-scrollbar">
            <ProcessingForm />
          </div>
        )}
        {activeTab === 'summary' && <DefectSummary jumpTo={summaryJump} />}
      </div>

      <nav className="fixed bottom-0 left-0 right-0 mx-auto h-20 bg-white/95 dark:bg-slate-900/95 backdrop-blur-xl border-t border-slate-200 dark:border-slate-800 flex items-center justify-around px-2 z-50 w-full md:max-w-5xl md:bottom-6 md:rounded-3xl md:shadow-2xl md:border">
        <button onClick={() => { setActiveTab('dashboard'); setSummaryJump(null); }} className={`flex flex-col items-center gap-1.5 flex-1 transition-all ${activeTab === 'dashboard' ? 'text-blue-600 scale-110' : 'text-slate-400'}`}>
          <div className={`p-2 rounded-xl ${activeTab === 'dashboard' ? 'bg-blue-600/10' : ''}`}><LayoutDashboard size={22} /></div>
          <span className="text-[8px] font-black uppercase tracking-tight">Tổng quan</span>
        </button>
        <button onClick={() => { setActiveTab('report'); setSummaryJump(null); }} className={`flex flex-col items-center gap-1.5 flex-1 transition-all ${activeTab === 'report' ? 'text-blue-600 scale-110' : 'text-slate-400'}`}>
          <div className={`p-2 rounded-xl ${activeTab === 'report' ? 'bg-blue-600/10' : ''}`}><ClipboardList size={22} /></div>
          <span className="text-[8px] font-black uppercase tracking-tight">Cập nhật tồn tại</span>
        </button>
        <button onClick={() => { setActiveTab('processing'); setSummaryJump(null); }} className={`flex flex-col items-center gap-1.5 flex-1 transition-all ${activeTab === 'processing' ? 'text-blue-600 scale-110' : 'text-slate-400'}`}>
          <div className={`p-2 rounded-xl ${activeTab === 'processing' ? 'bg-blue-600/10' : ''}`}><Camera size={22} /></div>
          <span className="text-[8px] font-black uppercase tracking-tight">Cập nhật xử lý</span>
        </button>
        <button onClick={() => setActiveTab('summary')} className={`flex flex-col items-center gap-1.5 flex-1 transition-all ${activeTab === 'summary' ? 'text-blue-600 scale-110' : 'text-slate-400'}`}>
          <div className={`p-2 rounded-xl ${activeTab === 'summary' ? 'bg-blue-600/10' : ''}`}><TableProperties size={22} /></div>
          <span className="text-[8px] font-black uppercase tracking-tight">Bảng tổng hợp</span>
        </button>
      </nav>
    </div>
  );
};
export default App;

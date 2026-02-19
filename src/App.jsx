import React, { useState, useMemo, useEffect } from 'react';
import { 
  Users, 
  ShoppingBag, 
  CreditCard, 
  Search, 
  Filter, 
  Download,
  AlertCircle,
  CheckCircle2,
  FileText,
  FileSpreadsheet,
  Hash,
  Calendar,
  RefreshCw,
  ReceiptText,
  History,
  TrendingUp,
  Info,
  Package
} from 'lucide-react';

const App = () => {
  const [allTransactions, setAllTransactions] = useState([]); 
  const [loading, setLoading] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [scriptLoaded, setScriptLoaded] = useState(false);

  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.async = true;
    script.onload = () => setScriptLoaded(true);
    document.body.appendChild(script);
    return () => {
      if (document.body.contains(script)) {
        document.body.removeChild(script);
      }
    };
  }, []);

  const parseDate = (dateVal) => {
    if (!dateVal) return null;
    if (typeof dateVal === 'number') {
      return new Date((dateVal - 25569) * 86400 * 1000);
    }
    const dateStr = dateVal.toString().trim();
    const parts = dateStr.split(/[/-]/);
    if (parts.length === 3) {
      let day, month, year;
      if (parts[0].length === 4) {
        year = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        day = parseInt(parts[2], 10);
      } else {
        day = parseInt(parts[0], 10);
        month = parseInt(parts[1], 10) - 1;
        year = parseInt(parts[2], 10);
      }
      if (year > 2400) year -= 543; 
      const date = new Date(year, month, day);
      return isNaN(date.getTime()) ? null : date;
    }
    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      if (d.getFullYear() > 2400) d.setFullYear(d.getFullYear() - 543);
      return d;
    }
    return null;
  };

  const formatThaiDate = (date) => {
    if (!date || isNaN(date.getTime())) return '-';
    const d = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const y = date.getFullYear() + 543;
    return `${d}/${m}/${y}`;
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file || !window.XLSX) return;

    setLoading(true);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = window.XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = window.XLSX.utils.sheet_to_json(worksheet);
        
        const processed = jsonData.filter(row => {
          const amount = parseFloat(row['ยอดชำระ'] || row['PaidAmount'] || 0);
          const type = (row['ประเภทชำระ'] || row['PaymentType'] || '').toString();
          const itemName = (row['รายการ'] || row['Item'] || '').toString();
          const isDebt = type.includes('ชำระหนี้');
          const isCutCourse = itemName.includes('[ตัดคอร์ส]');
          return amount > 0 && !isDebt && !isCutCourse;
        }).map(row => ({
          ...row,
          _parsedDate: parseDate(row['วันที่ขาย'] || row['Date'] || row['วันที่'])
        }));

        setAllTransactions(processed);
        
        if (processed.length > 0) {
          const dates = processed.map(t => t._parsedDate).filter(d => d);
          if (dates.length > 0) {
            const minDate = new Date(Math.min(...dates));
            const maxDate = new Date(Math.max(...dates));
            setStartDate(minDate.toISOString().split('T')[0]);
            setEndDate(maxDate.toISOString().split('T')[0]);
          }
        }
      } catch (err) {
        console.error("Error processing file:", err);
      } finally {
        setLoading(false);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const { filteredCustomers, totalTransactionsCount } = useMemo(() => {
    const start = startDate ? new Date(startDate) : null;
    const end = endDate ? new Date(endDate) : null;
    if (end) end.setHours(23, 59, 59, 999);

    const customerMap = {};
    let txCount = 0;

    allTransactions.forEach(row => {
      if (start && row._parsedDate && row._parsedDate < start) return;
      if (end && row._parsedDate && row._parsedDate > end) return;

      txCount++;
      const hn = row['HN']?.toString();
      if (!hn) return;

      const itemName = (row['รายการ'] || row['Item'] || '').toString().toLowerCase();
      const quantity = parseFloat(row['จำนวน'] || row['Quantity'] || 0);
      
      // Keywords that definitely mean a course
      const hasKeywords = itemName.includes('คอร์ส') || itemName.includes('course') || itemName.includes('package') || itemName.includes('pkg');
      
      // Special Logic: Botox or 1U items are exempt from "Quantity > 1" being a course
      const isBotoxOrUnit = itemName.includes('botox') || itemName.includes('1u');
      const hasHighQty = isBotoxOrUnit ? false : quantity > 1;

      if (!customerMap[hn]) {
        customerMap[hn] = {
          hn: hn,
          name: row['ชื่อ-สกุล'] || row['CustomerName'] || 'ไม่ระบุชื่อ',
          phone: row['เบอร์โทร'] || row['Phone'] || '-',
          totalPaid: 0,
          hasCourseItem: false,
          hasHighQty: false,
          lastDate: row._parsedDate,
          visitCount: 0,
          allTransactions: []
        };
      }

      customerMap[hn].totalPaid += parseFloat(row['ยอดชำระ'] || row['PaidAmount']) || 0;
      customerMap[hn].visitCount += 1;
      customerMap[hn].allTransactions.push({
        item: row['รายการ'] || row['Item'],
        qty: quantity,
        isCourse: hasKeywords || hasHighQty,
        date: row._parsedDate,
        amount: parseFloat(row['ยอดชำระ'] || row['PaidAmount']) || 0
      });

      if (hasKeywords) customerMap[hn].hasCourseItem = true;
      if (hasHighQty) customerMap[hn].hasHighQty = true;
      
      if (row._parsedDate && (!customerMap[hn].lastDate || row._parsedDate > customerMap[hn].lastDate)) {
        customerMap[hn].lastDate = row._parsedDate;
      }
    });

    const customers = Object.values(customerMap).map(customer => {
      let reason = "";
      let isCourseType = false;

      if (customer.hasCourseItem) {
        isCourseType = true;
        reason = "พบชื่อรายการ 'คอร์ส/Package'";
      } else if (customer.hasHighQty) {
        isCourseType = true;
        reason = "พบรายการ 'จำนวน > 1'";
      } else {
        isCourseType = false;
        reason = customer.visitCount >= 2 
          ? `ซื้อรายครั้งหลายรายการ (${customer.visitCount} ครั้ง)` 
          : "ชำระ 1 ครั้ง (สินค้าทั่วไป)";
      }

      return {
        ...customer,
        customerType: isCourseType ? 'ลูกค้าคอร์ส' : 'ลูกค้าซื้อครั้งเดียว',
        categoryReason: reason
      };
    });

    return { filteredCustomers: customers, totalTransactionsCount: txCount };
  }, [allTransactions, startDate, endDate]);

  const stats = useMemo(() => {
    const courseCount = filteredCustomers.filter(c => c.customerType === 'ลูกค้าคอร์ส').length;
    const singleCount = filteredCustomers.filter(c => c.customerType === 'ลูกค้าซื้อครั้งเดียว').length;
    return {
      totalCustomers: filteredCustomers.length,
      totalTransactions: totalTransactionsCount,
      course: courseCount,
      single: singleCount,
      revenue: filteredCustomers.reduce((acc, curr) => acc + curr.totalPaid, 0)
    };
  }, [filteredCustomers, totalTransactionsCount]);

  const displayList = useMemo(() => {
    return filteredCustomers.filter(c => {
      const searchStr = searchTerm.toLowerCase();
      const matchesSearch = (c.name?.toLowerCase() || '').includes(searchStr) || 
                            (c.hn?.toLowerCase() || '').includes(searchStr) || 
                            (c.phone?.toLowerCase() || '').includes(searchStr);
      const matchesFilter = filterType === 'all' || 
                            (filterType === 'course' && c.customerType === 'ลูกค้าคอร์ส') ||
                            (filterType === 'single' && c.customerType === 'ลูกค้าซื้อครั้งเดียว');
      return matchesSearch && matchesFilter;
    });
  }, [filteredCustomers, searchTerm, filterType]);

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-800">
      <div className="max-w-full mx-auto">
        
        {/* Header */}
        <header className="mb-6 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div>
            <h1 className="text-2xl font-bold text-slate-900 flex items-center gap-2">
              <TrendingUp className="text-blue-600" />
              วิเคราะห์กลุ่มลูกค้าและรายการที่ซื้อ
            </h1>
            <p className="text-slate-500 text-sm">ตรวจสอบรายการสินค้าเชิงลึก (Botox/1U จะไม่นับเป็นคอร์สแม้จำนวนมาก)</p>
          </div>
          
          <div className="flex items-center gap-3">
            <label className={`cursor-pointer ${!scriptLoaded ? 'bg-slate-300' : 'bg-blue-600 hover:bg-blue-700'} text-white px-6 py-2.5 rounded-xl font-bold transition flex items-center gap-2 shadow-lg shadow-blue-100`}>
              <Download size={18} />
              อัปโหลดรายงาน Excel
              <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleFileUpload} disabled={!scriptLoaded} />
            </label>
          </div>
        </header>

        {/* Stats Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 mb-8">
          <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 border-t-4 border-t-slate-400">
            <div className="text-slate-400 text-[10px] font-bold uppercase mb-1">ลูกค้าที่จ่ายเงิน</div>
            <div className="text-2xl font-bold">{stats.totalCustomers.toLocaleString()} <span className="text-xs font-normal text-slate-400">คน</span></div>
          </div>
          <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 border-t-4 border-t-blue-500">
            <div className="text-blue-400 text-[10px] font-bold uppercase mb-1">จำนวนบิล/รายการ</div>
            <div className="text-2xl font-bold text-blue-700">{stats.totalTransactions.toLocaleString()} <span className="text-xs font-normal text-slate-400">ครั้ง</span></div>
          </div>
          <div className="bg-purple-50 p-5 rounded-2xl shadow-sm border border-purple-100 border-t-4 border-t-purple-500">
            <div className="text-purple-600 text-[10px] font-bold uppercase mb-1">กลุ่มลูกค้าคอร์ส</div>
            <div className="text-2xl font-bold text-purple-700">{stats.course.toLocaleString()} <span className="text-xs font-normal text-slate-400">คน</span></div>
          </div>
          <div className="bg-blue-50 p-5 rounded-2xl shadow-sm border border-blue-100 border-t-4 border-t-blue-600">
            <div className="text-blue-600 text-[10px] font-bold uppercase mb-1">ซื้อครั้งเดียว</div>
            <div className="text-2xl font-bold text-blue-800">{stats.single.toLocaleString()} <span className="text-xs font-normal text-slate-400">คน</span></div>
          </div>
          <div className="bg-emerald-600 p-5 rounded-2xl shadow-lg text-white">
            <div className="text-white/70 text-[10px] font-bold uppercase mb-1">รายได้รวม</div>
            <div className="text-2xl font-bold">฿{stats.revenue.toLocaleString(undefined, { maximumFractionDigits: 0 })}</div>
          </div>
        </div>

        {/* Filters */}
        <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 mb-6 flex flex-wrap items-center gap-4">
          <div className="flex items-center gap-3 bg-slate-50 p-1 px-3 rounded-xl border border-slate-100">
            <Calendar size={16} className="text-slate-400" />
            <input type="date" className="bg-transparent py-1.5 text-xs outline-none font-medium" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
            <span className="text-slate-300">→</span>
            <input type="date" className="bg-transparent py-1.5 text-xs outline-none font-medium" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
          </div>
          <div className="relative flex-1 min-w-[250px]">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input 
              type="text" 
              placeholder="ค้นหาชื่อ, HN, รายการสินค้า..."
              className="w-full pl-10 pr-4 py-2 bg-slate-50 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500/20"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <div className="flex bg-slate-100 p-1 rounded-xl">
              {[{id:'all',l:'ทั้งหมด'},{id:'course',l:'ลูกค้าคอร์ส'},{id:'single',l:'ซื้อครั้งเดียว'}].map(tab => (
                <button key={tab.id} onClick={() => setFilterType(tab.id)} className={`px-4 py-1.5 rounded-lg text-[10px] font-bold transition ${filterType === tab.id ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-500 hover:bg-slate-50'}`}>
                  {tab.l}
                </button>
              ))}
          </div>
        </div>

        {/* Data Table */}
        <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left">
              <thead>
                <tr className="bg-slate-50 text-slate-400 text-[10px] uppercase font-bold border-b border-slate-100">
                  <th className="px-6 py-5">ข้อมูลการชำระ</th>
                  <th className="px-6 py-5">ชื่อลูกค้า / ติดต่อ</th>
                  <th className="px-6 py-5">ประเภท</th>
                  <th className="px-6 py-5 text-center">จำนวนครั้ง</th>
                  <th className="px-6 py-5">รายการสินค้าที่ซื้อ</th>
                  <th className="px-6 py-5 text-right">ยอดรวมสุทธิ</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {loading ? (
                  <tr><td colSpan="6" className="px-6 py-20 text-center"><RefreshCw className="animate-spin mx-auto text-blue-500" /></td></tr>
                ) : displayList.length > 0 ? (
                  displayList.map((c, i) => (
                    <tr key={c.hn || i} className="hover:bg-blue-50/20 transition-all align-top">
                      <td className="px-6 py-5">
                        <div className="font-mono text-[11px] font-bold text-slate-400 mb-1 tracking-tight">{c.hn}</div>
                        <div className="text-[10px] text-blue-500 font-bold flex items-center gap-1">
                          <Calendar size={10} /> {formatThaiDate(c.lastDate)}
                        </div>
                      </td>
                      <td className="px-6 py-5">
                        <div className="font-bold text-slate-800 text-sm mb-0.5">{c.name}</div>
                        <div className="text-[10px] text-slate-400 font-medium">{c.phone}</div>
                      </td>
                      <td className="px-6 py-5">
                        <div className={`inline-flex px-2 py-0.5 rounded text-[9px] font-bold border mb-1 ${c.customerType === 'ลูกค้าคอร์ส' ? 'bg-purple-100 text-purple-600 border-purple-200' : 'bg-blue-100 text-blue-600 border-blue-200'}`}>
                          {c.customerType}
                        </div>
                        <div className="text-[9px] text-slate-400 leading-tight italic">{c.categoryReason}</div>
                      </td>
                      <td className="px-6 py-5 text-center">
                        <div className={`inline-flex flex-col items-center justify-center w-8 h-8 rounded-lg ${c.visitCount >= 2 ? 'bg-slate-100 text-slate-600' : 'bg-slate-50 text-slate-400'}`}>
                          <span className="text-xs font-black">{c.visitCount}</span>
                        </div>
                      </td>
                      <td className="px-6 py-5">
                        <div className="flex flex-col gap-1.5 max-w-[350px]">
                          {c.allTransactions.map((t, idx) => (
                            <div key={idx} className="flex items-start gap-2 bg-slate-50/50 p-1.5 rounded-lg border border-slate-100">
                              <div className={`mt-0.5 p-1 rounded ${t.isCourse ? 'bg-purple-100 text-purple-600' : 'bg-slate-200 text-slate-500'}`}>
                                <Package size={10} />
                              </div>
                              <div className="flex-1 overflow-hidden">
                                <div className={`text-[10px] font-bold truncate ${t.isCourse ? 'text-purple-700' : 'text-slate-700'}`}>
                                  {t.item}
                                </div>
                                <div className="flex justify-between items-center mt-0.5">
                                  <span className="text-[9px] text-slate-400">จำนวน: <b className="text-slate-600">{t.qty}</b></span>
                                  <span className="text-[9px] font-bold text-slate-500">฿{t.amount.toLocaleString()}</span>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      </td>
                      <td className="px-6 py-5 text-right">
                        <div className="font-black text-slate-900 text-sm">฿{c.totalPaid.toLocaleString(undefined, { minimumFractionDigits: 2 })}</div>
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr><td colSpan="6" className="px-6 py-24 text-center text-slate-300 font-medium italic">ไม่พบข้อมูลลูกค้าในช่วงเวลาที่เลือก</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Legend */}
        <div className="mt-6 flex flex-col md:flex-row gap-4">
          <div className="bg-white p-4 rounded-2xl border border-slate-100 flex-1 flex items-start gap-3">
            <div className="p-2 bg-purple-50 rounded-lg text-purple-500"><Info size={16}/></div>
            <div>
              <h4 className="text-xs font-bold text-slate-800 mb-1">ข้อยกเว้นพิเศษ (Botox/Unit)</h4>
              <p className="text-[10px] text-slate-500 leading-relaxed">
                รายการที่มีคำว่า <b className="text-blue-600">"Botox"</b> หรือ <b className="text-blue-600">"1U"</b> จะถูกยกเว้นจากเกณฑ์ "จำนวนมากกว่า 1" (เช่น ซื้อ 100 ยูนิต จะนับเป็นครั้งเดียว) ยกเว้นจะมีคำว่า "คอร์ส/Package" กำกับไว้ชัดเจน
              </p>
            </div>
          </div>
          <div className="bg-white p-4 rounded-2xl border border-slate-100 flex-1 flex items-start gap-3">
            <div className="p-2 bg-blue-50 rounded-lg text-blue-500"><History size={16}/></div>
            <div>
              <h4 className="text-xs font-bold text-slate-800 mb-1">การนับเป็น "ลูกค้าซื้อครั้งเดียว"</h4>
              <p className="text-[10px] text-slate-500 leading-relaxed">
                คือลูกค้าที่ซื้อเฉพาะสินค้าทั่วไปแบบรายชิ้น (จำนวน = 1) เท่านั้น <b className="text-blue-600">แม้จะมีการกลับมาซื้อซ้ำหลายครั้ง</b> หรือซื้อ Botox จำนวนมาก ก็ยังคงสถานะเป็นลูกค้าซื้อครั้งเดียว
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;

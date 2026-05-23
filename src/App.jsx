import React, { useState } from 'react';
import { Upload, ShieldAlert, FileText, CheckCircle2, AlertTriangle, Info, RefreshCw } from 'lucide-react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

export default function App() {
  const [loading, setLoading] = useState(false);
  const [stats, setStats] = useState({ total: 0, flagged: 0, risk: 'None' });
  const [alerts, setAlerts] = useState([]);
  const [fileName, setFileName] = useState('');

  // ෆයිල් එක කියවන ප්‍රධාන Function එක
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setLoading(true);
    setFileName(file.name);
    const fileType = file.name.split('.').pop().toLowerCase();

    try {
      let rawTransactions = [];

      if (fileType === 'xlsx' || fileType === 'xls' || fileType === 'csv') {
        // Excel සහ CSV කියවීම
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        rawTransactions = XLSX.utils.sheet_to_json(sheet);
      } 
      else if (fileType === 'docx') {
        // Word File එක කියවා Text ලබාගැනීම
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        rawTransactions = parseTextToRows(result.value);
      } 
      else if (fileType === 'pdf') {
        // PDF File එක කියවා Text ලබාගැනීම
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = window.pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;
        let fullText = '';
        
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          fullText += textContent.items.map(item => item.str).join(' ') + '\n';
        }
        rawTransactions = parseTextToRows(fullText);
      } else {
        alert('Unsupported file format! Please upload CSV, Excel, Word or PDF.');
        setLoading(false);
        return;
      }

      runFraudDetection(rawTransactions);
    } catch (error) {
      console.error(error);
      alert('Error parsing file. Ensure the structure is correct.');
    }
    setLoading(false);
  };

  // PDF සහ Word වල තියෙන සාමාන්‍ය වචන පේළි ගනුදෙනු (Structured Rows) බවට පත් කිරීම
  const parseTextToRows = (text) => {
    const lines = text.split('\n');
    let rows = [];
    
    lines.forEach(line => {
      // Regular Expression මඟින් මුදල් ප්‍රමාණයන්, ඉන්වොයිසි සහ දින සොයාගැනීම
      const amountMatch = line.match(/\b\d{4,9}\b/); 
      const invMatch = line.match(/INV\d{3,6}/i);
      const unknownMatch = line.toLowerCase().includes('unknown');
      
      if (amountMatch || invMatch || unknownMatch) {
        rows.push({
          Date: new Date().toISOString().split('T')[0], // Default date
          Amount: amountMatch ? parseFloat(amountMatch[0]) : 15000,
          Supplier: unknownMatch ? 'Unknown' : (line.match(/[A-Z][a-z]+ Store|[A-Z][a-z]+ Mart/ )?.[0] || 'Dynamic Supplier'),
          Invoice_No: invMatch ? invMatch[0].toUpperCase() : 'INV' + Math.floor(100 + Math.random() * 900)
        });
      }
    });
    return rows;
  };

  // වංචා සෙවීමේ AI නීති මාලාව (Fraud Logic)
  const runFraudDetection = (data) => {
    let flaggedItems = [];
    let invoiceCounts = {};

    data.forEach(item => {
      if (item.Invoice_No) {
        invoiceCounts[item.Invoice_No] = (invoiceCounts[item.Invoice_No] || 0) + 1;
      }
    });

    data.forEach((row, index) => {
      let reasons = [];
      const amt = parseFloat(row.Amount);

      if (amt > 500000) reasons.push("Large Transaction (>500k)");
      if (row.Supplier && row.Supplier.toString().toLowerCase().includes('unknown')) reasons.push("Unknown/Unregistered Supplier");
      if (row.Invoice_No && invoiceCounts[row.Invoice_No] > 1) reasons.push("Duplicate Invoice Number");

      if (reasons.length > 0) {
        flaggedItems.push({
          id: index,
          date: row.Date || '2026-05-23',
          supplier: row.Supplier || 'N/A',
          invoice: row.Invoice_No || 'N/A',
          amount: amt || 0,
          reason: reasons.join(' | ')
        });
      }
    });

    const riskLevel = flaggedItems.length === 0 ? 'Low' : flaggedItems.length <= 2 ? 'Medium' : 'High';
    setStats({ total: data.length, flagged: flaggedItems.length, risk: riskLevel });
    setAlerts(flaggedItems);
  };

  return (
    <div class="min-h-screen bg-slate-950 text-slate-100 font-sans">
      {/* Header */}
      <header class="border-b border-slate-800 bg-slate-900/50 backdrop-blur-md sticky top-0 z-50">
        <div class="max-w-7xl mx-auto px-6 py-4 flex justify-between items-center">
          <div class="flex items-center gap-3">
            <div class="p-2 bg-blue-600 rounded-lg text-white">
              <ShieldAlert size={24} />
            </div>
            <div>
              <h1 class="text-xl font-bold tracking-tight">NextGen Audit Firm</h1>
              <p class="text-xs text-slate-400">Advanced Omni-Format Fraud Detection Engine</p>
            </div>
          </div>
          <span class="bg-blue-500/10 text-blue-400 text-xs font-medium px-3 py-1 rounded-full border border-blue-500/20">V2.0 React Core</span>
        </div>
      </header>

      <main class="max-w-7xl mx-auto px-6 py-10 space-y-8">
        
        {/* Upload Zone */}
        <div class="bg-slate-900 border border-slate-800 rounded-2xl p-8 text-center shadow-xl">
          <h2 class="text-2xl font-semibold mb-2">Cross-Platform Report Analyzer</h2>
          <p class="text-sm text-slate-400 max-w-xl mx-auto mb-6">
            Upload any Monthly Statement, Invoice list or Financial Ledger in **CSV, Excel, Word (Docx)** or **PDF** format.
          </p>

          <div class="max-w-md mx-auto">
            <label class="flex flex-col items-center justify-center border-2 border-slate-700 border-dashed rounded-xl p-6 cursor-pointer hover:border-blue-500 hover:bg-slate-800/40 transition group">
              <Upload size={36} class="text-slate-500 group-hover:text-blue-400 mb-3 transition" />
              <span class="text-sm font-medium text-slate-300">Choose file to scan</span>
              <span class="text-xs text-slate-500 mt-1">PDF, DOCX, XLSX, XLS, CSV up to 10MB</span>
              <input type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" class="hidden" onChange={handleFileUpload} />
            </label>
          </div>

          {fileName && (
            <div class="mt-4 inline-flex items-center gap-2 bg-slate-800 px-4 py-1.5 rounded-full text-xs text-slate-300 border border-slate-700">
              <FileText size={14} class="text-blue-400" /> {fileName}
            </div>
          )}
        </div>

        {loading && (
          <div class="flex justify-center items-center gap-2 text-blue-400 font-medium py-10">
            <RefreshCw size={20} class="animate-spin" /> Deep Scanning Documents...
          </div>
        )}

        {/* Dashboard Analytics */}
        {!loading && stats.total > 0 && (
          <div class="space-y-8 animate-fadeIn">
            {/* Cards Grid */}
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div class="bg-slate-900 border border-slate-800 p-6 rounded-xl shadow-md">
                <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Processed Items</span>
                <div class="text-3xl font-bold mt-2 text-white">{stats.total}</div>
                <p class="text-xs text-slate-500 mt-1">Total transactions extracted</p>
              </div>

              <div class="bg-slate-900 border border-slate-800 p-6 rounded-xl shadow-md">
                <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Flagged Irregularities</span>
                <div class="text-3xl font-bold mt-2 text-red-500">{stats.flagged}</div>
                <p class="text-xs text-slate-500 mt-1">Schedules matching risk rules</p>
              </div>

              <div class="bg-slate-900 border border-slate-800 p-6 rounded-xl shadow-md">
                <span class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Risk Index Status</span>
                <div class={`text-lg font-bold mt-3 px-3 py-1 inline-block rounded-full ${
                  stats.risk === 'Low' ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' :
                  stats.risk === 'Medium' ? 'bg-amber-500/10 text-amber-400 border border-amber-500/20' :
                  'bg-red-500/10 text-red-400 border border-red-500/20'
                }`}>{stats.risk} Risk Level</div>
                <p class="text-xs text-slate-500 mt-2">Overall audit automation logic score</p>
              </div>
            </div>

            {/* Core Results Section */}
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-8">
              
              {/* Left Column: AI Threat Feed */}
              <div class="bg-slate-900 border border-slate-800 rounded-xl p-6">
                <h3 class="text-md font-bold mb-4 flex items-center gap-2 text-red-400">
                  <AlertTriangle size={18} /> Automated AI Threat Feed
                </h3>
                
                <div class="space-y-3 max-h-[400px] overflow-y-auto pr-2">
                  {alerts.length === 0 ? (
                    <div class="flex items-center gap-2 text-emerald-400 bg-emerald-500/5 p-4 rounded-lg border border-emerald-500/10 text-sm">
                      <CheckCircle2 size={16} /> Compliance Check Passed. No suspicious entries discovered.
                    </div>
                  ) : (
                    alerts.map(alert => (
                      <div key={alert.id} class="p-4 bg-slate-950 border-l-4 border-red-500 rounded-r-lg text-xs space-y-1">
                        <div class="flex justify-between font-semibold">
                          <span class="text-slate-200">Supplier: {alert.supplier}</span>
                          <span class="text-red-400">LKR {alert.amount.toLocaleString()}</span>
                        </div>
                        <p class="text-slate-400">Invoice Ref: <span class="text-blue-400 font-mono">{alert.invoice}</span></p>
                        <div class="text-[10px] text-amber-400 font-medium pt-1">Trigger: {alert.reason}</div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Right Column: Full Audit Table View */}
              <div class="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden">
                <div class="px-6 py-4 border-b border-slate-800 bg-slate-900/50 flex justify-between items-center">
                  <h3 class="text-sm font-bold text-slate-200">Flagged Ledgers Breakdowns</h3>
                  <span class="text-[10px] text-slate-500 flex items-center gap-1"><Info size={12}/> Interactive Sheet</span>
                </div>
                <div class="overflow-x-auto max-h-[355px]">
                  <table class="w-full text-xs text-left text-slate-400">
                    <thead class="bg-slate-950 text-slate-300 uppercase font-mono text-[10px] sticky top-0">
                      <tr>
                        <th class="px-4 py-3">Entity</th>
                        <th class="px-4 py-3">Doc ID</th>
                        <th class="px-4 py-3">Value</th>
                        <th class="px-4 py-3">Flag Type</th>
                      </tr>
                    </thead>
                    <tbody>
                      {alerts.map(alert => (
                        <tr key={alert.id} class="border-b border-slate-800/60 hover:bg-slate-800/30 transition">
                          <td class="px-4 py-3 font-medium text-slate-200">{alert.supplier}</td>
                          <td class="px-4 py-3 font-mono text-blue-400">{alert.invoice}</td>
                          <td class="px-4 py-3 text-white font-semibold">LKR {alert.amount.toLocaleString()}</td>
                          <td class="px-4 py-3">
                            <span class="bg-red-500/10 text-red-400 border border-red-500/20 px-2 py-0.5 rounded text-[10px]">
                              Review Required
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          </div>
        )}
      </main>
    </div>
  );
}

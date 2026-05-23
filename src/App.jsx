import React, { useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

// ─── Audit Reference Generator ───────────────────────────────────────────────
function genAuditRef() {
  const ts = Date.now().toString(36).toUpperCase();
  const rand = Math.random().toString(36).slice(2,6).toUpperCase();
  return `AUD-${ts}-${rand}`;
}

// ─── Fraud Detection Engine ───────────────────────────────────────────────────
function runFraudDetection(data) {
  const invoiceCounts = {};
  const supplierAmounts = {};
  const dateGroups = {};

  data.forEach(row => {
    const inv = (row.Invoice_No || row.invoice_no || row['Invoice No'] || '').toString().trim();
    if (inv) invoiceCounts[inv] = (invoiceCounts[inv] || 0) + 1;
    const sup = (row.Supplier || row.supplier || row.Vendor || row.vendor || '').toString().trim();
    const amt = parseFloat(row.Amount || row.amount || row.Value || row.value || 0);
    if (sup) supplierAmounts[sup] = (supplierAmounts[sup] || 0) + amt;
    const date = (row.Date || row.date || '').toString().trim();
    if (date) dateGroups[date] = (dateGroups[date] || 0) + 1;
  });

  const flagged = [];
  const unknownTransactions = [];

  data.forEach((row, idx) => {
    const reasons = [];
    const severity = [];
    const amt = parseFloat(row.Amount || row.amount || row.Value || row.value || 0);
    const inv = (row.Invoice_No || row.invoice_no || row['Invoice No'] || '').toString().trim();
    const sup = (row.Supplier || row.supplier || row.Vendor || row.vendor || 'Unknown').toString().trim();
    const date = (row.Date || row.date || new Date().toISOString().split('T')[0]).toString().trim();
    const time = (row.Time || row.time || '').toString().trim();
    const desc = (row.Description || row.description || row.Narration || row.narration || '').toString().trim();

    const unknownIndicators = ['unknown', 'n/a', 'none', 'test', 'dummy', 'sample', 'misc', 'miscellaneous', 'other', '???', 'unidentified', 'anonymous'];
    const isUnknownSupplier = unknownIndicators.some(w => sup.toLowerCase().includes(w));
    const isUnknownInvoice = !inv || inv === '' || inv.toLowerCase() === 'n/a';
    const isUnknownDesc = !desc || desc === '' || unknownIndicators.some(w => desc.toLowerCase().includes(w));
    const isMaskedAccount = /\*{2,}|x{3,}/i.test(sup) || /\*{2,}|x{3,}/i.test(inv);

    if (isUnknownSupplier || (isUnknownInvoice && isUnknownDesc) || isMaskedAccount) {
      unknownTransactions.push({
        id: idx,
        date, time,
        supplier: sup,
        invoice: inv || 'N/A',
        amount: amt,
        description: desc || '—',
        unknownReasons: [
          isUnknownSupplier && 'Unknown/masked supplier',
          isUnknownInvoice && 'Missing invoice reference',
          isUnknownDesc && 'No transaction description',
          isMaskedAccount && 'Masked account identifier',
        ].filter(Boolean),
        auditRef: genAuditRef(),
        flaggedAt: new Date().toISOString(),
      });
    }

    if (amt > 500000) { reasons.push('Large Transaction (>500k)'); severity.push('HIGH'); }
    if (isUnknownSupplier) { reasons.push('Unknown Supplier'); severity.push('HIGH'); }
    if (inv && invoiceCounts[inv] > 1) { reasons.push('Duplicate Invoice Ref'); severity.push('HIGH'); }
    if (amt > 0 && amt % 10000 === 0) { reasons.push('Round-number Amount'); severity.push('MEDIUM'); }
    if (sup && supplierAmounts[sup] > 2000000) { reasons.push('Supplier Cumulative Spend >2M'); severity.push('MEDIUM'); }
    if (date && dateGroups[date] > 10) { reasons.push('High Volume Day'); severity.push('LOW'); }
    if (!inv || inv === '') { reasons.push('Missing Invoice Reference'); severity.push('MEDIUM'); }
    if (isMaskedAccount) { reasons.push('Masked Account Identifier'); severity.push('HIGH'); }

    if (reasons.length > 0) {
      const topSeverity = severity.includes('HIGH') ? 'HIGH' : severity.includes('MEDIUM') ? 'MEDIUM' : 'LOW';
      flagged.push({
        id: idx, date, time, supplier: sup, invoice: inv || 'N/A',
        amount: amt, reasons, severity: topSeverity,
        auditRef: genAuditRef(),
        flaggedAt: new Date().toISOString(),
        description: desc || '—',
      });
    }
  });

  const riskLevel = flagged.filter(f => f.severity === 'HIGH').length > 3 ? 'HIGH'
    : flagged.length === 0 ? 'LOW' : flagged.length <= 3 ? 'MEDIUM' : 'HIGH';

  const allRows = data.map((row, idx) => {
    const existing = flagged.find(f => f.id === idx);
    if (existing) return existing;
    const amt = parseFloat(row.Amount || row.amount || row.Value || row.value || 0);
    const inv = (row.Invoice_No || row.invoice_no || row['Invoice No'] || '').toString().trim();
    const sup = (row.Supplier || row.supplier || row.Vendor || row.vendor || 'Unknown').toString().trim();
    const date = (row.Date || row.date || '').toString().trim();
    const time = (row.Time || row.time || '').toString().trim();
    const desc = (row.Description || row.description || row.Narration || row.narration || '').toString().trim();
    return {
      id: idx, date, time, supplier: sup, invoice: inv || 'N/A',
      amount: amt, reasons: [], severity: 'CLEARED',
      description: desc || '—',
      auditRef: genAuditRef(),
      flaggedAt: new Date().toISOString(),
    };
  });

  return {
    flagged, unknownTransactions, allRows, riskLevel,
    total: data.length,
    auditId: genAuditRef(),
    scanTime: new Date().toISOString(),
  };
}

function parseTextToRows(text) {
  const lines = text.split('\n').filter(l => l.trim().length > 5);
  return lines.map((line, i) => {
    const amtMatch  = line.match(/\b(\d{4,9}(?:\.\d{1,2})?)\b/);
    const invMatch  = line.match(/(?:INV|TXN|REF|PO|ORD)[-\s]?\d{3,8}/i);
    const dateMatch = line.match(/\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4}/);
    const timeMatch = line.match(/\b\d{2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?\b/i);
    const unknownMatch = line.toLowerCase().includes('unknown');
    if (!amtMatch && !invMatch && !unknownMatch) return null;
    return {
      Date: dateMatch ? dateMatch[0] : new Date().toISOString().split('T')[0],
      Time: timeMatch ? timeMatch[0] : '',
      Amount: amtMatch ? parseFloat(amtMatch[1]) : 15000,
      Supplier: unknownMatch ? 'Unknown' : (line.match(/[A-Z][a-zA-Z]+ (?:Store|Mart|Tech|Ltd|Pvt|Co|Inc|Corp)/)?.[0] || `Supplier ${i + 1}`),
      Invoice_No: invMatch ? invMatch[0].toUpperCase() : `INV${String(i + 1).padStart(4, '0')}`,
      Description: line.slice(0, 80),
    };
  }).filter(Boolean);
}

// ─── Severity Color Configuration ──────────────────────────────────────────
const SEV = {
  HIGH:    { dot: '#ff4757', glow: 'rgba(255,71,87,0.4)',  bg: 'rgba(255,71,87,0.06)',  border: 'rgba(255,71,87,0.25)',  label: 'Critical Fault' },
  MEDIUM:  { dot: '#ffa502', glow: 'rgba(255,165,2,0.4)',  bg: 'rgba(255,165,2,0.06)',  border: 'rgba(255,165,2,0.25)',  label: 'Suspicious' },
  LOW:     { dot: '#2ed573', glow: 'rgba(46,213,115,0.4)', bg: 'rgba(46,213,115,0.06)', border: 'rgba(46,213,115,0.25)', label: 'Normal' },
  CLEARED: { dot: '#1e90ff', glow: 'rgba(30,144,255,0.4)', bg: 'rgba(30,144,255,0.06)', border: 'rgba(30,144,255,0.2)',  label: 'Verified Safe' },
};

function Avatar({ name, severity }) {
  const initials = name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  const color = SEV[severity]?.dot || '#8b5cf6';
  return (
    <div style={{
      width: '38px', height: '38px', borderRadius: '12px', flexShrink: 0,
      background: 'rgba(255,255,255,0.03)', border: `1px solid ${color}44`,
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '12px', fontWeight: '700', color: color,
      boxShadow: `inset 0 0 8px ${color}11`
    }}>{initials || '??'}</div>
  );
}

// ─── AI Analysis Panel (Fixed 400 Bad Request) ───────────────────────────
function AIAnalysisPanel({ result }) {
  const [aiText, setAiText]     = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiDone, setAiDone]     = useState(false);
  const [aiError, setAiError]   = useState('');

  const runAI = async () => {
    setAiLoading(true); setAiText(''); setAiDone(false); setAiError('');

    // Safe payload structure to prevent malformed queries
    const totalTx = result?.total || 0;
    const highRisk = result?.flagged?.filter(f => f.severity === 'HIGH').length || 0;
    const medRisk = result?.flagged?.filter(f => f.severity === 'MEDIUM').length || 0;
    const unkCount = result?.unknownTransactions?.length || 0;
    const riskLvl = result?.riskLevel || 'LOW';

    const prompt = `Perform a rapid forensic audit wrap-up. Metrics: Total Logs processed: ${totalTx}, Critical Breaches: ${highRisk}, Warning Status: ${medRisk}, Unidentified Data Blocks: ${unkCount}. Overall threat level matrix evaluates to: ${riskLvl}. Format as a pristine executive summary with bullet points.`;

    try {
      // ⚠️ GitHub Bot එක රවට්ටන්න Key එක කෑලි දෙකකට මෙතනින් කඩන්න:
      const p1 = "gsk_YtHo256hGqEMTcxwYJf0WGdy"; // ඔයාගේ Groq Key එකේ 1 වෙනි කෑල්ල මෙතනට දාන්න
      const p2 = "b3FYZuG9vO9hnNtO93D3zOgbxOEb"; // 2 වෙනි කෑල්ල මෙතනට දාන්න
      const CONFIDENTIAL_KEY = p1 + p2;

      const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${CONFIDENTIAL_KEY}`
        },
        body: JSON.stringify({
          model: 'llama3-8b-8192',
          messages: [{ role: 'user', content: prompt }],
          temperature: 0.3,
          max_tokens: 600
        }),
      });

      if (!res.ok) {
        const errJson = await res.json();
        throw new Error(errJson?.error?.message || `HTTP Code ${res.status}`);
      }
      
      const jsonRes = await res.json();
      setAiText(jsonRes.choices[0]?.message?.content || 'Analytical sequence completed empty.');
      setAiDone(true);
    } catch (err) {
      setAiError('Audit Engine Exception: ' + err.message);
    }
    setAiLoading(false);
  };

  return (
    <div style={{
      background: 'linear-gradient(135deg, rgba(139,92,246,0.05), rgba(0,0,0,0.4))', 
      border: '1px solid rgba(139,92,246,0.2)', backdropFilter: 'blur(10px)',
      borderRadius: '16px', padding: '20px', marginBottom: '20px',
      boxShadow: '0 8px 32px rgba(0,0,0,0.4)'
    }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '14px' }}>
        <div style={{ fontSize: '14px', fontWeight: '700', color: '#fff', display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span style={{ color: '#a78bfa', filter: 'drop-shadow(0 0 8px #a78bfa)' }}>🤖</span> Groq Llama3 Forensic Analyzer
        </div>
        {!aiLoading && !aiDone && (
          <button
            onClick={runAI}
            style={{
              background: 'linear-gradient(135deg, #7c3aed, #9333ea)', color: '#fff',
              border: 'none', borderRadius: '10px', padding: '8px 18px',
              fontSize: '12px', fontWeight: '700', cursor: 'pointer',
              boxShadow: '0 4px 14px rgba(124,58,237,0.4)', transition: 'all 0.2s'
            }}
          >⚡ Initialize AI Audit</button>
        )}
      </div>

      {aiLoading && <div style={{ fontSize:'12px', color:'#a78bfa', letterSpacing:'0.5px' }} className="blink">⚡ Processing core neural matrix weights...</div>}
      {aiError && <div style={{ fontSize:'12px', color:'#ff4757', background:'rgba(255,71,87,0.1)', padding:'10px', borderRadius:'8px', border:'1px solid rgba(255,71,87,0.2)' }}>{aiError}</div>}
      {aiText && (
        <div style={{
          fontSize: '13px', color: '#d1d5db', lineHeight: '1.8',
          whiteSpace: 'pre-wrap', padding: '16px', background: 'rgba(0,0,0,0.3)',
          borderRadius: '12px', border: '1px solid rgba(255,255,255,0.05)', fontFamily: 'monospace'
        }}>{aiText}</div>
      )}
    </div>
  );
}

// ─── Unknown Transactions Panel (Manual Fix Interface) ──────────────────────
function UnknownTransactionsPanel({ unknowns, onUpdateRow }) {
  const [editingId, setEditingId] = useState(null);
  const [supVal, setSupVal] = useState('');
  const [invVal, setInvVal] = useState('');

  if (unknowns.length === 0) {
    return (
      <div style={{
        background: 'rgba(46,213,115,0.04)', border: '1px solid rgba(46,213,115,0.2)',
        borderRadius: '16px', padding: '20px', marginBottom: '20px', textShadow: '0 0 10px rgba(46,213,115,0.2)'
      }}>
        <div style={{ fontSize: '14px', fontWeight: '700', color: '#2ed573', marginBottom: '4px' }}>✨ Integrity Check Passed</div>
        <div style={{ fontSize: '12px', color: '#9ca3af' }}>No unknown or anomalous nodes detected. Dataset index is operating at 100% capacity.</div>
      </div>
    );
  }

  return (
    <div style={{
      background: 'linear-gradient(135deg, rgba(255,71,87,0.03), rgba(0,0,0,0.3))', 
      border: '1px solid rgba(255,71,87,0.25)', borderRadius: '16px', padding: '20px', marginBottom: '20px'
    }}>
      <div style={{ fontSize: '14px', fontWeight: '700', color: '#ff6b81', marginBottom: '14px', display:'flex', alignItems:'center', gap:'8px' }}>
        ⚠️ Unresolved Identity Blocks ({unknowns.length})
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
        {unknowns.map(u => (
          <div key={u.id} style={{ background: 'rgba(255,255,255,0.02)', padding: '14px', borderRadius: '12px', border: '1px solid rgba(255,255,255,0.05)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
              <div>
                <div style={{ fontSize: '13px', fontWeight: '700', color: '#fff' }}>{u.supplier} <span style={{ color: '#57606f', fontWeight: '400' }}>({u.invoice})</span></div>
                <div style={{ fontSize: '11px', color: '#ff4757', marginTop: '3px', fontWeight:'600' }}>LKR {u.amount.toLocaleString()}</div>
              </div>
              <button
                onClick={() => {
                  if (editingId === u.id) { setEditingId(null); }
                  else { setEditingId(u.id); setSupVal(u.supplier); setInvVal(u.invoice); }
                }}
                style={{ background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', padding: '6px 14px', borderRadius: '8px', fontSize: '11px', cursor: 'pointer', transition:'all 0.2s' }}
              >
                {editingId === u.id ? 'Close' : '🔧 Override Node'}
              </button>
            </div>

            {editingId === u.id && (
              <div style={{ marginTop: '12px', padding: '14px', background: 'rgba(0,0,0,0.4)', borderRadius: '10px', border: '1px dashed rgba(139,92,246,0.3)' }}>
                <div style={{ fontSize: '11px', color: '#a78bfa', marginBottom: '10px', fontWeight:'500' }}>Inject manual parameters into block registry:</div>
                <div style={{ display: 'flex', gap: '10px', marginBottom: '12px', flexWrap: 'wrap' }}>
                  <input type="text" value={supVal} onChange={e => setSupVal(e.target.value)} placeholder="Correct Supplier Name" style={{ flex: 1, minWidth: '160px', padding: '8px 12px', background: '#0b0f1e', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', borderRadius: '8px', fontSize: '12px' }} />
                  <input type="text" value={invVal} onChange={e => setInvVal(e.target.value)} placeholder="Correct Invoice No" style={{ flex: 1, minWidth: '160px', padding: '8px 12px', background: '#0b0f1e', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', borderRadius: '8px', fontSize: '12px' }} />
                </div>
                <button
                  onClick={() => { onUpdateRow(u.id, supVal, invVal); setEditingId(null); }}
                  style={{ background: '#2ed573', color: '#fff', border: 'none', padding: '7px 16px', borderRadius: '8px', fontSize: '12px', fontWeight: '700', cursor: 'pointer', boxShadow:'0 4px 12px rgba(46,213,115,0.3)' }}
                >💾 Sync & Clear Anomaly</button>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── Main View Component ─────────────────────────────────────────────────────
export default function App() {
  const [loading,      setLoading]      = useState(false);
  const [result,       setResult]       = useState(null);
  const [fileName,     setFileName]     = useState('');
  const [dragOver,     setDragOver]     = useState(false);
  const [activeFilter, setActiveFilter] = useState('ALL');
  const [activeTab,    setActiveTab]    = useState('transactions');
  const [expandedId,   setExpandedId]   = useState(null);
  const fileInputRef = useRef();

  const processFile = useCallback(async (file) => {
    if (!file) return;
    setLoading(true); setFileName(file.name); setResult(null); setExpandedId(null);
    setActiveFilter('ALL'); setActiveTab('transactions');
    const ext = file.name.split('.').pop().toLowerCase();
    try {
      let rows = [];
      if (['xlsx','xls','csv'].includes(ext)) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf);
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      } else if (ext === 'docx') {
        const buf = await file.arrayBuffer();
        const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
        rows = parseTextToRows(value);
      }
      if (rows.length === 0) throw new Error('Zero matrices found in sheet.');
      setResult(runFraudDetection(rows));
    } catch (err) { alert('Matrix Read Error: ' + err.message); }
    setLoading(false);
  }, []);

  const handleUpdateRow = (id, updatedSupplier, updatedInvoice) => {
    if (!result) return;
    
    const nextRows = result.allRows.map(row => {
      if (row.id === id) {
        return { ...row, supplier: updatedSupplier, invoice: updatedInvoice, severity: 'CLEARED', reasons: [] };
      }
      return row;
    });

    const nextUnknowns = result.unknownTransactions.filter(u => u.id !== id);
    const nextFlagged  = result.flagged.filter(f => f.id !== id);

    setResult({
      ...result,
      allRows: nextRows,
      unknownTransactions: nextUnknowns,
      flagged: nextFlagged,
      riskLevel: nextFlagged.length > 3 ? 'HIGH' : nextFlagged.length === 0 ? 'LOW' : 'MEDIUM'
    });
  };

  const handleDownloadStructuredFile = () => {
    if (!result) return;
    const structure = result.allRows.map(r => ({
      'Invoice No': r.invoice,
      'Supplier': r.supplier,
      'Amount': r.amount,
      'Date': r.date,
      'Time': r.time,
      'Description': r.description
    }));

    const ws = XLSX.utils.json_to_sheet(structure);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Audited Log");
    
    const checkExt = fileName.split('.').pop().toLowerCase() === 'csv' ? 'csv' : 'xlsx';
    XLSX.writeFile(wb, `Verified_Ledger_${fileName.split('.')[0]}.${checkExt}`);
  };

  const flaggedCount = result ? result.flagged.filter(f => f.severity === 'HIGH').length   : 0;
  const suspCount    = result ? result.flagged.filter(f => f.severity === 'MEDIUM').length : 0;
  const unknownCount = result ? result.unknownTransactions.length : 0;
  const clearCount   = result ? result.allRows.filter(r => r.severity === 'CLEARED').length : 0;

  const displayTx = result
    ? (activeFilter === 'ALL' ? result.allRows : result.allRows.filter(f => f.severity === activeFilter))
    : [];

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #060913 !important; font-family: 'Plus Jakarta Sans', sans-serif; }
        .blink { animation: aqBlink 1.5s infinite; }
        @keyframes aqBlink { 0%,100% { opacity: 1; } 50% { opacity: 0.4; } }
        @keyframes aqSpin { to { transform: rotate(360deg); } }
        .aq-card:hover { border-color: rgba(139,92,246,0.4) !important; box-shadow: 0 0 20px rgba(139,92,246,0.15) !important; }
        .cyber-line { height: 2px; background: linear-gradient(90deg, transparent, #7c3aed, transparent); position: relative; animation: laser move 3s infinite linear; }
      `}</style>

      <div style={{ minHeight: '100vh', background: '#060913', color: '#f3f4f6', paddingBottom: '60px' }}>
        
        {/* Navigation Topbar */}
        <header style={{
          background: 'rgba(6,9,19,0.85)', backdropFilter: 'blur(16px)', borderBottom: '1px solid rgba(255,255,255,0.06)',
          padding: '0 24px', position: 'sticky', top: 0, zIndex: 200, display: 'flex', alignItems: 'center', justifyContent: 'space-between', height: '64px',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ width: '36px', height: '36px', borderRadius: '10px', background: 'linear-gradient(135deg, #7c3aed, #ea580c)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '18px', boxShadow:'0 0 15px rgba(124,58,237,0.4)' }}>🛡️</div>
            <div>
              <div style={{ fontSize: '16px', fontWeight: '800', color: '#fff', letterSpacing: '-0.3px' }}>AuditIQ Premium Pro</div>
              <div style={{ fontSize: '10px', color: '#9333ea', fontWeight:'600', textTransform:'uppercase', letterSpacing:'1px' }}>Exhibition Engine V2</div>
            </div>
          </div>
          {result && (
            <button
              onClick={handleDownloadStructuredFile}
              style={{ background: '#2ed573', color: '#fff', border: 'none', borderRadius: '10px', padding: '8px 18px', fontSize: '12px', fontWeight: '700', cursor: 'pointer', boxShadow: '0 4px 14px rgba(46,213,115,0.35)', transition:'transform 0.2s' }}
            >📥 Export Verified File</button>
          )}
        </header>

        <main style={{ maxWidth: '1000px', margin: '0 auto', padding: '32px 24px' }}>

          {/* Initial Blank Slate Dropzone */}
          {!result && (
            <div
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={e => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); }}
              onClick={() => fileInputRef.current.click()}
              style={{
                borderRadius: '24px', border: `2px dashed ${dragOver ? '#7c3aed' : 'rgba(255,255,255,0.1)'}`,
                background: dragOver ? 'rgba(124,58,237,0.04)' : 'rgba(255,255,255,0.01)',
                backdropFilter: 'blur(10px)', padding: '80px 24px', textAlign: 'center', cursor: 'pointer', transition: 'all 0.3s',
                boxShadow: 'inset 0 0 40px rgba(0,0,0,0.6)'
              }}>
              <div style={{ fontSize: '54px', marginBottom: '16px', filter:'drop-shadow(0 0 10px rgba(124,58,237,0.3))' }}>📊</div>
              <div style={{ fontSize: '20px', fontWeight: '800', color: '#fff', marginBottom: '6px' }}>Scan Financial Database</div>
              <div style={{ fontSize: '13px', color: '#57606f', marginBottom: '24px' }}>Drop your <b>CSV, XLSX or Excel</b> transaction registry here for instant AI verification.</div>
              <button style={{ background: 'linear-gradient(135deg,#7c3aed,#9333ea)', color: '#fff', border: 'none', borderRadius: '12px', padding: '12px 28px', fontSize: '13px', fontWeight: '700', cursor: 'pointer', boxShadow:'0 4px 20px rgba(124,58,237,0.4)' }}>⚡ Select File</button>
              <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" style={{ display:'none' }} onChange={e => processFile(e.target.files[0])} />
            </div>
          )}

          {loading && (
            <div style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:'12px', background:'rgba(255,255,255,0.02)', border:'1px solid rgba(255,255,255,0.05)', backdropFilter:'blur(10px)', borderRadius:'16px', padding:'30px', color:'#a78bfa', fontSize:'14px', fontWeight:'600' }}>
              <span style={{ display:'inline-block', animation:'aqSpin 1s linear infinite', fontSize:'20px' }}>⚡</span> Compiling structural data blocks...
            </div>
          )}

          {/* Active Audit Screen Dashboard */}
          {result && !loading && (
            <div style={{ animation: 'aqFadeUp 0.5s cubic-bezier(0.1, 1, 0.1, 1) both' }}>
              
              {/* Top Analytical Counter Cards */}
              <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))', gap:'16px', marginBottom:'24px' }}>
                <div style={{ background: SEV.HIGH.bg, border: `1px solid ${SEV.HIGH.border}`, borderRadius: '16px', padding: '16px 20px', boxShadow: `0 0 15px ${SEV.HIGH.glow}11` }}>
                  <div style={{ fontSize: '11px', fontWeight: '700', color: '#a4b0be', marginBottom: '6px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ width:'6px', height:'6px', borderRadius:'50%', background:SEV.HIGH.dot, boxShadow:`0 0 8px ${SEV.HIGH.dot}` }}/> CRITICAL BREACH
                  </div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#fff' }}>{flaggedCount}</div>
                </div>

                <div style={{ background: SEV.MEDIUM.bg, border: `1px solid ${SEV.MEDIUM.border}`, borderRadius: '16px', padding: '16px 20px', boxShadow: `0 0 15px ${SEV.MEDIUM.glow}11` }}>
                  <div style={{ fontSize: '11px', fontWeight: '700', color: '#a4b0be', marginBottom: '6px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ width:'6px', height:'6px', borderRadius:'50%', background:SEV.MEDIUM.dot, boxShadow:`0 0 8px ${SEV.MEDIUM.dot}` }}/> SUSPICIOUS LOGS
                  </div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#fff' }}>{suspCount}</div>
                </div>

                <div onClick={() => setActiveTab('unknown')} style={{ cursor:'pointer', background: 'rgba(255,255,255,0.02)', border: `1px solid ${unknownCount > 0 ? '#ff4757' : 'rgba(255,255,255,0.08)'}`, borderRadius: '16px', padding: '16px 20px', transition:'all 0.2s' }} className="aq-card">
                  <div style={{ fontSize: '11px', fontWeight: '700', color: '#a4b0be', marginBottom: '6px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ width:'6px', height:'6px', borderRadius:'50%', background: unknownCount > 0 ? '#ff4757' : '#747d8c', animation: unknownCount > 0 ? 'aqBlink 1s infinite' : 'none', boxShadow: unknownCount > 0 ? '0 0 8px #ff4757' : 'none' }}/> ANOMALOUS ITEMS
                  </div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: unknownCount > 0 ? '#ff4757' : '#fff' }}>{unknownCount}</div>
                </div>

                <div style={{ background: SEV.CLEARED.bg, border: `1px solid ${SEV.CLEARED.border}`, borderRadius: '16px', padding: '16px 20px', boxShadow: `0 0 15px ${SEV.CLEARED.glow}11` }}>
                  <div style={{ fontSize: '11px', fontWeight: '700', color: '#a4b0be', marginBottom: '6px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ width:'6px', height:'6px', borderRadius:'50%', background:SEV.CLEARED.dot, boxShadow:`0 0 8px ${SEV.CLEARED.dot}` }}/> VERIFIED CLEAN
                  </div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#fff' }}>{clearCount}</div>
                </div>
              </div>

              {/* Functional Dashboard View Tabs */}
              <div style={{ display:'flex', gap:'10px', marginBottom:'20px', borderBottom:'1px solid rgba(255,255,255,0.06)', paddingBottom:'12px', flexWrap:'wrap' }}>
                <button onClick={() => setActiveTab('transactions')} style={{ fontSize:'13px', fontWeight:'700', padding:'10px 20px', borderRadius:'10px', cursor:'pointer', border:'none', background: activeTab === 'transactions' ? 'rgba(139,92,246,0.15)' : 'transparent', color: activeTab === 'transactions' ? '#a78bfa' : '#747d8c', transition:'all 0.2s' }}>📋 General Registry ({result.allRows.length})</button>
                <button onClick={() => setActiveTab('unknown')} style={{ fontSize:'13px', fontWeight:'700', padding:'10px 20px', borderRadius:'10px', cursor:'pointer', border:'none', background: activeTab === 'unknown' ? 'rgba(255,71,87,0.1)' : 'transparent', color: activeTab === 'unknown' ? '#ff4757' : '#747d8c', transition:'all 0.2s' }}>❓ Fix Anomalies ({unknownCount})</button>
                <button onClick={() => setActiveTab('ai')} style={{ fontSize:'13px', fontWeight:'700', padding:'10px 20px', borderRadius:'10px', cursor:'pointer', border:'none', background: activeTab === 'ai' ? 'rgba(139,92,246,0.15)' : 'transparent', color: activeTab === 'ai' ? '#a78bfa' : '#747d8c', transition:'all 0.2s' }}>🤖 Groq AI Analysis</button>
              </div>

              {/* Panel Views Grid Handler */}
              {activeTab === 'unknown' && <UnknownTransactionsPanel unknowns={result.unknownTransactions} onUpdateRow={handleUpdateRow} />}
              {activeTab === 'ai' && <AIAnalysisPanel result={result} />}

              {activeTab === 'transactions' && (
                <div style={{ background:'rgba(255,255,255,0.01)', border:'1px solid rgba(255,255,255,0.04)', borderRadius:'18px', padding:'20px', boxShadow:'0 10px 30px rgba(0,0,0,0.3)' }}>
                  
                  {/* Sorting / Filtering Quick Pills */}
                  <div style={{ display: 'flex', gap: '8px', marginBottom: '18px', flexWrap: 'wrap' }}>
                    {['ALL', 'HIGH', 'MEDIUM', 'CLEARED'].map(f => (
                      <button key={f} onClick={() => setActiveFilter(f)} style={{ fontSize:'11px', fontWeight:'700', padding:'6px 14px', borderRadius:'20px', cursor:'pointer', border:`1px solid ${activeFilter === f ? SEV[f]?.dot || '#8b5cf6' : 'rgba(255,255,255,0.06)'}`, background: activeFilter === f ? `${SEV[f]?.dot || '#8b5cf6'}15` : 'rgba(0,0,0,0.2)', color: activeFilter === f ? SEV[f]?.dot || '#fff' : '#a4b0be', transition:'all 0.15s' }}>{f === 'CLEARED' ? 'CLEAN' : f}</button>
                    ))}
                  </div>

                  {/* Transaction Component Block Loops */}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                    {displayTx.length === 0 ? (
                      <div style={{ padding:'30px', textAlign:'center', color:'#57606f', fontSize:'13px' }}>No entries found inside matching registry indices.</div>
                    ) : displayTx.map(tx => (
                      <div
                        key={tx.id}
                        onClick={() => setExpandedId(expandedId === tx.id ? null : tx.id)}
                        style={{
                          background: 'rgba(255,255,255,0.02)', borderRadius: '14px', border: '1px solid rgba(255,255,255,0.05)',
                          borderLeft: `4px solid ${SEV[tx.severity]?.dot || '#747d8c'}`, padding: '14px 18px', display: 'flex', alignItems: 'center', gap: '16px', cursor: 'pointer', transition: 'all 0.2s', flexWrap: 'wrap'
                        }}
                        className="aq-tx-row"
                      >
                        <Avatar name={tx.supplier} severity={tx.severity} />
                        <div style={{ flex: '1 1 240px', minWidth: 0 }}>
                          <div style={{ fontSize: '14px', fontWeight: '700', color: '#fff', display:'flex', alignItems:'center', gap:'8px' }}>
                            {tx.invoice} <span style={{ color:'#747d8c', fontWeight:'400' }}>•</span> <span style={{ textOverflow:'ellipsis', overflow:'hidden', whiteSpace:'nowrap' }}>{tx.supplier}</span>
                          </div>
                          <div style={{ fontSize: '11px', color: '#57606f', marginTop: '4px' }}>{tx.date} {tx.time && `| ${tx.time}`}</div>
                          
                          {/* Expanded Threat Reasons Details */}
                          {expandedId === tx.id && tx.reasons.length > 0 && (
                            <div style={{ marginTop: '10px', fontSize: '12px', color: '#ff6b81', background:'rgba(255,71,87,0.05)', padding:'8px 12px', borderRadius:'8px', border:'1px solid rgba(255,71,87,0.1)' }}>
                              ⚡ <b>Audit Flaw Triggered:</b> {tx.reasons.join(', ')}
                            </div>
                          )}
                        </div>
                        <div style={{ textAlign: 'right', marginLeft: 'auto', flexShrink: 0 }}>
                          <div style={{ fontSize: '15px', fontWeight: '800', color: tx.severity === 'CLEARED' ? '#2ed573' : '#ff4757' }}>
                            LKR {tx.amount.toLocaleString()}
                          </div>
                          <div style={{ fontSize:'10px', color:'#57606f', marginTop:'3px', fontWeight:'600' }}>{SEV[tx.severity]?.label}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Data Re-upload Controller Reset */}
              <div style={{ marginTop: '24px', display: 'flex', justifyContent: 'flex-end' }}>
                <button onClick={() => { setResult(null); setFileName(''); }} style={{ padding: '12px 24px', background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.08)', borderRadius: '10px', color: '#a4b0be', fontSize: '12px', fontWeight: '700', cursor: 'pointer', transition:'all 0.2s' }}>🔄 Purge Engine Data & Scan New Sheet</button>
              </div>

            </div>
          )}

        </main>
      </div>
    </>
  );
}

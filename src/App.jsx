import React, { useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

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
  data.forEach((row, idx) => {
    const reasons = [];
    const severity = [];
    const amt = parseFloat(row.Amount || row.amount || row.Value || row.value || 0);
    const inv = (row.Invoice_No || row.invoice_no || row['Invoice No'] || '').toString().trim();
    const sup = (row.Supplier || row.supplier || row.Vendor || row.vendor || 'Unknown').toString().trim();
    const date = (row.Date || row.date || new Date().toISOString().split('T')[0]).toString().trim();
    const time = (row.Time || row.time || '').toString().trim();

    if (amt > 500000) { reasons.push('Large Transaction (>500k)'); severity.push('HIGH'); }
    const suspiciousWords = ['unknown', 'n/a', 'none', 'test', 'dummy', 'sample'];
    if (suspiciousWords.some(w => sup.toLowerCase().includes(w))) { reasons.push('Unknown Supplier'); severity.push('HIGH'); }
    if (inv && invoiceCounts[inv] > 1) { reasons.push(`Duplicate Invoice Ref`); severity.push('HIGH'); }
    if (amt > 0 && amt % 10000 === 0) { reasons.push('Round-number Amount'); severity.push('MEDIUM'); }
    if (sup && supplierAmounts[sup] > 2000000) { reasons.push('Supplier Cumulative Spend >2M'); severity.push('MEDIUM'); }
    if (date && dateGroups[date] > 10) { reasons.push(`High Volume Day`); severity.push('LOW'); }
    if (!inv || inv === '') { reasons.push('Missing Invoice Reference'); severity.push('MEDIUM'); }

    if (reasons.length > 0) {
      const topSeverity = severity.includes('HIGH') ? 'HIGH' : severity.includes('MEDIUM') ? 'MEDIUM' : 'LOW';
      flagged.push({ id: idx, date, time, supplier: sup, invoice: inv || 'N/A', amount: amt, reasons, severity: topSeverity });
    }
  });

  const riskLevel = flagged.filter(f => f.severity === 'HIGH').length > 3 ? 'HIGH'
    : flagged.length === 0 ? 'LOW' : flagged.length <= 3 ? 'MEDIUM' : 'HIGH';

  return { flagged, riskLevel, total: data.length };
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
      Invoice_No: invMatch ? invMatch[0].toUpperCase() : `INV${String(i + 1).padStart(4, '0')}`
    };
  }).filter(Boolean);
}

// ─── Severity config ──────────────────────────────────────────────────────────
const SEV = {
  HIGH:   { dot: '#ef4444', badge: { bg: 'rgba(239,68,68,0.18)',   color: '#fca5a5', border: 'rgba(239,68,68,0.35)'   }, leftBorder: '#ef4444', label: 'Flagged',    iconBg: 'rgba(239,68,68,0.15)',   icon: '🔴' },
  MEDIUM: { dot: '#f59e0b', badge: { bg: 'rgba(245,158,11,0.18)',  color: '#fcd34d', border: 'rgba(245,158,11,0.35)'  }, leftBorder: '#f59e0b', label: 'Suspicious', iconBg: 'rgba(245,158,11,0.15)',  icon: '🟡' },
  LOW:    { dot: '#22c55e', badge: { bg: 'rgba(34,197,94,0.18)',   color: '#86efac', border: 'rgba(34,197,94,0.35)'   }, leftBorder: '#22c55e', label: 'Normal',     iconBg: 'rgba(34,197,94,0.15)',   icon: '🟢' },
};

// Supplier initials avatar
function Avatar({ name, severity }) {
  const initials = name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  const bg = SEV[severity].iconBg;
  const color = SEV[severity].dot;
  return (
    <div style={{
      width: '38px', height: '38px', borderRadius: '10px', flexShrink: 0,
      background: bg, border: `1px solid ${color}33`,
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '13px', fontWeight: '700', color,
    }}>{initials || '??'}</div>
  );
}

// Pulse dot
function PulseDot({ color, animate }) {
  return <span style={{
    display: 'inline-block', width: '7px', height: '7px', borderRadius: '50%',
    background: color, flexShrink: 0,
    animation: animate ? 'aqPulse 1.5s infinite' : 'none',
  }} />;
}

// Risk bar
function RiskBar({ label, pct, color }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '9px' }}>
      <span style={{ fontSize: '12px', color: '#9ca3af', width: '140px', flexShrink: 0 }}>{label}</span>
      <div style={{ flex: 1, height: '5px', background: 'rgba(255,255,255,0.08)', borderRadius: '3px', overflow: 'hidden' }}>
        <div style={{ width: `${pct}%`, height: '100%', background: color, borderRadius: '3px', transition: 'width 0.9s cubic-bezier(.4,0,.2,1)' }} />
      </div>
      <span style={{ fontSize: '12px', fontWeight: '600', color: '#e5e7eb', width: '36px', textAlign: 'right' }}>{pct}%</span>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [loading,    setLoading]    = useState(false);
  const [result,     setResult]     = useState(null);
  const [fileName,   setFileName]   = useState('');
  const [dragOver,   setDragOver]   = useState(false);
  const [activeFilter, setActiveFilter] = useState('ALL');
  const [expandedId, setExpandedId] = useState(null);
  const fileInputRef = useRef();

  const processFile = useCallback(async (file) => {
    if (!file) return;
    setLoading(true); setFileName(file.name); setResult(null); setExpandedId(null); setActiveFilter('ALL');
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
      } else if (ext === 'pdf') {
        if (!window.pdfjsLib) throw new Error('PDF.js not loaded.');
        window.pdfjsLib.GlobalWorkerOptions.workerSrc =
          'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        const pdf = await window.pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
        let text = '';
        for (let p = 1; p <= pdf.numPages; p++) {
          const pg = await pdf.getPage(p);
          text += (await pg.getTextContent()).items.map(i => i.str).join(' ') + '\n';
        }
        rows = parseTextToRows(text);
      } else {
        alert('Unsupported format. Use CSV, XLSX, DOCX, or PDF.');
        setLoading(false); return;
      }
      if (rows.length === 0) throw new Error('No transaction data found.');
      setResult(runFraudDetection(rows));
    } catch (err) { alert('Error: ' + err.message); }
    setLoading(false);
  }, []);

  const handleChange = e => processFile(e.target.files[0]);
  const handleDrop   = e => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); };

  // Derived
  const flaggedCount = result ? result.flagged.filter(f => f.severity === 'HIGH').length   : 0;
  const suspCount    = result ? result.flagged.filter(f => f.severity === 'MEDIUM').length : 0;
  const clearCount   = result ? result.total - result.flagged.length : 0;
  const totalCount   = result ? result.total : 0;

  // Reason breakdown %
  const rb = { dup: 0, missing: 0, mismatch: 0, timing: 0 };
  if (result) result.flagged.forEach(f => f.reasons.forEach(r => {
    if (r.includes('Duplicate'))  rb.dup++;
    if (r.includes('Missing'))    rb.missing++;
    if (r.includes('Round'))      rb.mismatch++;
    if (r.includes('High Volume') || r.includes('Cumulative')) rb.timing++;
  }));
  const total_rb = Math.max(result ? result.flagged.length : 1, 1);
  const pct = v => Math.round((v / total_rb) * 100);

  const filtered = result
    ? (activeFilter === 'ALL' ? result.flagged
      : activeFilter === 'FLAGGED'    ? result.flagged.filter(f => f.severity === 'HIGH')
      : activeFilter === 'SUSPICIOUS' ? result.flagged.filter(f => f.severity === 'MEDIUM')
      : result.flagged.filter(f => f.severity === 'LOW'))
    : [];

  // Demo transactions to show when no file loaded
  const demoTx = [
    { id:0, supplier:'Apex Supplies Ltd',   invoice:'INV-2024-0871', amount:84500,  date:'2026-05-23', time:'10:42 AM', severity:'HIGH',   reasons:['Duplicate Invoice Ref'] },
    { id:1, supplier:'Account: 0029-***-4412', invoice:'TXN-9934',  amount:127000, date:'2026-05-23', time:'02:17 AM', severity:'MEDIUM', reasons:['High Volume Day'] },
    { id:2, supplier:'SriTech Pvt Ltd',     invoice:'PO-55221',     amount:45200,  date:'2026-05-23', time:'09:15 AM', severity:'LOW',    reasons:[] },
    { id:3, supplier:'Finance Dept',        invoice:'REF-0045',     amount:210000, date:'2026-05-23', time:'11:58 AM', severity:'HIGH',   reasons:['Missing Invoice Reference'] },
    { id:4, supplier:'Account: 0081-***-7723', invoice:'TXN-8820', amount:49900,  date:'2026-05-23', time:'01:30 PM', severity:'MEDIUM', reasons:['Round-number Amount'] },
  ];
  const displayTx = result ? filtered : demoTx;
  const showingDemo = !result;

  // ── Icon button ───────────────────────────────────────────────────────────
  const IBtn = ({ children, title, hasNotif }) => (
    <div title={title} style={{ position: 'relative', cursor: 'pointer' }}>
      <div style={{
        width: '32px', height: '32px', borderRadius: '8px',
        border: '1px solid rgba(139,92,246,0.3)', background: 'rgba(255,255,255,0.05)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '15px',
        transition: 'all 0.15s', color: '#a78bfa',
      }}
        onMouseEnter={e => { e.currentTarget.style.background = 'rgba(139,92,246,0.2)'; e.currentTarget.style.transform = 'scale(1.08)'; }}
        onMouseLeave={e => { e.currentTarget.style.background = 'rgba(255,255,255,0.05)'; e.currentTarget.style.transform = 'scale(1)'; }}
      >{children}</div>
      {hasNotif && <span style={{ position:'absolute', top:'4px', right:'4px', width:'7px', height:'7px', background:'#ef4444', borderRadius:'50%', border:'1.5px solid #0b0f1e' }} />}
    </div>
  );

  // ── Filter chip ───────────────────────────────────────────────────────────
  const FChip = ({ id, label, icon }) => {
    const active = activeFilter === id;
    return (
      <button onClick={() => setActiveFilter(id)} style={{
        fontSize: '12px', fontWeight: '600', padding: '5px 13px', borderRadius: '20px',
        cursor: 'pointer', border: `1px solid ${active ? 'rgba(139,92,246,0.6)' : 'rgba(255,255,255,0.15)'}`,
        background: active ? 'rgba(139,92,246,0.25)' : 'rgba(255,255,255,0.05)',
        color: active ? '#c4b5fd' : '#9ca3af',
        display: 'flex', alignItems: 'center', gap: '5px', transition: 'all 0.15s',
      }}>
        {icon && <span style={{ fontSize: '11px' }}>{icon}</span>}
        {label}
      </button>
    );
  };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0b0f1e !important; }
        @keyframes aqPulse  { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.4;transform:scale(1.4)} }
        @keyframes aqSpin   { to { transform: rotate(360deg); } }
        @keyframes aqFadeUp { from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)} }
        @keyframes aqSlide  { from{opacity:0;transform:translateX(-8px)} to{opacity:1;transform:translateX(0)} }
        .aq-stat:hover  { transform: translateY(-2px) !important; box-shadow: 0 8px 24px rgba(139,92,246,0.18) !important; }
        .aq-tx:hover    { border-color: rgba(139,92,246,0.35) !important; transform: translateX(3px); }
        .aq-results     { animation: aqFadeUp 0.4s ease both; }
      `}</style>

      <div style={{
        minHeight: '100vh', background: '#0b0f1e',
        fontFamily: "'Inter', ui-sans-serif, system-ui, sans-serif",
        color: '#f3f4f6',
      }}>

        {/* ── Top Bar ────────────────────────────────────────────────────── */}
        <header style={{
          background: 'rgba(11,15,30,0.95)', backdropFilter: 'blur(20px)',
          borderBottom: '1px solid rgba(139,92,246,0.2)',
          padding: '0 20px', position: 'sticky', top: 0, zIndex: 200,
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', height: '56px',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{
              width: '34px', height: '34px', borderRadius: '9px', flexShrink: 0,
              background: 'linear-gradient(135deg, #6d28d9, #a855f7)',
              display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '18px',
            }}>🛡️</div>
            <div>
              <div style={{ fontSize: '15px', fontWeight: '800', color: '#fff', letterSpacing: '-0.3px', lineHeight: 1.2 }}>AuditIQ</div>
              <div style={{ fontSize: '10px', color: '#8b5cf6', fontWeight: '500' }}>Transaction anomaly detection</div>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <IBtn title="Search">🔍</IBtn>
            <IBtn title="Settings">⚙️</IBtn>
            <IBtn title="Notifications" hasNotif={result && result.flagged.length > 0}>🔔</IBtn>
            <IBtn title="Export">📤</IBtn>
          </div>
        </header>

        <main style={{ maxWidth: '960px', margin: '0 auto', padding: '24px 20px' }}>

          {/* ── Upload zone (shows when no result) ──────────────────────── */}
          {!result && (
            <div
              onDragOver={e => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current.click()}
              style={{
                borderRadius: '16px',
                border: `2px dashed ${dragOver ? '#7c3aed' : 'rgba(139,92,246,0.35)'}`,
                background: dragOver ? 'rgba(124,58,237,0.08)' : 'rgba(255,255,255,0.03)',
                padding: '36px 24px', textAlign: 'center', cursor: 'pointer',
                transition: 'all 0.2s', marginBottom: '24px',
              }}>
              <div style={{ fontSize: '32px', marginBottom: '10px' }}>📂</div>
              <div style={{ fontSize: '16px', fontWeight: '700', color: '#fff', marginBottom: '6px' }}>Upload Transaction Document</div>
              <div style={{ fontSize: '13px', color: '#6b7280', marginBottom: '16px' }}>
                CSV, XLSX, PDF, or DOCX — drag & drop or click to browse
              </div>
              <button onClick={e => { e.stopPropagation(); fileInputRef.current.click(); }} style={{
                background: 'linear-gradient(135deg,#7c3aed,#9333ea)', color: '#fff',
                border: 'none', borderRadius: '10px', padding: '9px 22px',
                fontSize: '13px', fontWeight: '700', cursor: 'pointer',
              }}>⚡ Choose File</button>
              {fileName && <div style={{ marginTop:'12px', fontSize:'12px', color:'#8b5cf6', fontWeight:'600' }}>📄 {fileName}</div>}
              <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" style={{ display:'none' }} onChange={handleChange} />
            </div>
          )}

          {/* ── Loading ────────────────────────────────────────────────── */}
          {loading && (
            <div style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', background:'rgba(139,92,246,0.08)', border:'1px solid rgba(139,92,246,0.2)', borderRadius:'12px', padding:'16px', marginBottom:'24px', fontSize:'13px', fontWeight:'600', color:'#a78bfa' }}>
              <span style={{ display:'inline-block', animation:'aqSpin 0.9s linear infinite', fontSize:'18px' }}>🔄</span>
              Analysing transactions...
            </div>
          )}

          {/* ── Stat cards ─────────────────────────────────────────────── */}
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))', gap:'12px', marginBottom:'20px' }}>

            {/* Flagged */}
            <div className="aq-stat" style={{ background:'rgba(239,68,68,0.08)', border:'1px solid rgba(239,68,68,0.25)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#ef4444" animate={true} /> Flagged
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? 12 : flaggedCount}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(239,68,68,0.18)', color:'#fca5a5', border:'1px solid rgba(239,68,68,0.3)' }}>
                ↑ {showingDemo ? 4 : flaggedCount} new
              </span>
            </div>

            {/* Suspicious */}
            <div className="aq-stat" style={{ background:'rgba(245,158,11,0.08)', border:'1px solid rgba(245,158,11,0.25)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.07s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#f59e0b" animate={true} /> Suspicious
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? 38 : suspCount}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(245,158,11,0.18)', color:'#fcd34d', border:'1px solid rgba(245,158,11,0.3)' }}>
                Review needed
              </span>
            </div>

            {/* Cleared */}
            <div className="aq-stat" style={{ background:'rgba(34,197,94,0.07)', border:'1px solid rgba(34,197,94,0.2)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.14s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#22c55e" animate={false} /> Cleared
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? '1,204' : clearCount.toLocaleString()}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(34,197,94,0.15)', color:'#86efac', border:'1px solid rgba(34,197,94,0.3)' }}>
                {showingDemo ? '97.2% pass' : (totalCount > 0 ? `${Math.round((clearCount/totalCount)*100)}% pass` : '—')}
              </span>
            </div>

            {/* Total scanned */}
            <div className="aq-stat" style={{ background:'rgba(139,92,246,0.08)', border:'1px solid rgba(139,92,246,0.25)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.21s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                📊 Total scanned
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? '1,254' : totalCount.toLocaleString()}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(139,92,246,0.18)', color:'#c4b5fd', border:'1px solid rgba(139,92,246,0.3)' }}>
                Today
              </span>
            </div>
          </div>

          {/* ── Risk Breakdown ─────────────────────────────────────────── */}
          <div style={{ background:'rgba(255,255,255,0.03)', border:'1px solid rgba(139,92,246,0.15)', borderRadius:'14px', padding:'18px 20px', marginBottom:'20px' }}>
            <div style={{ fontSize:'13px', fontWeight:'700', color:'#e5e7eb', marginBottom:'14px', display:'flex', alignItems:'center', gap:'6px' }}>
              ⚡ Risk breakdown
            </div>
            <RiskBar label="Duplicate entries"  pct={showingDemo ? 72 : pct(rb.dup)}      color="#ef4444" />
            <RiskBar label="Missing refs"        pct={showingDemo ? 55 : pct(rb.missing)}  color="#f59e0b" />
            <RiskBar label="Amount mismatch"     pct={showingDemo ? 38 : pct(rb.mismatch)} color="#8b5cf6" />
            <RiskBar label="Timing anomalies"    pct={showingDemo ? 21 : pct(rb.timing)}   color="#6366f1" />
          </div>

          {/* ── Recent Transactions ────────────────────────────────────── */}
          <div style={{ background:'rgba(255,255,255,0.02)', border:'1px solid rgba(139,92,246,0.15)', borderRadius:'14px', padding:'18px 20px' }}>

            {/* Section header */}
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:'14px' }}>
              <div style={{ fontSize:'13px', fontWeight:'700', color:'#e5e7eb', display:'flex', alignItems:'center', gap:'6px' }}>
                ≡ Recent transactions
              </div>
              {result && (
                <button onClick={() => fileInputRef.current.click()} style={{ background:'none', border:'none', cursor:'pointer', fontSize:'12px', color:'#8b5cf6', fontWeight:'600', display:'flex', alignItems:'center', gap:'4px' }}>
                  View all →
                </button>
              )}
            </div>

            {/* Filter chips */}
            <div style={{ display:'flex', gap:'8px', marginBottom:'14px', flexWrap:'wrap' }}>
              <FChip id="ALL"        label="All"        icon="⊞" />
              <FChip id="FLAGGED"    label="Flagged"    icon="🚩" />
              <FChip id="SUSPICIOUS" label="Suspicious" icon="⚠️" />
              <FChip id="LOW"        label="Normal"     icon="✓" />
            </div>

            {/* Transaction cards */}
            <div style={{ display:'flex', flexDirection:'column', gap:'8px' }}>
              {displayTx.length === 0 ? (
                <div style={{ textAlign:'center', padding:'32px', color:'#6b7280', fontSize:'13px' }}>
                  No transactions match this filter.
                </div>
              ) : displayTx.map((tx, i) => (
                <div
                  key={tx.id}
                  className="aq-tx"
                  onClick={() => setExpandedId(expandedId === tx.id ? null : tx.id)}
                  style={{
                    background: 'rgba(255,255,255,0.03)',
                    borderRadius: '12px',
                    borderLeft: `3px solid ${SEV[tx.severity].dot}`,
                    border: '1px solid rgba(255,255,255,0.07)',
                    borderLeftWidth: '3px', borderLeftColor: SEV[tx.severity].dot,
                    padding: '12px 14px',
                    display: 'flex', alignItems: 'center', gap: '12px',
                    cursor: 'pointer', transition: 'all 0.18s ease',
                    animation: `aqSlide 0.25s ease ${i * 0.05}s both`,
                    boxShadow: expandedId === tx.id ? '0 4px 20px rgba(139,92,246,0.12)' : 'none',
                  }}
                >
                  <Avatar name={tx.supplier} severity={tx.severity} />

                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: '13px', fontWeight: '700', color: '#fff', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {tx.invoice} · {tx.reasons[0] || 'Cleared'}
                    </div>
                    <div style={{ fontSize: '11px', color: '#6b7280', marginTop: '3px' }}>
                      {tx.severity === 'MEDIUM' ? `Account: ${tx.supplier}` : `Vendor: ${tx.supplier}`}
                      {tx.time && <span style={{ marginLeft: '10px' }}>{tx.time}</span>}
                    </div>
                    {expandedId === tx.id && tx.reasons.length > 1 && (
                      <div style={{ marginTop: '8px', display: 'flex', flexWrap: 'wrap', gap: '5px' }}>
                        {tx.reasons.slice(1).map((r, ri) => (
                          <span key={ri} style={{ fontSize: '10px', fontWeight: '600', padding: '2px 8px', borderRadius: '6px', background: 'rgba(139,92,246,0.15)', color: '#c4b5fd', border: '1px solid rgba(139,92,246,0.25)' }}>⚡ {r}</span>
                        ))}
                      </div>
                    )}
                  </div>

                  <div style={{ textAlign: 'right', flexShrink: 0 }}>
                    <div style={{ fontSize: '14px', fontWeight: '800', color: tx.severity === 'LOW' ? '#86efac' : '#fca5a5' }}>
                      -LKR {tx.amount.toLocaleString()}
                    </div>
                    <span style={{
                      display: 'inline-block', marginTop: '4px',
                      fontSize: '10px', fontWeight: '700', padding: '2px 9px', borderRadius: '20px',
                      background: SEV[tx.severity].badge.bg,
                      color: SEV[tx.severity].badge.color,
                      border: `1px solid ${SEV[tx.severity].badge.border}`,
                    }}>{SEV[tx.severity].label}</span>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* ── Bottom action buttons ──────────────────────────────────── */}
          <div style={{ display:'flex', gap:'10px', marginTop:'16px', flexWrap:'wrap' }}>
            <button style={{
              flex:1, minWidth:'100px', padding:'11px', borderRadius:'10px',
              border:'1px solid rgba(139,92,246,0.3)', background:'rgba(255,255,255,0.04)',
              color:'#a78bfa', fontSize:'13px', fontWeight:'700', cursor:'pointer',
              display:'flex', alignItems:'center', justifyContent:'center', gap:'7px', transition:'all 0.15s',
            }}
              onMouseEnter={e => e.currentTarget.style.background='rgba(139,92,246,0.15)'}
              onMouseLeave={e => e.currentTarget.style.background='rgba(255,255,255,0.04)'}
            >📤 Export</button>

            <button style={{
              flex:1, minWidth:'100px', padding:'11px', borderRadius:'10px',
              border:'1px solid rgba(139,92,246,0.3)', background:'rgba(255,255,255,0.04)',
              color:'#a78bfa', fontSize:'13px', fontWeight:'700', cursor:'pointer',
              display:'flex', alignItems:'center', justifyContent:'center', gap:'7px', transition:'all 0.15s',
            }}
              onMouseEnter={e => e.currentTarget.style.background='rgba(139,92,246,0.15)'}
              onMouseLeave={e => e.currentTarget.style.background='rgba(255,255,255,0.04)'}
            >🔗 API status</button>

            <button
              onClick={() => fileInputRef.current.click()}
              style={{
                flex:1, minWidth:'100px', padding:'11px', borderRadius:'10px',
                border:'none', background:'linear-gradient(135deg,#7c3aed,#9333ea)',
                color:'#fff', fontSize:'13px', fontWeight:'700', cursor:'pointer',
                display:'flex', alignItems:'center', justifyContent:'center', gap:'7px',
                boxShadow:'0 4px 16px rgba(124,58,237,0.35)', transition:'all 0.15s',
              }}
              onMouseEnter={e => { e.currentTarget.style.transform='translateY(-1px)'; e.currentTarget.style.boxShadow='0 6px 20px rgba(124,58,237,0.45)'; }}
              onMouseLeave={e => { e.currentTarget.style.transform='translateY(0)'; e.currentTarget.style.boxShadow='0 4px 16px rgba(124,58,237,0.35)'; }}
            >▶ Run audit scan ↗</button>
          </div>

          {/* File input (hidden, triggered by Run audit scan too) */}
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" style={{ display:'none' }} onChange={handleChange} />

        </main>
      </div>
    </>
  );
}

import React, { useState, useCallback, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

// ─── Confidential API Key Security Matrix ───────────────────────────────────
const p1 = "gsk_YtHo256hGqEMTcxwYJf0";
const p2 = "WGdyb3FYZuG9vO9hnNtO93D3zOgbxOEb";
const CONFIDENTIAL_KEY = p1 + p2;

// ─── Audit Reference Generator ───────────────────────────────────────────────
function genAuditRef() {
  const ts = Date.now().toString(36).toUpperCase();
  const rand = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `AUD-${ts}-${rand}`;
}

// ─── Benford's Law Mathematical Analysis Engine ──────────────────────────────
function calculateBenfordsLaw(data) {
  const counts = Array(10).fill(0);
  let totalValid = 0;

  data.forEach(row => {
    const amtStr = (row.Amount || row.amount || row.Value || row.value || '').toString().trim().replace(/[^0-9.]/g, '');
    if (amtStr) {
      const firstDigit = parseInt(amtStr[0]);
      if (firstDigit >= 1 && firstDigit <= 9) {
        counts[firstDigit]++;
        totalValid++;
      }
    }
  });

  const ideal = [0, 30.1, 17.6, 12.5, 9.7, 7.9, 6.7, 5.8, 5.1, 4.6];
  const actual = counts.map((c, i) => i === 0 ? 0 : totalValid > 0 ? parseFloat(((c / totalValid) * 100).toFixed(1)) : 0);

  return { actual, ideal, totalValid };
}

// ─── Advanced Forensic Fraud Engine ──────────────────────────────────────────
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
    const isFixed = row.isFixed || false;

    const unknownIndicators = ['unknown', 'n/a', 'none', 'test', 'dummy', 'sample', 'misc', 'miscellaneous', 'other', '???', 'unidentified', 'anonymous'];
    const isUnknownSupplier = unknownIndicators.some(w => sup.toLowerCase().includes(w));
    const isUnknownInvoice = !inv || inv === '' || inv.toLowerCase() === 'n/a';
    const isUnknownDesc = !desc || desc === '' || unknownIndicators.some(w => desc.toLowerCase().includes(w));
    const isMaskedAccount = /\*{2,}|x{3,}/i.test(sup) || /\*{2,}|x{3,}/i.test(inv);

    // Ellyeta ganneth fix karapu nathi ewai
    if (!isFixed && (isUnknownSupplier || (isUnknownInvoice && isUnknownDesc) || isMaskedAccount)) {
      unknownTransactions.push({
        id: idx, date, time, supplier: sup, invoice: inv || 'N/A', amount: amt, description: desc || '—', isFixed: false,
        unknownReasons: [
          isUnknownSupplier && 'Unknown/masked supplier',
          isUnknownInvoice && 'Missing invoice reference',
          isUnknownDesc && 'No transaction description',
          isMaskedAccount && 'Masked account identifier',
        ].filter(Boolean),
        auditRef: genAuditRef(), flaggedAt: new Date().toISOString(),
      });
    }

    if (amt > 500000) { reasons.push('Large Transaction (>500k)'); severity.push('HIGH'); }
    if (isUnknownSupplier && !isFixed) { reasons.push('Unknown Supplier'); severity.push('HIGH'); }
    if (inv && invoiceCounts[inv] > 1) { reasons.push('Duplicate Invoice Ref'); severity.push('HIGH'); }
    if (amt > 0 && amt % 10000 === 0) { reasons.push('Round-number Amount'); severity.push('MEDIUM'); }
    if (sup && supplierAmounts[sup] > 2000000) { reasons.push('Supplier Cumulative Spend >2M'); severity.push('MEDIUM'); }
    if (date && dateGroups[date] > 10) { reasons.push('High Volume Day'); severity.push('LOW'); }
    if ((!inv || inv === '') && !isFixed) { reasons.push('Missing Invoice Reference'); severity.push('MEDIUM'); }
    if (isMaskedAccount) { reasons.push('Masked Account Identifier'); severity.push('HIGH'); }

    if (reasons.length > 0 && !isFixed) {
      const topSeverity = severity.includes('HIGH') ? 'HIGH' : severity.includes('MEDIUM') ? 'MEDIUM' : 'LOW';
      flagged.push({
        id: idx, date, time, supplier: sup, invoice: inv || 'N/A',
        amount: amt, reasons, severity: topSeverity, isFixed: false,
        auditRef: genAuditRef(), flaggedAt: new Date().toISOString(), description: desc || '—',
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
    const isFixed = row.isFixed || false;

    return {
      id: idx, date, time, supplier: sup, invoice: inv || 'N/A',
      amount: amt, reasons: [], severity: isFixed ? 'FIXED' : 'CLEARED', description: desc || '—', isFixed,
      auditRef: genAuditRef(), flaggedAt: new Date().toISOString(),
    };
  });

  const benford = calculateBenfordsLaw(data);

  return {
    flagged, unknownTransactions, allRows, riskLevel, benford,
    total: data.length, auditId: genAuditRef(), scanTime: new Date().toISOString(),
  };
}

function parseTextToRows(text) {
  const lines = text.split('\n').filter(l => l.trim().length > 5);
  return lines.map((line, i) => {
    const amtMatch = line.match(/\b(\d{4,9}(?:\.\d{1,2})?)\b/);
    const invMatch = line.match(/(?:INV|TXN|REF|PO|ORD)[-\s]?\d{3,8}/i);
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
      isFixed: false
    };
  }).filter(Boolean);
}

// ─── Severity Color Configurations ───────────────────────────────────────────
const SEV = {
  HIGH: { dot: '#ff4757', glow: 'rgba(255,71,87,0.4)', bg: 'rgba(255,71,87,0.06)', border: 'rgba(255,71,87,0.25)', label: 'Critical Fault' },
  MEDIUM: { dot: '#ffa502', glow: 'rgba(255,165,2,0.4)', bg: 'rgba(255,165,2,0.06)', border: 'rgba(255,165,2,0.25)', label: 'Suspicious' },
  LOW: { dot: '#2ed573', glow: 'rgba(46,213,115,0.4)', bg: 'rgba(46,213,115,0.06)', border: 'rgba(46,213,115,0.25)', label: 'Normal' },
  CLEARED: { dot: '#1e90ff', glow: 'rgba(30,144,255,0.4)', bg: 'rgba(30,144,255,0.06)', border: 'rgba(30,144,255,0.2)', label: 'Verified Safe' },
  FIXED: { dot: '#00d2d3', glow: 'rgba(0,210,211,0.4)', bg: 'rgba(0,210,211,0.06)', border: 'rgba(0,210,211,0.25)', label: 'Error Fixed' },
};

function Avatar({ name, severity }) {
  const initials = name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  const color = SEV[severity]?.dot || '#8b5cf6';
  return (
    <div style={{
      width: '38px', height: '38px', borderRadius: '12px', flexShrink: 0,
      background: 'rgba(255,255,255,0.03)', border: `1px solid ${color}44`,
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '12px', fontWeight: '700', color: color, boxShadow: `inset 0 0 8px ${color}11`
    }}>{initials || '??'}</div>
  );
}

function AIAnalysisPanel({ result }) {
  const [aiText, setAiText] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiDone, setAiDone] = useState(false);
  const [aiError, setAiError] = useState('');

  const runAI = async () => {
    setAiLoading(true); setAiText(''); setAiError('');
    const prompt = `Perform a forensic audit summary. Metrics - Total Logs: ${result?.total || 0}, High Risk Threats: ${result?.flagged?.filter(f => f.severity === 'HIGH').length || 0}, Medium Warnings: ${result?.flagged?.filter(f => f.severity === 'MEDIUM').length || 0}, Unresolved Anonymous Blocks: ${result?.unknownTransactions?.length || 0}. Strategic Threat Level Evaluated: ${result?.riskLevel || 'LOW'}. Give a brief executive report with bullet points.`;

    try {
      const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${CONFIDENTIAL_KEY}` },
        body: JSON.stringify({ model: 'llama-3.1-8b-instant', messages: [{ role: 'user', content: prompt }], temperature: 0.3, max_tokens: 600 }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const jsonRes = await res.json();
      setAiText(jsonRes.choices[0]?.message?.content || 'Empty intelligence sequence returned.');
      setAiDone(true);
    } catch (err) { setAiError('Audit Engine Error: ' + err.message); }
    setAiLoading(false);
  };

  return (
    <div style={{ background: 'linear-gradient(135deg, rgba(139,92,246,0.05), rgba(0,0,0,0.4))', border: '1px solid rgba(139,92,246,0.2)', backdropFilter: 'blur(10px)', borderRadius: '16px', padding: '20px', marginBottom: '20px' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '14px' }}>
        <div style={{ fontSize: '14px', fontWeight: '700', color: '#fff', display: 'flex', alignItems: 'center', gap: '8px' }}>🤖 Groq Forensic Executive Report</div>
        {!aiLoading && !aiDone && <button onClick={runAI} style={{ background: 'linear-gradient(135deg, #7c3aed, #9333ea)', color: '#fff', border: 'none', borderRadius: '10px', padding: '8px 18px', fontSize: '12px', fontWeight: '700', cursor: 'pointer' }}>⚡ Run Engine AI</button>}
      </div>
      {aiLoading && <div style={{ fontSize: '12px', color: '#a78bfa' }} className="blink">⚡ Processing neural core weights...</div>}
      {aiError && <div style={{ fontSize: '12px', color: '#ff4757' }}>{aiError}</div>}
      {aiText && <div style={{ fontSize: '13px', color: '#d1d5db', lineHeight: '1.7', whiteSpace: 'pre-wrap', padding: '14px', background: 'rgba(0,0,0,0.3)', borderRadius: '12px', border: '1px solid rgba(255,255,255,0.05)' }}>{aiText}</div>}
    </div>
  );
}

function AIChatAssistant({ result }) {
  const [messages, setMessages] = useState([
    { role: 'assistant', text: "Hello! I am your AI Auditor Bot. Ask me anything about the uploaded ledger dataset." }
  ]);
  const [input, setInput] = useState('');
  const [chatLoading, setChatLoading] = useState(false);
  const chatEndRef = useRef(null);

  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [messages]);

  const handleSendMessage = async (e) => {
    e.preventDefault();
    if (!input.trim() || chatLoading) return;

    const userMessage = input;
    setMessages(prev => [...prev, { role: 'user', text: userMessage }]);
    setInput('');
    setChatLoading(true);

    const contextPrompt = `You are a professional corporate forensic chat bot. Here is the active dataset context: Total Rows: ${result?.total || 0}, High Risk Faults: ${result?.flagged?.filter(f => f.severity === 'HIGH').length || 0}, Unresolved Identity Items: ${result?.unknownTransactions?.length || 0}. The auditor asks: "${userMessage}". Keep your response concise, expert-level, and directly actionable under 3 sentences.`;

    try {
      const res = await fetch('https://api.groq.com/openai/v1/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${CONFIDENTIAL_KEY}` },
        body: JSON.stringify({ model: 'llama-3.1-8b-instant', messages: [{ role: 'user', content: contextPrompt }], temperature: 0.5, max_tokens: 250 }),
      });
      const jsonRes = await res.ok ? await res.json() : null;
      const aiReply = jsonRes?.choices[0]?.message?.content || "Apologies, communication node timed out.";
      setMessages(prev => [...prev, { role: 'assistant', text: aiReply }]);
    } catch {
      setMessages(prev => [...prev, { role: 'assistant', text: "Error syncing with Llama core matrix network." }]);
    }
    setChatLoading(false);
  };

  return (
    <div style={{ background: 'rgba(255,255,255,0.01)', border: '1px solid rgba(255,255,255,0.05)', borderRadius: '16px', padding: '20px', marginBottom: '20px', display: 'flex', flexDirection: 'column', height: '350px' }}>
      <div style={{ fontSize: '14px', fontWeight: '700', color: '#a78bfa', marginBottom: '10px' }}>💬 Real-Time Auditor Bot Chat</div>
      <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '8px', paddingRight: '4px', marginBottom: '10px' }}>
        {messages.map((m, idx) => (
          <div key={idx} style={{ alignSelf: m.role === 'user' ? 'flex-end' : 'flex-start', background: m.role === 'user' ? '#7c3aed' : 'rgba(255,255,255,0.04)', color: '#fff', padding: '8px 12px', borderRadius: '10px', fontSize: '12px', maxWidth: '85%', lineHeight: '1.5' }}>
            {m.text}
          </div>
        ))}
        {chatLoading && <div style={{ fontSize: '11px', color: '#747d8c' }} className="blink">Bot typing...</div>}
        <div ref={chatEndRef} />
      </div>
      <form onSubmit={handleSendMessage} style={{ display: 'flex', gap: '8px' }}>
        <input type="text" value={input} onChange={e => setInput(e.target.value)} placeholder="Ask AI about suspicious suppliers..." style={{ flex: 1, padding: '8px 12px', background: '#0b0f1e', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', borderRadius: '8px', fontSize: '12px' }} />
        <button type="submit" style={{ background: '#7c3aed', color: '#fff', border: 'none', padding: '0 16px', borderRadius: '8px', fontSize: '12px', fontWeight: '700', cursor: 'pointer' }}>Send</button>
      </form>
    </div>
  );
}

// ─── Unknown Transactions & Fixed Workspace Panel ────────────────────────────────────
function UnknownTransactionsPanel({ unknowns, fixedRows, onUpdateRow, onReEditRow }) {
  const [editingId, setEditingId] = useState(null);
  const [supVal, setSupVal] = useState('');
  const [invVal, setInvVal] = useState('');

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
      
      {/* SECTION 1: INCORRECT ANOMALIES AREA */}
      <div style={{ background: 'linear-gradient(135deg, rgba(255,71,87,0.03), rgba(0,0,0,0.3))', border: '1px solid rgba(255,71,87,0.25)', borderRadius: '16px', padding: '20px' }}>
        <div style={{ fontSize: '14px', fontWeight: '700', color: '#ff6b81', marginBottom: '12px' }}>⚠️ Unresolved Identity Blocks ({unknowns.length})</div>
        {unknowns.length === 0 ? (
          <div style={{ fontSize: '12px', color: '#2ed573' }}>✨ All anomalies cleared from this layout!</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
            {unknowns.map(u => (
              <div key={u.id} style={{ background: 'rgba(255,255,255,0.02)', padding: '12px', borderRadius: '12px', border: '1px solid rgba(255,255,255,0.05)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <div style={{ fontSize: '13px', fontWeight: '700', color: '#fff' }}>{u.supplier}</div>
                    <div style={{ fontSize: '11px', color: '#ff4757', marginTop: '2px' }}>LKR {u.amount.toLocaleString()} <span style={{color:'#747d8c', marginLeft:'8px'}}>({u.invoice})</span></div>
                  </div>
                  <button onClick={() => { if (editingId === u.id) setEditingId(null); else { setEditingId(u.id); setSupVal(u.supplier); setInvVal(u.invoice); } }} style={{ background: '#ff475722', border: '1px solid #ff475744', color: '#ff4757', padding: '5px 12px', borderRadius: '6px', fontSize: '11px', cursor: 'pointer', fontWeight:'600' }}>
                    {editingId === u.id ? 'Close' : '🔧 Override'}
                  </button>
                </div>
                {editingId === u.id && (
                  <div style={{ marginTop: '10px', padding: '10px', background: 'rgba(0,0,0,0.3)', borderRadius: '8px' }}>
                    <div style={{ display: 'flex', gap: '8px', marginBottom: '8px' }}>
                      <div style={{ flex: 1 }}>
                        <label style={{ fontSize: '10px', color: '#a4b0be', display: 'block', marginBottom: '4px' }}>Correct Supplier</label>
                        <input type="text" value={supVal} onChange={e => setSupVal(e.target.value)} style={{ width: '100%', padding: '6px', background: '#060913', color: '#fff', border: '1px solid #333', borderRadius: '4px', fontSize: '11px' }} />
                      </div>
                      <div style={{ flex: 1 }}>
                        <label style={{ fontSize: '10px', color: '#a4b0be', display: 'block', marginBottom: '4px' }}>Correct Invoice</label>
                        <input type="text" value={invVal} onChange={e => setInvVal(e.target.value)} style={{ width: '100%', padding: '6px', background: '#060913', color: '#fff', border: '1px solid #333', borderRadius: '4px', fontSize: '11px' }} />
                      </div>
                    </div>
                    <button onClick={() => { onUpdateRow(u.id, supVal, invVal); setEditingId(null); }} style={{ background: '#2ed573', border: 'none', color: '#fff', padding: '6px 12px', borderRadius: '4px', fontSize: '11px', fontWeight: '700', cursor: 'pointer' }}>✔️ Save & Move to Fixed Section</button>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* NEW SECTION 2: ERROR FIXED WORKSPACE */}
      <div style={{ background: 'linear-gradient(135deg, rgba(0,210,211,0.03), rgba(0,0,0,0.3))', border: '1px solid rgba(0,210,211,0.25)', borderRadius: '16px', padding: '20px' }}>
        <div style={{ fontSize: '14px', fontWeight: '700', color: '#00d2d3', marginBottom: '12px' }}>🛠️ Error Fixed Workspace ({fixedRows.length})</div>
        {fixedRows.length === 0 ? (
          <div style={{ fontSize: '12px', color: '#747d8c' }}>No logs resolved yet. Hot-fix anomalies to populate this zone.</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
            {fixedRows.map(f => (
              <div key={f.id} style={{ background: 'rgba(255,255,255,0.02)', padding: '12px', borderRadius: '12px', border: '1px solid rgba(0,210,211,0.15)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <div style={{ fontSize: '13px', fontWeight: '700', color: '#fff' }}>{f.supplier} <span style={{fontSize:'10px', padding:'2px 6px', background:'#00d2d322', color:'#00d2d3', borderRadius:'4px', marginLeft:'6px'}}>Fixed</span></div>
                    <div style={{ fontSize: '11px', color: '#00d2d3', marginTop: '2px' }}>LKR {f.amount.toLocaleString()} <span style={{color:'#747d8c', marginLeft:'8px'}}>({f.invoice})</span></div>
                  </div>
                  <button onClick={() => onReEditRow(f.id)} style={{ background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)', color: '#fff', padding: '5px 12px', borderRadius: '6px', fontSize: '11px', cursor: 'pointer' }}>
                    ✏️ Re-Edit
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

    </div>
  );
}

// ─── Main View Application Core ──────────────────────────────────────────────
export default function App() {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [fileName, setFileName] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [activeFilter, setActiveFilter] = useState('ALL');
  const [activeTab, setActiveTab] = useState('transactions');
  const [expandedId, setExpandedId] = useState(null);
  const fileInputRef = useRef();

  const processFile = useCallback(async (file) => {
    if (!file) return;
    setLoading(true); setFileName(file.name); setResult(null);
    const ext = file.name.split('.').pop().toLowerCase();
    try {
      let rows = [];
      if (['xlsx', 'xls', 'csv'].includes(ext)) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf);
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      } else if (ext === 'docx') {
        const buf = await file.arrayBuffer();
        const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
        rows = parseTextToRows(value);
      }
      if (rows.length === 0) throw new Error('No compatible logs extracted.');
      
      // Injecting basic standard setup parameter keys
      const updatedRows = rows.map(r => ({...r, isFixed: false}));
      setResult(runFraudDetection(updatedRows));
    } catch (err) { alert('Data Matrix Error: ' + err.message); }
    setLoading(false);
  }, []);

  // 🔄 Hot-Fix & Move Logic: Item eka Incorrect list eken ain karala "Fixed Section" ekata danna
  const handleUpdateRow = (id, updatedSupplier, updatedInvoice) => {
    if (!result) return;
    
    const nextRows = result.allRows.map(row => 
      row.id === id 
        ? { ...row, supplier: updatedSupplier, invoice: updatedInvoice, isFixed: true } 
        : row
    );

    const dynamicInputData = nextRows.map(r => ({
      Date: r.date,
      Time: r.time,
      Amount: r.amount,
      Supplier: r.supplier,
      Invoice_No: r.invoice,
      Description: r.description,
      isFixed: r.isFixed
    }));

    setResult(runFraudDetection(dynamicInputData));
  };

  // ✏️ Re-Edit Logic: Fixed section eke thiyena ekak aayeth edit karanna oni unoth track karana filter eka
  const handleReEditRow = (id) => {
    if (!result) return;

    const nextRows = result.allRows.map(row => 
      row.id === id ? { ...row, isFixed: false } : row
    );

    const dynamicInputData = nextRows.map(r => ({
      Date: r.date,
      Time: r.time,
      Amount: r.amount,
      Supplier: r.supplier,
      Invoice_No: r.invoice,
      Description: r.description,
      isFixed: r.isFixed
    }));

    setResult(runFraudDetection(dynamicInputData));
  };

  // 📥 Excel Export Engine: Code structure for active downloads
  const downloadCorrectedFile = () => {
    if (!result || result.allRows.length === 0) return;

    const cleanedSheetData = result.allRows.map(r => ({
      'Date': r.date,
      'Time': r.time,
      'Supplier': r.supplier,
      'Invoice No': r.invoice,
      'Amount (LKR)': r.amount,
      'Description': r.description,
      'Audit Trace Status': r.isFixed ? 'ERROR FIXED' : r.severity
    }));

    const worksheet = XLSX.utils.json_to_sheet(cleanedSheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Audited_Ledger');
    XLSX.writeFile(workbook, `Corrected_Audit_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const flaggedCount = result ? result.flagged.filter(f => f.severity === 'HIGH').length : 0;
  const suspCount = result ? result.flagged.filter(f => f.severity === 'MEDIUM').length : 0;
  const unknownCount = result ? result.unknownTransactions.length : 0;
  const fixedRowsList = result ? result.allRows.filter(r => r.isFixed === true) : [];
  const clearCount = result ? result.allRows.filter(r => r.severity === 'CLEARED' || r.severity === 'FIXED').length : 0;

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #060913 !important; font-family: 'Plus Jakarta Sans', sans-serif; color: #f3f4f6; }
        .blink { animation: bAn 1.5s infinite; }
        @keyframes bAn { 0%,100% { opacity: 1; } 50% { opacity: 0.4; } }
      `}</style>

      <div style={{ minHeight: '100vh', background: '#060913', paddingBottom: '60px' }}>
        <header style={{ background: 'rgba(6,9,19,0.85)', borderBottom: '1px solid rgba(255,255,255,0.06)', padding: '0 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', height: '64px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ width: '36px', height: '36px', borderRadius: '10px', background: 'linear-gradient(135deg, #7c3aed, #ea580c)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>🛡️</div>
            <div>
              <div style={{ fontSize: '16px', fontWeight: '800', color: '#fff' }}>AuditIQ Premium Pro</div>
              <div style={{ fontSize: '10px', color: '#00d2d3', fontWeight: '600', letterSpacing: '0.5px' }}>EXHIBITION MATRIX UPDATE</div>
            </div>
          </div>
          {result && (
            <button onClick={downloadCorrectedFile} style={{ background: 'linear-gradient(135deg, #00d2d3, #01a3a4)', color: '#fff', border: 'none', borderRadius: '10px', padding: '8px 16px', fontSize: '12px', fontWeight: '700', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}>
              📥 Download Corrected Excel
            </button>
          )}
        </header>

        <main style={{ maxWidth: '1000px', margin: '0 auto', padding: '32px 24px' }}>
          {!result && (
            <div onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={e => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); }} onClick={() => fileInputRef.current.click()} style={{ borderRadius: '24px', border: `2px dashed ${dragOver ? '#7c3aed' : 'rgba(255,255,255,0.1)'}`, background: 'rgba(255,255,255,0.01)', padding: '80px 24px', textAlign: 'center', cursor: 'pointer' }}>
              <div style={{ fontSize: '54px', marginBottom: '16px' }}>📊</div>
              <div style={{ fontSize: '20px', fontWeight: '800', color: '#fff', marginBottom: '6px' }}>Scan Financial Databases & Files</div>
              <div style={{ fontSize: '13px', color: '#57606f', marginBottom: '20px' }}>Supports CSV, XLSX, XLS & Microsoft Word Docx</div>
              <button style={{ background: 'linear-gradient(135deg,#7c3aed,#9333ea)', color: '#fff', border: 'none', borderRadius: '12px', padding: '12px 28px', fontSize: '13px', fontWeight: '700' }}>⚡ Upload Ledger File</button>
              <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.docx" style={{ display: 'none' }} onChange={e => processFile(e.target.files[0])} />
            </div>
          )}

          {loading && <div style={{ color: '#a78bfa', textAlign: 'center', padding: '40px' }} className="blink">⚡ Processing cryptographic multi-format ledger blocks...</div>}

          {result && !loading && (
            <div>
              {/* Interactive Dashboard Counters */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(180px,1fr))', gap: '16px', marginBottom: '24px' }}>
                <div style={{ background: SEV.HIGH.bg, border: `1px solid ${SEV.HIGH.border}`, borderRadius: '16px', padding: '16px' }}>
                  <div style={{ fontSize: '11px', color: '#a4b0be', fontWeight: '700' }}>CRITICAL BREACH</div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#fff' }}>{flaggedCount}</div>
                </div>
                <div style={{ background: SEV.MEDIUM.bg, border: `1px solid ${SEV.MEDIUM.border}`, borderRadius: '16px', padding: '16px' }}>
                  <div style={{ fontSize: '11px', color: '#a4b0be', fontWeight: '700' }}>SUSPICIOUS LOGS</div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#fff' }}>{suspCount}</div>
                </div>
                <div style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.08)', borderRadius: '16px', padding: '16px' }}>
                  <div style={{ fontSize: '11px', color: '#a4b0be', fontWeight: '700' }}>ANOMALIES</div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#ff4757' }}>{unknownCount}</div>
                </div>
                <div style={{ background: SEV.FIXED.bg, border: `1px solid ${SEV.FIXED.border}`, borderRadius: '16px', padding: '16px' }}>
                  <div style={{ fontSize: '11px', color: '#a4b0be', fontWeight: '700' }}>ERRORS FIXED</div>
                  <div style={{ fontSize: '32px', fontWeight: '800', color: '#fff' }}>{fixedRowsList.length}</div>
                </div>
              </div>

              {/* Benford's Law and Distribution Overlays */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '16px', marginBottom: '24px' }}>
                <div style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.05)', borderRadius: '16px', padding: '20px' }}>
                  <div style={{ fontSize: '13px', fontWeight: '700', marginBottom: '14px', color: '#fff' }}>📈 Dataset Distribution Proportions</div>
                  <div style={{ display: 'flex', height: '24px', borderRadius: '6px', overflow: 'hidden', background: '#1e272e', marginBottom: '14px' }}>
                    <div style={{ width: `${result.total > 0 ? (flaggedCount / result.total) * 100 : 0}%`, background: '#ff4757' }} />
                    <div style={{ width: `${result.total > 0 ? (suspCount / result.total) * 100 : 0}%`, background: '#ffa502' }} />
                    <div style={{ width: `${result.total > 0 ? (fixedRowsList.length / result.total) * 100 : 0}%`, background: '#00d2d3' }} />
                    <div style={{ width: `${result.total > 0 ? (result.allRows.filter(r=>r.severity==='CLEARED').length / result.total) * 100 : 0}%`, background: '#1e90ff' }} />
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: '#a4b0be' }}>
                    <span>🔴 Fault: {result.total > 0 ? ((flaggedCount / result.total) * 100).toFixed(0) : 0}%</span>
                    <span>🟡 Fixed: {result.total > 0 ? ((fixedRowsList.length / result.total) * 100).toFixed(0) : 0}%</span>
                    <span>🔵 Clean: {result.total > 0 ? ((result.allRows.filter(r=>r.severity==='CLEARED').length / result.total) * 100).toFixed(0) : 0}%</span>
                  </div>
                </div>

                <div style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.05)', borderRadius: '16px', padding: '20px' }}>
                  <div style={{ fontSize: '13px', fontWeight: '700', marginBottom: '4px', color: '#fff' }}>📐 Benford's Law Digital Forensic Overlays</div>
                  <div style={{ fontSize: '10px', color: '#747d8c', marginBottom: '10px' }}>Checks first-digit variance for human tampering flags</div>
                  <div style={{ display: 'flex', alignItems: 'end', height: '80px', gap: '6px', paddingBottom: '5px', borderBottom: '1px solid #333' }}>
                    {result.benford.actual.slice(1, 7).map((val, i) => (
                      <div key={i} style={{ flex: 1, display: 'flex', flexDirection: 'column', justifyEnd: 'end', height: '100%' }}>
                        <div style={{ display: 'flex', gap: '2px', alignItems: 'end', height: '100%' }}>
                          <div style={{ width: '50%', height: `${val * 2.5}%`, background: '#a78bfa', borderRadius: '2px 2px 0 0' }} />
                          <div style={{ width: '50%', height: `${result.benford.ideal[i + 1] * 2.5}%`, background: 'rgba(255,255,255,0.15)', borderRadius: '2px 2px 0 0' }} />
                        </div>
                      </div>
                    ))}
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '10px', color: '#57606f', marginTop: '4px' }}>
                    <span>Digit: 1</span><span>Digit: 2</span><span>Digit: 3</span><span>Digit: 4</span><span>Digit: 5</span><span>Digit: 6</span>
                  </div>
                </div>
              </div>

              {/* Tabs Navigation */}
              <div style={{ display: 'flex', gap: '10px', marginBottom: '20px', borderBottom: '1px solid rgba(255,255,255,0.06)', paddingBottom: '12px' }}>
                <button onClick={() => setActiveTab('transactions')} style={{ padding: '8px 16px', background: activeTab === 'transactions' ? 'rgba(139,92,246,0.15)' : 'transparent', color: activeTab === 'transactions' ? '#a78bfa' : '#747d8c', border: 'none', borderRadius: '8px', fontSize: '13px', fontWeight: '700', cursor: 'pointer' }}>📋 Ledger Accounts</button>
                <button onClick={() => setActiveTab('unknown')} style={{ padding: '8px 16px', background: activeTab === 'unknown' ? 'rgba(0,210,211,0.15)' : 'transparent', color: activeTab === 'unknown' ? '#00d2d3' : '#747d8c', border: 'none', borderRadius: '8px', fontSize: '13px', fontWeight: '700', cursor: 'pointer' }}>⚙️ Fix Anomalies ({unknownCount})</button>
                <button onClick={() => setActiveTab('ai')} style={{ padding: '8px 16px', background: activeTab === 'ai' ? 'rgba(139,92,246,0.15)' : 'transparent', color: activeTab === 'ai' ? '#a78bfa' : '#747d8c', border: 'none', borderRadius: '8px', fontSize: '13px', fontWeight: '700', cursor: 'pointer' }}>🤖 Interactive Bot Suite</button>
              </div>

              {activeTab === 'unknown' && (
                <UnknownTransactionsPanel 
                  unknowns={result.unknownTransactions} 
                  fixedRows={fixedRowsList} 
                  onUpdateRow={handleUpdateRow} 
                  onReEditRow={handleReEditRow}
                />
              )}
              
              {activeTab === 'ai' && (
                <div>
                  <AIAnalysisPanel result={result} />
                  <AIChatAssistant result={result} />
                </div>
              )}

              {activeTab === 'transactions' && (
                <div style={{ background: 'rgba(255,255,255,0.01)', border: '1px solid rgba(255,255,255,0.04)', borderRadius: '18px', padding: '20px' }}>
                  <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
                    {['ALL', 'HIGH', 'MEDIUM', 'FIXED', 'CLEARED'].map(f => (
                      <button key={f} onClick={() => setActiveFilter(f)} style={{ fontSize: '11px', fontWeight: '700', padding: '6px 14px', borderRadius: '20px', cursor: 'pointer', border: `1px solid ${activeFilter === f ? SEV[f]?.dot || '#8b5cf6' : 'rgba(255,255,255,0.06)'}`, background: activeFilter === f ? `${SEV[f]?.dot || '#8b5cf6'}15` : 'transparent', color: activeFilter === f ? SEV[f]?.dot || '#fff' : '#a4b0be' }}>{f}</button>
                    ))}
                  </div>

                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                    {(activeFilter === 'ALL' ? result.allRows : result.allRows.filter(r => r.severity === activeFilter)).map(tx => (
                      <div key={tx.id} onClick={() => setExpandedId(expandedId === tx.id ? null : tx.id)} style={{ background: 'rgba(255,255,255,0.02)', borderRadius: '12px', borderLeft: `4px solid ${SEV[tx.severity]?.dot || '#747d8c'}`, padding: '12px 16px', display: 'flex', alignItems: 'center', gap: '14px', cursor: 'pointer' }}>
                        <Avatar name={tx.supplier} severity={tx.severity} />
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: '13px', fontWeight: '700', color: '#fff' }}>{tx.supplier} <span style={{ color: '#57606f', fontWeight: '400' }}>({tx.invoice})</span></div>
                          <div style={{ fontSize: '11px', color: '#57606f', marginTop: '2px' }}>{tx.date}</div>
                          {expandedId === tx.id && tx.reasons.length > 0 && (
                            <div style={{ marginTop: '8px', fontSize: '11px', color: '#ff6b81', background: 'rgba(255,71,87,0.05)', padding: '6px 10px', borderRadius: '6px' }}>
                              🚨 Triggered: {tx.reasons.join(', ')}
                            </div>
                          )}
                        </div>
                        <div style={{ textAlign: 'right' }}>
                          <div style={{ fontSize: '14px', fontWeight: '800', color: tx.severity === 'CLEARED' || tx.severity === 'FIXED' ? '#2ed573' : '#ff4757' }}>LKR {tx.amount.toLocaleString()}</div>
                          <div style={{ fontSize: '10px', color: '#57606f' }}>{SEV[tx.severity]?.label}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <div style={{ marginTop: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <button onClick={downloadCorrectedFile} style={{ background: '#00d2d3', color: '#fff', border: 'none', borderRadius: '8px', padding: '10px 20px', fontSize: '12px', fontWeight: '700', cursor: 'pointer' }}>
                  📥 Download Audited File
                </button>
                <button onClick={() => { setResult(null); setFileName(''); }} style={{ padding: '10px 20px', background: 'rgba(255,255,255,0.03)', border: '1px solid rgba(255,255,255,0.08)', borderRadius: '8px', color: '#a4b0be', fontSize: '12px', cursor: 'pointer' }}>🔄 Purge Ledger Matrix Data</button>
              </div>
            </div>
          )}
        </main>
      </div>
    </>
  );
}

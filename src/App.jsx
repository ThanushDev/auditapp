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

    // Unknown transaction detection — enhanced
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

    // Standard fraud rules
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

  // Build allRows — every transaction with severity (CLEARED for clean ones)
  const flaggedIds = new Set(flagged.map(f => f.id));
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

// ─── Severity config ──────────────────────────────────────────────────────────
const SEV = {
  HIGH:    { dot: '#ef4444', badge: { bg: 'rgba(239,68,68,0.18)',   color: '#fca5a5', border: 'rgba(239,68,68,0.35)'   }, leftBorder: '#ef4444', label: 'Flagged',    iconBg: 'rgba(239,68,68,0.15)',   icon: '🔴' },
  MEDIUM:  { dot: '#f59e0b', badge: { bg: 'rgba(245,158,11,0.18)',  color: '#fcd34d', border: 'rgba(245,158,11,0.35)'  }, leftBorder: '#f59e0b', label: 'Suspicious', iconBg: 'rgba(245,158,11,0.15)',  icon: '🟡' },
  LOW:     { dot: '#22c55e', badge: { bg: 'rgba(34,197,94,0.18)',   color: '#86efac', border: 'rgba(34,197,94,0.35)'   }, leftBorder: '#22c55e', label: 'Normal',     iconBg: 'rgba(34,197,94,0.15)',   icon: '🟢' },
  CLEARED: { dot: '#6b7280', badge: { bg: 'rgba(107,114,128,0.15)', color: '#d1d5db', border: 'rgba(107,114,128,0.3)'  }, leftBorder: '#374151', label: 'Cleared',    iconBg: 'rgba(107,114,128,0.1)',  icon: '✅' },
};

function Avatar({ name, severity }) {
  const initials = name.split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  const bg = SEV[severity]?.iconBg || 'rgba(139,92,246,0.15)';
  const color = SEV[severity]?.dot || '#8b5cf6';
  return (
    <div style={{
      width: '38px', height: '38px', borderRadius: '10px', flexShrink: 0,
      background: bg, border: `1px solid ${color}33`,
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '13px', fontWeight: '700', color,
    }}>{initials || '??'}</div>
  );
}

function PulseDot({ color, animate }) {
  return <span style={{
    display: 'inline-block', width: '7px', height: '7px', borderRadius: '50%',
    background: color, flexShrink: 0,
    animation: animate ? 'aqPulse 1.5s infinite' : 'none',
  }} />;
}

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

// ─── AI Analysis Component ────────────────────────────────────────────────────
function AIAnalysisPanel({ result, fileName }) {
  const [aiText, setAiText]     = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiDone, setAiDone]     = useState(false);
  const [aiError, setAiError]   = useState('');

  const runAI = async () => {
    setAiLoading(true); setAiText(''); setAiDone(false); setAiError('');

    const summary = {
      totalTransactions: result.total,
      flaggedHigh: result.flagged.filter(f => f.severity === 'HIGH').length,
      flaggedMedium: result.flagged.filter(f => f.severity === 'MEDIUM').length,
      unknownCount: result.unknownTransactions.length,
      riskLevel: result.riskLevel,
      auditId: result.auditId,
      topReasons: [...new Set(result.flagged.flatMap(f => f.reasons))].slice(0, 6),
      unknownSample: result.unknownTransactions.slice(0, 5).map(u => ({
        supplier: u.supplier, amount: u.amount, invoice: u.invoice,
        reasons: u.unknownReasons,
      })),
      flaggedSample: result.flagged.filter(f => f.severity === 'HIGH').slice(0, 5).map(f => ({
        supplier: f.supplier, amount: f.amount, invoice: f.invoice,
        reasons: f.reasons,
      })),
    };

    const prompt = `You are a senior forensic auditor preparing an internal audit memo. Analyze the following transaction scan results and produce a structured audit analysis.

SCAN RESULTS:
${JSON.stringify(summary, null, 2)}

Write a professional audit analysis covering:
1. EXECUTIVE SUMMARY — overall risk assessment in 2-3 sentences
2. KEY FINDINGS — bullet points of the most critical anomalies found
3. UNKNOWN TRANSACTION ANALYSIS — specific concerns about unidentified/masked transactions, why they're audit red flags, and what further steps are needed
4. RISK INDICATORS — explain each flagged reason and what fraud scheme it may indicate
5. AUDIT RECOMMENDATIONS — 4-6 concrete next steps the audit team should take
6. COMPLIANCE NOTES — reference to relevant audit standards (ISA 240, COSO, IIA standards) that apply

Format clearly with section headers. Be concise but thorough. Use professional audit language.`;

    try {
      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 1000,
          stream: true,
          messages: [{ role: 'user', content: prompt }],
        }),
      });

      const reader = res.body.getReader();
      const dec = new TextDecoder();
      let buffer = '';

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += dec.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop();
        for (const line of lines) {
          if (!line.startsWith('data: ')) continue;
          const raw = line.slice(6).trim();
          if (raw === '[DONE]') continue;
          try {
            const ev = JSON.parse(raw);
            if (ev.type === 'content_block_delta' && ev.delta?.text) {
              setAiText(t => t + ev.delta.text);
            }
          } catch {}
        }
      }
      setAiDone(true);
    } catch (err) {
      setAiError('AI analysis failed: ' + err.message);
    }
    setAiLoading(false);
  };

  return (
    <div style={{
      background: 'rgba(139,92,246,0.05)', border: '1px solid rgba(139,92,246,0.25)',
      borderRadius: '14px', padding: '18px 20px', marginBottom: '20px',
    }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '14px' }}>
        <div style={{ fontSize: '13px', fontWeight: '700', color: '#e5e7eb', display: 'flex', alignItems: 'center', gap: '6px' }}>
          🤖 AI Audit Analysis
          {aiDone && <span style={{ fontSize: '10px', padding: '2px 8px', borderRadius: '20px', background: 'rgba(34,197,94,0.15)', color: '#86efac', border: '1px solid rgba(34,197,94,0.3)' }}>Complete</span>}
        </div>
        {!aiLoading && !aiDone && (
          <button
            onClick={runAI}
            style={{
              background: 'linear-gradient(135deg,#7c3aed,#9333ea)', color: '#fff',
              border: 'none', borderRadius: '8px', padding: '7px 16px',
              fontSize: '12px', fontWeight: '700', cursor: 'pointer',
              boxShadow: '0 2px 12px rgba(124,58,237,0.4)',
            }}
          >⚡ Run AI Analysis</button>
        )}
        {aiDone && (
          <button
            onClick={() => { setAiText(''); setAiDone(false); }}
            style={{ background: 'none', border: '1px solid rgba(139,92,246,0.3)', color: '#8b5cf6', borderRadius: '8px', padding: '5px 12px', fontSize: '12px', cursor: 'pointer' }}
          >↺ Re-run</button>
        )}
      </div>

      {!aiText && !aiLoading && !aiError && (
        <div style={{ fontSize: '12px', color: '#6b7280', textAlign: 'center', padding: '20px 0' }}>
          Click "Run AI Analysis" to generate a detailed forensic audit report using Claude AI.
        </div>
      )}

      {aiLoading && !aiText && (
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '13px', color: '#a78bfa' }}>
          <span style={{ animation: 'aqSpin 0.9s linear infinite', display: 'inline-block' }}>🔄</span>
          Generating forensic audit analysis...
        </div>
      )}

      {aiError && (
        <div style={{ fontSize: '12px', color: '#fca5a5', padding: '10px', background: 'rgba(239,68,68,0.08)', borderRadius: '8px' }}>
          {aiError}
        </div>
      )}

      {aiText && (
        <div style={{
          fontSize: '12.5px', color: '#d1d5db', lineHeight: '1.75',
          whiteSpace: 'pre-wrap', maxHeight: '460px', overflowY: 'auto',
          padding: '14px 16px', background: 'rgba(0,0,0,0.25)',
          borderRadius: '10px', border: '1px solid rgba(139,92,246,0.15)',
          fontFamily: "'JetBrains Mono', 'Fira Code', monospace",
        }}>
          {aiText}
          {aiLoading && <span style={{ animation: 'aqPulse 1s infinite', display: 'inline-block', marginLeft: '2px' }}>▋</span>}
        </div>
      )}
    </div>
  );
}

// ─── Unknown Transactions Panel ───────────────────────────────────────────────
function UnknownTransactionsPanel({ unknowns }) {
  const [expanded, setExpanded] = useState(null);

  if (unknowns.length === 0) {
    return (
      <div style={{
        background: 'rgba(34,197,94,0.05)', border: '1px solid rgba(34,197,94,0.2)',
        borderRadius: '14px', padding: '18px 20px', marginBottom: '20px',
      }}>
        <div style={{ fontSize: '13px', fontWeight: '700', color: '#86efac', display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' }}>
          ✅ Unknown Transactions
        </div>
        <div style={{ fontSize: '12px', color: '#6b7280' }}>No unknown or unidentified transactions detected in this dataset.</div>
      </div>
    );
  }

  const totalAmount = unknowns.reduce((s, u) => s + u.amount, 0);

  return (
    <div style={{
      background: 'rgba(239,68,68,0.04)', border: '1px solid rgba(239,68,68,0.3)',
      borderRadius: '14px', padding: '18px 20px', marginBottom: '20px',
    }}>
      {/* Header */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '14px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span style={{ fontSize: '13px', fontWeight: '700', color: '#fca5a5', display: 'flex', alignItems: 'center', gap: '6px' }}>
            <PulseDot color="#ef4444" animate={true} />
            Unknown Transactions
          </span>
          <span style={{
            fontSize: '11px', fontWeight: '700', padding: '2px 10px', borderRadius: '20px',
            background: 'rgba(239,68,68,0.2)', color: '#fca5a5', border: '1px solid rgba(239,68,68,0.4)',
          }}>{unknowns.length} found</span>
        </div>
        <div style={{ fontSize: '12px', color: '#9ca3af' }}>
          Total exposure: <span style={{ color: '#fca5a5', fontWeight: '700' }}>LKR {totalAmount.toLocaleString()}</span>
        </div>
      </div>

      {/* Audit note */}
      <div style={{
        fontSize: '11px', color: '#fbbf24', background: 'rgba(245,158,11,0.08)',
        border: '1px solid rgba(245,158,11,0.2)', borderRadius: '8px',
        padding: '8px 12px', marginBottom: '12px', display: 'flex', gap: '6px',
      }}>
        ⚠️ These transactions require immediate investigation per ISA 240 (Auditor's Responsibilities Relating to Fraud). Each has been assigned a unique audit reference number.
      </div>

      {/* Transaction list */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
        {unknowns.map((u, i) => (
          <div
            key={u.id}
            onClick={() => setExpanded(expanded === u.id ? null : u.id)}
            style={{
              background: 'rgba(239,68,68,0.06)', borderRadius: '10px',
              border: '1px solid rgba(239,68,68,0.2)', borderLeftWidth: '3px',
              borderLeftColor: '#ef4444', padding: '11px 14px',
              cursor: 'pointer', transition: 'all 0.18s',
            }}
            onMouseEnter={e => e.currentTarget.style.borderColor = 'rgba(239,68,68,0.45)'}
            onMouseLeave={e => e.currentTarget.style.borderColor = 'rgba(239,68,68,0.2)'}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <div style={{
                width: '36px', height: '36px', borderRadius: '9px', flexShrink: 0,
                background: 'rgba(239,68,68,0.12)', border: '1px solid rgba(239,68,68,0.25)',
                display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '16px',
              }}>❓</div>

              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: '13px', fontWeight: '700', color: '#fca5a5', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.supplier}</span>
                  <span style={{ fontSize: '9px', fontFamily: 'monospace', background: 'rgba(239,68,68,0.15)', padding: '1px 6px', borderRadius: '4px', color: '#fca5a5', flexShrink: 0 }}>{u.auditRef}</span>
                </div>
                <div style={{ fontSize: '11px', color: '#9ca3af', marginTop: '2px', display: 'flex', gap: '10px' }}>
                  <span>📄 {u.invoice}</span>
                  <span>📅 {u.date}</span>
                  {u.time && <span>🕐 {u.time}</span>}
                </div>
                {expanded === u.id && (
                  <div style={{ marginTop: '10px' }}>
                    <div style={{ fontSize: '11px', color: '#d1d5db', marginBottom: '6px' }}>
                      <span style={{ color: '#9ca3af' }}>Description: </span>{u.description}
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px', marginBottom: '8px' }}>
                      {u.unknownReasons.map((r, ri) => (
                        <span key={ri} style={{
                          fontSize: '10px', fontWeight: '600', padding: '2px 8px', borderRadius: '6px',
                          background: 'rgba(239,68,68,0.15)', color: '#fca5a5', border: '1px solid rgba(239,68,68,0.3)',
                        }}>🚩 {r}</span>
                      ))}
                    </div>
                    <div style={{ fontSize: '10px', color: '#6b7280', fontFamily: 'monospace' }}>
                      Flagged at: {new Date(u.flaggedAt).toLocaleString()}
                    </div>
                  </div>
                )}
              </div>

              <div style={{ textAlign: 'right', flexShrink: 0 }}>
                <div style={{ fontSize: '14px', fontWeight: '800', color: '#fca5a5' }}>
                  -LKR {u.amount.toLocaleString()}
                </div>
                <div style={{ fontSize: '10px', color: '#ef4444', marginTop: '3px' }}>
                  {expanded === u.id ? '▲ collapse' : '▼ details'}
                </div>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── Audit Trail Header ───────────────────────────────────────────────────────
function AuditTrailBadge({ result }) {
  return (
    <div style={{
      background: 'rgba(139,92,246,0.06)', border: '1px solid rgba(139,92,246,0.2)',
      borderRadius: '10px', padding: '10px 16px', marginBottom: '16px',
      display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px',
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
        <span style={{ fontSize: '11px', color: '#8b5cf6', fontWeight: '700', display: 'flex', alignItems: 'center', gap: '5px' }}>
          📋 AUDIT ID
          <span style={{ fontFamily: 'monospace', fontSize: '11px', color: '#c4b5fd', background: 'rgba(139,92,246,0.15)', padding: '2px 8px', borderRadius: '5px' }}>{result.auditId}</span>
        </span>
        <span style={{ fontSize: '11px', color: '#6b7280' }}>
          Scanned: <span style={{ color: '#9ca3af' }}>{new Date(result.scanTime).toLocaleString()}</span>
        </span>
      </div>
      <span style={{
        fontSize: '10px', fontWeight: '700', padding: '3px 10px', borderRadius: '20px',
        background: result.riskLevel === 'HIGH' ? 'rgba(239,68,68,0.18)' : result.riskLevel === 'MEDIUM' ? 'rgba(245,158,11,0.18)' : 'rgba(34,197,94,0.15)',
        color: result.riskLevel === 'HIGH' ? '#fca5a5' : result.riskLevel === 'MEDIUM' ? '#fcd34d' : '#86efac',
        border: `1px solid ${result.riskLevel === 'HIGH' ? 'rgba(239,68,68,0.35)' : result.riskLevel === 'MEDIUM' ? 'rgba(245,158,11,0.35)' : 'rgba(34,197,94,0.3)'}`,
      }}>
        Overall Risk: {result.riskLevel}
      </span>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [loading,      setLoading]      = useState(false);
  const [result,       setResult]       = useState(null);
  const [fileName,     setFileName]     = useState('');
  const [dragOver,     setDragOver]     = useState(false);
  const [activeFilter, setActiveFilter] = useState('ALL');
  const [activeTab,    setActiveTab]    = useState('transactions'); // 'transactions' | 'unknown' | 'ai'
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
      const res = runFraudDetection(rows);
      setResult(res);
      // Auto-switch to unknown tab if unknowns found
      if (res.unknownTransactions.length > 0) setActiveTab('unknown');
    } catch (err) { alert('Error: ' + err.message); }
    setLoading(false);
  }, []);

  const handleChange = e => processFile(e.target.files[0]);
  const handleDrop   = e => { e.preventDefault(); setDragOver(false); processFile(e.dataTransfer.files[0]); };

  const flaggedCount = result ? result.flagged.filter(f => f.severity === 'HIGH').length   : 0;
  const suspCount    = result ? result.flagged.filter(f => f.severity === 'MEDIUM').length : 0;
  const clearCount   = result ? result.total - result.flagged.length : 0;
  const totalCount   = result ? result.total : 0;
  const unknownCount = result ? result.unknownTransactions.length : 0;

  const rb = { dup: 0, missing: 0, mismatch: 0, timing: 0 };
  if (result) result.flagged.forEach(f => f.reasons.forEach(r => {
    if (r.includes('Duplicate'))  rb.dup++;
    if (r.includes('Missing'))    rb.missing++;
    if (r.includes('Round'))      rb.mismatch++;
    if (r.includes('High Volume') || r.includes('Cumulative')) rb.timing++;
  }));
  const total_rb = Math.max(result ? result.flagged.length : 1, 1);
  const pct = v => Math.round((v / total_rb) * 100);

  // ALL transactions (flagged + cleared) from allRows
  const filtered = result
    ? (activeFilter === 'ALL'        ? result.allRows
      : activeFilter === 'FLAGGED'    ? result.allRows.filter(f => f.severity === 'HIGH')
      : activeFilter === 'SUSPICIOUS' ? result.allRows.filter(f => f.severity === 'MEDIUM')
      : activeFilter === 'UNKNOWN'    ? result.allRows.filter(f => f.severity === 'HIGH' && f.reasons.some(r => r.includes('Unknown') || r.includes('Masked')))
      : result.allRows.filter(f => f.severity === 'LOW' || f.severity === 'CLEARED'))
    : [];

  const demoTx = [
    { id:0, supplier:'Apex Supplies Ltd',      invoice:'INV-2024-0871', amount:84500,  date:'2026-05-23', time:'10:42 AM', severity:'HIGH',    reasons:['Duplicate Invoice Ref'],        auditRef:'AUD-DEMO-0001' },
    { id:1, supplier:'Account: 0029-***-4412', invoice:'TXN-9934',      amount:127000, date:'2026-05-23', time:'02:17 AM', severity:'MEDIUM',  reasons:['Masked Account Identifier'],    auditRef:'AUD-DEMO-0002' },
    { id:2, supplier:'SriTech Pvt Ltd',        invoice:'PO-55221',      amount:45200,  date:'2026-05-23', time:'09:15 AM', severity:'CLEARED', reasons:[],                               auditRef:'AUD-DEMO-0003' },
    { id:3, supplier:'Finance Dept',           invoice:'REF-0045',      amount:210000, date:'2026-05-23', time:'11:58 AM', severity:'HIGH',    reasons:['Missing Invoice Reference'],    auditRef:'AUD-DEMO-0004' },
    { id:4, supplier:'Unknown',                invoice:'N/A',           amount:49900,  date:'2026-05-23', time:'01:30 PM', severity:'HIGH',    reasons:['Unknown Supplier','Round-number Amount'], auditRef:'AUD-DEMO-0005' },
    { id:5, supplier:'MAS Holdings',           invoice:'INV-0092',      amount:32000,  date:'2026-05-23', time:'03:10 PM', severity:'CLEARED', reasons:[],                               auditRef:'AUD-DEMO-0006' },
    { id:6, supplier:'Dialog Axiata',          invoice:'INV-8821',      amount:18500,  date:'2026-05-23', time:'04:00 PM', severity:'CLEARED', reasons:[],                               auditRef:'AUD-DEMO-0007' },
  ];
  const displayTx = result ? filtered : demoTx;
  const showingDemo = !result;

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

  const Tab = ({ id, label, icon, count, highlight }) => {
    const active = activeTab === id;
    return (
      <button
        onClick={() => setActiveTab(id)}
        style={{
          fontSize: '12px', fontWeight: '700', padding: '7px 14px', borderRadius: '10px',
          cursor: 'pointer', border: `1px solid ${active ? (highlight ? 'rgba(239,68,68,0.5)' : 'rgba(139,92,246,0.5)') : 'rgba(255,255,255,0.1)'}`,
          background: active ? (highlight ? 'rgba(239,68,68,0.12)' : 'rgba(139,92,246,0.18)') : 'rgba(255,255,255,0.04)',
          color: active ? (highlight ? '#fca5a5' : '#c4b5fd') : '#9ca3af',
          display: 'flex', alignItems: 'center', gap: '6px', transition: 'all 0.15s',
          position: 'relative',
        }}
      >
        {icon} {label}
        {count > 0 && (
          <span style={{
            fontSize: '10px', fontWeight: '800', padding: '1px 6px', borderRadius: '20px',
            background: highlight ? 'rgba(239,68,68,0.25)' : 'rgba(139,92,246,0.25)',
            color: highlight ? '#fca5a5' : '#c4b5fd',
          }}>{count}</span>
        )}
      </button>
    );
  };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #0b0f1e !important; }
        @keyframes aqPulse  { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.4;transform:scale(1.4)} }
        @keyframes aqSpin   { to { transform: rotate(360deg); } }
        @keyframes aqFadeUp { from{opacity:0;transform:translateY(14px)} to{opacity:1;transform:translateY(0)} }
        @keyframes aqSlide  { from{opacity:0;transform:translateX(-8px)} to{opacity:1;transform:translateX(0)} }
        .aq-stat:hover  { transform: translateY(-2px) !important; box-shadow: 0 8px 24px rgba(139,92,246,0.18) !important; }
        .aq-tx:hover    { border-color: rgba(139,92,246,0.35) !important; transform: translateX(3px); }
        .aq-results     { animation: aqFadeUp 0.4s ease both; }
        ::-webkit-scrollbar { width: 4px; } ::-webkit-scrollbar-track { background: transparent; } ::-webkit-scrollbar-thumb { background: rgba(139,92,246,0.3); border-radius: 2px; }
      `}</style>

      <div style={{ minHeight: '100vh', background: '#0b0f1e', fontFamily: "'Inter', ui-sans-serif, system-ui, sans-serif", color: '#f3f4f6' }}>

        {/* ── Top Bar ───────────────────────────────────────────────────────── */}
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
            <IBtn title="Notifications" hasNotif={result && (result.flagged.length > 0 || unknownCount > 0)}>🔔</IBtn>
            <IBtn title="Export">📤</IBtn>
          </div>
        </header>

        <main style={{ maxWidth: '960px', margin: '0 auto', padding: '24px 20px' }}>

          {/* ── Upload zone ─────────────────────────────────────────────────── */}
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
              <div style={{ fontSize: '13px', color: '#6b7280', marginBottom: '16px' }}>CSV, XLSX, PDF, or DOCX — drag & drop or click to browse</div>
              <button onClick={e => { e.stopPropagation(); fileInputRef.current.click(); }} style={{
                background: 'linear-gradient(135deg,#7c3aed,#9333ea)', color: '#fff',
                border: 'none', borderRadius: '10px', padding: '9px 22px',
                fontSize: '13px', fontWeight: '700', cursor: 'pointer',
              }}>⚡ Choose File</button>
              {fileName && <div style={{ marginTop:'12px', fontSize:'12px', color:'#8b5cf6', fontWeight:'600' }}>📄 {fileName}</div>}
              <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" style={{ display:'none' }} onChange={handleChange} />
            </div>
          )}

          {/* ── Loading ──────────────────────────────────────────────────────── */}
          {loading && (
            <div style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', background:'rgba(139,92,246,0.08)', border:'1px solid rgba(139,92,246,0.2)', borderRadius:'12px', padding:'16px', marginBottom:'24px', fontSize:'13px', fontWeight:'600', color:'#a78bfa' }}>
              <span style={{ display:'inline-block', animation:'aqSpin 0.9s linear infinite', fontSize:'18px' }}>🔄</span>
              Analysing transactions...
            </div>
          )}

          {/* ── Stat cards ──────────────────────────────────────────────────── */}
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit,minmax(140px,1fr))', gap:'12px', marginBottom:'20px' }}>
            <div className="aq-stat" style={{ background:'rgba(239,68,68,0.08)', border:'1px solid rgba(239,68,68,0.25)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#ef4444" animate={true} /> Flagged
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? 12 : flaggedCount}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(239,68,68,0.18)', color:'#fca5a5', border:'1px solid rgba(239,68,68,0.3)' }}>↑ {showingDemo ? 4 : flaggedCount} new</span>
            </div>

            <div className="aq-stat" style={{ background:'rgba(245,158,11,0.08)', border:'1px solid rgba(245,158,11,0.25)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.07s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#f59e0b" animate={true} /> Suspicious
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? 38 : suspCount}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(245,158,11,0.18)', color:'#fcd34d', border:'1px solid rgba(245,158,11,0.3)' }}>Review needed</span>
            </div>

            {/* Unknown stat — NEW */}
            <div className="aq-stat" onClick={() => result && setActiveTab('unknown')} style={{ background:'rgba(239,68,68,0.05)', border:`1px solid ${result && unknownCount > 0 ? 'rgba(239,68,68,0.4)' : 'rgba(239,68,68,0.15)'}`, borderRadius:'14px', padding:'16px 18px', cursor: result ? 'pointer' : 'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.1s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#ef4444" animate={result && unknownCount > 0} /> Unknown
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color: result && unknownCount > 0 ? '#fca5a5' : '#fff', lineHeight:1 }}>{showingDemo ? 7 : unknownCount}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(239,68,68,0.15)', color:'#fca5a5', border:'1px solid rgba(239,68,68,0.3)' }}>Needs investigation</span>
            </div>

            <div className="aq-stat" style={{ background:'rgba(34,197,94,0.07)', border:'1px solid rgba(34,197,94,0.2)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.14s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                <PulseDot color="#22c55e" animate={false} /> Cleared
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? '1,204' : clearCount.toLocaleString()}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(34,197,94,0.15)', color:'#86efac', border:'1px solid rgba(34,197,94,0.3)' }}>
                {showingDemo ? '97.2% pass' : (totalCount > 0 ? `${Math.round((clearCount/totalCount)*100)}% pass` : '—')}
              </span>
            </div>

            <div className="aq-stat" style={{ background:'rgba(139,92,246,0.08)', border:'1px solid rgba(139,92,246,0.25)', borderRadius:'14px', padding:'16px 18px', cursor:'default', transition:'all 0.2s', animation:'aqFadeUp 0.3s ease 0.21s both' }}>
              <div style={{ fontSize:'11px', fontWeight:'700', color:'#9ca3af', textTransform:'uppercase', letterSpacing:'0.7px', marginBottom:'8px', display:'flex', alignItems:'center', gap:'5px' }}>
                📊 Total scanned
              </div>
              <div style={{ fontSize:'32px', fontWeight:'800', color:'#fff', lineHeight:1 }}>{showingDemo ? '1,254' : totalCount.toLocaleString()}</div>
              <span style={{ display:'inline-block', marginTop:'6px', fontSize:'10px', fontWeight:'700', padding:'2px 9px', borderRadius:'20px', background:'rgba(139,92,246,0.18)', color:'#c4b5fd', border:'1px solid rgba(139,92,246,0.3)' }}>Today</span>
            </div>
          </div>

          {/* ── Audit Trail Badge ────────────────────────────────────────────── */}
          {result && <AuditTrailBadge result={result} />}

          {/* ── Risk breakdown ───────────────────────────────────────────────── */}
          <div style={{ background:'rgba(255,255,255,0.03)', border:'1px solid rgba(139,92,246,0.15)', borderRadius:'14px', padding:'18px 20px', marginBottom:'20px' }}>
            <div style={{ fontSize:'13px', fontWeight:'700', color:'#e5e7eb', marginBottom:'14px', display:'flex', alignItems:'center', gap:'6px' }}>⚡ Risk breakdown</div>
            <RiskBar label="Duplicate entries"  pct={showingDemo ? 72 : pct(rb.dup)}      color="#ef4444" />
            <RiskBar label="Missing refs"        pct={showingDemo ? 55 : pct(rb.missing)}  color="#f59e0b" />
            <RiskBar label="Amount mismatch"     pct={showingDemo ? 38 : pct(rb.mismatch)} color="#8b5cf6" />
            <RiskBar label="Timing anomalies"    pct={showingDemo ? 21 : pct(rb.timing)}   color="#6366f1" />
          </div>

          {/* ── Tabs ─────────────────────────────────────────────────────────── */}
          {result && (
            <div style={{ display:'flex', gap:'8px', marginBottom:'16px', flexWrap:'wrap' }}>
              <Tab id="transactions" label="All Transactions" icon="≡" count={result.allRows.length} />
              <Tab id="unknown"      label="Unknown Transactions" icon="❓" count={unknownCount} highlight={unknownCount > 0} />
              <Tab id="ai"           label="AI Audit Report" icon="🤖" count={0} />
            </div>
          )}

          {/* ── Tab content ──────────────────────────────────────────────────── */}
          {result && activeTab === 'unknown' && (
            <UnknownTransactionsPanel unknowns={result.unknownTransactions} />
          )}

          {result && activeTab === 'ai' && (
            <AIAnalysisPanel result={result} fileName={fileName} />
          )}

          {/* Transactions tab (also shown when no result = demo) */}
          {(!result || activeTab === 'transactions') && (
            <div style={{ background:'rgba(255,255,255,0.02)', border:'1px solid rgba(139,92,246,0.15)', borderRadius:'14px', padding:'18px 20px' }}>
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:'14px' }}>
                <div style={{ fontSize:'13px', fontWeight:'700', color:'#e5e7eb', display:'flex', alignItems:'center', gap:'6px' }}>≡ Recent transactions</div>
                {result && (
                  <button onClick={() => fileInputRef.current.click()} style={{ background:'none', border:'none', cursor:'pointer', fontSize:'12px', color:'#8b5cf6', fontWeight:'600', display:'flex', alignItems:'center', gap:'4px' }}>
                    View all →
                  </button>
                )}
              </div>

              <div style={{ display:'flex', gap:'8px', marginBottom:'14px', flexWrap:'wrap' }}>
                <FChip id="ALL"        label={`All (${result ? result.allRows.length : demoTx.length})`} icon="⊞" />
                <FChip id="FLAGGED"    label="Flagged"    icon="🚩" />
                <FChip id="SUSPICIOUS" label="Suspicious" icon="⚠️" />
                <FChip id="LOW"        label="Cleared"    icon="✅" />
              </div>

              <div style={{ display:'flex', flexDirection:'column', gap:'8px' }}>
                {displayTx.length === 0 ? (
                  <div style={{ textAlign:'center', padding:'32px', color:'#6b7280', fontSize:'13px' }}>No transactions match this filter.</div>
                ) : displayTx.map((tx, i) => (
                  <div
                    key={tx.id}
                    className="aq-tx"
                    onClick={() => setExpandedId(expandedId === tx.id ? null : tx.id)}
                    style={{
                      background: 'rgba(255,255,255,0.03)', borderRadius: '12px',
                      borderLeft: `3px solid ${SEV[tx.severity].dot}`,
                      border: '1px solid rgba(255,255,255,0.07)', borderLeftWidth: '3px', borderLeftColor: SEV[tx.severity].dot,
                      padding: '12px 14px', display: 'flex', alignItems: 'center', gap: '12px',
                      cursor: 'pointer', transition: 'all 0.18s ease',
                      animation: `aqSlide 0.25s ease ${i * 0.05}s both`,
                      boxShadow: expandedId === tx.id ? '0 4px 20px rgba(139,92,246,0.12)' : 'none',
                    }}
                  >
                    <Avatar name={tx.supplier} severity={tx.severity} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '13px', fontWeight: '700', color: '#fff', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {tx.invoice} · {tx.reasons[0] || 'Cleared'}
                        </span>
                        {tx.auditRef && (
                          <span style={{ fontSize: '9px', fontFamily: 'monospace', background: 'rgba(139,92,246,0.12)', padding: '1px 6px', borderRadius: '4px', color: '#8b5cf6', flexShrink: 0 }}>{tx.auditRef}</span>
                        )}
                      </div>
                      <div style={{ fontSize: '11px', color: '#6b7280', marginTop: '3px' }}>
                        {tx.severity === 'MEDIUM' ? `Account: ${tx.supplier}` : `Vendor: ${tx.supplier}`}
                        {tx.time && <span style={{ marginLeft: '10px' }}>{tx.time}</span>}
                      </div>
                      {expandedId === tx.id && (
                        <div style={{ marginTop: '8px' }}>
                          {tx.reasons.length > 1 && (
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px', marginBottom: '6px' }}>
                              {tx.reasons.slice(1).map((r, ri) => (
                                <span key={ri} style={{ fontSize: '10px', fontWeight: '600', padding: '2px 8px', borderRadius: '6px', background: 'rgba(139,92,246,0.15)', color: '#c4b5fd', border: '1px solid rgba(139,92,246,0.25)' }}>⚡ {r}</span>
                              ))}
                            </div>
                          )}
                          {tx.flaggedAt && (
                            <div style={{ fontSize: '10px', color: '#6b7280', fontFamily: 'monospace' }}>
                              Flagged: {new Date(tx.flaggedAt).toLocaleString()}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                    <div style={{ textAlign: 'right', flexShrink: 0 }}>
                      <div style={{ fontSize: '14px', fontWeight: '800', color: (tx.severity === 'LOW' || tx.severity === 'CLEARED') ? '#86efac' : '#fca5a5' }}>
                        -LKR {tx.amount.toLocaleString()}
                      </div>
                      <span style={{
                        display: 'inline-block', marginTop: '4px', fontSize: '10px', fontWeight: '700',
                        padding: '2px 9px', borderRadius: '20px',
                        background: SEV[tx.severity].badge.bg, color: SEV[tx.severity].badge.color, border: `1px solid ${SEV[tx.severity].badge.border}`,
                      }}>{SEV[tx.severity].label}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ── Bottom buttons ───────────────────────────────────────────────── */}
          <div style={{ display:'flex', gap:'10px', marginTop:'16px', flexWrap:'wrap' }}>
            <button style={{ flex:1, minWidth:'100px', padding:'11px', borderRadius:'10px', border:'1px solid rgba(139,92,246,0.3)', background:'rgba(255,255,255,0.04)', color:'#a78bfa', fontSize:'13px', fontWeight:'700', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'7px', transition:'all 0.15s' }}
              onMouseEnter={e => e.currentTarget.style.background='rgba(139,92,246,0.15)'}
              onMouseLeave={e => e.currentTarget.style.background='rgba(255,255,255,0.04)'}
            >📤 Export</button>

            <button
              onClick={() => result && setActiveTab('ai')}
              style={{ flex:1, minWidth:'100px', padding:'11px', borderRadius:'10px', border:'1px solid rgba(139,92,246,0.3)', background:'rgba(255,255,255,0.04)', color:'#a78bfa', fontSize:'13px', fontWeight:'700', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'7px', transition:'all 0.15s' }}
              onMouseEnter={e => e.currentTarget.style.background='rgba(139,92,246,0.15)'}
              onMouseLeave={e => e.currentTarget.style.background='rgba(255,255,255,0.04)'}
            >🤖 AI Report</button>

            <button
              onClick={() => fileInputRef.current.click()}
              style={{ flex:1, minWidth:'100px', padding:'11px', borderRadius:'10px', border:'none', background:'linear-gradient(135deg,#7c3aed,#9333ea)', color:'#fff', fontSize:'13px', fontWeight:'700', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'7px', boxShadow:'0 4px 16px rgba(124,58,237,0.35)', transition:'all 0.15s' }}
              onMouseEnter={e => { e.currentTarget.style.transform='translateY(-1px)'; e.currentTarget.style.boxShadow='0 6px 20px rgba(124,58,237,0.45)'; }}
              onMouseLeave={e => { e.currentTarget.style.transform='translateY(0)'; e.currentTarget.style.boxShadow='0 4px 16px rgba(124,58,237,0.35)'; }}
            >▶ Run audit scan ↗</button>
          </div>

          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" style={{ display:'none' }} onChange={handleChange} />

        </main>
      </div>
    </>
  );
}

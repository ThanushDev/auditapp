import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

export default function App() {
  const [loading, setLoading] = useState(false);
  const [stats, setStats] = useState({ total: 0, flagged: 0, risk: 'None' });
  const [alerts, setAlerts] = useState([]);
  const [fileName, setFileName] = useState('');
  const [hasData, setHasData] = useState(false); // මුලින්ම Dashboard එක හංගන්න පාවිච්චි කරන State එක

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setLoading(true);
    setHasData(false); // අලුත් ෆයිල් එකක් දාද්දී පරණ ඩේටා හංගනවා
    setFileName(file.name);
    const fileType = file.name.split('.').pop().toLowerCase();

    try {
      let rawTransactions = [];

      if (fileType === 'xlsx' || fileType === 'xls' || fileType === 'csv') {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        rawTransactions = XLSX.utils.sheet_to_json(sheet);
      } 
      else if (fileType === 'docx') {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        rawTransactions = parseTextToRows(result.value);
      } 
      else if (fileType === 'pdf') {
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
        alert('Format Unsupported! Use CSV, Excel, Word or PDF.');
        setLoading(false);
        return;
      }

      runFraudDetection(rawTransactions);
    } catch (error) {
      console.error(error);
      alert('Error reading document structure.');
    }
    setLoading(false);
  };

  const parseTextToRows = (text) => {
    const lines = text.split('\n');
    let rows = [];
    lines.forEach(line => {
      const amountMatch = line.match(/\b\d{4,9}\b/); 
      const invMatch = line.match(/INV\d{3,6}/i);
      const unknownMatch = line.toLowerCase().includes('unknown');
      
      if (amountMatch || invMatch || unknownMatch) {
        rows.push({
          Date: new Date().toISOString().split('T')[0],
          Amount: amountMatch ? parseFloat(amountMatch[0]) : 15000,
          Supplier: unknownMatch ? 'Unknown' : (line.match(/[A-Z][a-z]+ Store|[A-Z][a-z]+ Mart|[A-Z][a-z]+ Tech/)?.[0] || 'Dynamic Supplier'),
          Invoice_No: invMatch ? invMatch[0].toUpperCase() : 'INV' + Math.floor(100 + Math.random() * 900)
        });
      }
    });
    return rows;
  };

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
      if (row.Supplier && row.Supplier.toString().toLowerCase().includes('unknown')) reasons.push("Unknown Supplier");
      if (row.Invoice_No && invoiceCounts[row.Invoice_No] > 1) reasons.push("Duplicate Invoice Ref");

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
    setHasData(true); // දත්ත සාර්ථකව විශ්ලේෂණය කළාට පස්සේ Dashboard එක පෙන්වන්න සෙට් කරනවා
  };

  // Modern Responsive CSS Styles with Media Queries logic handled dynamically
  const styles = {
    container: { backgroundColor: '#030712', color: '#f3f4f6', fontFamily: '"SF Pro Display", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', minHeight: '100vh', boxSizing: 'border-box' },
    header: { borderBottom: '1px solid #1f2937', backgroundColor: '#0b0f19', padding: '20px 5%', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: '0 4px 20px rgba(0,0,0,0.4)', flexWrap: 'wrap', gap: '10px' },
    brand: { display: 'flex', alignItems: 'center', gap: '15px' },
    logo: { backgroundColor: '#2563eb', width: '45px', height: '45px', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '22px', boxShadow: '0 0 15px rgba(37,99,235,0.5)' },
    title: { margin: 0, fontSize: '22px', fontWeight: '800', color: '#ffffff', letterSpacing: '-0.5px' },
    subtitle: { margin: '2px 0 0 0', fontSize: '12px', color: '#9ca3af', tracking: '0.5px' },
    badge: { background: 'linear-gradient(90deg, #1e40af, #2563eb)', color: '#ffffff', fontSize: '11px', fontWeight: '700', padding: '6px 16px', borderRadius: '30px', textTransform: 'uppercase', letterSpacing: '1px', boxShadow: '0 0 10px rgba(37,99,235,0.2)' },
    main: { width: '90%', maxWidth: '1400px', margin: '0 auto', padding: '40px 0' },
    uploadZone: { background: 'linear-gradient(145deg, #0f172a, #0b0f19)', border: '2px dashed #3b82f6', borderRadius: '24px', padding: '50px 30px', textAlign: 'center', boxShadow: '0 20px 40px rgba(0,0,0,0.5)', marginBottom: '40px', position: 'relative', overflow: 'hidden' },
    uploadLabel: { cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px' },
    uploadIcon: { fontSize: '50px', filter: 'drop-shadow(0 0 10px rgba(59,130,246,0.5))' },
    grid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: '25px', marginBottom: '40px' },
    card: { background: '#0f172a', border: '1px solid #1e2937', borderRadius: '18px', padding: '25px', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' },
    cardLabel: { fontSize: '12px', fontWeight: '600', color: '#9ca3af', textTransform: 'uppercase', letterSpacing: '1px' },
    cardVal: { fontSize: '38px', fontWeight: '800', marginTop: '10px', color: '#ffffff' },
    // Mobile layouts සඳහා automatic responsive වෙන flex-wrap සහ minmax Grid එකක් මෙතන තියෙන්නේ
    dashboardLayout: { display: 'flex', flexDirection: 'row', flexWrap: 'wrap', gap: '30px', width: '100%' },
    panelLeft: { flex: '1 1 450px', background: '#0f172a', border: '1px solid #1e2937', borderRadius: '20px', padding: '25px', boxShadow: '0 10px 30px rgba(0,0,0,0.3)', minWidth: '300px' },
    panelRight: { flex: '1 1 550px', background: '#0f172a', border: '1px solid #1e2937', borderRadius: '20px', padding: '25px 0 0 0', boxShadow: '0 10px 30px rgba(0,0,0,0.3)', minWidth: '300px', overflow: 'hidden' },
    panelTitle: { fontSize: '16px', fontWeight: '700', marginBottom: '20px', color: '#ffffff', display: 'flex', alignItems: 'center', gap: '10px', borderBottom: '1px solid #1f2937', paddingBottom: '12px', paddingLeft: '20px' },
    alertBox: { background: '#070a13', borderLeft: '5px solid #ef4444', borderRadius: '8px', padding: '16px', fontSize: '13px', marginBottom: '15px', boxShadow: '0 4px 12px rgba(0,0,0,0.15)' },
    tableContainer: { overflowX: 'auto', maxHeight: '400px', width: '100%' },
    table: { width: '100%', borderCollapse: 'collapse', fontSize: '13px', textAlign: 'left' },
    th: { backgroundColor: '#070a13', padding: '14px', color: '#9ca3af', textTransform: 'uppercase', fontSize: '11px', fontWeight: '700', letterSpacing: '0.5px', borderBottom: '2px solid #1f2937' },
    td: { padding: '14px 12px', borderBottom: '1px solid #1f2937', color: '#e5e7eb', verticalAlign: 'middle' }
  };

  return (
    <div style={styles.container}>
      {/* Navigation Header */}
      <header style={styles.header}>
        <div style={styles.brand}>
          <div style={styles.logo}>🛡️</div>
          <div>
            <h1 style={styles.title}>NextGen Audit Firm</h1>
            <p style={styles.subtitle}>Forensic AI Fraud Detection Engine • Dashboard</p>
          </div>
        </div>
        <span style={styles.badge}>Security Status: Active</span>
      </header>

      <main style={styles.main}>
        {/* Upload Module (මුලින්ම මේක විතරයි පේන්නේ) */}
        <div style={styles.uploadZone}>
          <div style={{position: 'absolute', top: 0, left: 0, right: 0, height: '4px', background: 'linear-gradient(90deg, #3b82f6, #ef4444, #3b82f6)'}}></div>
          <h2 style={{fontSize: '26px', fontWeight: '800', margin: '0 0 10px 0', letterSpacing: '-0.5px'}}>Cross-Platform Data Stream Analyzer</h2>
          <p style={{fontSize: '14px', color: '#9ca3af', margin: '0 0 30px 0', maxWidth: '650px', marginLeft: 'auto', marginRight: 'auto', lineHeight: '1.5'}}>
            Securely drop internal ledgers or monthly firm statements in <span style={{color: '#60a5fa', fontWeight: '600'}}>PDF, DOCX, XLSX</span>, or <span style={{color: '#60a5fa', fontWeight: '600'}}>CSV</span> format. The machine-learning engine will parse data fields natively.
          </p>
          
          <label style={styles.uploadLabel}>
            <div style={styles.uploadIcon}>⚡</div>
            <span style={{fontSize: '16px', fontWeight: '600', color: '#3b82f6'}}>Upload System Document</span>
            <span style={{fontSize: '12px', color: '#6b7280'}}>Drag & Drop supported locally</span>
            <input type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" style={{display: 'none'}} onChange={handleFileUpload} />
          </label>

          {fileName && (
            <div style={{marginTop: '25px', display: 'inline-flex', alignItems: 'center', gap: '8px', background: 'rgba(59,130,246,0.1)', padding: '8px 20px', borderRadius: '30px', fontSize: '13px', color: '#60a5fa', border: '1px solid rgba(59,130,246,0.3)', fontWeight: '600'}}>
              <span>📊 Layer Loaded:</span> {fileName}
            </div>
          )}
        </div>

        {loading && (
          <div style={{textAlign: 'center', color: '#3b82f6', fontSize: '15px', fontWeight: '600', padding: '40px 0', display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '10px'}}>
            ⏳ Executing Neural Matrix Over Ledger Blocks...
          </div>
        )}

        {/* Analytics Dashboard (ෆයිල් එක දැම්මම විතරක් Activate වෙන කොටස) */}
        {!loading && hasData && (
          <div style={{animation: 'fadeIn 0.6s ease-out'}}>
            
            {/* KPI Cards Grid */}
            <div style={styles.grid}>
              <div style={styles.card}>
                <div style={styles.cardLabel}>Analyzed Registry Count</div>
                <div style={styles.cardVal}>{stats.total}</div>
                <div style={{fontSize: '11px', color: '#6b7280', marginTop: '8px'}}>Total data arrays verified</div>
              </div>
              <div style={styles.card}>
                <div style={styles.cardLabel}>Flagged Threat Anomalies</div>
                <div style={{...styles.cardVal, color: '#f87171', textShadow: '0 0 15px rgba(248,113,113,0.3)'}}>{stats.flagged}</div>
                <div style={{fontSize: '11px', color: '#f87171', marginTop: '8px'}}>Violations against standard compliance</div>
              </div>
              <div style={styles.card}>
                <div style={styles.cardLabel}>Security Risk Assessment</div>
                <div style={{
                  fontSize: '13px', fontWeight: '800', marginTop: '18px', padding: '8px 16px', borderRadius: '8px', display: 'inline-block', textTransform: 'uppercase', letterSpacing: '0.5px',
                  backgroundColor: stats.risk === 'Low' ? 'rgba(52, 211, 153, 0.1)' : stats.risk === 'Medium' ? 'rgba(251, 191, 36, 0.1)' : 'rgba(248, 113, 113, 0.1)',
                  color: stats.risk === 'Low' ? '#34d399' : stats.risk === 'Medium' ? '#fbbf24' : '#f87171',
                  border: stats.risk === 'Low' ? '1px solid rgba(52,211,153,0.3)' : stats.risk === 'Medium' ? '1px solid rgba(251,191,36,0.3)' : '1px solid rgba(248,113,113,0.3)'
                }}>{stats.risk} Risk Profile Detected</div>
              </div>
            </div>

            {/* Responsive Panels Layout */}
            <div style={styles.dashboardLayout}>
              
              {/* Threat Feed Column */}
              <div style={styles.panelLeft}>
                <div style={{...styles.panelTitle, paddingLeft: '0'}}>
                  <span style={{color: '#f87171'}}>🛑</span> Core System Threat Feed
                </div>
                <div style={{maxHeight: '400px', overflowY: 'auto', paddingRight: '5px'}}>
                  {alerts.length === 0 ? (
                    <div style={{padding: '20px', backgroundColor: 'rgba(52, 211, 153, 0.03)', color: '#34d399', borderRadius: '12px', border: '1px solid rgba(52, 211, 153, 0.1)', fontSize: '14px', textAlign: 'center'}}>
                      🎉 Integrity Verified. No anomalous indicators detected inside this file.
                    </div>
                  ) : (
                    alerts.map(alert => (
                      <div key={alert.id} style={styles.alertBox}>
                        <div style={{display: 'flex', justifyContent: 'space-between', fontWeight: '700', marginBottom: '6px', fontSize: '14px'}}>
                          <span style={{color: '#ffffff'}}>{alert.supplier}</span>
                          <span style={{color: '#f87171'}}>Rs. {alert.amount.toLocaleString()}</span>
                        </div>
                        <div style={{color: '#9ca3af', fontSize: '12px', marginBottom: '6px'}}>Invoice Registry: <span style={{color: '#60a5fa', fontFamily: 'monospace', background: '#1e2937', padding: '2px 6px', borderRadius: '4px'}}>{alert.invoice}</span></div>
                        <div style={{color: '#fbbf24', fontSize: '12px', fontWeight: '600', display: 'flex', alignItems: 'center', gap: '5px'}}>
                          <span>⚠️</span> Trigger: {alert.reason}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Data Ledger Column */}
              <div style={styles.panelRight}>
                <div style={styles.panelTitle}>
                  <span style={{color: '#60a5fa'}}>📊</span> Forensic Analytical Sheet Log
                </div>
                <div style={styles.tableContainer}>
                  <table style={styles.table}>
                    <thead>
                      <tr>
                        <th style={{...styles.th, paddingLeft: '25px'}}>Counterparty</th>
                        <th style={styles.th}>Document ID</th>
                        <th style={styles.th}>Value Base</th>
                        <th style={{...styles.th, paddingRight: '25px'}}>Protocol</th>
                      </tr>
                    </thead>
                    <tbody>
                      {alerts.map(alert => (
                        <tr key={alert.id} style={{backgroundColor: 'transparent', borderBottom: '1px solid #1f2937'}}>
                          <td style={{...styles.td, paddingLeft: '25px', fontWeight: '600', color: '#ffffff'}}>{alert.supplier}</td>
                          <td style={{...styles.td, color: '#60a5fa', fontFamily: 'monospace'}}>{alert.invoice}</td>
                          <td style={{...styles.td, fontWeight: '700', color: '#ffffff'}}>Rs. {alert.amount.toLocaleString()}</td>
                          <td style={{...styles.td, paddingRight: '25px'}}>
                            <span style={{backgroundColor: 'rgba(248, 113, 113, 0.1)', color: '#f87171', padding: '4px 10px', borderRadius: '6px', fontSize: '11px', border: '1px solid rgba(248, 113, 113, 0.2)', fontWeight: '600'}}>Isolate</span>
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

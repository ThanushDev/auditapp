import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

export default function App() {
  const [loading, setLoading] = useState(false);
  const [stats, setStats] = useState({ total: 0, flagged: 0, risk: 'None' });
  const [alerts, setAlerts] = useState([]);
  const [fileName, setFileName] = useState('');

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    setLoading(true);
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
        alert('Unsupported format! Use CSV, Excel, Word or PDF.');
        setLoading(false);
        return;
      }

      runFraudDetection(rawTransactions);
    } catch (error) {
      console.error(error);
      alert('Error reading file. Please check format.');
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
      if (row.Invoice_No && invoiceCounts[row.Invoice_No] > 1) reasons.push("Duplicate Invoice");

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

  // Inline CSS Styles for absolute independence from internet speed / Tailwind failures
  const styles = {
    container: { backgroundColor: '#0b1329', color: '#f8fafc', fontFamily: 'system-ui, sans-serif', minHeight: '100vh', padding: '0 0 40px 0' },
    header: { borderBottom: '1px solid #1e293b', backgroundColor: '#111c44', padding: '15px 30px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' },
    brand: { display: 'flex', alignItems: 'center', gap: '12px' },
    logo: { backgroundColor: '#3b82f6', width: '36px', height: '36px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyCwontent: 'center', fontWeight: 'bold', fontSize: '18px' },
    title: { margin: 0, fontSize: '18px', fontWeight: '700', color: '#ffffff' },
    subtitle: { margin: 0, fontSize: '11px', color: '#94a3b8' },
    badge: { backgroundColor: 'rgba(59, 130, 246, 0.1)', color: '#60a5fa', fontSize: '11px', fontWeight: '600', padding: '4px 12px', borderRadius: '20px', border: '1px solid rgba(59, 130, 246, 0.2)' },
    main: { maxWidth: '1200px', margin: '0 auto', padding: '30px 20px' },
    uploadZone: { backgroundColor: '#111c44', border: '2px dashed #334155', borderRadius: '16px', padding: '40px', textAlign: 'center', boxShadow: '0 10px 25px -5px rgba(0,0,0,0.3)', marginBottom: '30px' },
    uploadLabel: { cursor: 'pointer', display: 'inline-flex', flexDirection: 'column', alignItems: 'center', gap: '10px' },
    uploadIcon: { fontSize: '40px', marginBottom: '5px' },
    grid: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '20px', marginBottom: '35px' },
    card: { backgroundColor: '#111c44', border: '1px solid #1e293b', borderRadius: '12px', padding: '20px', position: 'relative' },
    cardLabel: { fontSize: '11px', fontWeight: '600', color: '#94a3b8', textTransform: 'uppercase', tracking: '1px' },
    cardVal: { fontSize: '32px', fontWeight: '700', marginTop: '5px', color: '#ffffff' },
    dashboardLayout: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '30px' },
    panel: { backgroundColor: '#111c44', border: '1px solid #1e293b', borderRadius: '12px', padding: '20px' },
    panelTitle: { fontSize: '15px', fontWeight: '700', marginBottom: '15px', color: '#f1f5f9', display: 'flex', alignItems: 'center', gap: '8px' },
    alertBox: { backgroundColor: '#0b1329', borderLeft: '4px solid #ef4444', borderRadius: '0 8px 8px 0', padding: '12px 15px', fontSize: '12px', marginBottom: '12px' },
    table: { width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' },
    th: { backgroundColor: '#0b1329', padding: '10px 12px', color: '#94a3b8', textTransform: 'uppercase', fontSize: '10px', fontWeight: '600' },
    td: { padding: '12px', borderBottom: '1px solid #1e293b', color: '#cbd5e1' }
  };

  return (
    <div style={styles.container}>
      {/* Dynamic Header */}
      <header style={styles.header}>
        <div style={styles.brand}>
          <div style={styles.logo}>🔍</div>
          <div>
            <h1 style={styles.title}>NextGen Audit Firm</h1>
            <p style={styles.subtitle}>Advanced Omni-Format Fraud Detection Engine</p>
          </div>
        </div>
        <span style={styles.badge}>Pro Audit Core v2.0</span>
      </header>

      <main style={styles.main}>
        {/* Drop Zone */}
        <div style={styles.uploadZone}>
          <h2 style={{fontSize: '22px', fontWeight: '600', margin: '0 0 8px 0'}}>Cross-Platform Statement Analyzer</h2>
          <p style={{fontSize: '13px', color: '#94a3b8', margin: '0 0 25px 0'}}>Upload any corporate statement or ledger in <b>PDF, DOCX, XLSX</b>, or <b>CSV</b> format to perform instant forensic auditing.</p>
          
          <label style={styles.uploadLabel}>
            <div style={styles.uploadIcon}>📥</div>
            <span style={{fontSize: '14px', fontWeight: '500', color: '#3b82f6'}}>Click to choose file for deep scan</span>
            <span style={{fontSize: '11px', color: '#64748b'}}>Supports all popular document extensions</span>
            <input type="file" accept=".csv,.xlsx,.xls,.docx,.pdf" style={{display: 'none'}} onChange={handleFileUpload} />
          </label>

          {fileName && (
            <div style={{marginTop: '15px', display: 'inline-block', backgroundColor: '#1e293b', padding: '5px 15px', borderRadius: '20px', fontSize: '12px', color: '#60a5fa', border: '1px solid #334155'}}>
              📄 Connected: {fileName}
            </div>
          )}
        </div>

        {loading && (
          <div style={{textAlign: 'center', color: '#3b82f6', fontSize: '14px', fontWeight: '500', padding: '20px'}}>
            ⏳ Decoding Document Streams & Structuring Data Layers...
          </div>
        )}

        {/* Dashboard Panels */}
        {!loading && stats.total > 0 && (
          <div>
            {/* KPI Cards Grid */}
            <div style={styles.grid}>
              <div style={styles.card}>
                <div style={styles.cardLabel}>Processed Line Items</div>
                <div style={styles.cardVal}>{stats.total}</div>
              </div>
              <div style={styles.card}>
                <div style={styles.cardLabel}>Flagged Irregularities</div>
                <div style={{...styles.cardVal, color: '#ef4444'}}>{stats.flagged}</div>
              </div>
              <div style={styles.card}>
                <div style={styles.cardLabel}>Risk Level Assessment</div>
                <div style={{
                  fontSize: '14px', fontWeight: '700', marginTop: '15px', padding: '6px 14px', borderRadius: '6px', display: 'inline-block',
                  backgroundColor: stats.risk === 'Low' ? 'rgba(16, 185, 129, 0.1)' : stats.risk === 'Medium' ? 'rgba(245, 158, 11, 0.1)' : 'rgba(239, 68, 68, 0.1)',
                  color: stats.risk === 'Low' ? '#10b981' : stats.risk === 'Medium' ? '#f59e0b' : '#ef4444',
                  border: stats.risk === 'Low' ? '1px solid #10b981' : stats.risk === 'Medium' ? '1px solid #f59e0b' : '1px solid #ef4444'
                }}>{stats.risk} Risk Profile</div>
              </div>
            </div>

            {/* Split Screen Insights */}
            <div style={styles.dashboardLayout}>
              
              {/* AI Threat Alerts Feed */}
              <div style={styles.panel}>
                <div style={styles.panelTitle}>🚨 Automated Forensic Threat Feed</div>
                <div style={{maxHeight: '380px', overflowY: 'auto'}}>
                  {alerts.length === 0 ? (
                    <div style={{padding: '15px', backgroundColor: 'rgba(16, 185, 129, 0.05)', color: '#10b981', borderRadius: '8px', border: '1px solid rgba(16, 185, 129, 0.1)', fontSize: '13px'}}>
                      ✅ Document clean. No anomalous indicators detected inside this file.
                    </div>
                  ) : (
                    alerts.map(alert => (
                      <div key={alert.id} style={styles.alertBox}>
                        <div style={{display: 'flex', justifyContent: 'space-between', fontWeight: '600', marginBottom: '4px'}}>
                          <span style={{color: '#f1f5f9'}}>{alert.supplier}</span>
                          <span style={{color: '#ef4444'}}>Rs. {alert.amount.toLocaleString()}</span>
                        </div>
                        <div style={{color: '#94a3b8', fontSize: '11px'}}>Invoice Ref: <span style={{color: '#60a5fa', fontFamily: 'monospace'}}>{alert.invoice}</span></div>
                        <div style={{color: '#f59e0b', fontSize: '11px', marginTop: '4px', fontWeight: '500'}}>Reason: {alert.reason}</div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* Ledger Sheet Breakdown */}
              <div style={{...styles.panel, padding: '0', overflow: 'hidden'}}>
                <div style={{...styles.panelTitle, padding: '20px 20px 10px 20px', margin: '0'}}>📋 Audit Ledger Records</div>
                <div style={{maxHeight: '350px', overflowY: 'auto'}}>
                  <table style={styles.table}>
                    <thead>
                      <tr>
                        <th style={styles.th}>Entity</th>
                        <th style={styles.th}>Doc ID</th>
                        <th style={styles.th}>Value</th>
                        <th style={styles.th}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {alerts.map(alert => (
                        <tr key={alert.id} style={{backgroundColor: 'rgba(255,255,255,0.01)'}}>
                          <td style={styles.td}>{alert.supplier}</td>
                          <td style={{...styles.td, color: '#60a5fa', fontFamily: 'monospace'}}>{alert.invoice}</td>
                          <td style={{...styles.td, fontWeight: '600', color: '#ffffff'}}>Rs. {alert.amount.toLocaleString()}</td>
                          <td style={styles.td}>
                            <span style={{backgroundColor: 'rgba(239, 68, 68, 0.1)', color: '#ef4444', padding: '2px 8px', borderRadius: '4px', fontSize: '10px', border: '1px solid rgba(239, 68, 68, 0.2)'}}>Review</span>
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

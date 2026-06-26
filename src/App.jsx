import React, { useState, useMemo, useRef } from 'react';
import Papa from 'papaparse';
import { format, parse } from 'date-fns';
import { Upload, FileSpreadsheet, FileText, Image as ImageIcon } from 'lucide-react';
import { toPng } from 'html-to-image';
import { exportToExcel, exportToPDF } from './utils/exportUtils';
import './index.css';

const TOTAL_DAYS = 30;

function App() {
  const [data, setData] = useState([]);
  const [fileName, setFileName] = useState('');
  const [totalTarget, setTotalTarget] = useState(4300000);
  const dashboardRef = useRef(null);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        setData(results.data);
      }
    });
  };

  // Process data to generate summary
  const summary = useMemo(() => {
    if (data.length === 0) return [];

    let maxDay = 0;
    data.forEach(row => {
      const dateRaw = row['Date'];
      if (dateRaw) {
        try {
          const parsedDate = parse(dateRaw, 'dd/MM/yyyy', new Date());
          const day = parsedDate.getDate();
          if (day > maxDay) maxDay = day;
        } catch (e) {}
      }
    });
    
    // Fallback to today if parsing fails
    if (maxDay === 0) maxDay = new Date().getDate();

    const grouped = {};

    // Pre-fill days up to maxDay of June 2026
    for (let i = 1; i <= maxDay; i++) {
      const formattedDate = `${i}-Jun-26`;
      const sortKey = `${i.toString().padStart(2, '0')}/06/2026`; // For sorting if needed
      grouped[formattedDate] = {
        date: formattedDate,
        sortKey: sortKey,
        Commercial: { slip: 0, amount: 0 },
        Industrial: { slip: 0, amount: 0 },
        Institutional: { slip: 0, amount: 0 },
        Residential: { slip: 0, amount: 0 },
        Total: { slip: 0, amount: 0 }
      };
    }

    data.forEach(row => {
      const dateRaw = row['Date'];
      if (!dateRaw) return;

      // Ensure Date format
      let formattedDate = dateRaw;
      try {
        const parsedDate = parse(dateRaw, 'dd/MM/yyyy', new Date());
        formattedDate = format(parsedDate, 'd-MMM-yy');
      } catch (e) {
        // Fallback to raw if parsing fails
      }

      const type = row['Property Type Name']?.trim();
      const amountStr = row['Amount Collected'];
      const amount = parseFloat(amountStr) || 0;

      if (!grouped[formattedDate]) {
        grouped[formattedDate] = {
          date: formattedDate,
          sortKey: dateRaw, // Basic sorting based on raw DD/MM/YYYY
          Commercial: { slip: 0, amount: 0 },
          Industrial: { slip: 0, amount: 0 },
          Institutional: { slip: 0, amount: 0 },
          Residential: { slip: 0, amount: 0 },
          Total: { slip: 0, amount: 0 }
        };
      }

      const rowGroup = grouped[formattedDate];
      let matchedCategory = null;

      if (type === 'Commercial') matchedCategory = 'Commercial';
      else if (type === 'Industrial') matchedCategory = 'Industrial';
      else if (type === 'Institutional') matchedCategory = 'Institutional';
      else if (type === 'Residential') matchedCategory = 'Residential';

      if (matchedCategory) {
        rowGroup[matchedCategory].slip += 1;
        rowGroup[matchedCategory].amount += amount;
        rowGroup.Total.slip += 1;
        rowGroup.Total.amount += amount;
      }
    });

    // Sort by date (assuming all dates are in June 2026 for now, or just rely on simple parsing sort)
    const sorted = Object.values(grouped).sort((a, b) => {
      try {
        const dateA = parse(a.sortKey, 'dd/MM/yyyy', new Date());
        const dateB = parse(b.sortKey, 'dd/MM/yyyy', new Date());
        return dateA - dateB;
      } catch(e) {
        return 0;
      }
    });

    return sorted;
  }, [data]);

  const totals = useMemo(() => {
    const t = {
      Commercial: { slip: 0, amount: 0 },
      Industrial: { slip: 0, amount: 0 },
      Institutional: { slip: 0, amount: 0 },
      Residential: { slip: 0, amount: 0 },
      Total: { slip: 0, amount: 0 }
    };

    summary.forEach(row => {
      t.Commercial.slip += row.Commercial.slip;
      t.Commercial.amount += row.Commercial.amount;
      t.Industrial.slip += row.Industrial.slip;
      t.Industrial.amount += row.Industrial.amount;
      t.Institutional.slip += row.Institutional.slip;
      t.Institutional.amount += row.Institutional.amount;
      t.Residential.slip += row.Residential.slip;
      t.Residential.amount += row.Residential.amount;
      t.Total.slip += row.Total.slip;
      t.Total.amount += row.Total.amount;
    });

    return t;
  }, [summary]);

  const formatCurrency = (val) => {
    return Math.round(val).toString(); // Format without decimals for the table, as in the image
  };
  
  const formatComma = (val) => {
     return Math.round(val).toLocaleString('en-IN');
  }

  const handleExportExcel = () => {
    exportToExcel(summary, totals, { totalTarget, totalDays: TOTAL_DAYS });
  };

  const handleExportPDF = async () => {
    await exportToPDF(summary, totals, { totalTarget, totalDays: TOTAL_DAYS });
  };

  const handleExportImage = async () => {
    if (dashboardRef.current === null) return;
    try {
      const dataUrl = await toPng(dashboardRef.current, { 
        cacheBust: true, 
        pixelRatio: 3, 
        backgroundColor: '#f8fafc' 
      });
      const link = document.createElement('a');
      link.download = `UCC_Summary_Report_${format(new Date(), 'dd_MM_yyyy')}.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('Failed to export image', err);
    }
  };

  const totalCollection = totals.Total.amount;
  const perDayTarget = totalTarget / TOTAL_DAYS;
  const tillDateRequired = summary.length > 0 ? (summary.length * perDayTarget) : 0; // Or based on max date? Using row count as days passed.
  const extraShort = totalCollection - tillDateRequired;

  const renderSlip = (val) => val ? val : '-';
  const renderAmount = (val) => val ? formatCurrency(val) : '-';

  const formatExtraShort = (val) => {
    const rounded = Math.round(val);
    if (rounded < 0) return `- ${Math.abs(rounded).toLocaleString('en-IN')} (Short)`;
    if (rounded > 0) return `+ ${rounded.toLocaleString('en-IN')} (Extra)`;
    return `0 (Balanced)`;
  };

  return (
    <div className="app-container">
      <div className="header-controls">
        <div className="file-upload-wrapper">
          <label className="file-upload-btn">
            <Upload size={20} />
            {fileName ? `File: ${fileName}` : 'Upload CSV Report'}
            <input type="file" accept=".csv" className="file-input" onChange={handleFileUpload} />
          </label>
        </div>
        
        {data.length > 0 && (
          <div className="actions">
            <button className="btn btn-excel" onClick={handleExportExcel}>
              <FileSpreadsheet size={20} /> Export Excel
            </button>
            <button className="btn btn-pdf" onClick={handleExportPDF}>
              <FileText size={20} /> Export PDF
            </button>
            <button className="btn" style={{backgroundColor: '#8b5cf6', color: 'white', display: 'flex', alignItems: 'center', gap: '0.5rem'}} onClick={handleExportImage}>
              <ImageIcon size={20} /> Export Image
            </button>
          </div>
        )}
      </div>

      {data.length === 0 ? (
        <div className="empty-state">
          <FileSpreadsheet />
          <h3>No Data Loaded</h3>
          <p>Please upload the UCC Charge Collection Export CSV file to generate the summary.</p>
        </div>
      ) : (
        <div className="dashboard" ref={dashboardRef} style={{padding: '20px', borderRadius: '12px'}}>
          <div className="table-container">
            <table className="summary-table">
              <thead>
                <tr>
                  <th colSpan="11" className="header-orange" style={{padding: '0.5rem 1rem'}}>
                    <div style={{position: 'relative', display: 'flex', justifyContent: 'space-between', alignItems: 'center'}}>
                      <img src="/LOGO/NatureGreen_Logo.png" alt="Nature Green Logo" style={{height: '60px', background: 'white', padding: '4px', borderRadius: '4px', zIndex: 1}} />
                      <span style={{position: 'absolute', left: 0, right: 0, textAlign: 'center', pointerEvents: 'none'}}>Nature Green Tools & Machine Pvt Ltd</span>
                      <img src="/LOGO/nagar-nigam (1).png?v=2" alt="Nagar Nigam Logo" style={{height: '60px', background: 'white', padding: '4px', borderRadius: '4px', zIndex: 1}} />
                    </div>
                  </th>
                </tr>
                <tr>
                  <th colSpan="11" className="header-brown">NAGAR NIGAM MATHUR VRINDAVAN</th>
                </tr>
                <tr>
                  <th colSpan="11" className="header-yellow">USER CHARGE COLLECTION SUMMARY FOR THE MONTH OF JUNE-2026</th>
                </tr>
                <tr>
                  <th rowSpan="2" style={{width: '100px', backgroundColor: '#F5F5F5'}}>Date</th>
                  <th colSpan="2" className="bg-commercial">Commercial</th>
                  <th colSpan="2" className="bg-industrial">Industrial</th>
                  <th colSpan="2" className="bg-institutional">Institutional</th>
                  <th colSpan="2" className="bg-residential">Residential</th>
                  <th colSpan="2" className="bg-total">June-26</th>
                </tr>
                <tr>
                  <th className="bg-commercial">Slip</th>
                  <th className="bg-commercial">Amount</th>
                  <th className="bg-industrial">Slip</th>
                  <th className="bg-industrial">Amount</th>
                  <th className="bg-institutional">Slip</th>
                  <th className="bg-institutional">Amount</th>
                  <th className="bg-residential">Slip</th>
                  <th className="bg-residential">Amount</th>
                  <th className="bg-total">Slip</th>
                  <th className="bg-total">Amount</th>
                </tr>
              </thead>
              <tbody>
                {summary.map((row, idx) => (
                  <tr key={idx}>
                    <td>{row.date}</td>
                    <td>{renderSlip(row.Commercial.slip)}</td>
                    <td>{renderAmount(row.Commercial.amount)}</td>
                    <td>{renderSlip(row.Industrial.slip)}</td>
                    <td>{renderAmount(row.Industrial.amount)}</td>
                    <td>{renderSlip(row.Institutional.slip)}</td>
                    <td>{renderAmount(row.Institutional.amount)}</td>
                    <td>{renderSlip(row.Residential.slip)}</td>
                    <td>{renderAmount(row.Residential.amount)}</td>
                    <td>{renderSlip(row.Total.slip)}</td>
                    <td>{renderAmount(row.Total.amount)}</td>
                  </tr>
                ))}
                <tr>
                  <td>Total</td>
                  <td>{renderSlip(totals.Commercial.slip)}</td>
                  <td>{renderAmount(totals.Commercial.amount)}</td>
                  <td>{renderSlip(totals.Industrial.slip)}</td>
                  <td>{renderAmount(totals.Industrial.amount)}</td>
                  <td>{renderSlip(totals.Institutional.slip)}</td>
                  <td>{renderAmount(totals.Institutional.amount)}</td>
                  <td>{renderSlip(totals.Residential.slip)}</td>
                  <td>{renderAmount(totals.Residential.amount)}</td>
                  <td>{renderSlip(totals.Total.slip)}</td>
                  <td>{renderAmount(totals.Total.amount)}</td>
                </tr>
              </tbody>
            </table>
          </div>

          <div className="sidebar">
            <table className="sidebar-table">
              <thead>
                <tr>
                  <th colSpan="2" className="bg-light-blue" style={{textAlign: 'center', fontSize: '1.2rem', padding: '1rem'}}>
                    Target Detail For June-2026
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <th className="bg-yellow" style={{fontSize: '1.1rem'}}>Total Target</th>
                  <td className="bg-yellow" style={{fontSize: '1.1rem', padding: '0.5rem'}}>
                    <input 
                      type="number" 
                      value={totalTarget} 
                      onChange={(e) => setTotalTarget(Number(e.target.value))}
                      style={{
                        width: '100%', 
                        padding: '0.3rem', 
                        fontSize: '1.1rem', 
                        fontWeight: 'bold', 
                        textAlign: 'right',
                        border: '1px solid #ccc',
                        borderRadius: '4px'
                      }}
                    />
                  </td>
                </tr>
                <tr>
                  <th>Total Days in this Month</th>
                  <td>{TOTAL_DAYS}</td>
                </tr>
                <tr>
                  <th>Per Day Target</th>
                  <td>{formatComma(perDayTarget)}</td>
                </tr>
                <tr>
                  <th>Total Collection Till Date</th>
                  <td>{formatComma(totalCollection)}</td>
                </tr>
                <tr>
                  <th>Till Date Required</th>
                  <td>{tillDateRequired.toLocaleString('en-IN', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                </tr>
                <tr>
                  <td className="bg-green">Extra/ Short</td>
                  <td className="bg-green" style={{textAlign: 'center'}}>{formatExtraShort(extraShort)}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;

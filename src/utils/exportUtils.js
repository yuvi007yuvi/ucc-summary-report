import * as XLSX from 'xlsx-js-style';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

const formatExtraShort = (val) => {
  const rounded = Math.round(val);
  if (rounded < 0) return `- ${Math.abs(rounded).toLocaleString('en-IN')} (Short)`;
  if (rounded > 0) return `+ ${rounded.toLocaleString('en-IN')} (Extra)`;
  return `0 (Balanced)`;
};

export const exportToExcel = (summary, totals, config) => {
  const wb = XLSX.utils.book_new();

  // Create empty worksheet
  const ws = {};

  const colWidths = [
    { wch: 12 }, // Date
    { wch: 8 },  // Comm Slip
    { wch: 12 }, // Comm Amount
    { wch: 8 },  // Ind Slip
    { wch: 12 }, // Ind Amount
    { wch: 8 },  // Inst Slip
    { wch: 12 }, // Inst Amount
    { wch: 8 },  // Res Slip
    { wch: 12 }, // Res Amount
    { wch: 8 },  // OnDemand Slip
    { wch: 12 }, // OnDemand Amount
    { wch: 8 },  // Total Slip
    { wch: 12 }, // Total Amount
    { wch: 2 },  // Spacer
    { wch: 25 }, // Sidebar col 1
    { wch: 15 }, // Sidebar col 2
  ];

  // Helper to add a cell
  const addCell = (r, c, val, type = 's', style = {}) => {
    const cellRef = XLSX.utils.encode_cell({ r, c });
    ws[cellRef] = { v: val, t: type, s: style };
  };

  const borderAll = {
    top: { style: 'thin', color: { rgb: "000000" } },
    bottom: { style: 'thin', color: { rgb: "000000" } },
    left: { style: 'thin', color: { rgb: "000000" } },
    right: { style: 'thin', color: { rgb: "000000" } }
  };

  // Row 0: Urban Envirotech
  addCell(0, 0, "Nature Green Tools & Machine Pvt Ltd", "s", {
    font: { name: "Calibri", sz: 20, bold: true },
    alignment: { horizontal: "center", vertical: "center" },
    fill: { fgColor: { rgb: "FFCC80" } } // Light Orange
  });

  // Row 1: Nagar Nigam
  addCell(1, 0, "NAGAR NIGAM MATHUR VRINDAVAN", "s", {
    font: { name: "Calibri", sz: 18, bold: true, color: { rgb: "000000" } },
    alignment: { horizontal: "center", vertical: "center" },
    fill: { fgColor: { rgb: "C8E6C9" } } // Light Green
  });

  // Row 2: Title
  addCell(2, 0, `USER CHARGE COLLECTION SUMMARY FOR THE MONTH OF\n${config.monthName.toUpperCase()}-${config.yearLong}`, "s", {
    font: { name: "Calibri", sz: 14, bold: true },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    fill: { fgColor: { rgb: "FFF59D" } } // Light Yellow
  });

  // Header Row 3
  // Header Row 3
  const headerFill = (rgb) => ({ fill: { fgColor: { rgb } }, font: { bold: true }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll });

  addCell(3, 0, "Date", "s", headerFill("F5F5F5"));
  addCell(3, 1, "Commercial", "s", headerFill("FFE4E1"));
  addCell(3, 3, "Industrial", "s", headerFill("E3F2FD"));
  addCell(3, 5, "Institutional", "s", headerFill("F3E5F5"));
  addCell(3, 7, "Residential", "s", headerFill("FFF9C4"));
  addCell(3, 9, "On Demand", "s", headerFill("E2E8F0"));
  addCell(3, 11, `${config.monthAbbr}-${config.yearShort}`, "s", headerFill("E8F5E9"));

  // Header Row 4
  const cats = ["FFE4E1", "E3F2FD", "F3E5F5", "FFF9C4", "E2E8F0", "E8F5E9"];
  for (let i = 0; i < 6; i++) {
    addCell(4, 1 + i * 2, "Slip", "s", headerFill(cats[i]));
    addCell(4, 2 + i * 2, "Amount", "s", headerFill(cats[i]));
  }

  // Merges for main table
  const merges = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } }, // Company Name
    { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } }, // Location
    { s: { r: 2, c: 0 }, e: { r: 2, c: 12 } }, // Title
    { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } }, // Date
    { s: { r: 3, c: 1 }, e: { r: 3, c: 2 } }, // Commercial
    { s: { r: 3, c: 3 }, e: { r: 3, c: 4 } }, // Industrial
    { s: { r: 3, c: 5 }, e: { r: 3, c: 6 } }, // Institutional
    { s: { r: 3, c: 7 }, e: { r: 3, c: 8 } }, // Residential
    { s: { r: 3, c: 9 }, e: { r: 3, c: 10 } }, // OnDemand
    { s: { r: 3, c: 11 }, e: { r: 3, c: 12 } }, // Total
  ];

  // Fill empty merged cells with borders
  const applyMergeBorders = (r1, c1, r2, c2, style) => {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (!ws[XLSX.utils.encode_cell({ r, c })]) {
          addCell(r, c, "", "s", style);
        } else {
           ws[XLSX.utils.encode_cell({ r, c })].s = { ...ws[XLSX.utils.encode_cell({ r, c })].s, ...style };
        }
      }
    }
  }

  applyMergeBorders(0, 0, 0, 12, { fill: { fgColor: { rgb: "FFCC80" } } });
  applyMergeBorders(1, 0, 1, 12, { fill: { fgColor: { rgb: "C8E6C9" } } });
  applyMergeBorders(2, 0, 2, 12, { fill: { fgColor: { rgb: "FFF59D" } } });
  applyMergeBorders(3, 0, 4, 12, borderAll);

  // Data rows
  let r = 5;
  const dataStyle = { font: { bold: true }, alignment: { horizontal: "center" }, border: borderAll };
  const dateStyle = { font: { bold: true }, alignment: { horizontal: "center" }, border: borderAll };

  const addData = (r, c, val) => {
    if (val) {
      addCell(r, c, val, "n", dataStyle);
    } else {
      addCell(r, c, "-", "s", { ...dataStyle, alignment: { horizontal: "center" } });
    }
  };

  summary.forEach(row => {
    addCell(r, 0, row.date, "s", dateStyle);
    addData(r, 1, row.Commercial.slip);
    addData(r, 2, row.Commercial.amount);
    addData(r, 3, row.Industrial.slip);
    addData(r, 4, row.Industrial.amount);
    addData(r, 5, row.Institutional.slip);
    addData(r, 6, row.Institutional.amount);
    addData(r, 7, row.Residential.slip);
    addData(r, 8, row.Residential.amount);
    addData(r, 9, row.OnDemand.slip);
    addData(r, 10, row.OnDemand.amount);
    addData(r, 11, row.Total.slip);
    addData(r, 12, row.Total.amount);
    r++;
  });

  const addTotal = (r, c, val) => {
    if (val) {
      addCell(r, c, val, "n", tStyle);
    } else {
      addCell(r, c, "-", "s", { ...tStyle, alignment: { horizontal: "center" } });
    }
  };

  // Totals row
  const tStyle = { font: { bold: true }, alignment: { horizontal: "center" }, border: borderAll };
  const tLabelStyle = { font: { bold: true }, alignment: { horizontal: "center" }, border: borderAll };
  addCell(r, 0, "Total", "s", tLabelStyle);
  addTotal(r, 1, totals.Commercial.slip);
  addTotal(r, 2, totals.Commercial.amount);
  addTotal(r, 3, totals.Industrial.slip);
  addTotal(r, 4, totals.Industrial.amount);
  addTotal(r, 5, totals.Institutional.slip);
  addTotal(r, 6, totals.Institutional.amount);
  addTotal(r, 7, totals.Residential.slip);
  addTotal(r, 8, totals.Residential.amount);
  addTotal(r, 9, totals.OnDemand.slip);
  addTotal(r, 10, totals.OnDemand.amount);
  addTotal(r, 11, totals.Total.slip);
  addTotal(r, 12, totals.Total.amount);

  // --- Sidebar Table ---
  const sc = 14; // Start column for sidebar
  const perDay = config.totalTarget / config.daysInMonth;
  const colTillDate = totals.Total.amount;
  const tillDateReq = summary.length * perDay;
  const diff = colTillDate - tillDateReq;

  // Title
  addCell(0, sc, `Target Detail For ${config.monthName}-${config.yearLong}`, "s", {
    font: { bold: true, sz: 12 }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll, fill: { fgColor: { rgb: "E1F5FE" } }
  });
  merges.push({ s: { r: 0, c: sc }, e: { r: 1, c: sc + 1 } });
  applyMergeBorders(0, sc, 1, sc + 1, borderAll);

  // Target row
  const yStyle = { font: { bold: true }, fill: { fgColor: { rgb: "FFF59D" } }, border: borderAll };
  addCell(2, sc, "Total Target", "s", yStyle);
  addCell(2, sc + 1, config.totalTarget, "n", yStyle);

  // Empty row
  merges.push({ s: { r: 3, c: sc }, e: { r: 4, c: sc + 1 } });

  // Sidebar labels style
  const slStyle = { border: borderAll };
  const svStyle = { font: { bold: true }, alignment: { horizontal: "center" }, border: borderAll };
  
  addCell(5, sc, "Total Days in this Month", "s", slStyle);
  addCell(5, sc + 1, config.daysInMonth, "n", svStyle);

  addCell(6, sc, "Per Day Target", "s", slStyle);
  addCell(6, sc + 1, Math.round(perDay), "n", svStyle);

  addCell(7, sc, "Total Collection Till Date", "s", slStyle);
  addCell(7, sc + 1, colTillDate, "n", svStyle);

  addCell(8, sc, "Till Date Required", "s", slStyle);
  addCell(8, sc + 1, tillDateReq, "n", { ...svStyle, numFmt: "0.00" });

  const gStyle = { font: { bold: true }, fill: { fgColor: { rgb: "A5D6A7" } }, alignment: { horizontal: "center" }, border: borderAll };
  addCell(9, sc, "Extra/ Short", "s", { ...slStyle, fill: { fgColor: { rgb: "A5D6A7" } } });
  addCell(9, sc + 1, formatExtraShort(diff), "s", gStyle);

  // Finalize Worksheet
  const maxRow = Math.max(r, 9);
  const maxCol = sc + 1;
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
  ws['!merges'] = merges;
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, "Summary");
  XLSX.writeFile(wb, "UCC_Summary_Report.xlsx");
};


const loadImage = (url) => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = reject;
    img.src = url;
  });
};

export const exportToPDF = async (summary, totals, config) => {
  const doc = new jsPDF({ orientation: 'portrait' });

  let logo1 = null;
  let logo2 = null;
  try {
    logo1 = await loadImage('/LOGO/NatureGreen_Logo.png');
    logo2 = await loadImage('/LOGO/nagar-nigam (1).png?v=2');
  } catch (err) {
    console.error("Could not load logos for PDF", err);
  }

  // Main Table
  const tableData = summary.map(row => [
    row.date,
    row.Commercial.slip || '-',
    row.Commercial.amount ? Math.round(row.Commercial.amount) : '-',
    row.Industrial.slip || '-',
    row.Industrial.amount ? Math.round(row.Industrial.amount) : '-',
    row.Institutional.slip || '-',
    row.Institutional.amount ? Math.round(row.Institutional.amount) : '-',
    row.Residential.slip || '-',
    row.Residential.amount ? Math.round(row.Residential.amount) : '-',
    row.OnDemand.slip || '-',
    row.OnDemand.amount ? Math.round(row.OnDemand.amount) : '-',
    row.Total.slip || '-',
    row.Total.amount ? Math.round(row.Total.amount) : '-',
  ]);

  tableData.push([
    'Total',
    totals.Commercial.slip || '-',
    totals.Commercial.amount ? Math.round(totals.Commercial.amount) : '-',
    totals.Industrial.slip || '-',
    totals.Industrial.amount ? Math.round(totals.Industrial.amount) : '-',
    totals.Institutional.slip || '-',
    totals.Institutional.amount ? Math.round(totals.Institutional.amount) : '-',
    totals.Residential.slip || '-',
    totals.Residential.amount ? Math.round(totals.Residential.amount) : '-',
    totals.OnDemand.slip || '-',
    totals.OnDemand.amount ? Math.round(totals.OnDemand.amount) : '-',
    totals.Total.slip || '-',
    totals.Total.amount ? Math.round(totals.Total.amount) : '-'
  ]);

  autoTable(doc, {
    startY: 20,
    margin: { left: 5, right: 5 }, // Use small margins for portrait
    head: [
      [{ content: 'Nature Green Tools & Machine Pvt Ltd', colSpan: 13, styles: { halign: 'center', fillColor: [255, 204, 128], fontSize: 16, minCellHeight: 18, valign: 'middle' } }],
      [{ content: 'NAGAR NIGAM MATHUR VRINDAVAN', colSpan: 13, styles: { halign: 'center', fillColor: [200, 230, 201], textColor: [0, 0, 0], fontSize: 14 } }],
      [{ content: `USER CHARGE COLLECTION SUMMARY FOR THE MONTH OF ${config.monthName.toUpperCase()}-${config.yearLong}`, colSpan: 13, styles: { halign: 'center', fillColor: [255, 245, 157], textColor: [0, 0, 0] } }],
      [
        { content: 'Date', rowSpan: 2, styles: { fillColor: [245, 245, 245] } },
        { content: 'Commercial', colSpan: 2, styles: { fillColor: [255, 228, 225] } },
        { content: 'Industrial', colSpan: 2, styles: { fillColor: [227, 242, 253] } },
        { content: 'Institutional', colSpan: 2, styles: { fillColor: [243, 229, 245] } },
        { content: 'Residential', colSpan: 2, styles: { fillColor: [255, 249, 196] } },
        { content: 'On Demand', colSpan: 2, styles: { fillColor: [226, 232, 240] } },
        { content: `${config.monthAbbr}-${config.yearShort}`, colSpan: 2, styles: { fillColor: [232, 245, 233] } },
      ],
      [
        { content: 'Slip', styles: { fillColor: [255, 228, 225] } }, { content: 'Amount', styles: { fillColor: [255, 228, 225] } },
        { content: 'Slip', styles: { fillColor: [227, 242, 253] } }, { content: 'Amount', styles: { fillColor: [227, 242, 253] } },
        { content: 'Slip', styles: { fillColor: [243, 229, 245] } }, { content: 'Amount', styles: { fillColor: [243, 229, 245] } },
        { content: 'Slip', styles: { fillColor: [255, 249, 196] } }, { content: 'Amount', styles: { fillColor: [255, 249, 196] } },
        { content: 'Slip', styles: { fillColor: [226, 232, 240] } }, { content: 'Amount', styles: { fillColor: [226, 232, 240] } },
        { content: 'Slip', styles: { fillColor: [232, 245, 233] } }, { content: 'Amount', styles: { fillColor: [232, 245, 233] } }
      ]
    ],
    body: tableData,
    theme: 'grid',
    styles: { fontSize: 6.5, cellPadding: 1, lineColor: [0, 0, 0], lineWidth: 0.1 },
    headStyles: { fillColor: [179, 229, 252], textColor: [0, 0, 0], halign: 'center', valign: 'middle' },
    bodyStyles: { fontStyle: 'bold' },
    columnStyles: {
      0: { halign: 'center' },
      1: { halign: 'center' },
      2: { halign: 'center' },
      3: { halign: 'center' },
      4: { halign: 'center' },
      5: { halign: 'center' },
      6: { halign: 'center' },
      7: { halign: 'center' },
      8: { halign: 'center' },
      9: { halign: 'center' },
      10: { halign: 'center' },
      11: { halign: 'center' },
      12: { halign: 'center' }
    },
    willDrawCell: (data) => {
      if (data.section === 'body' && data.row.index === tableData.length - 1) {
        doc.setFillColor(255, 245, 157); // Yellow for totals row
      }
    },
    didDrawCell: (data) => {
      // Draw logos inside the very first header cell (the company name)
      if (data.section === 'head' && data.row.index === 0 && data.column.index === 0) {
        const logoSize = 14;
        const paddingY = (data.cell.height - logoSize) / 2;
        if (logo1) {
          doc.addImage(logo1, 'PNG', data.cell.x + 4, data.cell.y + paddingY, logoSize, logoSize);
        }
        if (logo2) {
          doc.addImage(logo2, 'PNG', data.cell.x + data.cell.width - logoSize - 4, data.cell.y + paddingY, logoSize, logoSize);
        }
      }
    }
  });

  // Sidebar Table
  const perDay = config.totalTarget / config.daysInMonth;
  const colTillDate = totals.Total.amount;
  const tillDateReq = summary.length * perDay;
  const diff = colTillDate - tillDateReq;

  const finalY = doc.lastAutoTable.finalY || 20;

  autoTable(doc, {
    startY: finalY + 10,
    margin: { left: 5, right: 5 },
    head: [
      [{ content: `Target Detail For ${config.monthName}-${config.yearLong}`, colSpan: 2, styles: { halign: 'center', fillColor: [225, 245, 254], fontSize: 10 } }]
    ],
    body: [
      [{ content: 'Total Target', styles: { fillColor: [255, 245, 157], fontStyle: 'bold' } }, { content: config.totalTarget, styles: { fillColor: [255, 245, 157], fontStyle: 'bold', halign: 'center' } }],
      ['Total Days in this Month', { content: config.daysInMonth, styles: { halign: 'center', fontStyle: 'bold' } }],
      ['Per Day Target', { content: Math.round(perDay).toLocaleString('en-IN'), styles: { halign: 'center', fontStyle: 'bold' } }],
      ['Total Collection Till Date', { content: Math.round(colTillDate).toLocaleString('en-IN'), styles: { halign: 'center', fontStyle: 'bold' } }],
      ['Till Date Required', { content: tillDateReq.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }), styles: { halign: 'center', fontStyle: 'bold' } }],
      [
        { content: 'Extra/ Short', styles: { fillColor: [165, 214, 167], fontStyle: 'bold' } }, 
        { content: formatExtraShort(diff), styles: { fillColor: [165, 214, 167], halign: 'center', fontStyle: 'bold' } }
      ]
    ],
    theme: 'grid',
    styles: { fontSize: 9, cellPadding: 2, textColor: [0, 0, 0] },
    headStyles: { textColor: [0, 0, 0], lineWidth: 0.1, lineColor: [0, 0, 0] },
    bodyStyles: { lineWidth: 0.1, lineColor: [0, 0, 0] },
  });

  doc.save('UCC_Summary_Report.pdf');
};

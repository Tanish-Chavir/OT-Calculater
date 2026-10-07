import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { createRequire } from 'module'

function excelExportPlugin() {
  return {
    name: 'excel-export-plugin',
    configureServer(server: any) {
      server.middlewares.use('/api/export-excel', (req: any, res: any) => {
        try {
          // Use xlsx-js-style for styled output via require
          const require2 = createRequire(import.meta.url);
          const XLSX = require2('xlsx-js-style');

          const urlObj = new URL(req.url, 'http://localhost:5173');
          const payloadStr = urlObj.searchParams.get('payload');
          let payload: any = {};
          if (payloadStr) {
            try {
              payload = JSON.parse(payloadStr);
            } catch (e) {}
          }

          const {
            rows = [],
            salary = 32000,
            monthLabel = 'Report',
            selectedMonth = 7,
            selectedYear = 2026,
            totalOTHours = 0,
            totalOTPayment = 0,
            paidAmount = 32000,
            balanceAmount = 0,
            otRateWorkingDay = 100,
            otRateSaturday = 100,
            otRateNonWorkingDay = 200,
            otRateHoliday = 0,
          } = payload;

          const daysInMonth = new Date(selectedYear, selectedMonth, 0).getDate();
          const dailyBasicAmount = Math.round(salary / daysInMonth);

          // Short month names
          const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          const monthYearLabel = `${shortMonths[selectedMonth - 1]}-${selectedYear}`;

          // --- Style Definitions ---
          // Header row style (dark green bg, white bold text)
          const headerStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11, name: "Calibri" },
            fill: { fgColor: { rgb: "4F6228" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "808080" } },
              bottom: { style: "thin", color: { rgb: "808080" } },
              left: { style: "thin", color: { rgb: "808080" } },
              right: { style: "thin", color: { rgb: "808080" } },
            }
          };

          // Month-year title style (olive bg, white bold, larger)
          const titleStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12, name: "Calibri" },
            fill: { fgColor: { rgb: "4F6228" } },
            alignment: { horizontal: "center", vertical: "center" },
          };

          // Data row even (light yellow bg)
          const dataStyleEven = {
            font: { color: { rgb: "333300" }, sz: 10, name: "Calibri" },
            fill: { fgColor: { rgb: "FFFFCC" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "C0C0C0" } },
              bottom: { style: "thin", color: { rgb: "C0C0C0" } },
              left: { style: "thin", color: { rgb: "C0C0C0" } },
              right: { style: "thin", color: { rgb: "C0C0C0" } },
            }
          };

          // Data row odd (white bg)
          const dataStyleOdd = {
            font: { color: { rgb: "333300" }, sz: 10, name: "Calibri" },
            fill: { fgColor: { rgb: "FFFFFF" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "C0C0C0" } },
              bottom: { style: "thin", color: { rgb: "C0C0C0" } },
              left: { style: "thin", color: { rgb: "C0C0C0" } },
              right: { style: "thin", color: { rgb: "C0C0C0" } },
            }
          };

          // Amount column style (currency, green text)
          const amountStyleEven = {
            ...dataStyleEven,
            font: { color: { rgb: "006100" }, sz: 10, name: "Calibri", bold: true },
            numFmt: "₹#,##0",
          };
          const amountStyleOdd = {
            ...dataStyleOdd,
            font: { color: { rgb: "006100" }, sz: 10, name: "Calibri", bold: true },
            numFmt: "₹#,##0",
          };

          // Summary label style (olive bg, white text)
          const summaryLabelStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 10, name: "Calibri" },
            fill: { fgColor: { rgb: "4F6228" } },
            alignment: { horizontal: "right", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "808080" } },
              bottom: { style: "thin", color: { rgb: "808080" } },
              left: { style: "thin", color: { rgb: "808080" } },
              right: { style: "thin", color: { rgb: "808080" } },
            }
          };

          // Summary value style (yellow bg, green bold text)
          const summaryValueStyle = {
            font: { bold: true, color: { rgb: "006100" }, sz: 11, name: "Calibri" },
            fill: { fgColor: { rgb: "FFFF00" } },
            alignment: { horizontal: "center", vertical: "center" },
            numFmt: "₹#,##0",
            border: {
              top: { style: "thin", color: { rgb: "808080" } },
              bottom: { style: "thin", color: { rgb: "808080" } },
              left: { style: "thin", color: { rgb: "808080" } },
              right: { style: "thin", color: { rgb: "808080" } },
            }
          };

          // Total hours label style (olive bg, white text, left aligned)
          const totalHoursLabelStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 10, name: "Calibri" },
            fill: { fgColor: { rgb: "4F6228" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "808080" } },
              bottom: { style: "thin", color: { rgb: "808080" } },
              left: { style: "thin", color: { rgb: "808080" } },
              right: { style: "thin", color: { rgb: "808080" } },
            }
          };

          // Balance label style
          const balanceLabelStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 10, name: "Calibri" },
            fill: { fgColor: { rgb: "4F6228" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
              top: { style: "thin", color: { rgb: "808080" } },
              bottom: { style: "thin", color: { rgb: "808080" } },
              left: { style: "thin", color: { rgb: "808080" } },
              right: { style: "thin", color: { rgb: "808080" } },
            }
          };

          // --- Build worksheet manually with styles ---
          const ws: any = {};
          let currentRow = 0;

          // Helper to set a cell with style
          const setCell = (r: number, c: number, value: any, style: any) => {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (typeof value === 'number') {
              ws[cellRef] = { v: value, t: 'n', s: style };
            } else {
              ws[cellRef] = { v: value || '', t: 's', s: style };
            }
          };

          // Row 0: Month-Year title
          setCell(0, 0, monthYearLabel, titleStyle);
          for (let c = 1; c <= 7; c++) {
            setCell(0, c, '', titleStyle);
          }
          currentRow = 1;

          // Row 1: Header row
          const headers = ['Date', '', 'Ez Hours', 'OT Rate X Hours', '', 'In Time', 'Out Time', 'Amount'];
          headers.forEach((h, c) => {
            setCell(currentRow, c, h, headerStyle);
          });
          currentRow = 2;

          // Data rows
          const dataRows = rows.length > 0 ? rows : [];
          dataRows.forEach((row: any, idx: number) => {
            const isEven = idx % 2 === 0;
            const dStyle = isEven ? dataStyleEven : dataStyleOdd;
            const aStyle = isEven ? amountStyleEven : amountStyleOdd;

            // Format date as DD-MM-YYYY
            let dateFormatted = row.dateStr || '';
            if (row.dateStr) {
              const parts = row.dateStr.split('-'); // YYYY-MM-DD
              if (parts.length === 3) {
                dateFormatted = `${parts[2]}-${parts[1]}-${parts[0]}`;
              }
            }

            setCell(currentRow, 0, dateFormatted, { ...dStyle, alignment: { horizontal: "left", vertical: "center" } });
            setCell(currentRow, 1, '', dStyle);
            setCell(currentRow, 2, row.otHoursHHMM || '00:00', dStyle);
            setCell(currentRow, 3, '', dStyle);
            setCell(currentRow, 4, '', dStyle);
            setCell(currentRow, 5, row.inTime || '', dStyle);
            setCell(currentRow, 6, row.outTime || '', dStyle);
            setCell(currentRow, 7, dailyBasicAmount, aStyle);
            currentRow++;
          });

          // Empty row
          currentRow++;

          // Summary: Basic payment row
          setCell(currentRow, 5, '', summaryLabelStyle);
          setCell(currentRow, 6, 'Basic payment', summaryLabelStyle);
          setCell(currentRow, 7, salary, summaryValueStyle);
          currentRow++;

          // Summary: Total Hours + O.T. Rs. row
          setCell(currentRow, 0, '', totalHoursLabelStyle);
          setCell(currentRow, 1, `Total Hours ${totalOTHours.toFixed(2)} hrs`, totalHoursLabelStyle);
          setCell(currentRow, 2, '', totalHoursLabelStyle);
          setCell(currentRow, 3, '', totalHoursLabelStyle);
          setCell(currentRow, 4, '', totalHoursLabelStyle);
          setCell(currentRow, 5, '', summaryLabelStyle);
          setCell(currentRow, 6, 'O. T.  Rs.', summaryLabelStyle);
          setCell(currentRow, 7, totalOTPayment, summaryValueStyle);
          currentRow++;

          // Summary: Paid amount row
          setCell(currentRow, 5, '', summaryLabelStyle);
          setCell(currentRow, 6, 'Paid amount', summaryLabelStyle);
          setCell(currentRow, 7, paidAmount, summaryValueStyle);
          currentRow++;

          // Empty row
          currentRow++;

          // Balance Amount row
          const balanceText = `Balance Amount  ${paidAmount - salary + totalOTPayment > 0 ? (paidAmount - salary) + '+' + totalOTPayment : totalOTPayment}`;
          setCell(currentRow, 0, '', balanceLabelStyle);
          setCell(currentRow, 1, '', balanceLabelStyle);
          setCell(currentRow, 2, '', balanceLabelStyle);
          setCell(currentRow, 3, balanceText, balanceLabelStyle);
          setCell(currentRow, 4, '', balanceLabelStyle);
          setCell(currentRow, 5, '', balanceLabelStyle);
          setCell(currentRow, 6, '', balanceLabelStyle);
          setCell(currentRow, 7, balanceAmount, summaryValueStyle);

          // Set the range
          ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: currentRow, c: 7 } });

          // Column widths
          ws['!cols'] = [
            { wch: 14 },  // A: Date
            { wch: 4 },   // B: empty
            { wch: 10 },  // C: Ez Hours
            { wch: 18 },  // D: OT Rate X Hours
            { wch: 4 },   // E: empty
            { wch: 10 },  // F: In Time
            { wch: 14 },  // G: Out Time / labels
            { wch: 12 },  // H: Amount
          ];

          // Row heights
          const rowHeights: any[] = [];
          for (let r = 0; r <= currentRow; r++) {
            rowHeights.push({ hpt: 18 });
          }
          ws['!rows'] = rowHeights;

          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

          const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
          const safeMonth = String(monthLabel).replace(/[^a-zA-Z0-9]/g, '_');
          const fileName = `Overtime_Report_${safeMonth}_${selectedYear}.xlsx`;

          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
          res.end(buffer);
        } catch (e) {
          console.error("Server excel export error:", e);
          res.statusCode = 500;
          res.end('Error generating excel: ' + String(e));
        }
      });
    }
  }
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    excelExportPlugin(),
  ],
  resolve: {
    alias: {
      stream: 'stream-browserify',
      buffer: 'buffer',
      events: 'events',
      util: 'util',
    },
  },
})

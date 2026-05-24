import { useState, useEffect } from 'react';
// @ts-ignore
import XLSX from 'xlsx-js-style';

type DayType = 'Working Day' | 'Saturday' | 'Non-working Day' | 'Holiday';

interface RowData {
  dateStr: string;      // e.g. "2026-04-01"
  dayName: string;      // e.g. "Wednesday"
  dayType: DayType;
  inTime: string;       // e.g. "10:30"
  outTime: string;      // e.g. "20:00"
  otHours: string;      // decimal format e.g. "4.75" or "" for non-working
  otHoursHHMM: string;  // HH:MM format e.g. "04:45" or "" for non-working
}

const MONTHS = [
  { value: 1, label: 'January' },
  { value: 2, label: 'February' },
  { value: 3, label: 'March' },
  { value: 4, label: 'April' },
  { value: 5, label: 'May' },
  { value: 6, label: 'June' },
  { value: 7, label: 'July' },
  { value: 8, label: 'August' },
  { value: 9, label: 'September' },
  { value: 10, label: 'October' },
  { value: 11, label: 'November' },
  { value: 12, label: 'December' },
];

const YEARS = Array.from({ length: 11 }, (_, i) => 2020 + i);

// Preset data for April 2026 matching the user's Excel sheet
const APRIL_2026_PRESETS: { [key: number]: { in: string; out: string; type: DayType } } = {
  1: { in: '09:00', out: '23:45', type: 'Working Day' },
  2: { in: '08:00', out: '22:45', type: 'Working Day' },
  3: { in: '09:00', out: '21:30', type: 'Working Day' },
  4: { in: '06:45', out: '20:00', type: 'Saturday' },
  5: { in: '10:30', out: '20:30', type: 'Non-working Day' },
  6: { in: '09:00', out: '00:45', type: 'Working Day' },
  7: { in: '09:00', out: '22:45', type: 'Working Day' },
  8: { in: '09:00', out: '23:00', type: 'Working Day' },
  9: { in: '09:00', out: '23:30', type: 'Working Day' },
  10: { in: '09:00', out: '23:45', type: 'Working Day' },
  11: { in: '09:00', out: '21:00', type: 'Saturday' },
  12: { in: '10:30', out: '20:30', type: 'Non-working Day' },
  13: { in: '09:00', out: '22:30', type: 'Working Day' },
  14: { in: '09:00', out: '21:15', type: 'Working Day' },
  15: { in: '06:45', out: '21:15', type: 'Working Day' },
  16: { in: '08:00', out: '23:45', type: 'Working Day' },
  17: { in: '09:00', out: '23:15', type: 'Working Day' },
  18: { in: '09:00', out: '20:45', type: 'Saturday' },
  19: { in: '10:30', out: '20:30', type: 'Non-working Day' },
  20: { in: '09:00', out: '22:15', type: 'Working Day' },
  21: { in: '09:00', out: '21:30', type: 'Working Day' },
  22: { in: '06:45', out: '21:00', type: 'Working Day' },
  23: { in: '06:45', out: '21:00', type: 'Working Day' },
  24: { in: '06:45', out: '21:15', type: 'Working Day' },
  25: { in: '10:00', out: '20:45', type: 'Saturday' },
  26: { in: '10:30', out: '20:30', type: 'Non-working Day' },
  27: { in: '09:00', out: '21:45', type: 'Working Day' },
  28: { in: '09:00', out: '22:00', type: 'Working Day' },
  29: { in: '09:00', out: '21:45', type: 'Working Day' },
  30: { in: '09:15', out: '22:30', type: 'Working Day' },
};

// Helper to parse typed time strings into standard 24h "HH:MM" format
const parseTimeString = (val: string, format: '12h' | '24h'): string | null => {
  const clean = val.trim().toUpperCase().replace(/\s+/g, ' ');
  if (!clean) return '00:00';

  if (format === '24h') {
    // Match formats like "20:30", "20.30", "2030", "9:15", "9.15"
    let match = clean.match(/^(\d{1,2})[:.](\d{2})$/);
    if (match) {
      const h = Math.min(23, Number(match[1]));
      const m = Math.min(59, Number(match[2]));
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }
    
    match = clean.match(/^(\d{2})(\d{2})$/);
    if (match) {
      const h = Math.min(23, Number(match[1]));
      const m = Math.min(59, Number(match[2]));
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }

    match = clean.match(/^(\d{1,2})$/);
    if (match) {
      const h = Math.min(23, Number(match[1]));
      return `${String(h).padStart(2, '0')}:00`;
    }

    return null;
  } else {
    // 12h format: e.g. "08:30 PM", "8:30 PM", "8.30PM", "8:30PM", "8PM", "8 AM"
    let match = clean.match(/^(\d{1,2})[:.](\d{2})\s*(AM|PM)$/);
    if (match) {
      let h = Number(match[1]);
      const m = Math.min(59, Number(match[2]));
      const ampm = match[3];
      if (ampm === 'PM' && h < 12) h += 12;
      if (ampm === 'AM' && h === 12) h = 0;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }

    match = clean.match(/^(\d{1,2})[:.](\d{2})$/);
    if (match) {
      let h = Number(match[1]);
      const m = Math.min(59, Number(match[2]));
      if (h > 23) h = 23;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }

    match = clean.match(/^(\d{1,2})\s*(AM|PM)$/);
    if (match) {
      let h = Number(match[1]);
      const ampm = match[2];
      if (ampm === 'PM' && h < 12) h += 12;
      if (ampm === 'AM' && h === 12) h = 0;
      return `${String(h).padStart(2, '0')}:00`;
    }

    match = clean.match(/^(\d{1,2})$/);
    if (match) {
      let h = Number(match[1]);
      if (h > 23) h = 23;
      return `${String(h).padStart(2, '0')}:00`;
    }

    return null;
  }
};

// Unified custom Time Input that is 100% typing-only (text input)
function TimeInput({
  value,
  onChange,
  format,
  disabled
}: {
  value: string;
  onChange: (val: string) => void;
  format: '12h' | '24h';
  disabled?: boolean;
}) {
  const toDisplay = (val24: string): string => {
    if (!val24) return '';
    const [hStr, mStr] = val24.split(':');
    const h = Number(hStr) || 0;
    const m = Number(mStr) || 0;
    if (format === '24h') {
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    } else {
      let displayHour = h % 12;
      if (displayHour === 0) displayHour = 12;
      const period = h >= 12 ? 'PM' : 'AM';
      return `${String(displayHour).padStart(2, '0')}:${String(m).padStart(2, '0')} ${period}`;
    }
  };

  const [localVal, setLocalVal] = useState<string>('');

  useEffect(() => {
    setLocalVal(toDisplay(value));
  }, [value, format]);

  const handleBlur = () => {
    const parsed = parseTimeString(localVal, format);
    if (parsed) {
      onChange(parsed);
      setLocalVal(toDisplay(parsed));
    } else {
      setLocalVal(toDisplay(value));
    }
  };

  const placeholder = format === '24h' ? 'e.g. 20:30' : 'e.g. 08:30 PM';

  return (
    <input
      type="text"
      placeholder={placeholder}
      value={localVal}
      disabled={disabled}
      onChange={(e) => setLocalVal(e.target.value)}
      onBlur={handleBlur}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          handleBlur();
        }
      }}
      className={`w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-200 placeholder-slate-600 focus:outline-none focus:ring-1 focus:ring-emerald-500 transition-all text-center ${
        disabled ? 'opacity-30 cursor-not-allowed' : ''
      }`}
    />
  );
}

export default function App() {
  const [selectedMonth, setSelectedMonth] = useState<number>(4); // Default: April
  const [selectedYear, setSelectedYear] = useState<number>(2026); // Default: 2026
  const [salary, setSalary] = useState<number>(27000); // Default: 27000
  const [rows, setRows] = useState<RowData[]>([]);

  // Default time configuration state
  const [defaultInTime, setDefaultInTime] = useState<string>('10:30');
  const [defaultOutTime, setDefaultOutTime] = useState<string>('20:30');

  // Duty hours configuration state
  const [monFriDutyHours, setMonFriDutyHours] = useState<number>(10); // Default: 10h
  const [satDutyHours, setSatDutyHours] = useState<number>(5);     // Default: 5h

  // Time format state: '12h' or '24h'
  const [timeFormat, setTimeFormat] = useState<'12h' | '24h'>('24h');

  // Save Progress status notification
  const [saveStatus, setSaveStatus] = useState<string>('');

  // Function to calculate overtime based on row data
  const calculateRowOT = (inTime: string, outTime: string, dayType: DayType) => {
    if (dayType === 'Non-working Day') {
      return { decimal: '', hhmm: '' };
    }

    if (!inTime || !outTime) {
      return { decimal: '0.00', hhmm: '00:00' };
    }

    const [inH, inM] = inTime.split(':').map(Number);
    const [outH, outM] = outTime.split(':').map(Number);

    let inDecimal = inH + inM / 60;
    let outDecimal = outH + outM / 60;

    if (outDecimal < inDecimal) {
      outDecimal += 24;
    }

    const workHours = outDecimal - inDecimal;
    
    let dutyHours = 0;
    if (dayType === 'Working Day') {
      dutyHours = monFriDutyHours;
    } else if (dayType === 'Saturday') {
      dutyHours = satDutyHours;
    } else if (dayType === 'Holiday') {
      dutyHours = 0;
    }

    const otDecimal = Math.max(0, workHours - dutyHours);

    const otH = Math.floor(otDecimal);
    const otM = Math.round((otDecimal - otH) * 60);
    const hhmm = `${String(otH).padStart(2, '0')}:${String(otM).padStart(2, '0')}`;
    
    return {
      decimal: otDecimal.toFixed(2),
      hhmm,
    };
  };

  // Load state from localStorage on month/year changes
  useEffect(() => {
    const key = `ot_calc_data_${selectedYear}_${selectedMonth}`;
    const saved = localStorage.getItem(key);
    
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed.rows)) {
          setRows(parsed.rows);
          if (typeof parsed.salary === 'number') setSalary(parsed.salary);
          if (parsed.defaultInTime) setDefaultInTime(parsed.defaultInTime);
          if (parsed.defaultOutTime) setDefaultOutTime(parsed.defaultOutTime);
          if (typeof parsed.monFriDutyHours === 'number') setMonFriDutyHours(parsed.monFriDutyHours);
          if (typeof parsed.satDutyHours === 'number') setSatDutyHours(parsed.satDutyHours);
          if (parsed.timeFormat) setTimeFormat(parsed.timeFormat);
          return;
        }
      } catch (e) {
        console.error("Failed parsing saved storage data", e);
      }
    }

    // If no saved data, generate initial rows
    const daysInMonth = new Date(selectedYear, selectedMonth, 0).getDate();
    const newRows: RowData[] = [];

    for (let day = 1; day <= daysInMonth; day++) {
      const date = new Date(selectedYear, selectedMonth - 1, day);
      const dayOfWeek = date.getDay();
      const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      const dayName = dayNames[dayOfWeek];
      
      let defaultDayType: DayType = 'Working Day';
      if (dayOfWeek === 0) {
        defaultDayType = 'Non-working Day';
      } else if (dayOfWeek === 6) {
        defaultDayType = 'Saturday';
      }

      let inTime = defaultInTime;
      let outTime = defaultOutTime;
      let dayType: DayType = defaultDayType;

      if (dayType === 'Saturday') {
        const [inH, inM] = inTime.split(':').map(Number);
        const outDecimal = (inH + inM / 60 + satDutyHours) % 24;
        const outH = Math.floor(outDecimal);
        const outM = Math.round((outDecimal - outH) * 60);
        outTime = `${String(outH).padStart(2, '0')}:${String(outM).padStart(2, '0')}`;
      }

      if (selectedMonth === 4 && selectedYear === 2026 && APRIL_2026_PRESETS[day]) {
        inTime = APRIL_2026_PRESETS[day].in;
        outTime = APRIL_2026_PRESETS[day].out;
        dayType = APRIL_2026_PRESETS[day].type;
      }

      const ot = calculateRowOT(inTime, outTime, dayType);
      const yyyy = selectedYear;
      const mm = String(selectedMonth).padStart(2, '0');
      const dd = String(day).padStart(2, '0');
      const dateStr = `${yyyy}-${mm}-${dd}`;

      newRows.push({
        dateStr,
        dayName,
        dayType,
        inTime,
        outTime,
        otHours: ot.decimal,
        otHoursHHMM: ot.hhmm,
      });
    }

    setRows(newRows);
  }, [selectedMonth, selectedYear]);

  // Recalculate row overtime when custom duty hours change
  const handleDutyHoursChange = (type: 'mon-fri' | 'sat', val: number) => {
    if (type === 'mon-fri') {
      setMonFriDutyHours(val);
    } else {
      setSatDutyHours(val);
    }

    const updated = rows.map(row => {
      let dutyHours = 0;
      if (row.dayType === 'Working Day') {
        dutyHours = type === 'mon-fri' ? val : monFriDutyHours;
      } else if (row.dayType === 'Saturday') {
        dutyHours = type === 'sat' ? val : satDutyHours;
      }

      if (row.dayType === 'Non-working Day') {
        return row;
      }

      const [inH, inM] = row.inTime.split(':').map(Number);
      const [outH, outM] = row.outTime.split(':').map(Number);
      let inDecimal = inH + inM / 60;
      let outDecimal = outH + outM / 60;
      if (outDecimal < inDecimal) outDecimal += 24;

      const workHours = outDecimal - inDecimal;
      const otDecimal = Math.max(0, workHours - dutyHours);
      const otH = Math.floor(otDecimal);
      const otM = Math.round((otDecimal - otH) * 60);
      const hhmm = `${String(otH).padStart(2, '0')}:${String(otM).padStart(2, '0')}`;

      return {
        ...row,
        otHours: otDecimal.toFixed(2),
        otHoursHHMM: hhmm,
      };
    });
    setRows(updated);
  };

  // Save progress manually
  const saveCurrentWork = () => {
    const key = `ot_calc_data_${selectedYear}_${selectedMonth}`;
    const payload = {
      rows,
      salary,
      defaultInTime,
      defaultOutTime,
      monFriDutyHours,
      satDutyHours,
      timeFormat,
    };
    localStorage.setItem(key, JSON.stringify(payload));
    setSaveStatus('Work saved successfully! 💾');
    setTimeout(() => setSaveStatus(''), 3000);
  };

  // Clear saved storage
  const clearSavedData = () => {
    const key = `ot_calc_data_${selectedYear}_${selectedMonth}`;
    localStorage.removeItem(key);
    window.location.reload();
  };

  // Handle value changes and trigger recalculation
  const updateRow = (index: number, updatedFields: Partial<RowData>) => {
    const updated = [...rows];
    const currentRow = { ...updated[index], ...updatedFields };
    
    const ot = calculateRowOT(currentRow.inTime, currentRow.outTime, currentRow.dayType);
    currentRow.otHours = ot.decimal;
    currentRow.otHoursHHMM = ot.hhmm;

    updated[index] = currentRow;
    setRows(updated);
  };

  // Bulk apply default times to all empty/unedited rows or all rows
  const applyDefaultTimesToAll = () => {
    const updated = rows.map(row => {
      let inTime = defaultInTime;
      let outTime = defaultOutTime;
      if (row.dayType === 'Saturday') {
        const [inH, inM] = inTime.split(':').map(Number);
        const outDecimal = (inH + inM / 60 + satDutyHours) % 24;
        const outH = Math.floor(outDecimal);
        const outM = Math.round((outDecimal - outH) * 60);
        outTime = `${String(outH).padStart(2, '0')}:${String(outM).padStart(2, '0')}`;
      }
      const updatedRow = { ...row, inTime, outTime };
      const ot = calculateRowOT(updatedRow.inTime, updatedRow.outTime, updatedRow.dayType);
      updatedRow.otHours = ot.decimal;
      updatedRow.otHoursHHMM = ot.hhmm;
      return updatedRow;
    });
    setRows(updated);
  };

  // Calculations for Summary
  const totalOTHours = rows.reduce((sum, row) => sum + (row.otHours ? Number(row.otHours) : 0), 0);
  const totalOTPayment = Math.round(totalOTHours * 150); // ₹150 rate rounded to whole number
  const paidAmount = salary;
  const balanceAmount = totalOTPayment;

  // EXCEL EXPORT WITH DYNAMIC 100% SOLID FOOTER ROW STYLING (YELLOW & GREEN)
  const exportToExcel = () => {
    const ws: any = { '!ref': `A1:H${rows.length + 10}` };
    
    // Header styling
    const headerStyle = {
      fill: { fgColor: { rgb: '2E7D32' } }, // Deep emerald green
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: 'FFFFFF' } }, // White text
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: '1B5E20' } },
        bottom: { style: 'medium', color: { rgb: '1B5E20' } },
        left: { style: 'thin', color: { rgb: '1B5E20' } },
        right: { style: 'thin', color: { rgb: '1B5E20' } },
      }
    };

    const normalStyle = {
      font: { name: 'Segoe UI', size: 10 },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'E0E0E0' } },
        bottom: { style: 'thin', color: { rgb: 'E0E0E0' } },
        left: { style: 'thin', color: { rgb: 'E0E0E0' } },
        right: { style: 'thin', color: { rgb: 'E0E0E0' } },
      }
    };

    const dateStyle = {
      font: { name: 'Segoe UI', size: 10 },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'E0E0E0' } },
        bottom: { style: 'thin', color: { rgb: 'E0E0E0' } },
        left: { style: 'thin', color: { rgb: 'E0E0E0' } },
        right: { style: 'thin', color: { rgb: 'E0E0E0' } },
      }
    };

    const amountStyle = {
      font: { name: 'Segoe UI', size: 10 },
      alignment: { horizontal: 'right', vertical: 'center' },
      numFmt: '₹#,##0',
      border: {
        top: { style: 'thin', color: { rgb: 'E0E0E0' } },
        bottom: { style: 'thin', color: { rgb: 'E0E0E0' } },
        left: { style: 'thin', color: { rgb: 'E0E0E0' } },
        right: { style: 'thin', color: { rgb: 'E0E0E0' } },
      }
    };

    // Solid Yellow Footer Strip Styles (A34-H34, A36-H36, A37-H37)
    const yellowRowLeftStyle = {
      fill: { fgColor: { rgb: 'FFF2CC' } }, // Solid light yellow background
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: '7F6000' } },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'F0D980' } },
        bottom: { style: 'thin', color: { rgb: 'F0D980' } },
        left: { style: 'thin', color: { rgb: 'F0D980' } },
        right: { style: 'thin', color: { rgb: 'F0D980' } },
      }
    };

    const yellowRowRightStyle = {
      fill: { fgColor: { rgb: 'FFF2CC' } },
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: '7F6000' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      numFmt: '₹#,##0',
      border: {
        top: { style: 'thin', color: { rgb: 'F0D980' } },
        bottom: { style: 'thin', color: { rgb: 'F0D980' } },
        left: { style: 'thin', color: { rgb: 'F0D980' } },
        right: { style: 'thin', color: { rgb: 'F0D980' } },
      }
    };

    const yellowRowCenterStyle = {
      fill: { fgColor: { rgb: 'FFF2CC' } },
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: '7F6000' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'F0D980' } },
        bottom: { style: 'thin', color: { rgb: 'F0D980' } },
        left: { style: 'thin', color: { rgb: 'F0D980' } },
        right: { style: 'thin', color: { rgb: 'F0D980' } },
      }
    };

    // Solid Green Footer Strip Styles (A35-H35, A38-H38)
    const greenRowLeftStyle = {
      fill: { fgColor: { rgb: 'E2EFDA' } }, // Solid light green background
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: '375623' } },
      alignment: { horizontal: 'left', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'A9D18E' } },
        bottom: { style: 'thin', color: { rgb: 'A9D18E' } },
        left: { style: 'thin', color: { rgb: 'A9D18E' } },
        right: { style: 'thin', color: { rgb: 'A9D18E' } },
      }
    };

    const greenRowRightStyle = {
      fill: { fgColor: { rgb: 'E2EFDA' } },
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: '375623' } },
      alignment: { horizontal: 'right', vertical: 'center' },
      numFmt: '₹#,##0',
      border: {
        top: { style: 'thin', color: { rgb: 'A9D18E' } },
        bottom: { style: 'thin', color: { rgb: 'A9D18E' } },
        left: { style: 'thin', color: { rgb: 'A9D18E' } },
        right: { style: 'thin', color: { rgb: 'A9D18E' } },
      }
    };

    const greenRowCenterStyle = {
      fill: { fgColor: { rgb: 'E2EFDA' } },
      font: { name: 'Segoe UI', size: 10, bold: true, color: { rgb: '375623' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'A9D18E' } },
        bottom: { style: 'thin', color: { rgb: 'A9D18E' } },
        left: { style: 'thin', color: { rgb: 'A9D18E' } },
        right: { style: 'thin', color: { rgb: 'A9D18E' } },
      }
    };

    // Helper to set cell
    const setCell = (col: string, rowNum: number, value: any, type: string = 's', style: any = null) => {
      const cellRef = `${col}${rowNum}`;
      ws[cellRef] = { v: value, t: type };
      if (style) {
        ws[cellRef].s = style;
      }
    };

    // Row 1: Month Name and Year in cell A1 (formatted as exactly e.g. "Apr-2026")
    const monthLabel = MONTHS.find(m => m.value === selectedMonth)?.label || 'Month';
    const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthYearLabel = `${shortMonths[selectedMonth - 1]}-${selectedYear}`;
    setCell('A', 1, monthYearLabel, 's', {
      font: { name: 'Segoe UI', size: 12, bold: true, color: { rgb: '1B5E20' } }
    });

    // Row 2: Header Column titles
    setCell('A', 2, 'Date', 's', headerStyle);
    setCell('B', 2, '', 's', headerStyle);
    setCell('C', 2, 'Ex Hours', 's', headerStyle);
    setCell('D', 2, '150Rs. X Hours ', 's', headerStyle);
    setCell('E', 2, '', 's', headerStyle);
    setCell('F', 2, 'In Time', 's', headerStyle);
    setCell('G', 2, 'Out Time', 's', headerStyle);
    setCell('H', 2, 'Amount', 's', headerStyle);

    // Rows 3 to N+2: Daily Rows
    rows.forEach((row, i) => {
      const rNum = i + 3;
      const dateObj = new Date(row.dateStr);

      setCell('A', rNum, dateObj, 'd', { ...dateStyle, numFmt: 'yyyy-mm-dd' });
      setCell('B', rNum, '', 's', normalStyle);

      if (row.dayType === 'Non-working Day') {
        setCell('C', rNum, '', 's', normalStyle);
      } else {
        setCell('C', rNum, row.otHoursHHMM, 's', normalStyle);
      }

      setCell('D', rNum, '', 's', normalStyle);
      setCell('E', rNum, '', 's', normalStyle);

      if (row.dayType === 'Non-working Day') {
        setCell('F', rNum, '', 's', normalStyle);
        setCell('G', rNum, '', 's', normalStyle);
      } else {
        setCell('F', rNum, row.inTime ? `${row.inTime}:00` : '', 's', normalStyle);
        setCell('G', rNum, row.outTime ? `${row.outTime}:00` : '', 's', normalStyle);
      }

      setCell('H', rNum, 900, 'n', amountStyle);
    });

    const n = rows.length;

    // Row N + 4 (A34 to H34): Solid Yellow Row Strip ("Basic payment")
    const rBasic = n + 4;
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => setCell(col, rBasic, '', 's', yellowRowCenterStyle));
    setCell('G', rBasic, 'Basic payment', 's', yellowRowLeftStyle);
    setCell('H', rBasic, salary, 'n', yellowRowRightStyle);

    // Row N + 5 (A35 to H35): Solid Green Row Strip ("Total Hours =")
    const rOT = n + 5;
    ['A', 'D', 'E', 'F'].forEach(col => setCell(col, rOT, '', 's', greenRowCenterStyle));
    setCell('B', rOT, 'Total Hours =', 's', greenRowLeftStyle);
    setCell('C', rOT, `${totalOTHours.toFixed(2)}x150`, 's', greenRowLeftStyle);
    setCell('G', rOT, 'O. T.  Rs.', 's', greenRowLeftStyle);
    setCell('H', rOT, totalOTPayment, 'n', greenRowRightStyle);

    // Row N + 6 (A36 to H36): Solid Yellow Row Strip ("Paid amount")
    const rPaid = n + 6;
    ['A', 'B', 'C', 'D', 'E', 'F'].forEach(col => setCell(col, rPaid, '', 's', yellowRowCenterStyle));
    setCell('G', rPaid, 'Paid amount', 's', yellowRowLeftStyle);
    setCell('H', rPaid, paidAmount, 'n', yellowRowRightStyle);

    // Row N + 7 (A37 to H37): Solid Yellow Row Strip (Empty Separator Strip)
    const rEmptySep = n + 7;
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].forEach(col => setCell(col, rEmptySep, '', 's', yellowRowCenterStyle));

    // Row N + 8 (A38 to H38): Solid Green Row Strip ("Balance Amount")
    const rBalance = n + 8;
    ['A', 'B', 'C', 'D', 'E', 'G'].forEach(col => setCell(col, rBalance, '', 's', greenRowCenterStyle));
    setCell('F', rBalance, 'Balance Amount', 's', greenRowLeftStyle);
    setCell('H', rBalance, balanceAmount, 'n', greenRowRightStyle);

    ws['!cols'] = [
      { wch: 14 }, // A (Date)
      { wch: 14 }, // B (Total Hours label)
      { wch: 14 }, // C (OT Hours / formula)
      { wch: 16 }, // D (150Rs. X Hours header)
      { wch: 5 },  // E (empty)
      { wch: 16 }, // F (In Time / Balance label)
      { wch: 16 }, // G (Out Time / Basic/OT/Paid labels)
      { wch: 16 }  // H (Amount)
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, `Overtime_Report_${monthLabel}_${selectedYear}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-12">
      {/* Global CSS to disable browser spin buttons/arrows for clean typography */}
      <style>{`
        input::-webkit-outer-spin-button,
        input::-webkit-inner-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
        input[type=number] {
          -moz-appearance: textfield;
        }
      `}</style>

      <div className="max-w-7xl mx-auto space-y-6 md:space-y-8">
        
        {/* Header Section */}
        <header className="flex flex-col xl:flex-row justify-between items-start xl:items-center gap-4 pb-6 border-b border-slate-800">
          <div>
            <h1 className="text-2xl md:text-3xl font-extrabold tracking-tight bg-gradient-to-r from-emerald-400 via-teal-400 to-cyan-400 bg-clip-text text-transparent">
              Overtime Calculator Pro
            </h1>
            <p className="text-xs md:text-sm text-slate-400 mt-1">
              Effortlessly track shifts, calculate overtime, and manage payments.
            </p>
          </div>
          
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3 w-full xl:w-auto">
            {/* 12h / 24h Clock Format Switcher */}
            <div className="flex items-center justify-between sm:justify-start gap-3 bg-slate-900 border border-slate-800 rounded-xl px-4 py-2 text-xs">
              <span className="text-slate-400 font-semibold uppercase tracking-wider">Clock format:</span>
              <div className="flex bg-slate-950 p-0.5 rounded-lg border border-slate-800">
                <button
                  onClick={() => setTimeFormat('12h')}
                  className={`px-3 py-1.5 rounded-md font-bold transition-all cursor-pointer ${
                    timeFormat === '12h'
                      ? 'bg-emerald-500 text-slate-950 shadow-md'
                      : 'text-slate-400 hover:text-slate-200'
                  }`}
                >
                  12-Hour
                </button>
                <button
                  onClick={() => setTimeFormat('24h')}
                  className={`px-3 py-1.5 rounded-md font-bold transition-all cursor-pointer ${
                    timeFormat === '24h'
                      ? 'bg-emerald-500 text-slate-950 shadow-md'
                      : 'text-slate-400 hover:text-slate-200'
                  }`}
                >
                  24-Hour
                </button>
              </div>
            </div>

            {/* Save Current Work Action */}
            <button
              onClick={saveCurrentWork}
              className="bg-gradient-to-r from-cyan-500 to-blue-500 hover:from-cyan-600 hover:to-blue-600 active:scale-[0.97] text-slate-950 font-bold py-2.5 px-4 rounded-xl shadow-lg shadow-cyan-500/20 transition-all text-xs flex items-center justify-center gap-2 cursor-pointer w-full sm:w-auto"
            >
              💾 Save Progress
            </button>

            {/* Reset Saved Data */}
            <button
              onClick={clearSavedData}
              className="bg-slate-900 border border-slate-800 hover:bg-slate-850 hover:border-slate-700 active:scale-[0.97] text-slate-300 font-bold py-2.5 px-4 rounded-xl transition-all text-xs flex items-center justify-center gap-2 cursor-pointer w-full sm:w-auto"
            >
              🔄 Reset Default
            </button>

            {/* Download Excel Report */}
            <button
              onClick={exportToExcel}
              className="bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 active:scale-[0.97] text-slate-950 font-bold py-2.5 px-5 rounded-xl shadow-lg shadow-emerald-500/20 hover:shadow-emerald-500/30 transition-all text-xs flex items-center justify-center gap-2 cursor-pointer w-full sm:w-auto"
            >
              📥 Excel Report
            </button>
          </div>
        </header>

        {/* Save Status Notification Banner */}
        {saveStatus && (
          <div className="bg-cyan-950/60 border border-cyan-500/30 text-cyan-400 px-5 py-3.5 rounded-2xl text-sm font-medium animate-bounce flex items-center gap-3">
            <span className="w-2.5 h-2.5 bg-cyan-500 rounded-full animate-pulse"></span>
            {saveStatus}
          </div>
        )}

        {/* Configurations Bar Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Main Configurations Panel */}
          <section className="lg:col-span-2 grid grid-cols-1 sm:grid-cols-3 gap-4 md:gap-6 bg-slate-900/60 border border-slate-800/80 p-5 md:p-6 rounded-2xl backdrop-blur-xl">
            {/* Month Dropdown */}
            <div className="flex flex-col gap-2">
              <label className="text-xs font-semibold text-slate-400 uppercase tracking-wider">
                Select Month
              </label>
              <select
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(Number(e.target.value))}
                className="bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-100 focus:outline-none focus:ring-2 focus:ring-emerald-500/50 transition-all cursor-pointer text-sm"
              >
                {MONTHS.map((m) => (
                  <option key={m.value} value={m.value}>
                    {m.label}
                  </option>
                ))}
              </select>
            </div>

            {/* Year Dropdown */}
            <div className="flex flex-col gap-2">
              <label className="text-xs font-semibold text-slate-400 uppercase tracking-wider">
                Select Year
              </label>
              <select
                value={selectedYear}
                onChange={(e) => setSelectedYear(Number(e.target.value))}
                className="bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-100 focus:outline-none focus:ring-2 focus:ring-emerald-500/50 transition-all cursor-pointer text-sm"
              >
                {YEARS.map((y) => (
                  <option key={y} value={y}>
                    {y}
                  </option>
                ))}
              </select>
            </div>

            {/* Salary Input */}
            <div className="flex flex-col gap-2">
              <label className="text-xs font-semibold text-slate-400 uppercase tracking-wider">
                Monthly Salary (₹)
              </label>
              <div className="relative">
                <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 font-medium">₹</span>
                <input
                  type="number"
                  value={salary}
                  onChange={(e) => setSalary(Number(e.target.value))}
                  placeholder="27000"
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-8 pr-4 py-3 text-slate-100 focus:outline-none focus:ring-2 focus:ring-emerald-500/50 transition-all text-sm"
                />
              </div>
            </div>
          </section>

          {/* Default Time & Duty Hours Panel */}
          <section className="bg-slate-900/60 border border-slate-800/80 p-5 md:p-6 rounded-2xl backdrop-blur-xl flex flex-col justify-between gap-5">
            <div className="grid grid-cols-2 gap-4">
              {/* Default In Time */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">
                  Default In Time
                </label>
                <TimeInput
                  value={defaultInTime}
                  format={timeFormat}
                  onChange={setDefaultInTime}
                />
              </div>

              {/* Default Out Time */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">
                  Default Out Time
                </label>
                <TimeInput
                  value={defaultOutTime}
                  format={timeFormat}
                  onChange={setDefaultOutTime}
                />
              </div>

              {/* Mon-Fri Duty Hours Selection */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">
                  Mon-Fri Duty (h)
                </label>
                <input
                  type="number"
                  min="0"
                  max="24"
                  value={monFriDutyHours}
                  onChange={(e) => handleDutyHoursChange('mon-fri', Number(e.target.value))}
                  className="bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-slate-100 focus:outline-none focus:ring-2 focus:ring-emerald-500/50 transition-all text-xs"
                />
              </div>

              {/* Saturday Duty Hours Selection */}
              <div className="flex flex-col gap-1.5">
                <label className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider">
                  Saturday Duty (h)
                </label>
                <input
                  type="number"
                  min="0"
                  max="24"
                  value={satDutyHours}
                  onChange={(e) => handleDutyHoursChange('sat', Number(e.target.value))}
                  className="bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-slate-100 focus:outline-none focus:ring-2 focus:ring-emerald-500/50 transition-all text-xs"
                />
              </div>
            </div>

            {/* Auto-fill Button */}
            <button
              onClick={applyDefaultTimesToAll}
              className="w-full bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 active:scale-[0.98] text-slate-950 font-bold py-2.5 px-4 rounded-xl shadow-lg shadow-emerald-500/10 hover:shadow-emerald-500/20 transition-all text-sm flex items-center justify-center gap-2 cursor-pointer font-bold"
            >
              ⚡ Apply Defaults to All Rows
            </button>
          </section>
        </div>

        {/* Dynamic Table Section (Desktop View) */}
        <section className="hidden md:block bg-slate-900/60 border border-slate-800/80 rounded-2xl overflow-hidden backdrop-blur-xl">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="border-b border-slate-800 bg-slate-950/80 text-xs font-semibold text-slate-400 uppercase tracking-wider">
                  <th className="py-4 px-6">Date</th>
                  <th className="py-4 px-6">Day Name</th>
                  <th className="py-4 px-6">Day Type</th>
                  <th className="py-4 px-6">In Time</th>
                  <th className="py-4 px-6">Out Time</th>
                  <th className="py-4 px-6 text-right">Overtime Hours</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/50">
                {rows.map((row, index) => {
                  const dateObj = new Date(row.dateStr);
                  const isWeekend = dateObj.getDay() === 0 || dateObj.getDay() === 6;
                  
                  return (
                    <tr 
                      key={row.dateStr}
                      className={`hover:bg-slate-800/30 transition-colors ${
                        row.dayType === 'Non-working Day'
                          ? 'bg-slate-950/40 text-slate-400 opacity-75'
                          : isWeekend
                          ? 'bg-slate-950/20'
                          : ''
                      }`}
                    >
                      {/* Date */}
                      <td className="py-4 px-6 font-medium text-slate-200">
                        {dateObj.toLocaleDateString('en-GB', {
                          day: '2-digit',
                          month: 'short',
                          year: 'numeric',
                        })}
                      </td>

                      {/* Day Name */}
                      <td className="py-4 px-6">
                        <span 
                          className={`text-sm px-2.5 py-1 rounded-lg font-medium inline-block ${
                            row.dayName === 'Sunday'
                              ? 'bg-rose-500/10 text-rose-400'
                              : row.dayName === 'Saturday'
                              ? 'bg-amber-500/10 text-amber-400'
                              : 'bg-slate-800 text-slate-300'
                          }`}
                        >
                          {row.dayName}
                        </span>
                      </td>

                      {/* Day Type Dropdown */}
                      <td className="py-3 px-6">
                        <select
                          value={row.dayType}
                          onChange={(e) => updateRow(index, { dayType: e.target.value as DayType })}
                          className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500 transition-all cursor-pointer"
                        >
                          <option value="Working Day">Working Day</option>
                          <option value="Saturday">Saturday</option>
                          <option value="Non-working Day">Non-working Day</option>
                          <option value="Holiday">Holiday</option>
                        </select>
                      </td>

                      {/* In Time Input */}
                      <td className="py-3 px-6">
                        <TimeInput
                          value={row.inTime}
                          format={timeFormat}
                          onChange={(val) => updateRow(index, { inTime: val })}
                        />
                      </td>

                      {/* Out Time Input */}
                      <td className="py-3 px-6">
                        <TimeInput
                          value={row.outTime}
                          format={timeFormat}
                          onChange={(val) => updateRow(index, { outTime: val })}
                        />
                      </td>

                      {/* OT Hours (HH:MM and decimal) */}
                      <td className="py-3 px-6 text-right">
                        {row.dayType === 'Non-working Day' ? (
                          <span className="text-slate-600 text-xs italic font-medium px-4">—</span>
                        ) : (
                          <div className="flex flex-col items-end">
                            <span className="font-semibold text-emerald-400 text-sm">
                              {row.otHoursHHMM}
                            </span>
                            <span className="text-slate-400 text-xs mt-0.5">
                              ({row.otHours} hrs)
                            </span>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>

        {/* Dynamic Table Section (Mobile Card List View) */}
        <section className="block md:hidden space-y-4">
          <h2 className="text-sm font-semibold uppercase tracking-wider text-slate-400 mb-2 px-1">Daily Log</h2>
          {rows.map((row, index) => {
            const dateObj = new Date(row.dateStr);
            const isWeekend = dateObj.getDay() === 0 || dateObj.getDay() === 6;
            
            return (
              <div 
                key={row.dateStr}
                className={`p-4 rounded-xl border transition-all ${
                  row.dayType === 'Non-working Day'
                    ? 'bg-slate-950/30 border-slate-900/60 opacity-80'
                    : isWeekend
                    ? 'bg-slate-900/40 border-slate-800/60 shadow-md shadow-slate-950/20'
                    : 'bg-slate-900/70 border-slate-800 shadow-md shadow-slate-950/20'
                }`}
              >
                {/* Header: Date & Day Type Dropdown */}
                <div className="flex justify-between items-center pb-3 border-b border-slate-800/50">
                  <div>
                    <h4 className="font-bold text-slate-100 text-sm">
                      {dateObj.toLocaleDateString('en-GB', {
                        day: '2-digit',
                        month: 'short',
                        year: 'numeric',
                      })}
                    </h4>
                    <span className="text-[10px] text-slate-400 font-medium tracking-wide mt-0.5 block">{row.dayName}</span>
                  </div>
                  
                  <select
                    value={row.dayType}
                    onChange={(e) => updateRow(index, { dayType: e.target.value as DayType })}
                    className="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500 transition-all cursor-pointer font-medium"
                  >
                    <option value="Working Day">Working Day</option>
                    <option value="Saturday">Saturday</option>
                    <option value="Non-working Day">Non-working Day</option>
                    <option value="Holiday">Holiday</option>
                  </select>
                </div>

                {/* Inputs Grid */}
                <div className="grid grid-cols-2 gap-3 pt-3 items-center">
                  {/* In Time */}
                  <div className="flex flex-col gap-1">
                    <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">In Time</span>
                    <TimeInput
                      value={row.inTime}
                      format={timeFormat}
                      onChange={(val) => updateRow(index, { inTime: val })}
                    />
                  </div>

                  {/* Out Time */}
                  <div className="flex flex-col gap-1">
                    <span className="text-[9px] font-semibold text-slate-400 uppercase tracking-wider">Out Time</span>
                    <TimeInput
                      value={row.outTime}
                      format={timeFormat}
                      onChange={(val) => updateRow(index, { outTime: val })}
                    />
                  </div>
                </div>

                {/* Calculated Overtime output */}
                <div className="mt-3 pt-3 border-t border-slate-800/50 flex justify-between items-center">
                  <span className="text-[10px] text-slate-400 font-semibold tracking-wider uppercase">Overtime</span>
                  {row.dayType === 'Non-working Day' ? (
                    <span className="text-slate-500 text-xs italic font-medium">Non-working Day</span>
                  ) : (
                    <div className="flex items-center gap-2">
                      <span className="font-bold text-emerald-400 text-sm">
                        {row.otHoursHHMM}
                      </span>
                      <span className="text-slate-400 text-xs">
                        ({row.otHours} hrs)
                      </span>
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </section>

        {rows.length === 0 && (
          <div className="p-8 text-center text-slate-500 bg-slate-900/40 border border-slate-800/80 rounded-2xl">
            No days generated. Please select a valid month and year.
          </div>
        )}

        {/* Totals & Payment Calculation Summary Section */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 md:gap-6">
          {/* Basic Salary */}
          <div className="bg-slate-900/60 border border-slate-800/80 p-5 rounded-2xl backdrop-blur-xl hover:border-slate-700/60 transition-all">
            <h3 className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Basic Salary</h3>
            <p className="text-2xl font-black text-slate-100 mt-2">₹{salary.toLocaleString('en-IN')}</p>
            <span className="text-slate-500 text-xs mt-1 block">Monthly standard wage</span>
          </div>

          {/* Total OT Hours */}
          <div className="bg-slate-900/60 border border-slate-800/80 p-5 rounded-2xl backdrop-blur-xl hover:border-slate-700/60 transition-all">
            <h3 className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Total OT Hours</h3>
            <p className="text-2xl font-black text-teal-400 mt-2">{totalOTHours.toFixed(2)} hrs</p>
            <span className="text-slate-500 text-xs mt-1 block">({totalOTHours.toFixed(2)} × ₹150)</span>
          </div>

          {/* Total OT Payment (OT Rs.) */}
          <div className="bg-slate-900/60 border border-slate-800/80 p-5 rounded-2xl backdrop-blur-xl hover:border-teal-500/30 transition-all ring-1 ring-teal-500/10">
            <h3 className="text-xs font-semibold text-teal-400 uppercase tracking-wider">O. T. Amount</h3>
            <p className="text-2xl font-black text-emerald-400 mt-2">₹{totalOTPayment.toLocaleString('en-IN')}</p>
            <span className="text-slate-500 text-xs mt-1 block">Overtime accrued</span>
          </div>

          {/* Paid Amount */}
          <div className="bg-slate-900/60 border border-slate-800/80 p-5 rounded-2xl backdrop-blur-xl hover:border-slate-700/60 transition-all">
            <h3 className="text-xs font-semibold text-slate-400 uppercase tracking-wider">Paid Amount</h3>
            <p className="text-2xl font-black text-slate-100 mt-2">₹{paidAmount.toLocaleString('en-IN')}</p>
            <span className="text-emerald-500/80 text-xs mt-1 block font-medium">✓ Salary paid</span>
          </div>

          {/* Balance Amount */}
          <div className="bg-emerald-950/20 border border-emerald-500/30 p-5 rounded-2xl backdrop-blur-xl hover:border-emerald-500/50 transition-all ring-2 ring-emerald-500/20 sm:col-span-2 lg:col-span-1">
            <h3 className="text-xs font-bold text-emerald-400 uppercase tracking-wider">Balance Due</h3>
            <p className="text-2xl font-black text-emerald-300 mt-2">₹{balanceAmount.toLocaleString('en-IN')}</p>
            <span className="text-emerald-400/80 text-xs mt-1 block font-medium">⚠ Pending payment</span>
          </div>
        </section>

      </div>
    </div>
  );
}

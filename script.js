const fileInput = document.getElementById('fileInput');
const regularStartInput = document.getElementById('regularStart');
const regularEndInput = document.getElementById('regularEnd');
const morningBoundary = document.getElementById('morningBoundary');
const eveningBoundary = document.getElementById('eveningBoundary');
const calculateBtn = document.getElementById('calculateBtn');
const downloadBtn = document.getElementById('downloadBtn');
const message = document.getElementById('message');
const resultTableBody = document.querySelector('#resultTable tbody');
const totalAmountEl = document.getElementById('totalAmount');

const MEAL_RATE = 8000;
const DOUBLE_MEAL_RATE = 16000;
const MORNING_OFFSET = 2;
const EVENING_OFFSET = 2;

let parsedRows = [];

function showMessage(text, type = 'success') {
  message.textContent = text;
  message.className = `message ${type}`;
}

function parseTimeToMinutes(value) {
  if (!value) return null;

  const cleaned = String(value).trim();
  if (!cleaned) return null;

  const match = cleaned.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return null;

  const hours = Number(match[1]);
  const minutes = Number(match[2]);

  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
  return hours * 60 + minutes;
}

function formatTime(minutes) {
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
}

function updateBoundaries() {
  const regularStart = parseTimeToMinutes(regularStartInput.value);
  const regularEnd = parseTimeToMinutes(regularEndInput.value);

  if (regularStart === null || regularEnd === null) {
    showMessage('근무 시간 형식이 올바르지 않습니다.', 'error');
    return;
  }

  const morning = regularStart - MORNING_OFFSET * 60;
  const evening = regularEnd + EVENING_OFFSET * 60;

  morningBoundary.textContent = formatTime(morning);
  eveningBoundary.textContent = formatTime(evening);
  showMessage('기준선이 갱신되었습니다.', 'success');
}

function normalizeHeader(header) {
  return String(header).trim().toLowerCase().replace(/\s+/g, '');
}

function normalizeRecord(record) {
  const normalized = {};
  Object.entries(record).forEach(([key, value]) => {
    normalized[normalizeHeader(key)] = value;
  });

  return normalized;
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('엑셀 파일을 읽는 중 문제가 발생했습니다.'));
    reader.readAsArrayBuffer(file);
  });
}

function parseUploadedFile(file) {
  return new Promise((resolve, reject) => {
    if (!file) {
      reject(new Error('파일을 선택해 주세요.'));
      return;
    }

    const extension = file.name.split('.').pop()?.toLowerCase();

    if (extension === 'csv') {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (result) => {
          if (result.errors?.length) {
            reject(new Error('CSV 형식을 확인해 주세요.'));
            return;
          }
          resolve(result.data);
        },
        error: () => reject(new Error('CSV 파일을 읽는 중 문제가 발생했습니다.'))
      });
      return;
    }

    if (extension === 'xlsx') {
      readExcelFile(file).then(resolve).catch(reject);
      return;
    }

    reject(new Error('xlsx 또는 csv 파일만 지원합니다.'));
  });
}

function getValue(record, possibleKeys) {
  for (const key of possibleKeys) {
    if (record[key] !== undefined) return record[key];
  }
  return '';
}

function calculateRows(rawRows, regularStart, regularEnd) {
  const morningBoundaryMinutes = regularStart - MORNING_OFFSET * 60;
  const eveningBoundaryMinutes = regularEnd + EVENING_OFFSET * 60;

  return rawRows
    .map((rawRow) => {
      const record = normalizeRecord(rawRow);
      const date = getValue(record, ['날짜', 'date', '일자']);
      const startTime = getValue(record, ['출근시각', '출근시간', 'starttime', 'start']);
      const endTime = getValue(record, ['퇴근시각', '퇴근시간', 'endtime', 'end']);
      const holiday = getValue(record, ['휴일여부', 'holiday', '휴일']);

      const startMinutes = parseTimeToMinutes(startTime);
      const endMinutes = parseTimeToMinutes(endTime);

      if (startMinutes === null || endMinutes === null || endMinutes <= startMinutes) {
        return {
          date,
          startTime,
          endTime,
          holiday: String(holiday).toUpperCase(),
          durationText: '오류',
          morningCondition: 'N',
          eveningCondition: 'N',
          mealCount: 0,
          amount: 0,
          remark: '시간 형식 오류',
          holidayCondition: 'N'
        };
      }

      const durationMinutes = endMinutes - startMinutes;
      const durationHours = durationMinutes / 60;
      const durationText = `${durationHours.toFixed(1)}시간`;
      const morningCondition = startMinutes <= morningBoundaryMinutes ? 'Y' : 'N';
      const eveningCondition = endMinutes >= eveningBoundaryMinutes ? 'Y' : 'N';
      const holidayFlag = String(holiday).toUpperCase();
      const isHoliday = holidayFlag === 'Y' || holidayFlag === 'YES' || holidayFlag === '휴일';
      const holidayCondition = isHoliday && durationMinutes >= 120 ? 'Y' : 'N';

      let mealCount = 0;
      let amount = 0;
      let remark = '';

      if (isHoliday) {
        if (durationMinutes >= 360) {
          mealCount = 2;
          amount = DOUBLE_MEAL_RATE;
          remark = `휴일 근무시간 ${durationHours.toFixed(1)}시간으로 2식`;
        } else if (durationMinutes >= 120) {
          mealCount = 1;
          amount = MEAL_RATE;
          remark = `휴일 근무시간 ${durationHours.toFixed(1)}시간으로 1식`;
        } else {
          remark = `휴일 근무시간 ${durationHours.toFixed(1)}시간으로 조건 미달`;
        }
      } else {
        if (morningCondition === 'Y' || eveningCondition === 'Y') {
          mealCount = 1;
          amount = MEAL_RATE;
          remark = `${morningCondition === 'Y' ? '아침' : ''}${morningCondition === 'Y' && eveningCondition === 'Y' ? '/' : ''}${eveningCondition === 'Y' ? '저녁' : ''} 조건 충족`;
        } else {
          const reasons = [];
          if (startMinutes > morningBoundaryMinutes) {
            reasons.push(`출근 ${formatTime(startMinutes)} 기준선 미달`);
          }
          if (endMinutes < eveningBoundaryMinutes) {
            reasons.push(`퇴근 ${formatTime(endMinutes)} 기준선 미달`);
          }
          remark = reasons.join('; ');
        }
      }

      return {
        date,
        startTime,
        endTime,
        holiday: holidayFlag,
        durationText,
        morningCondition,
        eveningCondition,
        mealCount,
        amount,
        remark,
        holidayCondition
      };
    })
    .filter((row) => row.date !== undefined && row.date !== '');
}

function renderRows(rows) {
  resultTableBody.innerHTML = '';
  const totalAmount = rows.reduce((sum, row) => sum + row.amount, 0);
  totalAmountEl.textContent = `${totalAmount.toLocaleString()}원`;

  if (!rows.length) {
    resultTableBody.innerHTML = '<tr><td colspan="10">표시할 결과가 없습니다.</td></tr>';
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.date}</td>
      <td>${row.startTime}</td>
      <td>${row.endTime}</td>
      <td>${row.holiday}</td>
      <td>${row.durationText}</td>
      <td>${row.morningCondition}</td>
      <td>${row.eveningCondition}</td>
      <td>${row.mealCount}</td>
      <td>${row.amount.toLocaleString()}원</td>
      <td>${row.remark}</td>
    `;
    resultTableBody.appendChild(tr);
  });
}

function downloadCsv() {
  if (!parsedRows.length) {
    showMessage('먼저 결과를 계산해 주세요.', 'error');
    return;
  }

  const header = ['날짜', '출근', '퇴근', '휴일여부', '근무시간', '아침조건', '저녁조건', '식수', '금액', '비고'];
  const rows = parsedRows.map((row) => [
    row.date,
    row.startTime,
    row.endTime,
    row.holiday,
    row.durationText,
    row.morningCondition,
    row.eveningCondition,
    row.mealCount,
    row.amount,
    row.remark
  ]);

  const csvContent = [header, ...rows]
    .map((line) => line.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(','))
    .join('\n');

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'special_meal_allowance.csv';
  a.click();
  URL.revokeObjectURL(url);
}

async function handleCalculate() {
  try {
    if (!fileInput.files.length) {
      showMessage('업로드할 파일을 선택해 주세요.', 'error');
      return;
    }

    const regularStart = parseTimeToMinutes(regularStartInput.value);
    const regularEnd = parseTimeToMinutes(regularEndInput.value);

    if (regularStart === null || regularEnd === null || regularEnd <= regularStart) {
      showMessage('정규근무 시간 설정이 올바르지 않습니다.', 'error');
      return;
    }

    const rawRows = await parseUploadedFile(fileInput.files[0]);
    parsedRows = calculateRows(rawRows, regularStart, regularEnd);
    renderRows(parsedRows);
    downloadBtn.disabled = false;
    showMessage('계산이 완료되었습니다.', 'success');
  } catch (error) {
    console.error(error);
    showMessage(error.message || '계산 중 오류가 발생했습니다.', 'error');
  }
}

regularStartInput.addEventListener('change', updateBoundaries);
regularEndInput.addEventListener('change', updateBoundaries);
calculateBtn.addEventListener('click', handleCalculate);
downloadBtn.addEventListener('click', downloadCsv);

updateBoundaries();

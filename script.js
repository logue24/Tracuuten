/* ============================================================
   TRA CỨU THÔNG TIN — logic đọc dữ liệu & tra cứu
   ============================================================
   Cách dùng:
   1. Tạo file Excel tên chính xác "data.xlsx", tải lên cùng
      thư mục với index.html trên GitHub. Dòng đầu tiên là tên
      cột (VD: Tên, Mã số, Ngày sinh, Ghi chú...), các dòng sau
      là dữ liệu — mỗi dòng một hồ sơ.
   2. Mặc định ô tìm kiếm sẽ dò trên TẤT CẢ các cột. Nếu chỉ
      muốn dò theo một vài cột cụ thể, điền chính xác tên cột
      vào SEARCH_FIELDS bên dưới, ví dụ: ['Tên', 'Mã số'].
   3. Cột mã số có số 0 ở đầu (VD: 0091) nên định dạng kiểu
      Text trong Excel, nếu không Excel/trình đọc sẽ tự bỏ số 0.
   ============================================================ */

const DATA_FILE = 'data.xlsx';
const SEARCH_FIELDS = []; // để trống [] = tìm trong tất cả các cột

let headers = [];
let records = [];

const els = {
  input: document.getElementById('search-input'),
  button: document.getElementById('search-btn'),
  status: document.getElementById('status'),
  results: document.getElementById('results'),
  updated: document.getElementById('updated-time'),
  count: document.getElementById('record-count'),
};

function stripDiacritics(str) {
  return str
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D');
}

function normalize(value) {
  return stripDiacritics(String(value ?? '')).toLowerCase().trim();
}

function setStatus(message, isError) {
  els.status.textContent = message || '';
  els.status.classList.toggle('error', Boolean(isError));
}

async function loadData() {
  setStatus('Đang tải dữ liệu…', false);
  try {
    const res = await fetch(DATA_FILE, { cache: 'no-store' });
    if (!res.ok) throw new Error('HTTP ' + res.status);

    const lastModified = res.headers.get('Last-Modified');
    const buffer = await res.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

    if (!rows.length) throw new Error('File dữ liệu trống');

    headers = rows[0].map(h => String(h).trim()).filter(Boolean);
    records = rows.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => {
        const record = {};
        headers.forEach((h, i) => { record[h] = row[i] !== undefined ? row[i] : ''; });
        return record;
      });

    els.updated.textContent = lastModified
      ? new Date(lastModified).toLocaleString('vi-VN', { dateStyle: 'long', timeStyle: 'short' })
      : 'Không xác định';
    els.count.textContent = String(records.length);

    setStatus('', false);
    els.results.innerHTML = '<p class="hint">Nhập tên hoặc mã số rồi nhấn “Tra cứu” để mở hồ sơ.</p>';
    els.input.disabled = false;
    els.button.disabled = false;
    els.input.focus();
  } catch (err) {
    console.error(err);
    setStatus('Chưa tải được dữ liệu. Kiểm tra file "' + DATA_FILE + '" đã có trong repo và đúng tên chưa.', true);
  }
}

function renderResults(matches, query) {
  els.results.innerHTML = '';

  if (!matches.length) {
    const hint = document.createElement('p');
    hint.className = 'hint';
    hint.textContent = `Không có hồ sơ nào khớp với “${query}”. Kiểm tra lại chính tả hoặc thử mã số đầy đủ.`;
    els.results.appendChild(hint);
    return;
  }

  if (matches.length > 1) {
    const summary = document.createElement('p');
    summary.className = 'hint';
    summary.textContent = `Tìm thấy ${matches.length} hồ sơ khớp với “${query}”.`;
    els.results.appendChild(summary);
  }

  matches.forEach(record => {
    const card = document.createElement('div');
    card.className = 'result-card';

    const stamp = document.createElement('span');
    stamp.className = 'stamp';
    stamp.textContent = 'Đã xác nhận';
    card.appendChild(stamp);

    headers.forEach(h => {
      const row = document.createElement('div');
      row.className = 'result-row';

      const label = document.createElement('span');
      label.className = 'result-label';
      label.textContent = h;

      const value = document.createElement('span');
      value.className = 'result-value';
      const raw = record[h];
      value.textContent = raw !== undefined && String(raw).trim() !== '' ? raw : '—';

      row.appendChild(label);
      row.appendChild(value);
      card.appendChild(row);
    });

    els.results.appendChild(card);
  });
}

function search() {
  const query = els.input.value.trim();
  if (!query) {
    els.results.innerHTML = '<p class="hint">Vui lòng nhập tên hoặc mã số cần tra cứu.</p>';
    return;
  }

  const q = normalize(query);
  const fields = SEARCH_FIELDS.length ? SEARCH_FIELDS : headers;
  const matches = records.filter(record =>
    fields.some(field => normalize(record[field]).includes(q))
  );

  renderResults(matches, query);
}

els.button.addEventListener('click', search);
els.input.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') search();
});

loadData();

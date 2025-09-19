// server.js
require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const qs = require('qs');
const path = require('path');
const https = require('https');
const fs = require('fs');
const fse = require('fs-extra');
const ExcelJS = require('exceljs');
const { Octokit } = require('@octokit/rest');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// ===== Cấu hình API TLU =====
const BASE_URL = 'https://sinhvien1.tlu.edu.vn/education';
const CLIENT_ID = 'education_client';
const CLIENT_SECRET = 'password';

// Dev-only TLS
const INSECURE_TLS = true;
const httpsAgent = new https.Agent({ rejectUnauthorized: !INSECURE_TLS });

// Excel & admin
const EXPORT_DIR = path.join(__dirname, 'exports');
const EXCEL_PATH = path.join(EXPORT_DIR, 'gpa.xlsx');
fse.ensureDirSync(EXPORT_DIR);
const ADMIN_TOKEN = process.env.ADMIN_TOKEN || 'change-me-very-strong';

// ====== GitHub Storage ======
const GITHUB_TOKEN  = process.env.GITHUB_TOKEN;
const GH_REPO_OWNER = process.env.GH_REPO_OWNER;
const GH_REPO_NAME  = process.env.GH_REPO_NAME;
const GH_FILE_PATH  = process.env.GH_FILE_PATH || 'gpa.xlsx';

function getOctokit() {
  if (!GITHUB_TOKEN || !GH_REPO_OWNER || !GH_REPO_NAME) {
    throw new Error('Thiếu biến môi trường GitHub (GITHUB_TOKEN, GH_REPO_OWNER, GH_REPO_NAME)');
  }
  return new Octokit({ auth: GITHUB_TOKEN });
}
async function ghGetDefaultBranch(octokit) {
  const { data } = await octokit.repos.get({ owner: GH_REPO_OWNER, repo: GH_REPO_NAME });
  return data.default_branch || 'main';
}
async function ghGetFileSha(octokit, path, ref) {
  try {
    const { data } = await octokit.repos.getContent({ owner: GH_REPO_OWNER, repo: GH_REPO_NAME, path, ref });
    if (Array.isArray(data)) return null;
    return data.sha;
  } catch { return null; }
}
async function uploadExcelToGitHub(buffer) {
  const octokit = getOctokit();
  const branch = await ghGetDefaultBranch(octokit);
  const sha = await ghGetFileSha(octokit, GH_FILE_PATH, branch);
  const content = buffer.toString('base64');
  const message = `chore: update ${GH_FILE_PATH} at ${new Date().toISOString()}`;
  await octokit.repos.createOrUpdateFileContents({
    owner: GH_REPO_OWNER, repo: GH_REPO_NAME, path: GH_FILE_PATH,
    message, content, sha: sha || undefined, branch
  });
}
async function syncLocalFromGitHubOnce() {
  try {
    const octokit = getOctokit();
    const branch = await ghGetDefaultBranch(octokit);
    const sha = await ghGetFileSha(octokit, GH_FILE_PATH, branch);
    if (!sha) { console.log('[GitHub] Chưa có file trên repo, sẽ tạo khi flush.'); return; }
    const { data } = await octokit.repos.getContent({ owner: GH_REPO_OWNER, repo: GH_REPO_NAME, path: GH_FILE_PATH, ref: branch });
    if (Array.isArray(data)) return;
    const buf = Buffer.from(data.content || '', 'base64');
    if (buf.length === 0) { console.warn('[GitHub] File trên repo 0 byte, bỏ qua không ghi local.'); return; }
    fs.writeFileSync(EXCEL_PATH, buf);
    console.log('[GitHub] Đã tải gpa.xlsx từ repo về local.');
  } catch (e) {
    console.warn('[GitHub] Sync lần đầu thất bại:', e.message);
  }
}

// ===== Helpers =====
const VN_TZ = 'Asia/Ho_Chi_Minh';
function nowVN() {
  return new Intl.DateTimeFormat('sv-SE', {
    timeZone: VN_TZ, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  }).format(new Date());
}
function safeGet(obj, path, fallback = null) {
  try { return path.split('.').reduce((a, k) => (a == null ? a : a[k]), obj) ?? fallback; }
  catch { return fallback; }
}
function extractFromPayload(payload, fallbackUsername) {
  const name  = safeGet(payload, 'student.displayName', 'N/A');
  const clazz = safeGet(payload, 'student.enrollmentClass.className', 'N/A');
  const code  = safeGet(payload, 'student.studentCode', fallbackUsername || 'N/A');
  let gpa = safeGet(payload, 'learningMark4', null);
  if (typeof gpa !== 'number') {
    for (const k of ['mark4','firstLearningMark4']) { const v = payload?.[k]; if (typeof v === 'number') { gpa = v; break; } }
  }
  return { name, clazz, code, gpa };
}

// ===== Chuẩn hoá Job/Hobby từ FE =====
const JOB_LABELS = {
  'backend': 'Backend Developer',
  'frontend': 'Frontend Developer',
  'fullstack': 'Full-stack Developer',
  'mobile': 'Mobile Developer',
  'data-engineer': 'Data Engineer',
  'data-scientist': 'Data Scientist / Analyst',
  'devops': 'DevOps / Cloud Engineer',
  'security': 'Security Engineer',
  'qa': 'QA / Tester',
  'pm': 'Product / Project Manager',
  'uiux': 'UI/UX Designer',
  'other': 'Khác'
};
const HOBBY_LABELS = {
  'bong-da': 'Bóng đá',
  'cau-long': 'Cầu lông',
  'bong-ro': 'Bóng rổ',
  'bong-chuyen': 'Bóng chuyền',
  'chay-bo': 'Chạy bộ',
  'gym': 'Gym',
  'ca-hat': 'Ca hát',
  'nhac': 'Nghe nhạc',
  'nhiep-anh': 'Nhiếp ảnh',
  'doc-sach': 'Đọc sách',
  'du-lich': 'Du lịch',
  'game': 'Chơi game',
  'lap-trinh': 'Lập trình',
  'sang-tao-noi-dung': 'Sáng tạo nội dung',
  'tu-thien': 'Tình nguyện',
  'other': 'Khác'
};

function normalizeJob(job, jobOther) {
  if (!job) return null;
  const label = JOB_LABELS[job] || job;
  if (job === 'other') {
    const extra = (jobOther || '').trim();
    return extra ? `Khác: ${extra}` : 'Khác';
  }
  return label;
}
function normalizeHobbies(hobbies, hobbyOther) {
  const list = Array.isArray(hobbies) ? hobbies : (typeof hobbies === 'string' && hobbies ? [hobbies] : []);
  const labels = list.map(v => HOBBY_LABELS[v] || v);
  if (list.includes('other')) {
    const extra = (hobbyOther || '').trim();
    labels[labels.indexOf(HOBBY_LABELS['other'])] = extra ? `Khác: ${extra}` : 'Khác';
  }
  return labels.length ? labels.join(', ') : null;
}

// ===== Ca học từ kỳ 13 =====
const SCHEDULE_PATH = '/api/StudentCourseSubject/studentLoginUser/13';
const CODE_TO_CA = {
  '251071_CSE414_64HTTT1_1': 'Ca1',
  '251071_CSE414_64HTTT1_2': 'Ca2'
};
async function fetchCaHocLabel(accessToken) {
  const resp = await axios.get(`${BASE_URL}${SCHEDULE_PATH}`, {
    headers: { Authorization: `Bearer ${accessToken}` },
    timeout: 20000, validateStatus: () => true, httpsAgent,
  });
  if (resp.status >= 400) throw new Error(`Schedule API ${resp.status}: ${JSON.stringify(resp.data)}`);
  const items = Array.isArray(resp.data) ? resp.data : (resp.data?.data || []);
  const labels = new Set();
  for (const it of items) {
    const code = it?.courseSubject?.code || it?.classCode || it?.code || '';
    const ca = CODE_TO_CA[String(code).trim()];
    if (ca) labels.add(ca);
  }
  if (labels.size === 0) return null;
  if (labels.size === 1) return [...labels][0];
  return [...labels].join(', ');
}

// ===== Buffer + Flush theo lô =====
const FLUSH_EVERY_MS  = parseInt(process.env.FLUSH_EVERY_MS || '60000', 10);
const FLUSH_THRESHOLD = parseInt(process.env.FLUSH_THRESHOLD || '25', 10);

let wbCache = null, wsCache = null;
const pending = new Map(); // username -> record
let flushing = false;

async function initWorkbook() {
  if (wbCache) return;
  const wb = new ExcelJS.Workbook();

  if (fs.existsSync(EXCEL_PATH)) {
    const size = fs.statSync(EXCEL_PATH).size;
    if (size > 0) {
      try { await wb.xlsx.readFile(EXCEL_PATH); }
      catch (e) {
        console.warn('[Excel] File hỏng, tạo mới:', e.message);
        try { fs.renameSync(EXCEL_PATH, EXCEL_PATH + `.corrupt.${Date.now()}.xlsx`); } catch {}
      }
    } else {
      console.warn('[Excel] gpa.xlsx = 0 byte, xóa và tạo mới.');
      try { fs.unlinkSync(EXCEL_PATH); } catch {}
    }
  }

  let ws = wb.getWorksheet('GPA');
  if (!ws) ws = wb.addWorksheet('GPA');

  // KHÔNG có cột password — thêm 2 cột mới: Công việc IT, Sở thích
  ws.columns = [
    { header: 'Timestamp',   key: 'timestamp', width: 22 },
    { header: 'Username',    key: 'username',  width: 18 },
    { header: 'Mã SV',       key: 'code',      width: 16 },
    { header: 'Tên',         key: 'name',      width: 24 },
    { header: 'Lớp',         key: 'clazz',     width: 18 },
    { header: 'GPA',         key: 'gpa',       width: 10 },
    { header: 'Ca học',      key: 'caHoc',     width: 14 },
    { header: 'Công việc IT',key: 'job',       width: 22 },
    { header: 'Sở thích',    key: 'hobbies',   width: 30 },
  ];
  if (ws.rowCount === 0) {
    ws.addRow({
      timestamp: 'Timestamp', username: 'Username', code: 'Mã SV', name: 'Tên',
      clazz: 'Lớp', gpa: 'GPA', caHoc: 'Ca học', job: 'Công việc IT', hobbies: 'Sở thích'
    });
  }
  wbCache = wb; wsCache = ws;
}

function upsertRowInSheet(ws, rec) {
  const usernameCol = ws.getColumn('username').number;
  let found = null;
  for (let r = 2; r <= ws.rowCount; r++) {
    const v = ws.getCell(r, usernameCol).value;
    if (String(v || '').trim() === String(rec.username || '').trim()) { found = r; break; }
  }
  if (found) {
    ws.getCell(found, ws.getColumn('timestamp').number).value = rec.timestamp;
    ws.getCell(found, usernameCol).value = rec.username;
    ws.getCell(found, ws.getColumn('code').number).value = rec.code;
    ws.getCell(found, ws.getColumn('name').number).value = rec.name;
    ws.getCell(found, ws.getColumn('clazz').number).value = rec.clazz;
    ws.getCell(found, ws.getColumn('gpa').number).value = (typeof rec.gpa === 'number' ? rec.gpa : null);
    ws.getCell(found, ws.getColumn('caHoc').number).value = rec.caHoc || null;
    ws.getCell(found, ws.getColumn('job').number).value = rec.job || null;
    ws.getCell(found, ws.getColumn('hobbies').number).value = rec.hobbies || null;
  } else {
    ws.addRow({
      timestamp: rec.timestamp, username: rec.username,
      code: rec.code, name: rec.name, clazz: rec.clazz,
      gpa: (typeof rec.gpa === 'number' ? rec.gpa : null),
      caHoc: rec.caHoc || null,
      job: rec.job || null,
      hobbies: rec.hobbies || null
    });
  }
}

async function upsertInMemory({ username, code, name, clazz, gpa, caHoc, job, hobbies }) {
  await initWorkbook();
  const rec = {
    timestamp: nowVN(),
    username, code, name, clazz, gpa, caHoc,
    job: job || null,
    hobbies: hobbies || null
  };
  pending.set(username, rec); // upsert theo username
}

async function flushBatch() {
  if (flushing) return;
  if (pending.size === 0) return;
  flushing = true;
  try {
    await initWorkbook();
    for (const rec of pending.values()) upsertRowInSheet(wsCache, rec);

    await wbCache.xlsx.writeFile(EXCEL_PATH);
    const buf = await wbCache.xlsx.writeBuffer();
    try {
      await uploadExcelToGitHub(buf);
      console.log(`[flush] Đã ghi ${pending.size} bản ghi & upload GitHub.`);
      pending.clear();
    } catch (e) {
      console.warn('[flush] Upload GitHub lỗi, giữ pending để thử lần sau:', e.message);
    }
  } catch (e) {
    console.error('[flush] Lỗi ghi Excel:', e.message);
  } finally {
    flushing = false;
  }
}

// Flush định kỳ
setInterval(() => { flushBatch().catch(()=>{}); }, FLUSH_EVERY_MS);

// ====== APIs ======
app.post('/api/gpa/save', async (req, res) => {
  try {
    const { username, password, job, jobOther, hobbies, hobbyOther } = req.body || {};
    if (!username || !password) return res.status(400).json({ message: 'Thiếu username hoặc password' });

    // 1) token
    const data = qs.stringify({ grant_type:'password', client_id:CLIENT_ID, client_secret:CLIENT_SECRET, username, password });
    const tokenResp = await axios.post(`${BASE_URL}/oauth/token`, data, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      timeout: 15000, validateStatus: () => true, httpsAgent,
    });
    if (tokenResp.status >= 400) {
      return res.status(tokenResp.status).json({ message: 'Tài khoản mật khẩu không chính xác', detail: tokenResp.data });
    }
    const accessToken = tokenResp.data?.access_token;
    if (!accessToken) return res.status(500).json({ message: 'Không nhận được access_token' });

    // 2) GPA
    const gpaResp = await axios.get(`${BASE_URL}/api/studentsummarymark/getbystudent`, {
      headers: { Authorization: `Bearer ${accessToken}` },
      timeout: 15000, validateStatus: () => true, httpsAgent,
    });
    if (gpaResp.status >= 400) {
      return res.status(gpaResp.status).json({ message: 'Gọi API GPA thất bại', detail: gpaResp.data });
    }

    // 3) Parse
    const payload = gpaResp.data;
    const { name, clazz, code, gpa } = extractFromPayload(payload, username);

    // 4) Ca học
    let caHoc = null;
    try { caHoc = await fetchCaHocLabel(accessToken); } catch (_) { caHoc = null; }

    // 5) Chuẩn hoá 2 trường mới
    const jobText = normalizeJob(job, jobOther);
    const hobbiesText = normalizeHobbies(hobbies, hobbyOther);

    // 6) Upsert vào bộ đệm
    await upsertInMemory({ username, code, name, clazz, gpa, caHoc, job: jobText, hobbies: hobbiesText });

    if (pending.size >= FLUSH_THRESHOLD) { flushBatch().catch(()=>{}); }

    return res.json({ ok: true, message: 'Đã lưu tạm; sẽ đồng bộ vào Excel & GitHub ở lần flush gần nhất.' });
  } catch (err) {
    return res.status(500).json({ message: 'Lỗi hệ thống', error: err?.message });
  }
});

// Admin: export -> FLUSH NGAY rồi trả file
app.get('/admin/export', async (req, res) => {
  const token = req.header('x-admin-token');
  if (token !== ADMIN_TOKEN) return res.status(401).json({ message: 'Unauthorized' });

  await flushBatch(); // luôn flush ngay khi export

  if (!fs.existsSync(EXCEL_PATH)) return res.status(404).json({ message: 'Chưa có dữ liệu' });
  res.download(EXCEL_PATH, 'gpa.xlsx');
});

// Admin: xem pending
app.get('/admin/stats', async (req, res) => {
  const token = req.header('x-admin-token');
  if (token !== ADMIN_TOKEN) return res.status(401).json({ message: 'Unauthorized' });
  let rows = 0;
  if (fs.existsSync(EXCEL_PATH)) {
    const wb = new ExcelJS.Workbook();
    try {
      if (fs.statSync(EXCEL_PATH).size > 0) {
        await wb.xlsx.readFile(EXCEL_PATH);
        const ws = wb.getWorksheet('GPA'); rows = Math.max(0, (ws?.rowCount || 1) - 1);
      }
    } catch {}
  }
  res.json({ rows, pending: pending.size, flushing, file: EXCEL_PATH });
});

// Khởi động
(async () => {
  await syncLocalFromGitHubOnce();
  await initWorkbook();
  app.listen(process.env.PORT || 3000, () => {
    console.log('Server listening on http://localhost:3000');
    console.log(`Flush every: ${FLUSH_EVERY_MS}ms, threshold: ${FLUSH_THRESHOLD}`);
  });
})();

// Flush lần cuối khi thoát
async function gracefulExit() {
  try { await flushBatch(); } finally { process.exit(0); }
}
process.on('SIGINT', gracefulExit);
process.on('SIGTERM', gracefulExit);

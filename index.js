const crypto = require('crypto');
const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const { Readable } = require('stream');

const app = express();

const LINE_ACCESS_TOKEN  = process.env.LINE_CHANNEL_ACCESS_TOKEN || '';
const LINE_SECRET        = process.env.LINE_CHANNEL_SECRET || '';
const SHEET_ID           = process.env.SHEET_ID || '';
const DRIVE_FOLDER_ID    = process.env.DRIVE_FOLDER_ID || '';
const SHEET_NAME         = process.env.SHEET_NAME || '派工紀錄';
const EMPLOYEE_SHEET     = process.env.EMPLOYEE_SHEET_NAME || '員工清單';
const REPORT_EXPIRY_MS   = (parseInt(process.env.REPORT_EXPIRY_MINUTES) || 10) * 60 * 1000;

const reportMode = {};

app.get('/', (req, res) => res.send('OK'));

app.post('/', express.json({
  verify: (req, res, buf) => { req.rawBody = buf.toString(); },
}), async (req, res) => {
  const signature = req.headers['x-line-signature'] || '';
  const body = req.rawBody || JSON.stringify(req.body);
  if (!verifySignature(body, signature)) {
    return res.status(200).json({ error: 'Invalid signature' });
  }
  try {
    const events = req.body.events || [];
    for (const event of events) {
      try { await handleEvent(event); } catch (err) {
        console.error('handleEvent error:', err.message);
      }
    }
  } catch (err) {
    console.error('parse error:', err.message);
  }
  res.status(200).json({ ok: true });
});

async function handleEvent(event) {
  if (event.type !== 'message') return;
  const msg = event.message;
  const src = event.source;
  const rt  = event.replyToken;
  if (msg.type === 'text') {
    const text = msg.text.trim();
    if (/^#派工/.test(text)) {
      if (src.type !== 'user') return;
      await handleDispatch(text, src, rt);
    } else if (/^#請假/.test(text)) {
      if (src.type !== 'user') return;
      await handleLeave(text, rt);
    } else if (/^#加班/.test(text)) {
      if (src.type !== 'user') return;
      await handleOT(text, rt);
    } else if (/^#回報/.test(text)) {
      if (src.type !== 'group') return;
      await handleReportText(src, rt);
    }
  } else if (msg.type === 'image') {
    if (src.type === 'group') {
      await handleReportPhoto(msg.id, src, rt);
    }
  }
}

function verifySignature(body, signature) {
  if (!LINE_SECRET || !signature) return false;
  try {
    const hmac = crypto.createHmac('sha256', LINE_SECRET);
    hmac.update(Buffer.from(body, 'utf-8'));
    return hmac.digest('base64') === signature;
  } catch (e) {
    return false;
  }
}

const LINE_API = 'https://api.line.me/v2/bot';
const LINE_DATA = 'https://api-data.line.me/v2/bot';

async function getLineDisplayName(userId) {
  try {
    const res = await axios.get(`${LINE_API}/profile/${encodeURIComponent(userId)}`, {
      headers: { Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
      timeout: 5000,
    });
    return res.data.displayName || '未知';
  } catch (e) {
    return '未知';
  }
}

async function downloadContent(messageId) {
  const res = await axios.get(`${LINE_DATA}/message/${messageId}/content`, {
    headers: { Authorization: `Bearer ${LINE_ACCESS_TOKEN}` },
    responseType: 'arraybuffer',
    timeout: 15000,
  });
  return { data: Buffer.from(res.data), contentType: res.headers['content-type'] || 'image/jpeg' };
}

async function replyMessage(replyToken, text) {
  if (!replyToken || !LINE_ACCESS_TOKEN) return;
  try {
    await axios.post(`${LINE_API}/message/reply`, {
      replyToken,
      messages: [{ type: 'text', text }],
    }, {
      headers: { Authorization: `Bearer ${LINE_ACCESS_TOKEN}`, 'Content-Type': 'application/json' },
      timeout: 5000,
    });
  } catch (e) {}
}

async function getAuth() {
  const credsJson = process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON;
  if (credsJson) {
    const credentials = JSON.parse(credsJson);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file'],
    });
    return auth.getClient();
  }
  return google.auth.getClient({
    scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file'],
  });
}

async function getSheetsClient() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

async function getOrCreateSheet() {
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  let sheet = meta.data.sheets.find(s => s.properties.title === SHEET_NAME);
  if (!sheet) {
    const headers = ['序號', '類型', '日期', '人員', '地點', '工作', '時數(H)', '備註', '照片連結', '建立時間'];
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{ addSheet: { properties: { title: SHEET_NAME } } }]
      }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [headers] },
    });
    return { rowCount: 1 };
  }
  return { rowCount: (sheet.properties.gridProperties.rowCount || 1) };
}

async function getNextSeq() {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:A`,
  });
  const vals = res.data.values || [];
  if (vals.length <= 1) return 1;
  const last = parseInt(vals[vals.length - 1][0]);
  return isNaN(last) ? vals.length : last + 1;
}

async function appendRecord(type, date, person, location, work, hours, photoUrl, note) {
  const seq = await getNextSeq();
  const now = new Date().toISOString().slice(0, 19).replace('T', ' ');
  const sheets = await getSheetsClient();
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:J`,
    valueInputOption: 'RAW',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[seq, type, date, person, location, work, hours || '', photoUrl || '', note || '', now]],
    },
  });
  return seq;
}

async function appendPhotoToRecord(rowNum, photoUrl) {
  const sheets = await getSheetsClient();
  const range = `${SHEET_NAME}!I${rowNum}`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
  });
  const old = (res.data.values && res.data.values[0] && res.data.values[0][0]) || '';
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'RAW',
    requestBody: { values: [[old ? old + '\n' + photoUrl : photoUrl]] },
  });
}

async function findTodayAssignment(person, dateStr) {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:J`,
  });
  const rows = res.data.values || [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    if (String(r[1]) === '派工' && String(r[2]) === dateStr && String(r[3]) === person) {
      return { row: i + 1, location: String(r[4] || ''), work: String(r[5] || '') };
    }
  }
  return null;
}

async function getEmployeeList() {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${EMPLOYEE_SHEET}!A:A`,
  });
  const rows = res.data.values || [];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const name = (rows[i][0] || '').trim();
    if (name) out.push(name);
  }
  return out.length > 0 ? out : ['全體'];
}

function matchEmployeeName(lineName, employees) {
  for (const e of employees) { if (e === lineName) return e; }
  for (const e of employees) { if (lineName.includes(e)) return e; }
  for (const e of employees) { if (e.includes(lineName)) return e; }
  return null;
}

const DAY_MAP = { '日':0, '一':1, '二':2, '三':3, '四':4, '五':5, '六':6 };

function weekdayToDateStr(weekday) {
  const key = weekday.replace(/^(星期|週|周)/, '');
  const target = DAY_MAP[key];
  if (target === undefined) return todayStr();
  const now = new Date();
  const cur = now.getDay();
  let diff = target - cur;
  if (diff < 0) diff += 7;
  const d = new Date(now.getTime() + diff * 86400000);
  return dateStr(d);
}

function dateStr(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}

function todayStr() {
  return dateStr(new Date());
}

function mmddStr(d) {
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return m + day;
}

async function saveToDrive(buffer, fileName, mimeType) {
  const auth = await getAuth();
  const drive = google.drive({ version: 'v3', auth });
  const fileMetadata = { name: fileName };
  if (DRIVE_FOLDER_ID) fileMetadata.parents = [DRIVE_FOLDER_ID];
  const stream = new Readable();
  stream.push(buffer);
  stream.push(null);
  const response = await drive.files.create({
    requestBody: fileMetadata,
    media: { mimeType: mimeType || 'image/jpeg', body: stream },
    fields: 'id',
  });
  await drive.permissions.create({
    fileId: response.data.id,
    requestBody: { role: 'reader', type: 'anyone' },
  });
  return `https://drive.google.com/uc?id=${response.data.id}`;
}

async function handleDispatch(text, source, replyToken) {
  const lines = text.split('\n');
  const parsedAll = [];
  let inheritedDay = null;
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const result = parseDispatchLine(line, inheritedDay);
    if (!result) { continue; }
    inheritedDay = result.weekday;
    let names = result.names;
    if (names.length === 0) {
      names = await getEmployeeList();
    }
    for (const name of names) {
      await appendRecord('派工', result.dateStr, name, result.location, result.work, '', '', '');
    }
    const displayNames = result.names.length > 0 ? result.names.join('.') : '全體';
    parsedAll.push(`${result.dateStr} ${result.location ? result.location + '/' : ''}${result.work} → ${displayNames}`);
  }
  const reply = '✅ 已記錄派工 ' + parsedAll.length + ' 筆：\n' + parsedAll.join('\n');
  await replyMessage(replyToken, reply);
}

function parseDispatchLine(line, inheritedDay) {
  let text = line;
  const dayMatch = text.match(/^(星期[一二三四五六日日]|週[一二三四五六日日])/);
  let weekday = null;
  if (dayMatch) {
    weekday = dayMatch[1];
    text = text.substring(dayMatch[0].length).trim();
  } else {
    weekday = inheritedDay;
  }
  if (!weekday) return null;
  const names = [];
  const re = /@([^\s@]+)/g;
  let m;
  while ((m = re.exec(text)) !== null) names.push(m[1]);
  text = text.replace(/@[^\s@]+/g, '').trim();
  let location = '';
  let work = text;
  const si = text.indexOf(' ');
  if (si > -1) {
    location = text.substring(0, si).trim();
    work = text.substring(si + 1).trim();
  }
  return { weekday, dateStr: weekdayToDateStr(weekday), names, location, work };
}

async function handleLeave(text, replyToken) {
  const lines = text.split('\n');
  const details = [];
  const today = todayStr();
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const parts = line.split(/\s+/);
    const name = parts[0] || '';
    const reason = parts.slice(1).join(' ') || '請假';
    if (name) {
      await appendRecord('請假', today, name, '', '', '', '', reason);
      details.push(name + '(' + reason + ')');
    }
  }
  const reply = details.length > 0
    ? '✅ 已記錄請假 ' + details.length + ' 人：' + details.join('、')
    : '⚠️ 請假格式：\n#請假\n小名 事假';
  await replyMessage(replyToken, reply);
}

async function handleOT(text, replyToken) {
  const lines = text.split('\n');
  const details = [];
  const today = todayStr();
  for (let i = 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    const parts = line.split(/\s+/);
    const name = parts[0] || '';
    const hours = parts[1] || '';
    if (name && hours) {
      await appendRecord('加班', today, name, '', '', hours, '', '');
      details.push(name + ' ' + hours + 'H');
    }
  }
  const reply = details.length > 0
    ? '✅ 已記錄加班 ' + details.length + ' 筆：' + details.join('、')
    : '⚠️ 加班格式：\n#加班\n阿豪 4';
  await replyMessage(replyToken, reply);
}

async function handleReportText(source, replyToken) {
  const userId = source.userId;
  reportMode[userId] = Date.now();
  await replyMessage(replyToken, '📋 已開啟回報模式，請傳送照片。');
}

async function handleReportPhoto(messageId, source, replyToken) {
  const userId = source.userId;
  const ts = reportMode[userId];
  if (!ts) return;
  if (Date.now() - ts > REPORT_EXPIRY_MS) {
    delete reportMode[userId];
    return;
  }
  const displayName = await getLineDisplayName(userId);
  const employees = await getEmployeeList();
  const empName = matchEmployeeName(displayName, employees) || displayName;
  const today = todayStr();
  const assignment = await findTodayAssignment(empName, today);
  const site = assignment ? assignment.location : '';
  const work = assignment ? assignment.work : '';
  const { data: buffer, contentType } = await downloadContent(messageId);
  const datePart = mmddStr(new Date());
  const seq = getNextSeq(site || '未指派', empName, datePart);
  const fileName = (site || '未指派') + '_' + empName + '_' + datePart + '_' + seq + '.jpg';
  const photoUrl = await saveToDrive(buffer, fileName, contentType);
  if (assignment) {
    await appendPhotoToRecord(assignment.row, photoUrl);
    await replyMessage(replyToken,
      '📷 收到回報 ' + empName + '\n案場：' +
      (site ? site + (work ? '/' + work : '') : work) +
      '\n→ ' + fileName);
  } else {
    await appendRecord('回報', today, empName, site, work, '', photoUrl, '無對應派工');
    await replyMessage(replyToken,
      '📷 收到 ' + empName + ' 照片\n⚠️ 今日無對應派工紀錄\n→ ' + fileName);
  }
}

const photoCounters = {};

function getNextSeq(site, person, date) {
  const key = site + '_' + person + '_' + date;
  photoCounters[key] = (photoCounters[key] || 0) + 1;
  return photoCounters[key];
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('LINE bot running on port ' + PORT));

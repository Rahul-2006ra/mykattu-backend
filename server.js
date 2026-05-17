// =====================================================================
// MYKATTU GYM — Backend Server
// =====================================================================
// Secrets are loaded from .env file — never hardcoded here.
// This file is safe to share / push to GitHub.
// =====================================================================

require('dotenv').config();   // loads .env file automatically

const express    = require('express');
const cors       = require('cors');
const nodemailer = require('nodemailer');
const ExcelJS    = require('exceljs');
const path       = require('path');
const fs         = require('fs');
const https      = require('https');

const app  = express();
const PORT = process.env.PORT || 3001;

// ── SECRETS (loaded from .env — never hardcoded) ──────────────────────
const OWNER_EMAIL       = process.env.OWNER_EMAIL;
const GMAIL_USER        = process.env.GMAIL_USER;
const GMAIL_APP_PASS    = process.env.GMAIL_APP_PASS;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
// ─────────────────────────────────────────────────────────────────────

const EXCEL_FILE = path.join(__dirname, 'mykattu_leads.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

// ── NODEMAILER SETUP ──────────────────────────────────────────────────
const transporter = nodemailer.createTransport({
  host:   'smtp.gmail.com',
  port:   465,
  secure: true,
  auth: {
    user: GMAIL_USER,
    pass: GMAIL_APP_PASS,
  },
  tls: {
    rejectUnauthorized: false,
  },
});

transporter.verify((error) => {
  if (error) {
    console.error('❌ Email setup error:', error.message);
    console.error('   → Check GMAIL_USER and GMAIL_APP_PASS in your .env file');
  } else {
    console.log('✅ Email ready — Gmail connected');
  }
});

// ── EXCEL HELPER ──────────────────────────────────────────────────────
const HEADERS = [
  'S.No', 'Date', 'Time', 'First Name', 'Last Name',
  'Phone', 'Email', 'Fitness Goal', 'Preferred Plan', 'Message', 'Source'
];

async function appendToExcel(lead) {
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(EXCEL_FILE)) await workbook.xlsx.readFile(EXCEL_FILE);

  let sheet = workbook.getWorksheet('Leads');
  if (!sheet) {
    sheet = workbook.addWorksheet('Leads');
    const hr = sheet.addRow(HEADERS);
    hr.eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0099FF' } };
      cell.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FF0077CC' } } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    hr.height = 22;
    [6, 12, 10, 14, 14, 16, 28, 20, 20, 36, 14].forEach((w, i) => {
      sheet.getColumn(i + 1).width = w;
    });
  }

  const rowNum  = sheet.rowCount;
  const now     = new Date();
  const dateStr = now.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  const timeStr = now.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', hour12: true });

  const dr = sheet.addRow([
    rowNum, dateStr, timeStr,
    lead.fname   || '',
    lead.lname   || '',
    lead.phone   || '',
    lead.email   || '',
    lead.goal    || '',
    lead.plan    || '',
    lead.message || '',
    lead.source  || 'Website Form',
  ]);

  const isEven = (rowNum % 2 === 0);
  dr.eachCell(cell => {
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEven ? 'FFF0F8FF' : 'FFFFFFFF' } };
    cell.border    = { bottom: { style: 'hair', color: { argb: 'FFE2E8F0' } } };
    cell.alignment = { vertical: 'middle', wrapText: true };
  });
  dr.height = 18;

  await workbook.xlsx.writeFile(EXCEL_FILE);
  return rowNum;
}

// ── OWNER ALERT EMAIL ─────────────────────────────────────────────────
async function sendOwnerEmail(lead, sn) {
  await transporter.sendMail({
    from:    `"MYKATTU GYM" <${GMAIL_USER}>`,
    to:      OWNER_EMAIL,
    subject: `🔥 New Lead #${sn}: ${lead.fname} ${lead.lname} — ${lead.plan || 'Interested'}`,
    html: `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:linear-gradient(135deg,#0099ff,#0077cc);padding:24px 32px;border-radius:12px 12px 0 0">
        <h1 style="color:#fff;margin:0;font-size:20px">🏋️ New Lead — MYKATTU GYM</h1>
        <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px">Lead #${sn} • ${new Date().toLocaleString('en-IN')}</p>
      </div>
      <div style="background:#fff;padding:24px 32px;border:1px solid #e2e8f0">
        <table style="width:100%;border-collapse:collapse;font-size:14px">
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff;width:35%">Name</td><td style="padding:10px 14px">${lead.fname} ${lead.lname}</td></tr>
          <tr><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Phone</td><td style="padding:10px 14px">${lead.phone}</td></tr>
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Email</td><td style="padding:10px 14px">${lead.email || '—'}</td></tr>
          <tr><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Goal</td><td style="padding:10px 14px">${lead.goal || '—'}</td></tr>
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Plan</td><td style="padding:10px 14px">${lead.plan || '—'}</td></tr>
          <tr><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Message</td><td style="padding:10px 14px">${lead.message || '—'}</td></tr>
          <tr style="background:#eff6ff"><td style="padding:10px 14px;font-weight:bold;color:#0099ff">Source</td><td style="padding:10px 14px">${lead.source || 'Website Form'}</td></tr>
        </table>
      </div>
      <div style="background:#eff6ff;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center">
        <p style="margin:0;font-size:13px;color:#334155">⚡ Contact this lead within <strong>1 hour</strong> for best results. Lead saved to Excel.</p>
      </div>
    </div>`,
  });
}

// ── USER THANK-YOU EMAIL ──────────────────────────────────────────────
async function sendUserEmail(lead) {
  if (!lead.email) return;
  await transporter.sendMail({
    from:    `"MYKATTU GYM" <${GMAIL_USER}>`,
    to:      lead.email,
    subject: `🏋️ Thanks for your interest in MYKATTU GYM — We'll contact you within 12 hrs!`,
    html: `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
      <div style="background:linear-gradient(135deg,#0099ff,#0077cc);padding:36px 32px;text-align:center;border-radius:12px 12px 0 0">
        <div style="font-size:48px">🏋️</div>
        <h1 style="color:#fff;margin:8px 0 0;font-size:24px">Welcome to MYKATTU!</h1>
        <p style="color:rgba(255,255,255,0.85);margin:6px 0 0">Bengaluru's #1 AI-Powered Gym</p>
      </div>
      <div style="background:#fff;padding:32px;border:1px solid #e2e8f0">
        <p style="font-size:16px;color:#0f172a">Hi <strong>${lead.fname}</strong>,</p>
        <p style="font-size:15px;color:#334155;line-height:1.7">
          Thank you for showing interest in MYKATTU GYM! 🙌<br><br>
          Our team will connect with you within <strong>12 hours</strong> to discuss your fitness goals and kick-start your transformation.
        </p>
        <div style="background:#eff6ff;border-radius:10px;padding:20px 24px;margin:20px 0">
          <p style="font-size:13px;font-weight:bold;color:#0099ff;margin:0 0 10px;text-transform:uppercase;letter-spacing:1px">Your Details</p>
          <p style="font-size:14px;color:#334155;margin:4px 0">📋 <strong>Goal:</strong> ${lead.goal || 'General Fitness'}</p>
          <p style="font-size:14px;color:#334155;margin:4px 0">💎 <strong>Plan:</strong> ${lead.plan || 'To be discussed'}</p>
        </div>
        <p style="font-size:14px;color:#334155;line-height:1.8">
          📍 State Highway 39, opp. St. Anne's High School, T.B Cross, Hesaraghatta, Bengaluru 560088<br>
          📞 +91 98765 43210 (WhatsApp available)<br>
          🕐 Mon–Fri: 5 AM–11 PM | Sat: 6 AM–10 PM | Sun: 7 AM–8 PM
        </p>
        <div style="text-align:center;margin-top:24px">
          <a href="https://wa.me/919876543210" style="display:inline-block;background:linear-gradient(135deg,#0099ff,#0077cc);color:#fff;text-decoration:none;padding:14px 32px;border-radius:10px;font-weight:bold">💬 Chat on WhatsApp</a>
        </div>
      </div>
      <div style="background:#eff6ff;padding:16px 32px;border-radius:0 0 12px 12px;text-align:center">
        <p style="margin:0;font-size:13px;color:#334155">Follow us: <strong>@MYKATTUGym</strong> on Instagram & YouTube</p>
        <p style="margin:4px 0 0;font-size:12px;color:#94a3b8">© MYKATTU GYM, Hesaraghatta, Bengaluru</p>
      </div>
    </div>`,
  });
}

// ── API: SUBMIT LEAD FORM ─────────────────────────────────────────────
app.post('/api/lead', async (req, res) => {
  try {
    const lead = req.body;
    if (!lead.fname || !lead.phone) {
      return res.status(400).json({ success: false, message: 'Name and phone are required.' });
    }

    const sn = await appendToExcel(lead);
    console.log(`✅ Lead #${sn} saved: ${lead.fname} ${lead.lname} (${lead.phone})`);

    const [ownerR, userR] = await Promise.allSettled([
      sendOwnerEmail(lead, sn),
      sendUserEmail(lead),
    ]);

    if (ownerR.status === 'rejected') console.error('❌ Owner email:', ownerR.reason?.message);
    else console.log('✅ Owner email sent');

    if (userR.status === 'rejected') console.error('❌ User email:', userR.reason?.message);
    else if (lead.email) console.log('✅ User email sent to', lead.email);

    res.json({ success: true, message: 'Lead saved!', serialNo: sn });

  } catch (err) {
    console.error('Lead error:', err);
    res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
});

// ── API: CLAUDE AI CHAT PROXY ─────────────────────────────────────────
app.post('/api/chat', (req, res) => {
  try {
    const { messages } = req.body;
    if (!messages || !Array.isArray(messages)) {
      return res.status(400).json({ error: 'messages array required' });
    }

    if (!ANTHROPIC_API_KEY || ANTHROPIC_API_KEY === 'YOUR_ANTHROPIC_API_KEY') {
      return res.json({ reply: "AI chat needs setup! Add your Anthropic API key to the .env file. For now call us: +91 98765 43210! 💪" });
    }

    const bodyData = JSON.stringify({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 500,
      system: "You are MYKATTU AI Coach for MYKATTU GYM in Hesaraghatta, Bengaluru. " +
              "Personality: friendly energetic gym coach, call user 'bro'/'boss'/'champ', keep replies short (2-5 sentences), use fitness emojis. " +
              "You are expert in: workouts, fat loss, muscle building, diet plans, macros, protein, supplements (whey/creatine/BCAA), " +
              "training programs (PPL/full body/5x5), HIIT, injury prevention, recovery, motivation. " +
              "MYKATTU plans: Silver ₹999/mo (basic), Gold ₹1999/mo (all equipment+classes), Platinum ₹3499/mo (PT+diet+all access), Annual ₹29999/yr. " +
              "Address: State Highway 39, opp. St. Anne's High School, T.B Cross, Hesaraghatta, Bengaluru 560088. " +
              "Phone: +91 98765 43210. Timings: Mon-Fri 5AM-11PM, Sat 6AM-10PM, Sun 7AM-8PM. " +
              "If user wants to join, suggest booking a free trial. Answer all fitness questions confidently.",
      messages: messages,
    });

    const options = {
      hostname: 'api.anthropic.com',
      path:     '/v1/messages',
      method:   'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'Content-Length':    Buffer.byteLength(bodyData),
      },
    };

    const apiReq = https.request(options, (apiRes) => {
      let data = '';
      apiRes.on('data', chunk => { data += chunk; });
      apiRes.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.content && parsed.content[0]?.text) {
            res.json({ reply: parsed.content[0].text });
          } else {
            console.error('Anthropic response:', data);
            res.json({ reply: "Bro, AI is having a moment 😅 Try again! 💪" });
          }
        } catch (e) {
          res.json({ reply: "AI response error. Try again! 💪" });
        }
      });
    });

    apiReq.on('error', (e) => {
      console.error('Anthropic request error:', e.message);
      res.json({ reply: "Can't reach AI right now. Check internet and try again! 💪" });
    });

    apiReq.write(bodyData);
    apiReq.end();

  } catch (err) {
    console.error('Chat route error:', err);
    res.json({ reply: "Something went wrong. Try again! 💪" });
  }
});

// ── HEALTH CHECK ──────────────────────────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// ── START SERVER ──────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅ MYKATTU Backend running → http://localhost:${PORT}`);
  console.log(`📊 Excel file: ${EXCEL_FILE}`);
  console.log(`📧 Owner email: ${OWNER_EMAIL}`);
  console.log(`🤖 AI Chat: ${ANTHROPIC_API_KEY && ANTHROPIC_API_KEY !== 'YOUR_ANTHROPIC_API_KEY' ? 'Claude ready ✨' : '⚠️  Add ANTHROPIC_API_KEY to .env'}\n`);
});

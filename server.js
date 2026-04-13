const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Admin credentials
const ADMIN_USER = process.env.ADMIN_USER || 'emad';
const ADMIN_PASS = process.env.ADMIN_PASS || 'changan2026';

// Store latest data in memory + file
const DATA_FILE = path.join(__dirname, 'latest_data.json');
let latestData = null;

// Load existing data if available
if (fs.existsSync(DATA_FILE)) {
  try {
    latestData = JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
    console.log('Loaded existing data from file');
  } catch(e) { console.log('No existing data'); }
}

app.use(cors());
app.use(express.json());

// Multer for file upload
const storage = multer.memoryStorage();
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

// ── Auth endpoint ──
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_USER && password === ADMIN_PASS) {
    res.json({ success: true, token: 'admin-' + Date.now() });
  } else {
    res.status(401).json({ success: false, message: 'Invalid credentials' });
  }
});

// ── Get latest data ──
app.get('/api/data', (req, res) => {
  if (!latestData) {
    return res.json({ success: false, message: 'No data uploaded yet' });
  }
  res.json({ success: true, data: latestData });
});

// ── Upload Excel ──
app.post('/api/upload', upload.single('excel'), (req, res) => {
  try {
    // Simple token check
    const auth = req.headers.authorization || '';
    if (!auth.startsWith('admin-')) {
      return res.status(401).json({ success: false, message: 'Unauthorized' });
    }

    if (!req.file) {
      return res.status(400).json({ success: false, message: 'No file uploaded' });
    }

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });

    // Parse Day Results (2)
    const shDay = wb.Sheets['Day Results (2)'];
    const dayRows = shDay ? XLSX.utils.sheet_to_json(shDay, { header: 1, defval: null }) : [];

    let budget = 0, actual = 0, mtdPct = 0, runRate = 0, runAmt = 0;
    let branches = { alrai: 0, pdi: 0, jahra: 0, ahmadi: 0 };
    let budgets = { alrai: 0, pdi: 0, jahra: 0, ahmadi: 0 };
    let throughput = { ws: 0, qs: 0, pdi: 0, jahra: 0, ahmadi: 0, total: 0 };
    let expectedMTD = 0;

    // Read V2 cell
    if (shDay && shDay['V2']) expectedMTD = shDay['V2'].v || 0;

    dayRows.forEach(r => {
      if (!r || !r[0]) return;
      const lbl = typeof r[0] === 'string' ? r[0].trim() : '';
      if (lbl === 'Budget') {
        budgets = { alrai: r[1]||0, pdi: r[2]||0, jahra: r[3]||0, ahmadi: r[4]||0 };
        budget = r[5] || 0;
      }
      if (lbl === 'Actual Sales') {
        branches = { alrai: r[1]||0, pdi: r[2]||0, jahra: r[3]||0, ahmadi: r[4]||0 };
        actual = r[5] || 0;
      }
      if (lbl === 'MTD %') mtdPct = r[5] || 0;
      if (lbl === 'Run Rate') runRate = r[5] || 0;
      if (lbl === 'Run Amount KD') runAmt = r[5] || 0;
      if (lbl === 'Cars Served') {
        throughput = { ws: r[1]||0, qs: r[2]||0, pdi: r[3]||0, jahra: r[4]||0, ahmadi: r[5]||0, total: r[6]||0 };
      }
    });

    // Parse SA sheet
    const shSA = wb.Sheets['SA'];
    const saRows = shSA ? XLSX.utils.sheet_to_json(shSA, { header: 1 }) : [];
    const saData = [];
    const saBranchMap = { 'Al-Rai': [], 'Jahra': [], 'Ahmadi': [] };
    let curBranch = 'Al-Rai', inSA = false, nc = 0;

    saRows.forEach(r => {
      if (!r || !r[0]) return;
      const s = typeof r[0] === 'string' ? r[0].trim() : '';
      if (s.includes('Branch')) {
        curBranch = s.includes('Rai') ? 'Al-Rai' : s.includes('Jahra') ? 'Jahra' : 'Ahmadi';
        inSA = false; nc = 0; return;
      }
      const sn = s.toLowerCase().replace(/[^a-z]/g, '');
      if (sn.startsWith('saname')) {
        const ns = r.slice(1).find(v => typeof v === 'string') || '';
        inSA = ns.toLowerCase().includes('ros');
        nc = 0; return;
      }
      if (inSA && s.length > 2 && typeof r[nc+1] === 'number' && r[nc+1] > 0) {
        const sa = {
          name: s, branch: curBranch,
          ros: r[nc+1]||0, labor: r[nc+4]||0, parts: r[nc+5]||0,
          budget: r[nc+8]||0, pct: r[nc+9] ? parseFloat((r[nc+9]*100).toFixed(1)) : 0, wip: 0
        };
        saData.push(sa);
        if (saBranchMap[curBranch]) saBranchMap[curBranch].push(s);
      }
    });

    // Parse Productivity
    const shProd = wb.Sheets['Productivity'];
    const prodRows = shProd ? XLSX.utils.sheet_to_json(shProd, { header: 1 }) : [];
    const prod = [];
    const techsByBranch = { 'Al-Rai': [], 'Jahra': [], 'Ahmadi': [] };
    const branchOrder = ['Al-Rai', 'Jahra', 'Ahmadi'];
    let sectionIdx = -1, curSection = 'Al-Rai';
    const summaryRows = [];

    prodRows.forEach(r => {
      if (!r) return;
      if (r[0] === 'Code' && typeof r[1] === 'string' && r[1].trim() === 'Technician') {
        sectionIdx++; curSection = branchOrder[sectionIdx] || curSection; return;
      }
      if (typeof r[0] === 'number' && r[0] > 1000 && typeof r[1] === 'string') {
        const tech = { code: r[0], name: r[1].trim(), sold: r[2]||0, taken: r[3]||0, eff: parseFloat((r[4]||0).toFixed(1)), util: parseFloat((r[8]||0).toFixed(1)), prod: parseFloat((r[9]||0).toFixed(1)) };
        if (techsByBranch[curSection]) techsByBranch[curSection].push(tech);
      }
      if ((r[0]===null||r[0]===undefined) && typeof r[2]==='number' && r[2]>0 && typeof r[9]==='number' && r[9]>0) {
        summaryRows.push({ branch: curSection, prod: parseFloat(r[9].toFixed(1)), eff: parseFloat((r[4]||0).toFixed(1)), util: parseFloat((r[8]||0).toFixed(1)), sold: parseFloat(r[2].toFixed(2)), taken: parseFloat((r[3]||0).toFixed(2)), techs: techsByBranch[curSection]?.length || 0 });
      }
    });
    const branchMap = {};
    summaryRows.forEach(s => { if (!branchMap[s.branch]) branchMap[s.branch] = s; });
    branchOrder.forEach(b => { if (branchMap[b]) prod.push(branchMap[b]); });

    // Parse WIP
    const shWip = wb.Sheets['Pivot WIP per Adv'];
    const wipData = [];
    let alraiWip = 0, jahraWip = 0, ahmadiWip = 0;

    if (shWip) {
      const wipRows = XLSX.utils.sheet_to_json(shWip, { header: 1 });
      wipRows.forEach(r => {
        if (!r || !r[0]) return;
        if (r[0] === 'Grand Total') alraiWip = r[1] || 0;
        else if (typeof r[0] === 'string' && r[0] !== 'Row Labels') {
          wipData.push({ name: r[0], wip: r[1]||0, branch: 'Al-Rai' });
          saData.forEach(sa => { if (sa.name.substring(0,8) === r[0].substring(0,8)) sa.wip = r[1]||0; });
        }
      });
    }
    const shJahra = wb.Sheets['Jahra WIP'];
    if (shJahra) { jahraWip = XLSX.utils.sheet_to_json(shJahra).length; }
    const shAhmadi = wb.Sheets['Ahmadi WIP'];
    if (shAhmadi) { ahmadiWip = XLSX.utils.sheet_to_json(shAhmadi).length; }

    // Parse All Wips
    const shAllWips = wb.Sheets['All Wips'];
    const allWips = [];
    if (shAllWips) {
      const allWipRows = XLSX.utils.sheet_to_json(shAllWips, { header: 1, defval: '' });
      allWipRows.slice(1).forEach(r => {
        if (!r[0] || typeof r[0] !== 'number') return;
        let dueIn = '—';
        if (r[10]) {
          if (typeof r[10] === 'string') dueIn = r[10];
          else if (typeof r[10] === 'number') dueIn = new Date(Math.round((r[10]-25569)*86400*1000)).toLocaleDateString('en-GB');
        }
        allWips.push({ wipNo: r[0], sa: r[4]||'', reg: r[5]||'', model: r[7]||'', dueIn, inOut: r[14]||'', ageing: r[15]||0, status: r[16]||'' });
      });
    }

    // Build final data object
    latestData = {
      updatedAt: new Date().toISOString(),
      expectedMTD,
      mtd: { actual, budget, mtdPct, runRate, runAmt, branches, budgets },
      throughput,
      saData,
      saBranchMap,
      prod,
      techsByBranch,
      wipSummary: { alrai: alraiWip, jahra: jahraWip, ahmadi: ahmadiWip, total: alraiWip + jahraWip + ahmadiWip },
      wipData,
      allWips
    };

    // Save to file
    fs.writeFileSync(DATA_FILE, JSON.stringify(latestData));
    console.log('Data updated at', latestData.updatedAt);
    res.json({ success: true, message: 'Data updated successfully', updatedAt: latestData.updatedAt });

  } catch(err) {
    console.error('Upload error:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/', (req, res) => res.json({ status: 'Changan Dashboard API running', updatedAt: latestData?.updatedAt || 'No data yet' }));

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

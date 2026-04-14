const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

const ADMIN_USER = process.env.ADMIN_USER || 'emad';
const ADMIN_PASS = process.env.ADMIN_PASS || 'changan2026';

const DATA_FILE = path.join(__dirname, 'latest_data.json');
let latestData = null;

if (fs.existsSync(DATA_FILE)) {
  try { latestData = JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8')); } catch(e) {}
}

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } });

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_USER && password === ADMIN_PASS) {
    res.json({ success: true, token: 'admin-' + Date.now() });
  } else {
    res.status(401).json({ success: false, message: 'Invalid credentials' });
  }
});

app.get('/api/data', (req, res) => {
  if (!latestData) return res.json({ success: false, message: 'No data uploaded yet' });
  res.json({ success: true, data: latestData });
});

app.post('/api/upload', upload.single('excel'), (req, res) => {
  try {
    const auth = req.headers.authorization || '';
    if (!auth.startsWith('admin-')) return res.status(401).json({ success: false, message: 'Unauthorized' });
    if (!req.file) return res.status(400).json({ success: false, message: 'No file uploaded' });

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });

    // ── 1. Day Results (2) ──
    const shDay = wb.Sheets['Day Results (2)'];
    const dayRows = shDay ? XLSX.utils.sheet_to_json(shDay, { header: 1, defval: null }) : [];

    let budget=0, actual=0, mtdPct=0, runRate=0, runAmt=0;
    let branches = { alrai:0, pdi:0, jahra:0, ahmadi:0 };
    let budgets  = { alrai:0, pdi:0, jahra:0, ahmadi:0 };
    let throughput = { ws:0, qs:0, pdi:0, jahra:0, ahmadi:0, total:0 };
    let expectedMTD = 0;

    if (shDay && shDay['V2']) expectedMTD = shDay['V2'].v || 0;

    dayRows.forEach(r => {
      if (!r || !r[0]) return;
      const lbl = typeof r[0]==='string' ? r[0].trim() : '';
      // Cols: r[0]=label, r[1]=AlRaiWS&QS, r[2]=PDI, r[3]=Jahra, r[4]=Ahmadi, r[5]=Total
      if (lbl==='Budget')       { budgets={alrai:r[1]||0,pdi:r[2]||0,jahra:r[3]||0,ahmadi:r[4]||0}; budget=r[5]||0; }
      if (lbl==='Actual Sales') { branches={alrai:r[1]||0,pdi:r[2]||0,jahra:r[3]||0,ahmadi:r[4]||0}; actual=r[5]||0; }
      if (lbl==='MTD %')        { mtdPct=r[5]||0; }
      if (lbl==='Run Rate')     { runRate=r[5]||0; }
      if (lbl==='Run Amount KD'){ runAmt=r[5]||0; }
      // Throughput: r[0]=label, r[1]=WS, r[2]=QS, r[3]=PDI, r[4]=Jahra, r[5]=Ahmadi, r[6]=Total
      if (lbl==='Cars Served')  { throughput={ws:r[1]||0,qs:r[2]||0,pdi:r[3]||0,jahra:r[4]||0,ahmadi:r[5]||0,total:r[6]||0}; }
    });

    // ── 2. SA Sheet ──
    // Structure: col[5]=name, col[6]=ROs, col[7]=RO/Day, col[8]=pendingWIP,
    //            col[9]=laborOnly, col[10]=totalLabor, col[11]=parts,
    //            col[12]=hoursSold, col[13]=totalWithParts, col[14]=budget, col[15]=mtdPct
    const shSA = wb.Sheets['SA'];
    const saRows = shSA ? XLSX.utils.sheet_to_json(shSA, { header: 1, defval: null }) : [];
    const saData = [];
    const saBranchMap = { 'Al-Rai':[], 'Jahra':[], 'Ahmadi':[] };
    let curBranch = 'Al-Rai';

    saRows.forEach(r => {
      if (!r || !r[5]) return;
      const s = typeof r[5]==='string' ? r[5].trim() : '';
      if (!s) return;

      // Branch header
      if (s.includes('Branch')) {
        if (s.includes('Jahra'))  curBranch = 'Jahra';
        else if (s.includes('Ahmadi')) curBranch = 'Ahmadi';
        else curBranch = 'Al-Rai';
        return;
      }
      // Skip header row
      if (s === 'SA Name' || s === 'SA Name ') return;

      // SA data row - must have ROs (col 6) as a number > 0
      if (typeof r[6]==='number' && r[6] > 0) {
        const sa = {
          name:   s,
          branch: curBranch,
          ros:    r[6]  || 0,
          labor:  r[10] || 0,
          parts:  r[11] || 0,
          budget: r[14] || 0,
          pct:    r[15] ? parseFloat((r[15]*100).toFixed(1)) : 0,
          wip:    r[8]  || 0
        };
        saData.push(sa);
        if (saBranchMap[curBranch]) saBranchMap[curBranch].push(s);
      }
    });

    // ── 3. Productivity ──
    const shProd = wb.Sheets['Productivity'];
    const prodRows = shProd ? XLSX.utils.sheet_to_json(shProd, { header: 1, defval: null }) : [];
    const prod = [];
    const techsByBranch = { 'Al-Rai':[], 'Jahra':[], 'Ahmadi':[] };
    const branchOrder = ['Al-Rai','Jahra','Ahmadi'];
    let sectionIdx = -1, curSection = 'Al-Rai';
    const summaryMap = {};

    prodRows.forEach(r => {
      if (!r) return;
      if (typeof r[0]==='string' && r[0]==='Code' && typeof r[1]==='string' && r[1].trim()==='Technician') {
        sectionIdx++; curSection = branchOrder[sectionIdx] || curSection; return;
      }
      if (typeof r[0]==='number' && r[0]>1000 && typeof r[1]==='string' && r[1].trim()) {
        if (techsByBranch[curSection]) techsByBranch[curSection].push({
          code:r[0], name:r[1].trim(), sold:r[2]||0, taken:r[3]||0,
          eff:parseFloat((r[4]||0).toFixed(1)), util:parseFloat((r[8]||0).toFixed(1)),
          prod:parseFloat((r[9]||0).toFixed(1))
        });
      }
      if ((r[0]===null||r[0]===undefined) && typeof r[2]==='number' && r[2]>0 && typeof r[9]==='number') {
        if (!summaryMap[curSection]) summaryMap[curSection] = {
          branch:curSection, prod:parseFloat(r[9].toFixed(1)), eff:parseFloat((r[4]||0).toFixed(1)),
          util:parseFloat((r[8]||0).toFixed(1)), sold:parseFloat(r[2].toFixed(2)),
          taken:parseFloat((r[3]||0).toFixed(2)), techs:techsByBranch[curSection]?.length||0
        };
      }
    });
    branchOrder.forEach(b => { if (summaryMap[b]) prod.push(summaryMap[b]); });

    // ── 4. WIP ──
    const shWip = wb.Sheets['Pivot WIP per Adv'];
    const wipData = [];
    let alraiWip=0, jahraWip=0, ahmadiWip=0;

    if (shWip) {
      const wipRows = XLSX.utils.sheet_to_json(shWip, { header:1, defval:null });
      wipRows.forEach(r => {
        if (!r || !r[0]) return;
        const lbl = typeof r[0]==='string' ? r[0].trim() : '';
        if (lbl==='Grand Total') { alraiWip = r[1]||0; return; }
        if (lbl && lbl!=='Row Labels' && typeof r[1]==='number') {
          wipData.push({ name:lbl, wip:r[1]||0, branch:'Al-Rai' });
          // Update SA WIP
          saData.forEach(sa => {
            if (sa.branch==='Al-Rai' && (sa.name.substring(0,6)===lbl.substring(0,6))) sa.wip=r[1]||0;
          });
        }
      });
    }

    const shJahra = wb.Sheets['Jahra WIP'];
    if (shJahra) {
      const jr = XLSX.utils.sheet_to_json(shJahra, { header:1, defval:null });
      jahraWip = jr.filter(r => r && typeof r[0]==='number').length;
    }

    const shAhmadi = wb.Sheets['Ahmadi WIP'];
    if (shAhmadi) {
      const ar = XLSX.utils.sheet_to_json(shAhmadi, { header:1, defval:null });
      ahmadiWip = ar.filter(r => r && typeof r[0]==='number').length;
    }

    // ── 5. All Wips ──
    const shAllWips = wb.Sheets['All Wips'];
    const allWips = [];
    if (shAllWips) {
      const allWipRows = XLSX.utils.sheet_to_json(shAllWips, { header:1, defval:'' });
      allWipRows.slice(1).forEach(r => {
        if (!r[0] || typeof r[0]!=='number') return;
        let dueIn = '—';
        if (r[10]) {
          if (typeof r[10]==='string') dueIn=r[10];
          else if (typeof r[10]==='number') dueIn=new Date(Math.round((r[10]-25569)*86400*1000)).toLocaleDateString('en-GB');
        }
        allWips.push({ wipNo:r[0], sa:r[4]||'', reg:r[5]||'', model:r[7]||'', dueIn, inOut:r[14]||'', ageing:r[15]||0, status:r[16]||'' });
      });
    }

    latestData = {
      updatedAt: new Date().toISOString(),
      expectedMTD,
      mtd: { actual, budget, mtdPct, runRate, runAmt, branches, budgets },
      throughput,
      saData,
      saBranchMap,
      prod,
      techsByBranch,
      wipSummary: { alrai:alraiWip, jahra:jahraWip, ahmadi:ahmadiWip, total:alraiWip+jahraWip+ahmadiWip },
      wipData,
      allWips
    };

    fs.writeFileSync(DATA_FILE, JSON.stringify(latestData));
    console.log('✅ Data updated:', latestData.updatedAt, '| SAs:', saData.length, '| WIPs:', allWips.length);
    res.json({ success:true, message:'Data updated!', updatedAt:latestData.updatedAt, saCount:saData.length });

  } catch(err) {
    console.error('Upload error:', err);
    res.status(500).json({ success:false, message:err.message });
  }
});

app.get('/', (req, res) => res.json({ status:'Changan Dashboard API ✅', updatedAt:latestData?.updatedAt||'No data yet' }));
app.listen(PORT, () => console.log(`Server on port ${PORT}`));

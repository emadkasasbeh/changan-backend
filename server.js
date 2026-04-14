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
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50*1024*1024 } });

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_USER && password === ADMIN_PASS) {
    res.json({ success: true, token: 'admin-' + Date.now() });
  } else {
    res.status(401).json({ success: false, message: 'Invalid credentials' });
  }
});

app.get('/api/data', (req, res) => {
  if (!latestData) return res.json({ success: false, message: 'No data yet' });
  res.json({ success: true, data: latestData });
});

// Debug endpoint - shows raw SA rows
app.get('/api/debug', (req, res) => {
  res.json({ lastParse: latestData ? latestData.debugInfo : null });
});

app.post('/api/upload', upload.single('excel'), (req, res) => {
  try {
    const auth = req.headers.authorization || '';
    if (!auth.startsWith('admin-')) return res.status(401).json({ success: false, message: 'Unauthorized' });
    if (!req.file) return res.status(400).json({ success: false, message: 'No file' });

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const debugInfo = { sheets: wb.SheetNames, saRows: [], dayRows: [] };

    // ── Day Results (2) ──
    const shDay = wb.Sheets['Day Results (2)'];
    const dayRows = shDay ? XLSX.utils.sheet_to_json(shDay, { header:1, defval:null }) : [];

    let budget=0, actual=0, mtdPct=0, runRate=0, runAmt=0;
    let branches={alrai:0,pdi:0,jahra:0,ahmadi:0};
    let budgets={alrai:0,pdi:0,jahra:0,ahmadi:0};
    let throughput={ws:0,qs:0,pdi:0,jahra:0,ahmadi:0,total:0};
    let expectedMTD=0;

    if (shDay && shDay['V2']) expectedMTD = shDay['V2'].v || 0;

    // Find the label column by scanning for 'Budget'
    let dayLabelCol = 0;
    for (let r of dayRows) {
      for (let j=0; j<(r||[]).length; j++) {
        if (r[j] === 'Budget') { dayLabelCol = j; break; }
      }
      if (dayLabelCol > 0) break;
    }
    debugInfo.dayLabelCol = dayLabelCol;

    dayRows.forEach(r => {
      if (!r || !r[dayLabelCol]) return;
      const lbl = typeof r[dayLabelCol]==='string' ? r[dayLabelCol].trim() : '';
      const c = dayLabelCol;
      if (lbl==='Budget')        { budgets={alrai:r[c+1]||0,pdi:r[c+2]||0,jahra:r[c+3]||0,ahmadi:r[c+4]||0}; budget=r[c+5]||0; }
      if (lbl==='Actual Sales')  { branches={alrai:r[c+1]||0,pdi:r[c+2]||0,jahra:r[c+3]||0,ahmadi:r[c+4]||0}; actual=r[c+5]||0; }
      if (lbl==='MTD %')         { mtdPct=r[c+5]||0; }
      if (lbl==='Run Rate')      { runRate=r[c+5]||0; }
      if (lbl==='Run Amount KD') { runAmt=r[c+5]||0; }
      if (lbl==='Cars Served')   { throughput={ws:r[c+1]||0,qs:r[c+2]||0,pdi:r[c+3]||0,jahra:r[c+4]||0,ahmadi:r[c+5]||0,total:r[c+6]||0}; }
    });

    // ── SA Sheet ── using sheet_to_json with header:1 and auto-detect columns
    const shSA = wb.Sheets['SA'];
    const saRows = shSA ? XLSX.utils.sheet_to_json(shSA, { header:1, defval:null }) : [];
    const saData = [];
    const saBranchMap = { 'Al-Rai':[], 'Jahra':[], 'Ahmadi':[] };
    let curBranch = 'Al-Rai';

    // Find the column where SA Name header appears
    let saNameCol = -1, saRosCol = -1, saLaborCol = -1, saPartsCol = -1, saBudgetCol = -1, saPctCol = -1, saWipCol = -1;

    // Scan rows to find header
    for (let i=0; i<saRows.length; i++) {
      const r = saRows[i];
      if (!r) continue;
      for (let j=0; j<r.length; j++) {
        if (typeof r[j]==='string' && r[j].trim().toLowerCase().startsWith('sa name')) {
          // Found header row - detect all columns
          saNameCol = j;
          for (let k=j+1; k<r.length; k++) {
            if (!r[k]) continue;
            const h = typeof r[k]==='string' ? r[k].trim().toLowerCase() : '';
            if (h.includes('ros') || h.includes('no of ros')) saRosCol = k;
            else if (h.includes('pending') || h.includes('wip')) saWipCol = k;
            else if (h.includes('total labor') || h==='total labor') saLaborCol = k;
            else if (h.includes('parts')) saPartsCol = k;
            else if (h==='budget') saBudgetCol = k;
            else if (h.includes('mtd')) saPctCol = k;
          }
          debugInfo.saHeaderRow = i;
          debugInfo.cols = { saNameCol, saRosCol, saLaborCol, saPartsCol, saBudgetCol, saPctCol, saWipCol };
          break;
        }
      }
    }

    // If header not found, use known fixed positions from Excel analysis
    if (saNameCol === -1) {
      saNameCol=5; saRosCol=6; saWipCol=8; saLaborCol=10; saPartsCol=11; saBudgetCol=14; saPctCol=15;
      debugInfo.cols = { note:'used fallback', saNameCol, saRosCol, saLaborCol, saPartsCol, saBudgetCol, saPctCol, saWipCol };
    }

    saRows.forEach((r, i) => {
      if (!r) return;
      const nameVal = r[saNameCol];
      if (!nameVal) return;
      const s = typeof nameVal==='string' ? nameVal.trim() : String(nameVal).trim();
      if (!s) return;

      // Branch header
      if (s.toLowerCase().includes('branch')) {
        if (s.toLowerCase().includes('jahra')) curBranch = 'Jahra';
        else if (s.toLowerCase().includes('ahmadi')) curBranch = 'Ahmadi';
        else curBranch = 'Al-Rai';
        return;
      }

      // Skip header rows
      if (s.toLowerCase().startsWith('sa name')) return;

      // Must have ROs as a number
      const ros = r[saRosCol];
      if (typeof ros !== 'number' || ros <= 0) return;

      const pctRaw = r[saPctCol];
      const pct = typeof pctRaw === 'number' ? parseFloat((pctRaw*100).toFixed(1)) : 0;

      const sa = {
        name:   s,
        branch: curBranch,
        ros:    ros,
        labor:  r[saLaborCol]  || 0,
        parts:  r[saPartsCol]  || 0,
        budget: r[saBudgetCol] || 0,
        pct:    pct,
        wip:    r[saWipCol]    || 0
      };
      saData.push(sa);
      if (saBranchMap[curBranch]) saBranchMap[curBranch].push(s);
      debugInfo.saRows.push({ row:i, name:s, branch:curBranch, ros, pct });
    });

    // ── Productivity ──
    const shProd = wb.Sheets['Productivity'];
    const prodRows = shProd ? XLSX.utils.sheet_to_json(shProd, { header:1, defval:null }) : [];
    const prod = [];
    const techsByBranch = { 'Al-Rai':[], 'Jahra':[], 'Ahmadi':[] };
    const branchOrder = ['Al-Rai','Jahra','Ahmadi'];
    let sectionIdx = -1, curSection = 'Al-Rai';
    const summaryMap = {};

    prodRows.forEach(r => {
      if (!r) return;
      // Section header: Code | Technician
      if (typeof r[0]==='string' && r[0].trim()==='Code') {
        sectionIdx++; curSection = branchOrder[sectionIdx] || curSection; return;
      }
      // Tech row: code is big number
      if (typeof r[0]==='number' && r[0]>100 && typeof r[1]==='string' && r[1].trim()) {
        if (techsByBranch[curSection]) techsByBranch[curSection].push({
          code:r[0], name:r[1].trim(),
          sold:r[2]||0, taken:r[3]||0,
          eff:parseFloat((r[4]||0).toFixed(1)),
          util:parseFloat((r[8]||0).toFixed(1)),
          prod:parseFloat((r[9]||0).toFixed(1))
        });
      }
      // Summary row: no code, has sold hours
      if ((r[0]===null||r[0]===undefined) && typeof r[2]==='number' && r[2]>0 && typeof r[9]==='number') {
        if (!summaryMap[curSection]) summaryMap[curSection] = {
          branch:curSection,
          prod:parseFloat(r[9].toFixed(1)),
          eff:parseFloat((r[4]||0).toFixed(1)),
          util:parseFloat((r[8]||0).toFixed(1)),
          sold:parseFloat(r[2].toFixed(2)),
          taken:parseFloat((r[3]||0).toFixed(2)),
          techs:techsByBranch[curSection]?.length||0
        };
      }
    });
    branchOrder.forEach(b => { if (summaryMap[b]) prod.push(summaryMap[b]); });

    // ── WIP ──
    const shWip = wb.Sheets['Pivot WIP per Adv'];
    const wipData = [];
    let alraiWip=0, jahraWip=0, ahmadiWip=0;

    if (shWip) {
      const wipRows = XLSX.utils.sheet_to_json(shWip, { header:1, defval:null });
      wipRows.forEach(r => {
        if (!r || !r[0]) return;
        const lbl = typeof r[0]==='string' ? r[0].trim() : '';
        if (lbl==='Grand Total') { alraiWip=r[1]||0; return; }
        if (lbl && lbl!=='Row Labels' && typeof r[1]==='number') {
          wipData.push({ name:lbl, wip:r[1]||0, branch:'Al-Rai' });
          saData.forEach(sa => { if (sa.branch==='Al-Rai' && sa.name.substring(0,6)===lbl.substring(0,6)) sa.wip=r[1]||0; });
        }
      });
    }
    const shJahra = wb.Sheets['Jahra WIP'];
    if (shJahra) jahraWip = XLSX.utils.sheet_to_json(shJahra,{header:1}).filter(r=>r&&typeof r[0]==='number').length;
    const shAhmadi = wb.Sheets['Ahmadi WIP'];
    if (shAhmadi) ahmadiWip = XLSX.utils.sheet_to_json(shAhmadi,{header:1}).filter(r=>r&&typeof r[0]==='number').length;

    // ── All Wips ──
    const shAllWips = wb.Sheets['All Wips'];
    const allWips = [];
    if (shAllWips) {
      const allWipRows = XLSX.utils.sheet_to_json(shAllWips, { header:1, defval:'' });
      allWipRows.slice(1).forEach(r => {
        if (!r[0]||typeof r[0]!=='number') return;
        let dueIn='—';
        if (r[10]) {
          if (typeof r[10]==='string') dueIn=r[10];
          else if (typeof r[10]==='number') dueIn=new Date(Math.round((r[10]-25569)*86400*1000)).toLocaleDateString('en-GB');
        }
        allWips.push({wipNo:r[0],sa:r[4]||'',reg:r[5]||'',model:r[7]||'',dueIn,inOut:r[14]||'',ageing:r[15]||0,status:r[16]||''});
      });
    }

    latestData = {
      updatedAt: new Date().toISOString(),
      expectedMTD,
      mtd: { actual, budget, mtdPct, runRate, runAmt, branches, budgets },
      throughput, saData, saBranchMap, prod, techsByBranch,
      wipSummary: { alrai:alraiWip, jahra:jahraWip, ahmadi:ahmadiWip, total:alraiWip+jahraWip+ahmadiWip },
      wipData, allWips, debugInfo
    };

    fs.writeFileSync(DATA_FILE, JSON.stringify(latestData));
    console.log(`✅ Updated: SAs=${saData.length}, WIPs=${allWips.length}, Actual=${actual}`);
    res.json({ success:true, message:'Data updated!', saCount:saData.length, actual, updatedAt:latestData.updatedAt });

  } catch(err) {
    console.error('Upload error:', err);
    res.status(500).json({ success:false, message:err.message });
  }
});

app.get('/', (req, res) => res.json({ status:'Changan Dashboard API ✅', updatedAt: latestData?.updatedAt||'No data yet' }));
app.listen(PORT, () => console.log(`Server on port ${PORT}`));

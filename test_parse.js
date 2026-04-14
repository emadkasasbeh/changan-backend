const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('/mnt/user-data/uploads/1776145304503_Dashboard_April_2026.xlsx', {cellDates:true});
const shSA = wb.Sheets['SA'];
const rows = XLSX.utils.sheet_to_json(shSA, {header:1, defval:null});

console.log('=== SA Sheet - first 30 rows (XLSX.js) ===');
rows.slice(0,30).forEach((r,i) => {
  const nonNull = r.map((v,j) => v!==null ? `[${j}]=${JSON.stringify(v)}` : null).filter(Boolean);
  if (nonNull.length) console.log(`Row ${i}: ${nonNull.join(' | ')}`);
});

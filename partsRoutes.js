// partsRoutes.js
// Add this file to your backend and register it in server.js

const express = require('express');
const router = express.Router();

// In-memory store (replace with DB later to persist across deploys)
let partsData = null;

/**
 * GET /api/parts-data
 * Returns the latest uploaded parts dashboard data
 */
router.get('/parts-data', (req, res) => {
  if (!partsData) {
    return res.status(404).json({ error: 'No parts data uploaded yet' });
  }
  res.json(partsData);
});

/**
 * POST /api/parts-upload
 * Receives parsed JSON data from the frontend XLSX parser
 * Body: { overall, trvsach, monthly, dailyTotal, monthlyTotal, stock, comparative, wipAlrai, wipJahra, uploadedAt }
 */
router.post('/parts-upload', express.json({ limit: '20mb' }), (req, res) => {
  try {
    const data = req.body;
    if (!data || typeof data !== 'object') {
      return res.status(400).json({ error: 'Invalid data' });
    }
    partsData = {
      ...data,
      receivedAt: new Date().toISOString()
    };
    console.log(`[Parts] Data received at ${partsData.receivedAt}`);
    res.json({ success: true, receivedAt: partsData.receivedAt });
  } catch (err) {
    console.error('[Parts] Upload error:', err);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;

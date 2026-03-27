const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const { runPowerShell } = require('../utils/powershell');

const SNAPSHOT_PATH = path.join(__dirname, '../data/dashboard-snapshot.json');

const ensureDataDir = () => {
  const dir = path.dirname(SNAPSHOT_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
};

const readSnapshot = () => {
  ensureDataDir();
  if (!fs.existsSync(SNAPSHOT_PATH)) return null;
  try {
    const data = fs.readFileSync(SNAPSHOT_PATH, 'utf8');
    return JSON.parse(data);
  } catch {
    return null;
  }
};

const writeSnapshot = (stats) => {
  ensureDataDir();
  fs.writeFileSync(SNAPSHOT_PATH, JSON.stringify({
    timestamp: new Date().toISOString(),
    ...stats
  }, null, 2));
};

const computeTrends = (current, previous) => {
  if (!previous || !previous.kpis) return {};
  const trends = {};
  const prev = previous.kpis;
  const curr = current.kpis || {};
  if (typeof curr.totalUsers === 'number' && typeof prev.totalUsers === 'number') {
    trends.totalUsersDelta = curr.totalUsers - prev.totalUsers;
  }
  if (typeof curr.disabledAccounts === 'number' && typeof prev.disabledAccounts === 'number') {
    trends.disabledAccountsDelta = curr.disabledAccounts - prev.disabledAccounts;
  }
  if (typeof curr.lockedAccounts === 'number' && typeof prev.lockedAccounts === 'number') {
    trends.lockedAccountsDelta = curr.lockedAccounts - prev.lockedAccounts;
  }
  if (typeof curr.passwordExpiring === 'number' && typeof prev.passwordExpiring === 'number') {
    trends.passwordExpiringDelta = curr.passwordExpiring - prev.passwordExpiring;
  }
  if (typeof curr.inactiveUsers === 'number' && typeof prev.inactiveUsers === 'number') {
    trends.inactiveUsersDelta = curr.inactiveUsers - prev.inactiveUsers;
  }
  return trends;
};

router.get('/stats', async (req, res) => {
  try {
    const previous = readSnapshot();
    const result = await runPowerShell('Get-DashboardStats.ps1', { InactiveDays: 90 }, '');
    const stats = result.data;

    if (!stats || stats.Error) {
      return res.status(500).json({ error: stats?.Error || 'Failed to fetch dashboard stats' });
    }

    const trends = computeTrends(stats, previous);
    writeSnapshot(stats);

    res.json({
      data: stats,
      trends: Object.keys(trends).length > 0 ? trends : null
    });
  } catch (error) {
    console.error('Dashboard stats error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

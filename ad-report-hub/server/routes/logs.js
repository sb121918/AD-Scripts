const express = require('express');
const router = express.Router();
const { getActivityLogs, getLogCount, getUniqueUsers, getUniqueActionTypes } = require('../utils/database');

// GET /api/activity-logs - Get activity logs with filtering
router.get('/', async (req, res) => {
  try {
    const { from, to, actionType, user, status, limit = 50, offset = 0 } = req.query;
    
    const filters = {
      from,
      to,
      actionType,
      user,
      status,
      limit: parseInt(limit),
      offset: parseInt(offset)
    };

    const logs = getActivityLogs(filters);
    const total = getLogCount(filters);

    res.json({
      data: logs,
      total,
      limit: parseInt(limit),
      offset: parseInt(offset)
    });
  } catch (error) {
    console.error('Activity logs error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/activity-logs/users - Get unique users for filter dropdown
router.get('/users', async (req, res) => {
  try {
    const users = getUniqueUsers();
    res.json({ data: users.map(u => u.user_name) });
  } catch (error) {
    console.error('Get users error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/activity-logs/action-types - Get unique action types for filter dropdown
router.get('/action-types', async (req, res) => {
  try {
    const actionTypes = getUniqueActionTypes();
    res.json({ data: actionTypes.map(a => a.action_type) });
  } catch (error) {
    console.error('Get action types error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/activity-logs/export - Export logs
router.post('/export', async (req, res) => {
  try {
    const { from, to, actionType, user, status, format = 'json' } = req.body;
    
    const filters = { from, to, actionType, user, status };
    const logs = getActivityLogs(filters);

    if (format === 'csv') {
      // Convert to CSV format
      const headers = ['ID', 'Timestamp', 'User', 'Action', 'Target', 'Status', 'Duration (ms)', 'IP Address'];
      const rows = logs.map(log => [
        log.id,
        log.timestamp,
        log.user_name,
        log.action_type,
        log.target,
        log.status,
        log.duration_ms,
        log.ip_address
      ]);
      
      const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
      res.setHeader('Content-Type', 'text/csv');
      res.setHeader('Content-Disposition', 'attachment; filename=activity_logs.csv');
      res.send(csv);
    } else {
      res.json({ data: logs });
    }
  } catch (error) {
    console.error('Export logs error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

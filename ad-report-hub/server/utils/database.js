const path = require('path');
const fs = require('fs');

// Ensure data directory exists
const dataDir = path.join(__dirname, '../data');
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}

const logsPath = path.join(dataDir, 'activity-logs.json');

// Initialize logs file if it doesn't exist
if (!fs.existsSync(logsPath)) {
  fs.writeFileSync(logsPath, JSON.stringify({ logs: [], nextId: 1 }, null, 2));
}

// Read logs from file
const readLogs = () => {
  try {
    const data = fs.readFileSync(logsPath, 'utf8');
    return JSON.parse(data);
  } catch (err) {
    return { logs: [], nextId: 1 };
  }
};

// Write logs to file
const writeLogs = (data) => {
  fs.writeFileSync(logsPath, JSON.stringify(data, null, 2));
};

// Helper functions
const logActivity = (userName, actionType, target, details, ipAddress, status, durationMs) => {
  const data = readLogs();
  const newLog = {
    id: data.nextId,
    timestamp: new Date().toISOString(),
    user_name: userName,
    action_type: actionType,
    target: target,
    details: typeof details === 'string' ? details : JSON.stringify(details),
    ip_address: ipAddress,
    status: status || 'Success',
    duration_ms: durationMs
  };
  
  data.logs.unshift(newLog); // Add to beginning
  data.nextId++;
  
  // Keep only last 10000 logs to prevent file from growing too large
  if (data.logs.length > 10000) {
    data.logs = data.logs.slice(0, 10000);
  }
  
  writeLogs(data);
  return { changes: 1, lastInsertRowid: newLog.id };
};

const getActivityLogs = (filters = {}) => {
  const data = readLogs();
  let logs = data.logs;

  // Apply filters
  if (filters.from) {
    const fromDate = new Date(filters.from);
    logs = logs.filter(log => new Date(log.timestamp) >= fromDate);
  }
  if (filters.to) {
    const toDate = new Date(filters.to);
    logs = logs.filter(log => new Date(log.timestamp) <= toDate);
  }
  if (filters.actionType && filters.actionType !== 'all') {
    logs = logs.filter(log => log.action_type === filters.actionType);
  }
  if (filters.user && filters.user !== 'all') {
    logs = logs.filter(log => log.user_name === filters.user);
  }
  if (filters.status && filters.status !== 'all') {
    logs = logs.filter(log => log.status === filters.status);
  }

  // Apply pagination
  if (filters.offset) {
    logs = logs.slice(filters.offset);
  }
  if (filters.limit) {
    logs = logs.slice(0, filters.limit);
  }

  return logs;
};

const getLogCount = (filters = {}) => {
  const data = readLogs();
  let logs = data.logs;

  // Apply filters (same as getActivityLogs but without pagination)
  if (filters.from) {
    const fromDate = new Date(filters.from);
    logs = logs.filter(log => new Date(log.timestamp) >= fromDate);
  }
  if (filters.to) {
    const toDate = new Date(filters.to);
    logs = logs.filter(log => new Date(log.timestamp) <= toDate);
  }
  if (filters.actionType && filters.actionType !== 'all') {
    logs = logs.filter(log => log.action_type === filters.actionType);
  }
  if (filters.user && filters.user !== 'all') {
    logs = logs.filter(log => log.user_name === filters.user);
  }
  if (filters.status && filters.status !== 'all') {
    logs = logs.filter(log => log.status === filters.status);
  }

  return logs.length;
};

const getUniqueUsers = () => {
  const data = readLogs();
  const users = [...new Set(data.logs.map(log => log.user_name))].sort();
  return users.map(user => ({ user_name: user }));
};

const getUniqueActionTypes = () => {
  const data = readLogs();
  const actions = [...new Set(data.logs.map(log => log.action_type))].sort();
  return actions.map(action => ({ action_type: action }));
};

module.exports = {
  logActivity,
  getActivityLogs,
  getLogCount,
  getUniqueUsers,
  getUniqueActionTypes
};

const { logActivity } = require('../utils/database');
const os = require('os');

/**
 * Middleware to log all API activity
 */
const activityLogger = (req, res, next) => {
  // Skip logging for certain paths
  const skipPaths = ['/api/health', '/api/activity-logs'];
  if (skipPaths.some(p => req.path.startsWith(p) && req.method === 'GET')) {
    return next();
  }

  const startTime = Date.now();
  
  // Get user info (from Windows environment or default)
  const userName = process.env.USERNAME || process.env.USER || os.userInfo().username || 'anonymous';
  
  // Get client IP
  const ipAddress = req.ip || req.connection.remoteAddress || 'unknown';

  // Determine action type from path and method
  const getActionType = () => {
    const path = req.path.toLowerCase();
    const method = req.method;

    if (path.includes('/export')) return 'Export';
    if (path.includes('/users') && path.includes('/groups')) return 'View Groups';
    if (path.includes('/users')) return method === 'GET' ? 'Search User' : 'Export Users';
    if (path.includes('/groups') && path.includes('/members')) return 'View Members';
    if (path.includes('/groups')) return method === 'GET' ? 'Search Group' : 'Export Groups';
    if (path.includes('/computers')) return 'Query Computers';
    if (path.includes('/containers')) return 'Query Containers';
    if (path.includes('/gpos')) return 'Query GPOs';
    if (path.includes('/printers')) return 'Query Printers';
    if (path.includes('/contacts')) return 'Query Contacts';
    if (path.includes('/search')) return 'Global Search';
    if (path.includes('/compliance')) return 'Compliance Audit';
    
    return 'API Request';
  };

  // Get target from query or params
  const getTarget = () => {
    return req.query.q || req.query.query || req.params.name || req.params.id || req.query.type || '';
  };

  // Log on response finish
  res.on('finish', () => {
    const duration = Date.now() - startTime;
    const status = res.statusCode < 400 ? 'Success' : 'Failed';
    
    const details = {
      method: req.method,
      path: req.path,
      query: req.query,
      statusCode: res.statusCode,
      duration
    };

    try {
      logActivity(
        userName,
        getActionType(),
        getTarget(),
        details,
        ipAddress,
        status,
        duration
      );
    } catch (err) {
      console.error('Failed to log activity:', err);
    }
  });

  next();
};

module.exports = activityLogger;

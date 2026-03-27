import React, { useState, useEffect } from 'react';
import axios from 'axios';
import ExportButton from '../common/ExportButton';

const ActivityLogs = () => {
  const [logs, setLogs] = useState([]);
  const [total, setTotal] = useState(0);
  const [loading, setLoading] = useState(true);
  const [page, setPage] = useState(0);
  const [filters, setFilters] = useState({
    actionType: 'all',
    status: 'all'
  });

  const limit = 25;

  useEffect(() => {
    fetchLogs();
  }, [page, filters]);

  const fetchLogs = async () => {
    setLoading(true);
    try {
      const response = await axios.get('/api/activity-logs', {
        params: {
          ...filters,
          limit,
          offset: page * limit
        }
      });
      setLogs(response.data.data || []);
      setTotal(response.data.total || 0);
    } catch (err) {
      console.error('Failed to fetch logs:', err);
    } finally {
      setLoading(false);
    }
  };


  return (
    <div className="activity-logs-page">
      <h1>Activity Logs</h1>

      <div className="filters card" style={{ display: 'flex', gap: '20px', alignItems: 'center', marginBottom: '20px' }}>
        <div>
          <label style={{ marginRight: '8px' }}>Action Type:</label>
          <select
            value={filters.actionType}
            onChange={(e) => setFilters({ ...filters, actionType: e.target.value })}
          >
            <option value="all">All Actions</option>
            <option value="Search User">Search User</option>
            <option value="Search Group">Search Group</option>
            <option value="Query Computers">Query Computers</option>
            <option value="Export">Export</option>
          </select>
        </div>
        <div>
          <label style={{ marginRight: '8px' }}>Status:</label>
          <select
            value={filters.status}
            onChange={(e) => setFilters({ ...filters, status: e.target.value })}
          >
            <option value="all">All</option>
            <option value="Success">Success</option>
            <option value="Failed">Failed</option>
          </select>
        </div>
        <ExportButton data={logs} filename="activity_logs" title="Activity Logs Report" />
      </div>

      {loading ? (
        <div className="loading">
          <div className="spinner"></div>
          <span>Loading...</span>
        </div>
      ) : (
        <div className="results-panel card">
          <div className="results-header">
            Total: {total} entries
          </div>
          <table className="data-table">
            <thead>
              <tr>
                <th>Timestamp</th>
                <th>User</th>
                <th>Action</th>
                <th>Target</th>
                <th>Status</th>
                <th>Duration</th>
              </tr>
            </thead>
            <tbody>
              {logs.map((log) => (
                <tr key={log.id}>
                  <td>{new Date(log.timestamp).toLocaleString()}</td>
                  <td>{log.user_name}</td>
                  <td>{log.action_type}</td>
                  <td>{log.target || '-'}</td>
                  <td>
                    <span className={`status-badge ${log.status === 'Success' ? 'enabled' : 'disabled'}`}>
                      {log.status}
                    </span>
                  </td>
                  <td>{log.duration_ms ? `${log.duration_ms}ms` : '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="pagination">
            <span className="pagination-info">
              Showing {page * limit + 1} - {Math.min((page + 1) * limit, total)} of {total}
            </span>
            <div className="pagination-buttons">
              <button onClick={() => setPage(p => Math.max(0, p - 1))} disabled={page === 0}>
                Previous
              </button>
              <button onClick={() => setPage(p => p + 1)} disabled={(page + 1) * limit >= total}>
                Next
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ActivityLogs;

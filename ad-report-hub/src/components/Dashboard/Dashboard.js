import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import axios from 'axios';
import './Dashboard.css';
import FavoritesPanel from './FavoritesPanel';

// Custom colored SVG icons
const UserIcon = ({ color = '#F59E0B', size = 28 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={color} xmlns="http://www.w3.org/2000/svg">
    <circle cx="12" cy="8" r="4"/>
    <path d="M12 14c-4.42 0-8 1.79-8 4v2h16v-2c0-2.21-3.58-4-8-4z"/>
  </svg>
);

const GroupIcon = ({ color = '#EA580C', size = 28 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill={color} xmlns="http://www.w3.org/2000/svg">
    <circle cx="9" cy="8" r="3.5"/>
    <circle cx="16" cy="9" r="2.5"/>
    <path d="M9 13c-3.31 0-6 1.34-6 3v2h12v-2c0-1.66-2.69-3-6-3z"/>
    <path d="M16 13.5c-1.18 0-2.24.32-3 .82.93.6 1.5 1.38 1.5 2.18v1.5h6v-1.5c0-1.38-2.01-3-4.5-3z"/>
  </svg>
);

const ComputerIcon = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>
  </svg>
);

const ServerIcon = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#dc2626" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="2" width="20" height="8" rx="2"/><rect x="2" y="14" width="20" height="8" rx="2"/><circle cx="6" cy="6" r="1" fill="#dc2626"/><circle cx="6" cy="18" r="1" fill="#dc2626"/>
  </svg>
);

const ShieldIcon = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#0891b2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M9 12l2 2 4-4" stroke="#0891b2" strokeWidth="2"/>
  </svg>
);

const TopologyIcon = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#059669" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="5" r="3"/><circle cx="5" cy="19" r="3"/><circle cx="19" cy="19" r="3"/><line x1="12" y1="8" x2="5" y2="16"/><line x1="12" y1="8" x2="19" y2="16"/>
  </svg>
);

const ComplianceIcon = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/>
  </svg>
);

const LogsIcon = ({ size = 24 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#d97706" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/>
  </svg>
);

const Dashboard = () => {
  const [recentLogs, setRecentLogs] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchRecentActivity();
  }, []);

  const fetchRecentActivity = async () => {
    try {
      const response = await axios.get('/api/activity-logs?limit=5');
      setRecentLogs(response.data.data || []);
    } catch (error) {
      console.error('Failed to fetch recent activity:', error);
    } finally {
      setLoading(false);
    }
  };

  const quickActions = [
    { label: 'Search Users', path: '/users', icon: <UserIcon color="#F59E0B" />, color: '#F59E0B' },
    { label: 'Search Groups', path: '/groups', icon: <GroupIcon color="#EA580C" />, color: '#EA580C' },
    { label: 'Query Computers', path: '/computers', icon: <ComputerIcon />, color: '#7c3aed' },
    { label: 'Domain Controllers', path: '/domain-controllers', icon: <ServerIcon />, color: '#dc2626' },
    { label: 'Security Reports', path: '/security', icon: <ShieldIcon />, color: '#0891b2' },
    { label: 'AD Topology', path: '/topology', icon: <TopologyIcon />, color: '#059669' },
    { label: 'Compliance', path: '/compliance', icon: <ComplianceIcon />, color: '#7c3aed' },
    { label: 'Activity Logs', path: '/activity-logs', icon: <LogsIcon />, color: '#d97706' },
  ];

  return (
    <div className="dashboard">
      {/* Hero Section */}
      <div className="dashboard-hero">
        <div className="hero-content">
          <h1>AD Report Hub</h1>
          <p className="hero-tagline">
            Unified Active Directory visibility and reporting for your enterprise.
          </p>
          <p className="hero-subtitle">
            Real-time reporting across identity, infrastructure, and security — built for enterprise scale.
          </p>
        </div>
      </div>

      <FavoritesPanel />

      <div className="quick-actions">
        <h2>Quick Actions</h2>
        <div className="actions-grid">
          {quickActions.map((action) => (
            <Link
              key={action.path}
              to={action.path}
              className="action-card"
              style={{ borderLeftColor: action.color }}
            >
              <span className="action-icon">{action.icon}</span>
              <span className="action-label">{action.label}</span>
            </Link>
          ))}
        </div>
      </div>

      <div className="recent-activity">
        <h2>Recent Activity</h2>
        {loading ? (
          <div className="loading">
            <div className="spinner"></div>
            <span>Loading...</span>
          </div>
        ) : recentLogs.length > 0 ? (
          <table className="data-table">
            <thead>
              <tr>
                <th>Timestamp</th>
                <th>User</th>
                <th>Action</th>
                <th>Target</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {recentLogs.map((log) => (
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
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p className="no-activity">No recent activity</p>
        )}
        <div className="view-all">
          <Link to="/activity-logs" className="btn btn-secondary">
            View All Activity
          </Link>
        </div>
      </div>

      <div className="info-cards">
        <div className="info-card">
          <h3>Key Capabilities</h3>
          <ul>
            <li><strong>Identity Reports</strong> — Users, Groups, Service Accounts</li>
            <li><strong>Infrastructure</strong> — DCs, Sites, Replication, GPOs</li>
            <li><strong>Topology</strong> — Forest structure, FSMO, Trusts</li>
            <li><strong>Export</strong> — CSV, Excel, PDF formats</li>
          </ul>
        </div>
        <div className="info-card">
          <h3>Quick Tips</h3>
          <ul>
            <li>Click column headers to sort reports A-Z or Z-A</li>
            <li>Use the <strong>Recent</strong> button to revisit pages</li>
            <li>All queries run in real-time against AD</li>
            <li>Activity is logged for audit purposes</li>
          </ul>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;

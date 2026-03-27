import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { useApp } from '../../context/AppContext';
import { formatDate } from '../../utils/dateUtils';
import ExportButton from '../common/ExportButton';
import './Security.css';

const Security = () => {
  const { selectedDomain } = useApp();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [data, setData] = useState([]);
  const [summary, setSummary] = useState(null);
  const [summaryLoading, setSummaryLoading] = useState(true);
  
  // Pagination
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  const securityQueries = [
    { 
      id: 'adminsdholder', 
      label: 'AdminSDHolder Protected Users', 
      description: 'Users with adminCount=1 (protected by AdminSDHolder)',
      riskLevel: 'info'
    },
    { 
      id: 'unconstrained-delegation', 
      label: 'Unconstrained Delegation', 
      description: 'Users/Computers with unconstrained Kerberos delegation',
      riskLevel: 'critical'
    },
    { 
      id: 'kerberos-preauth-disabled', 
      label: 'Kerberos Pre-Auth Disabled', 
      description: 'Users vulnerable to AS-REP Roasting attacks',
      riskLevel: 'high'
    },
    { 
      id: 'reversible-encryption', 
      label: 'Reversible Encryption', 
      description: 'Users with password stored using reversible encryption',
      riskLevel: 'high'
    },
    { 
      id: 'nested-privileged', 
      label: 'Nested Privileged Members', 
      description: 'Users with nested membership in privileged groups',
      riskLevel: 'medium'
    },
    { 
      id: 'stale-admins', 
      label: 'Stale Admin Accounts', 
      description: 'Inactive privileged accounts (90+ days)',
      riskLevel: 'high'
    },
    { 
      id: 'disabled-in-groups', 
      label: 'Disabled Users in Groups', 
      description: 'Disabled users still members of security groups',
      riskLevel: 'medium'
    }
  ];

  useEffect(() => {
    loadSummary();
  }, [selectedDomain]);

  const loadSummary = async () => {
    setSummaryLoading(true);
    try {
      const params = {};
      if (selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/security/summary', { params });
      setSummary(response.data.data);
    } catch (err) {
      console.error('Failed to load security summary:', err);
    } finally {
      setSummaryLoading(false);
    }
  };

  const handleQuery = async (queryId) => {
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    
    try {
      const params = { query: queryId };
      if (selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/security', { params });
      setData(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  // Pagination helpers
  const totalPages = Math.ceil(data.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = data.slice(startIndex, endIndex);

  const handlePageSizeChange = (newSize) => {
    setPageSize(newSize);
    setCurrentPage(1);
  };

  const getRiskBadgeClass = (level) => {
    switch (level?.toLowerCase()) {
      case 'critical': return 'risk-critical';
      case 'high': return 'risk-high';
      case 'medium': return 'risk-medium';
      default: return 'risk-low';
    }
  };

  const renderTable = () => {
    if (!activeQuery || paginatedData.length === 0) return null;

    const query = securityQueries.find(q => q.id === activeQuery);
    
    if (activeQuery === 'nested-privileged') {
      return (
        <table className="data-table">
          <thead>
            <tr>
              <th>Privileged Group</th>
              <th>Member</th>
              <th>Account</th>
              <th>Domain</th>
              <th>Title</th>
              <th>Enabled</th>
              <th>Nested</th>
              <th>Path</th>
            </tr>
          </thead>
          <tbody>
            {paginatedData.map((item, i) => (
              <tr key={i}>
                <td>{item.PrivilegedGroup}</td>
                <td>{item.MemberName}</td>
                <td>{item.SamAccountName}</td>
                <td>{item.Domain || '-'}</td>
                <td>{item.Title || '-'}</td>
                <td>
                  <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                    {item.Enabled ? 'Yes' : 'No'}
                  </span>
                </td>
                <td>{item.IsNested ? 'Yes' : 'No'}</td>
                <td className="membership-path">{item.MembershipPath}</td>
              </tr>
            ))}
          </tbody>
        </table>
      );
    }

    if (activeQuery === 'unconstrained-delegation') {
      return (
        <table className="data-table">
          <thead>
            <tr>
              <th>Type</th>
              <th>Name</th>
              <th>Account</th>
              <th>Domain</th>
              <th>Enabled</th>
              <th>SPN Count</th>
              <th>Last Logon</th>
              <th>Risk</th>
            </tr>
          </thead>
          <tbody>
            {paginatedData.map((item, i) => (
              <tr key={i}>
                <td>
                  <span className={`type-badge type-${item.ObjectType?.toLowerCase()}`}>
                    {item.ObjectType}
                  </span>
                </td>
                <td>{item.Name}</td>
                <td>{item.SamAccountName}</td>
                <td>{item.Domain || '-'}</td>
                <td>
                  <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                    {item.Enabled ? 'Yes' : 'No'}
                  </span>
                </td>
                <td>{item.SPNCount}</td>
                <td>{formatDate(item.LastLogon)}</td>
                <td>
                  <span className={`risk-badge ${getRiskBadgeClass(item.RiskLevel)}`}>
                    {item.RiskLevel}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      );
    }

    return (
      <table className="data-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Account</th>
            <th>Domain</th>
            <th>Email</th>
            <th>Department</th>
            <th>Enabled</th>
            <th>Last Logon</th>
            <th>Risk</th>
          </tr>
        </thead>
        <tbody>
          {paginatedData.map((item, i) => (
            <tr key={i}>
              <td>{item.Name}</td>
              <td>{item.SamAccountName}</td>
              <td>{item.Domain || '-'}</td>
              <td>{item.Email || '-'}</td>
              <td>{item.Department || '-'}</td>
              <td>
                <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                  {item.Enabled ? 'Yes' : 'No'}
                </span>
              </td>
              <td>{formatDate(item.LastLogon)}</td>
              <td>
                <span className={`risk-badge ${getRiskBadgeClass(item.RiskLevel)}`}>
                  {item.RiskLevel}
                </span>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  const renderPagination = () => {
    if (data.length === 0) return null;
    const startItem = startIndex + 1;
    const endItem = Math.min(endIndex, data.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {data.length} results
        </div>
        <div className="pagination-actions">
          <div className="pagination-page-size">
            <label>Per page:</label>
            <select 
              value={pageSize} 
              onChange={(e) => handlePageSizeChange(parseInt(e.target.value))}
            >
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
            </select>
          </div>
          <div className="pagination-nav">
            <button 
              onClick={() => setCurrentPage(p => Math.max(1, p - 1))} 
              disabled={currentPage <= 1}
            >
              Back
            </button>
            <span className="page-indicator">
              Page {currentPage} of {totalPages}
            </span>
            <button 
              onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} 
              disabled={currentPage >= totalPages}
            >
              Next
            </button>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="security-page">
      <h1>Security & Compliance</h1>

      {/* Security Summary Dashboard - Clickable Cards */}
      <div className="security-summary card">
        <h3>Security Overview <span className="summary-hint">Click any card to view details</span></h3>
        {summaryLoading ? (
          <div className="loading-inline">Loading security summary...</div>
        ) : summary ? (
          <div className="summary-cards">
            <div className="summary-card critical clickable-card" onClick={() => handleQuery('adminsdholder')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('adminsdholder')}>
              <div className="summary-value">{summary.criticalIssues}</div>
              <div className="summary-label">Critical Issues</div>
              <div className="summary-action">View Details</div>
            </div>
            <div className="summary-card high clickable-card" onClick={() => handleQuery('stale-admins')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('stale-admins')}>
              <div className="summary-value">{summary.highIssues}</div>
              <div className="summary-label">High Risk</div>
              <div className="summary-action">View Details</div>
            </div>
            <div className="summary-card medium clickable-card" onClick={() => handleQuery('disabled-in-groups')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('disabled-in-groups')}>
              <div className="summary-value">{summary.mediumIssues}</div>
              <div className="summary-label">Medium Risk</div>
              <div className="summary-action">View Details</div>
            </div>
            <div className="summary-card info clickable-card" onClick={() => handleQuery('adminsdholder')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('adminsdholder')}>
              <div className="summary-value">{summary.adminSDHolderCount}</div>
              <div className="summary-label">AdminSDHolder Users</div>
              <div className="summary-action">View Report</div>
            </div>
            <div className="summary-card info clickable-card" onClick={() => handleQuery('unconstrained-delegation')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('unconstrained-delegation')}>
              <div className="summary-value">{summary.unconstrainedDelegationCount}</div>
              <div className="summary-label">Unconstrained Delegation</div>
              <div className="summary-action">View Report</div>
            </div>
            <div className="summary-card info clickable-card" onClick={() => handleQuery('stale-admins')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('stale-admins')}>
              <div className="summary-value">{summary.staleAdminCount}</div>
              <div className="summary-label">Stale Admins</div>
              <div className="summary-action">View Report</div>
            </div>
            <div className="summary-card info clickable-card" onClick={() => handleQuery('kerberos-preauth-disabled')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('kerberos-preauth-disabled')}>
              <div className="summary-value">{summary.kerberosPreAuthDisabledCount}</div>
              <div className="summary-label">Kerberos Pre-Auth Off</div>
              <div className="summary-action">View Report</div>
            </div>
            <div className="summary-card info clickable-card" onClick={() => handleQuery('disabled-in-groups')} role="button" tabIndex={0} onKeyDown={(e) => (e.key === 'Enter') && handleQuery('disabled-in-groups')}>
              <div className="summary-value">{summary.disabledInGroupsCount}</div>
              <div className="summary-label">Disabled in Groups</div>
              <div className="summary-action">View Report</div>
            </div>
          </div>
        ) : (
          <div className="summary-error">Unable to load security summary</div>
        )}
      </div>

      {/* Security Queries */}
      <div className="query-panel card">
        <h3>Security Reports</h3>
        <div className="query-cards" role="list">
          {securityQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''} risk-${q.riskLevel}`}
              onClick={() => handleQuery(q.id)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  handleQuery(q.id);
                }
              }}
              role="button"
              tabIndex={0}
              aria-label={`${q.label}: ${q.description}`}
            >
              <div className="query-card-title">
                <span>{q.label}</span>
                <span className={`risk-indicator risk-${q.riskLevel}`}></span>
              </div>
              <div className="query-card-desc">{q.description}</div>
            </div>
          ))}
        </div>
      </div>

      {error && (
        <div className="error-message">{error}</div>
      )}

      {loading && (
        <div className="loading">
          <div className="spinner"></div>
          <span>Running security analysis...</span>
        </div>
      )}

      {!loading && data.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <span>Found {data.length} result(s)</span>
            <ExportButton 
              data={data} 
              filename={`security_${activeQuery || 'export'}`} 
              title="Security Report" 
            />
          </div>
          <div style={{ overflowX: 'auto' }}>
            {renderTable()}
          </div>
          {renderPagination()}
        </div>
      )}

      {!loading && activeQuery && data.length === 0 && !error && (
        <div className="no-results card">
          <p>No security issues found for this query.</p>
        </div>
      )}
    </div>
  );
};

export default Security;

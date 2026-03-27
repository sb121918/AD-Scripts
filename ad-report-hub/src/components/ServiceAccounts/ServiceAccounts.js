import React, { useState } from 'react';
import axios from 'axios';
import './ServiceAccounts.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { formatDate } from '../../utils/dateUtils';
import { useSortableData } from '../../utils/useSortableData';
import { useApp } from '../../context/AppContext';

const ServiceAccounts = () => {
  const { selectedDomain } = useApp();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [domain, setDomain] = useState('');
  const [inactiveDomain, setInactiveDomain] = useState('');
  const [days, setDays] = useState(365);
  const [inactiveDays, setInactiveDays] = useState(90);
  const [orphanedDomain, setOrphanedDomain] = useState('');
  const [orphanedDays, setOrphanedDays] = useState(30);
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  // Sorting
  const { sortedItems, requestSort, getSortIndicator, getSortClass } = useSortableData(data);

  const predefinedQueries = [
    { id: 'all-forest', label: 'All Service Accounts (Forest-Wide)', description: 'All service accounts across the entire AD forest' },
    { id: 'all', label: 'Service Accounts per Domain', description: 'Service accounts for a specific domain', needsDomain: true },
    { id: 'orphaned', label: 'Orphaned Service Accounts', description: 'Accounts without a manager (forest-wide)' },
    { id: 'orphaned-recent', label: 'Recent Created - Orphaned Service Accounts per Domain', description: 'Orphaned accounts created in X days (max 90)', needsOrphaned: true },
    { id: 'inactive', label: 'Inactive Service Accounts', description: 'Service accounts not logged on for X days (per domain)', needsInactive: true },
    { id: 'pwd-never-expires', label: 'Password Never Expires', description: 'Service accounts with non-expiring passwords — security hygiene risk' },
    { id: 'interactive-logon', label: 'Interactive Logon Detected', description: 'Service accounts with recent interactive logon activity — potential misuse' },
  ];

  const handleQuery = async (queryId) => {
    if (queryId === 'all' && !domain.trim() && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    if (queryId === 'inactive' && !inactiveDomain.trim() && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    if (queryId === 'orphaned-recent' && !orphanedDomain.trim() && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    
    try {
      const params = { query: queryId };
      if (queryId === 'all') params.domain = domain || selectedDomain;
      if (queryId === 'orphaned-recent') {
        params.domain = orphanedDomain || selectedDomain;
        params.days = Math.min(orphanedDays, 90); // Max 90 days
      }
      if (queryId === 'inactive') {
        params.domain = inactiveDomain || selectedDomain;
        params.days = inactiveDays;
      }
      if (!params.domain && selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/service-accounts', { params });
      setData(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  // Pagination helpers
  const totalPages = Math.ceil(sortedItems.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = sortedItems.slice(startIndex, endIndex);

  const handlePageSizeChange = (newSize) => {
    setPageSize(newSize);
    setCurrentPage(1);
  };

  const handleNextPage = () => {
    if (currentPage < totalPages) {
      setCurrentPage(currentPage + 1);
    }
  };

  const handlePrevPage = () => {
    if (currentPage > 1) {
      setCurrentPage(currentPage - 1);
    }
  };

  const renderPagination = () => {
    if (sortedItems.length === 0) return null;
    const startItem = startIndex + 1;
    const endItem = Math.min(endIndex, sortedItems.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {sortedItems.length} results
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
              onClick={handlePrevPage} 
              disabled={currentPage <= 1}
            >
              Back
            </button>
            <span className="page-indicator">
              Page {currentPage} of {totalPages}
            </span>
            <button 
              onClick={handleNextPage} 
              disabled={currentPage >= totalPages}
            >
              Next
            </button>
          </div>
        </div>
      </div>
    );
  };

  const renderSortableHeader = (key, label) => (
    <th 
      className={getSortClass(key)}
      onClick={() => requestSort(key)}
      style={{ cursor: 'pointer' }}
    >
      {label}
      <span className="sort-indicator">{getSortIndicator(key)}</span>
    </th>
  );

  const renderTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          {renderSortableHeader('Name', 'Name')}
          {renderSortableHeader('SamAccountName', 'SAM Account')}
          {renderSortableHeader('AccountType', 'Account Type')}
          {renderSortableHeader('Enabled', 'Status')}
          {renderSortableHeader('Domain', 'Domain')}
          {renderSortableHeader('Created', 'Created')}
          {activeQuery === 'inactive' && renderSortableHeader('DaysSinceLogon', 'Days Inactive')}
          {activeQuery === 'inactive' && renderSortableHeader('LastLogonDate', 'Last Logon')}
          {renderSortableHeader('EA6_Value', 'EA6 Value')}
          {renderSortableHeader('ManagerStatus', 'Manager Status')}
          {renderSortableHeader('Department', 'Department')}
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((account, i) => (
          <tr key={i}>
            <td>{account.Name}</td>
            <td>{account.SamAccountName}</td>
            <td>{account.AccountType || 'Service Account'}</td>
            <td>
              <span className={`status-badge ${account.Enabled ? 'enabled' : 'disabled'}`}>
                {account.Enabled ? 'Enabled' : 'Disabled'}
              </span>
            </td>
            <td>{account.Domain || '-'}</td>
            <td>{formatDate(account.Created)}</td>
            {activeQuery === 'inactive' && (
              <td>
                <span className={`days-badge ${account.DaysSinceLogon === 'Never' ? 'never' : account.DaysSinceLogon > 180 ? 'critical' : account.DaysSinceLogon > 90 ? 'warning' : 'normal'}`}>
                  {account.DaysSinceLogon}
                </span>
              </td>
            )}
            {activeQuery === 'inactive' && (
              <td>{account.LastLogonDate ? formatDate(account.LastLogonDate) : 'Never'}</td>
            )}
            <td className="ea6-value">{account.EA6_Value || account.extensionAttribute6 || '-'}</td>
            <td>
              <span className={`manager-badge ${account.ManagerStatus === 'NOT ASSIGNED' || account.ManagerStatus === 'MISSING' ? 'missing' : 'assigned'}`}>
                {account.ManagerStatus || 'Unknown'}
              </span>
            </td>
            <td>{account.Department || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  return (
    <div className="svc-page">
      <h1>Service Accounts</h1>

      <div className="query-panel card">
        <h3>Predefined Queries</h3>
        <div className="query-cards" style={{
          display: 'flex',
          flexWrap: 'wrap',
          justifyContent: 'center',
          gap: '16px'
        }}>
          {predefinedQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => !q.needsDomain && !q.needsOrphaned && !q.needsInactive && handleQuery(q.id)}
              style={{
                flex: '0 0 calc(33.33% - 16px)',
                maxWidth: 'calc(33.33% - 16px)',
                minWidth: '300px'
              }}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/service-accounts" queryId={q.id} label={q.label} />
              </div>
              <div className="query-card-desc">{q.description}</div>
              
              {q.needsDomain && (
                <div className="query-input" onClick={(e) => e.stopPropagation()}>
                  <input
                    type="text"
                    placeholder="Enter domain (e.g., corp.domain.com)"
                    value={domain}
                    onChange={(e) => setDomain(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleQuery(q.id)}
                  />
                  <button className="btn btn-primary btn-sm" onClick={() => handleQuery(q.id)}>
                    Run
                  </button>
                </div>
              )}
              
              {q.needsOrphaned && (
                <div className="query-input" onClick={(e) => e.stopPropagation()}>
                  <input
                    type="text"
                    placeholder="Domain (e.g., corp.domain.com)"
                    value={orphanedDomain}
                    onChange={(e) => setOrphanedDomain(e.target.value)}
                    style={{ flex: 2 }}
                  />
                  <input
                    type="number"
                    placeholder="Days"
                    value={orphanedDays}
                    onChange={(e) => setOrphanedDays(Math.min(parseInt(e.target.value) || 0, 90))}
                    min="1"
                    max="90"
                    style={{ width: '80px', flex: 'none' }}
                  />
                  <button className="btn btn-primary btn-sm" onClick={() => handleQuery(q.id)}>
                    Run
                  </button>
                </div>
              )}
              
              {q.needsInactive && (
                <div className="query-input" onClick={(e) => e.stopPropagation()}>
                  <input
                    type="text"
                    placeholder="Domain (e.g., corp.domain.com)"
                    value={inactiveDomain}
                    onChange={(e) => setInactiveDomain(e.target.value)}
                    style={{ flex: 2 }}
                  />
                  <input
                    type="number"
                    placeholder="Days"
                    value={inactiveDays}
                    onChange={(e) => setInactiveDays(e.target.value)}
                    min="1"
                    max="9999"
                    style={{ width: '80px', flex: 'none' }}
                  />
                  <button className="btn btn-primary btn-sm" onClick={() => handleQuery(q.id)}>
                    Run
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>
      </div>

      {error && <div className="error-message">{error}</div>}

      {loading && (
        <div className="loading">
          <div className="spinner"></div>
          <span>Querying Service Accounts...</span>
        </div>
      )}

      {!loading && data.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <span>Found {data.length} total service account(s)</span>
            <ExportButton data={data} filename={`service_accounts_${activeQuery || 'export'}`} title="Service Accounts Report" />
          </div>
          <p style={{ padding: '8px 20px', margin: 0, fontSize: '12px', color: 'var(--text-secondary)' }}>
            Click on column headers to sort A-Z or Z-A
          </p>
          <div style={{ overflowX: 'auto' }}>
            {renderTable()}
          </div>
          {renderPagination()}
        </div>
      )}

      {!loading && activeQuery && data.length === 0 && !error && (
        <div className="no-results card">
          <p>No service accounts found matching the criteria.</p>
        </div>
      )}
    </div>
  );
};

export default ServiceAccounts;

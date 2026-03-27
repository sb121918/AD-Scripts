import React, { useState } from 'react';
import axios from 'axios';
import './Users.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { formatDate } from '../../utils/dateUtils';
import { useApp } from '../../context/AppContext';

const Users = () => {
  const { selectedDomain } = useApp();
  const [searchValue, setSearchValue] = useState('');
  const [searchType, setSearchType] = useState('wildcard');
  const [users, setUsers] = useState([]);
  const [selectedUser, setSelectedUser] = useState(null);
  const [userGroups, setUserGroups] = useState(null);
  const [loading, setLoading] = useState(false);
  const [showGroups, setShowGroups] = useState(false);
  const [selectedGroup, setSelectedGroup] = useState(null);
  const [groupMembers, setGroupMembers] = useState(null);
  const [showGroupMembers, setShowGroupMembers] = useState(false);
  const [loadingGroupMembers, setLoadingGroupMembers] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [domainInputs, setDomainInputs] = useState({});
  const [lockedOutUserInput, setLockedOutUserInput] = useState('');
  const [lookbackHoursInput, setLookbackHoursInput] = useState(24);
  const [daysInput, setDaysInput] = useState(30);
  const [queryResults, setQueryResults] = useState([]);
  const [accountTypeSummary, setAccountTypeSummary] = useState(null);
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  const [searchPageSize, setSearchPageSize] = useState(25);
  const [searchCurrentPage, setSearchCurrentPage] = useState(1);

  const forestWideQueries = [
    { id: 'enabled-forest', label: 'All Enabled Users', description: 'All enabled accounts across the forest' },
    { id: 'disabled-forest', label: 'All Disabled Users', description: 'All disabled accounts across the forest' },
    { id: 'enabled-human-forest', label: 'Enabled Users with EA6 "Human"', description: 'Forest-wide enabled users where EA6 contains "Human"' },
    { id: 'empty-ea6', label: 'Users with Empty EA6', description: 'Forest-wide users missing EA6 attribute' },
    { id: 'password-expired', label: 'Password Expired', description: 'Users with expired passwords' },
    { id: 'password-never-expires', label: 'Password Never Expires', description: 'Non-compliant password settings' },
    { id: 'never-logged-on', label: 'Never Logged On', description: 'Users who have never authenticated' },
    { id: 'created-x-days', label: 'Users Created in X Days', description: 'Recently created user accounts', needsDays: true },
    { id: 'locked-out', label: 'Locked Out Users', description: 'Lookup a user and show lockout source (DC/caller/IP)', needsLookup: true },
  ];

  const domainSpecificQueries = [
    { id: 'enabled-by-domain', label: 'Enabled Users', description: 'All enabled users in specific domain', needsDomain: true },
    { id: 'disabled-by-domain', label: 'Disabled Users', description: 'All disabled users in specific domain', needsDomain: true },
    { id: 'enabled-human', label: 'Enabled Human Accounts', description: 'Employees matched to HR (EA6 = Human Primary Identity SF Match)', needsDomain: true },
    { id: 'enabled-contractors', label: 'Enabled Contractors', description: 'Contractor accounts (EA6 = Human Primary Identity Contractor)', needsDomain: true },
    { id: 'employee-type-count', label: 'Employee Type Count', description: 'Count of enabled users by EmployeeType attribute', needsDomain: true },
    { id: 'ea6-value-count', label: 'EA6 Value Count', description: 'Count of enabled users by extensionAttribute6 value', needsDomain: true },
  ];

  const predefinedQueries = [...forestWideQueries, ...domainSpecificQueries];

  const handlePredefinedQuery = async (queryId) => {
    const query = predefinedQueries.find(q => q.id === queryId);
    const domainValue = domainInputs[queryId]?.trim() || '';
    
    // Validate required inputs
    if (query.needsDomain && !domainValue && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    if (query.needsLookup && !lockedOutUserInput.trim()) {
      setError('Please enter a user identity (samAccountName/UPN/email)');
      return;
    }

    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setSelectedUser(null);
    setUsers([]);
    setAccountTypeSummary(null);
    setCurrentPage(1); // Reset pagination on new query
    
    try {
      const params = { query: queryId };
      if (query.needsDomain) params.domain = domainValue || selectedDomain;
      if (query.needsDays) params.days = daysInput;
      if (queryId === 'locked-out') {
        params.value = lockedOutUserInput.trim();
        params.lookbackHours = lookbackHoursInput;
      }
      if (!params.domain && selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/users', { params });
      
      const summaryQueries = ['account-type-count', 'employee-type-count', 'ea6-value-count'];
      if (summaryQueries.includes(queryId) && response.data.data) {
        const summaryData = Array.isArray(response.data.data) ? response.data.data[0] : response.data.data;
        setAccountTypeSummary(summaryData);
        setQueryResults([]);
      } else {
        setQueryResults(response.data.data || []);
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setQueryResults([]);
    } finally {
      setLoading(false);
    }
  };

  const handleSearch = async (type) => {
    if (!searchValue.trim()) return;
    
    setLoading(true);
    setError(null);
    setSelectedUser(null);
    setSearchCurrentPage(1); // Reset pagination on new search
    
    try {
      const response = await axios.get('/api/users/search', {
        params: { q: searchValue, type }
      });
      
      // Handle nested response structure from PowerShell
      let userData = response.data?.data;
      
      // If data is an object with a nested data property (from PS script), extract it
      if (userData && typeof userData === 'object' && !Array.isArray(userData)) {
        if (userData.data && Array.isArray(userData.data)) {
          userData = userData.data;
        } else if (userData.error) {
          // PowerShell returned an error
          setError(userData.error);
          setUsers([]);
          return;
        } else {
          // Single object result, wrap in array
          userData = [userData];
        }
      }
      
      // Ensure we have an array
      setUsers(Array.isArray(userData) ? userData : []);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
      setUsers([]);
    } finally {
      setLoading(false);
    }
  };

  // Pagination helpers for query results
  const totalPages = Math.ceil(queryResults.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedQueryResults = queryResults.slice(startIndex, endIndex);

  // Pagination helpers for search results
  const searchTotalPages = Math.ceil(users.length / searchPageSize);
  const searchStartIndex = (searchCurrentPage - 1) * searchPageSize;
  const searchEndIndex = searchStartIndex + searchPageSize;
  const paginatedUsers = users.slice(searchStartIndex, searchEndIndex);

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

  const handleSearchPageSizeChange = (newSize) => {
    setSearchPageSize(newSize);
    setSearchCurrentPage(1);
  };

  const handleSearchNextPage = () => {
    if (searchCurrentPage < searchTotalPages) {
      setSearchCurrentPage(searchCurrentPage + 1);
    }
  };

  const handleSearchPrevPage = () => {
    if (searchCurrentPage > 1) {
      setSearchCurrentPage(searchCurrentPage - 1);
    }
  };

  const renderPagination = (data, page, totalPg, startIdx, endIdx, onPageSizeChange, onPrev, onNext, pageSz) => {
    if (data.length === 0) return null;
    const startItem = startIdx + 1;
    const endItem = Math.min(endIdx, data.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {data.length} results
        </div>
        <div className="pagination-actions">
          <div className="pagination-page-size">
            <label>Per page:</label>
            <select 
              value={pageSz} 
              onChange={(e) => onPageSizeChange(parseInt(e.target.value))}
            >
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
            </select>
          </div>
          <div className="pagination-nav">
            <button 
              onClick={onPrev} 
              disabled={page <= 1}
            >
              Back
            </button>
            <span className="page-indicator">
              Page {page} of {totalPg}
            </span>
            <button 
              onClick={onNext} 
              disabled={page >= totalPg}
            >
              Next
            </button>
          </div>
        </div>
      </div>
    );
  };

  const handleSelectUser = async (user) => {
    setSelectedUser(user);
    setShowGroups(false);
    setUserGroups(null);
  };

  const handleViewGroups = async () => {
    if (!selectedUser) return;
    
    try {
      setError(null);
      const response = await axios.get(`/api/users/${selectedUser.SamAccountName}/groups`);
      
      // Handle nested response structure
      let groupsData = response.data?.data;
      
      // If data is further nested, extract it
      if (groupsData && typeof groupsData === 'object') {
        // Check for error in response
        if (groupsData.Error) {
          setError(groupsData.Error);
          setUserGroups(null);
          setShowGroups(false);
          return;
        }
        
        // If the structure is { User, TotalGroups, Groups }
        if (groupsData.TotalGroups !== undefined || groupsData.Groups) {
          setUserGroups(groupsData);
          setShowGroups(true);
          return;
        }
      }
      
      // Fallback: try direct response.data
      if (response.data?.TotalGroups !== undefined || response.data?.Groups) {
        setUserGroups(response.data);
        setShowGroups(true);
        return;
      }
      
      // If we still don't have data, set empty groups
      setUserGroups({ TotalGroups: 0, Groups: [] });
      setShowGroups(true);
    } catch (err) {
      console.error('Failed to load groups:', err);
      setError('Failed to load groups: ' + (err.response?.data?.error || err.message));
      setUserGroups(null);
      setShowGroups(false);
    }
  };

  const handleViewGroupMembers = async (group) => {
    if (!group || !group.DistinguishedName) return;
    
    setSelectedGroup(group);
    setLoadingGroupMembers(true);
    setGroupMembers(null);
    setShowGroupMembers(true);
    
    try {
      const response = await axios.get('/api/users/group-members', {
        params: {
          groupDN: group.DistinguishedName,
          domain: group.Domain
        }
      });
      
      // Handle nested response structure
      let membersData = response.data?.data;
      if (membersData && typeof membersData === 'object') {
        setGroupMembers(membersData);
      } else if (response.data?.Members !== undefined) {
        setGroupMembers(response.data);
      } else {
        setGroupMembers({ TotalMembers: 0, Members: [] });
      }
    } catch (err) {
      console.error('Failed to load group members:', err);
      setGroupMembers({ 
        Error: 'Failed to load group members: ' + (err.response?.data?.error || err.message),
        TotalMembers: 0, 
        Members: [] 
      });
    } finally {
      setLoadingGroupMembers(false);
    }
  };

  const closeGroupMembers = () => {
    setShowGroupMembers(false);
    setSelectedGroup(null);
    setGroupMembers(null);
  };

  // Single user export is handled inline with ExportButton

  const renderQueryResultsTable = () => {
    if (activeQuery === 'locked-out') {
      const diag = Array.isArray(queryResults) ? queryResults[0] : null;
      const lockouts = diag?.lockouts || [];
      const ad = diag?.ad || {};

      return (
        <div style={{ padding: '16px 20px' }}>
          <div className="card" style={{ marginBottom: '16px' }}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '16px', padding: '12px 14px' }}>
              <div><strong>User</strong><div style={{ color: 'var(--text-secondary)' }}>{diag?.samAccountName || '-'}</div></div>
              <div><strong>Locked</strong><div style={{ color: 'var(--text-secondary)' }}>{ad?.isLocked ? 'Yes' : 'No'}</div></div>
              <div><strong>Lockout time</strong><div style={{ color: 'var(--text-secondary)' }}>{ad?.lockoutTime ? formatDate(ad.lockoutTime) : '-'}</div></div>
              <div><strong>Last bad password</strong><div style={{ color: 'var(--text-secondary)' }}>{ad?.lastBadPasswordAttempt ? formatDate(ad.lastBadPasswordAttempt) : '-'}</div></div>
              <div><strong>badPwdCount</strong><div style={{ color: 'var(--text-secondary)' }}>{ad?.badPwdCount ?? '-'}</div></div>
              <div><strong>PDC</strong><div style={{ color: 'var(--text-secondary)' }}>{diag?.pdcEmulator || '-'}</div></div>
            </div>
            {(diag?.warnings || []).length > 0 && (
              <div style={{ padding: '0 14px 12px 14px', color: '#b45309', fontSize: '12px' }}>
                {(diag.warnings || []).map((w, i) => <div key={i}>{w}</div>)}
              </div>
            )}
          </div>

          <table className="data-table">
            <thead>
              <tr>
                <th>Time</th>
                <th>DC (4740)</th>
                <th>Caller Computer</th>
                <th>IP (best effort)</th>
                <th>Workstation</th>
                <th>Matched Event</th>
                <th>Correlated DC</th>
              </tr>
            </thead>
            <tbody>
              {lockouts.length === 0 ? (
                <tr><td colSpan="7">No 4740 lockout events found in lookback window.</td></tr>
              ) : (
                lockouts.map((l, i) => (
                  <tr key={i}>
                    <td>{l.time ? formatDate(l.time) : '-'}</td>
                    <td>{l.dc || '-'}</td>
                    <td>{l.callerComputer || '-'}</td>
                    <td>{l.ipAddress || '-'}</td>
                    <td>{l.workstation || '-'}</td>
                    <td>{l.eventIdMatched || '-'}</td>
                    <td>{l.correlatedDc || '-'}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      );
    }

    return (
      <table className="data-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>SAM Account</th>
            <th>Email</th>
            <th>Department</th>
            <th>Domain</th>
            <th>Created</th>
          </tr>
        </thead>
        <tbody>
          {paginatedQueryResults.map((user, i) => (
            <tr key={i}>
              <td>{user.Name || user.DisplayName}</td>
              <td>{user.SamAccountName}</td>
              <td>{user.Email || '-'}</td>
              <td>{user.Department || '-'}</td>
              <td>{user.Domain || '-'}</td>
              <td>{formatDate(user.Created || user.WhenCreated)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  return (
    <div className="users-page">
      <h1>Users</h1>

      {/* Forest-Wide Queries */}
      <div className="query-panel card">
        <h3>Forest-Wide Reports</h3>
        <div className="query-cards" role="list">
          {forestWideQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => !q.needsDomain && !q.needsDays && !q.needsLookup && handlePredefinedQuery(q.id)}
              onKeyDown={(e) => {
                if ((e.key === 'Enter' || e.key === ' ') && !q.needsDomain && !q.needsDays && !q.needsLookup) {
                  e.preventDefault();
                  handlePredefinedQuery(q.id);
                }
              }}
              role={!q.needsDomain && !q.needsDays && !q.needsLookup ? 'button' : 'listitem'}
              tabIndex={!q.needsDomain && !q.needsDays && !q.needsLookup ? 0 : undefined}
              aria-label={`${q.label}: ${q.description}`}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/users" queryId={q.id} label={q.label} />
              </div>
              <div className="query-card-desc">{q.description}</div>
              
              {q.needsLookup && q.id === 'locked-out' && (
                <div className="query-input" onClick={(e) => e.stopPropagation()} style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <input
                    type="text"
                    placeholder="Enter username / UPN / email"
                    value={lockedOutUserInput}
                    onChange={(e) => setLockedOutUserInput(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handlePredefinedQuery(q.id)}
                  />
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <input
                      type="number"
                      placeholder="Lookback hours"
                      value={lookbackHoursInput}
                      onChange={(e) => setLookbackHoursInput(parseInt(e.target.value || '24', 10))}
                      min="1"
                      max="168"
                      style={{ width: '140px' }}
                    />
                    <button className="btn btn-primary btn-sm" onClick={() => handlePredefinedQuery(q.id)}>
                      Run
                    </button>
                  </div>
                </div>
              )}
              
              {q.needsDays && (
                <div className="query-input" onClick={(e) => e.stopPropagation()}>
                  <input
                    type="number"
                    placeholder="Days"
                    value={daysInput}
                    onChange={(e) => setDaysInput(e.target.value)}
                    min="1"
                    max="365"
                  />
                  <button className="btn btn-primary btn-sm" onClick={() => handlePredefinedQuery(q.id)}>
                    Run
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>
      </div>

      {/* Domain-Specific Queries */}
      <div className="query-panel card">
        <h3>Domain-Specific Reports</h3>
        <div className="query-cards" role="list">
          {domainSpecificQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              aria-label={`${q.label}: ${q.description}`}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/users" queryId={q.id} label={q.label} />
              </div>
              <div className="query-card-desc">{q.description}</div>
              <div className="query-input" onClick={(e) => e.stopPropagation()}>
                <input
                  type="text"
                  placeholder="Enter domain (e.g., corp.domain.com)"
                  value={domainInputs[q.id] || ''}
                  onChange={(e) => setDomainInputs(prev => ({ ...prev, [q.id]: e.target.value }))}
                  onKeyPress={(e) => e.key === 'Enter' && handlePredefinedQuery(q.id)}
                />
                <button className="btn btn-primary btn-sm" onClick={() => handlePredefinedQuery(q.id)}>
                  Run
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Account Type Summary */}
      {accountTypeSummary && (() => {
        const reportType = accountTypeSummary.ReportType || 'Account Type';
        const counts = accountTypeSummary.Counts;
        const hasEmployeeType = accountTypeSummary.EmployeeTypeCounts;
        const hasEA6 = accountTypeSummary.EA6Counts;

        const exportData = counts
          ? counts.map(item => ({
              [reportType]: item.Type || '(Empty)',
              Count: item.Count,
              Domain: accountTypeSummary.Domain
            }))
          : [
              ...(hasEmployeeType || []).map(item => ({ Category: 'Employee Type', Type: item.Type || '(Empty)', Count: item.Count, Domain: accountTypeSummary.Domain })),
              ...(hasEA6 || []).map(item => ({ Category: 'EA6 Value', Type: item.Type || '(Empty)', Count: item.Count, Domain: accountTypeSummary.Domain }))
            ];

        return (
          <div className="summary-panel card">
            <div className="summary-header">
              <div>
                <h3>{reportType} Summary - {accountTypeSummary.Domain}</h3>
                <p>Total Enabled Users: <strong>{accountTypeSummary.TotalEnabledUsers}</strong></p>
              </div>
              <ExportButton 
                data={exportData}
                filename={`${reportType.toLowerCase().replace(/\s+/g, '_')}_summary_${accountTypeSummary.Domain}`} 
                title={`${reportType} Summary - ${accountTypeSummary.Domain}`} 
              />
            </div>

            {/* New single-category format (Employee Type or EA6) */}
            {counts && (
              <div className="summary-section">
                <h4>By {reportType}</h4>
                <ul>
                  {counts.map((item, i) => (
                    <li key={i}><span>{item.Type || '(Empty)'}</span><strong>{item.Count}</strong></li>
                  ))}
                </ul>
              </div>
            )}

            {/* Legacy combined format (account-type-count) */}
            {!counts && (hasEmployeeType || hasEA6) && (
              <div className="summary-grid">
                {hasEmployeeType && (
                  <div className="summary-section">
                    <h4>By Employee Type</h4>
                    <ul>
                      {hasEmployeeType.map((item, i) => (
                        <li key={i}><span>{item.Type || '(Empty)'}</span><strong>{item.Count}</strong></li>
                      ))}
                    </ul>
                  </div>
                )}
                {hasEA6 && (
                  <div className="summary-section">
                    <h4>By EA6 Value</h4>
                    <ul>
                      {hasEA6.map((item, i) => (
                        <li key={i}><span>{item.Type || '(Empty)'}</span><strong>{item.Count}</strong></li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })()}

      {/* Query Results */}
      {!loading && queryResults.length > 0 && !selectedUser && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <span>Found {queryResults.length} total result(s)</span>
            <ExportButton data={queryResults} filename={`users_${activeQuery || 'export'}`} title="Users Report" />
          </div>
          <div style={{ overflowX: 'auto' }}>
            {renderQueryResultsTable()}
          </div>
          {renderPagination(queryResults, currentPage, totalPages, startIndex, endIndex, handlePageSizeChange, handlePrevPage, handleNextPage, pageSize)}
        </div>
      )}
      
      <h3 style={{ marginTop: '30px' }}>Search User</h3>
      
      <div className="search-panel card">
        <div className="search-row">
          <div className="search-field">
            <label>Exact Email</label>
            <div className="search-input-group">
              <input
                type="text"
                placeholder="user@domain.com"
                value={searchType === 'exact' ? searchValue : ''}
                onChange={(e) => { setSearchValue(e.target.value); setSearchType('exact'); }}
                onKeyPress={(e) => e.key === 'Enter' && handleSearch('exact')}
              />
              <button className="btn btn-primary" onClick={() => handleSearch('exact')}>
                SEARCH
              </button>
            </div>
          </div>
          
          <div className="search-field">
            <label>Wildcard Search (name, email, UPN)</label>
            <div className="search-input-group">
              <input
                type="text"
                placeholder="Enter search term..."
                value={searchType === 'wildcard' ? searchValue : ''}
                onChange={(e) => { setSearchValue(e.target.value); setSearchType('wildcard'); }}
                onKeyPress={(e) => e.key === 'Enter' && handleSearch('wildcard')}
              />
              <button className="btn btn-primary" onClick={() => handleSearch('wildcard')}>
                FIND MATCHES
              </button>
            </div>
          </div>
        </div>
      </div>

      {error && (
        <div className="error-message">{error}</div>
      )}

      {loading && (
        <div className="loading">
          <div className="spinner"></div>
          <span>Searching Active Directory...</span>
        </div>
      )}

      {!loading && users.length > 0 && !selectedUser && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header">
            <span>Found {users.length} total user(s)</span>
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th>Display Name</th>
                  <th>SAM Account</th>
                  <th>Email</th>
                  <th>Department</th>
                  <th>Status</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                {paginatedUsers.map((user, index) => (
                  <tr key={index}>
                    <td>{user.DisplayName}</td>
                    <td>{user.SamAccountName}</td>
                    <td>{user.Email}</td>
                    <td>{user.Department || '-'}</td>
                    <td>
                      <span className={`status-badge ${user.Enabled ? 'enabled' : 'disabled'}`}>
                        {user.Enabled ? 'Enabled' : 'Disabled'}
                      </span>
                    </td>
                    <td>
                      <button 
                        className="btn btn-secondary"
                        onClick={() => handleSelectUser(user)}
                      >
                        View Details
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {renderPagination(users, searchCurrentPage, searchTotalPages, searchStartIndex, searchEndIndex, handleSearchPageSizeChange, handleSearchPrevPage, handleSearchNextPage, searchPageSize)}
        </div>
      )}

      {selectedUser && (
        <div className="user-details card">
          <div className="user-header">
            <div className="user-avatar">
              {selectedUser.DisplayName?.charAt(0) || 'U'}
            </div>
            <div className="user-info">
              <h2>{selectedUser.DisplayName}</h2>
              <p className="user-email">{selectedUser.Email}</p>
              <p className="user-title">{selectedUser.Title}</p>
            </div>
            <div className="user-actions">
              <span className={`status-badge ${selectedUser.Enabled ? 'enabled' : 'disabled'}`}>
                {selectedUser.Enabled ? 'Enabled' : 'Disabled'}
              </span>
              <button className="btn btn-primary" onClick={handleViewGroups}>
                VIEW GROUPS
              </button>
              <ExportButton data={selectedUser ? [selectedUser] : []} filename={`user_${selectedUser?.SamAccountName || 'export'}`} title="User Details" />
              <button className="btn btn-secondary" onClick={() => setSelectedUser(null)}>
                HIDE DETAILS
              </button>
            </div>
          </div>

          <div className="user-details-grid">
            <div className="details-section">
              <h3>Basic Information</h3>
              <div className="details-list">
                <div className="detail-row">
                  <span className="detail-label">Display Name</span>
                  <span className="detail-value">{selectedUser.DisplayName}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Username (SAM)</span>
                  <span className="detail-value">{selectedUser.SamAccountName}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">User Principal Name</span>
                  <span className="detail-value">{selectedUser.UserPrincipalName}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Employee ID</span>
                  <span className="detail-value">{selectedUser.EmployeeID || '-'}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Employee Number</span>
                  <span className="detail-value">{selectedUser.EmployeeNumber || '-'}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Distinguished Name</span>
                  <span className="detail-value dn">{selectedUser.DistinguishedName}</span>
                </div>
              </div>
            </div>

            <div className="details-section">
              <h3>Contact Information</h3>
              <div className="details-list">
                <div className="detail-row">
                  <span className="detail-label">Email Address</span>
                  <span className="detail-value">{selectedUser.Email}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Telephone</span>
                  <span className="detail-value">{selectedUser.Telephone || '-'}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Mobile Phone</span>
                  <span className="detail-value">{selectedUser.Mobile || '-'}</span>
                </div>
                <div className="detail-row">
                  <span className="detail-label">Proxy Addresses ({selectedUser.ProxyAddresses?.length || 0})</span>
                  <div className="proxy-addresses">
                    {selectedUser.ProxyAddresses?.slice(0, 5).map((addr, i) => (
                      <span key={i} className="proxy-address">{addr}</span>
                    ))}
                    {selectedUser.ProxyAddresses?.length > 5 && (
                      <span className="proxy-more">+{selectedUser.ProxyAddresses.length - 5} more</span>
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {showGroups && (
        <div className="modal-overlay" onClick={() => setShowGroups(false)}>
          <div className="modal modal-large" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Group Memberships - {selectedUser?.DisplayName || selectedUser?.SamAccountName}</h3>
              <button className="modal-close" onClick={() => setShowGroups(false)} aria-label="Close groups modal">×</button>
            </div>
            <div className="modal-body">
              {userGroups?.Error ? (
                <div className="error-message">{userGroups.Error}</div>
              ) : (
                <>
                  <div className="groups-header">
                    <p className="groups-count">Total Groups: {userGroups?.TotalGroups || 0}</p>
                    {userGroups?.Groups?.length > 0 && (
                      <ExportButton 
                        data={userGroups.Groups} 
                        filename={`user_groups_${selectedUser?.SamAccountName || 'export'}`} 
                        title={`Group Memberships - ${selectedUser?.DisplayName || selectedUser?.SamAccountName}`} 
                      />
                    )}
                  </div>
                  <p className="groups-hint" style={{ color: 'var(--text-secondary)', fontSize: '0.85rem', marginBottom: '10px' }}>
                    Click on a group name to view its members
                  </p>
                  {(!userGroups?.Groups || userGroups.Groups.length === 0) ? (
                    <p style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '20px' }}>
                      No group memberships found for this user.
                    </p>
                  ) : (
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Group Name</th>
                          <th>Type</th>
                          <th>Scope</th>
                          <th>Domain</th>
                          <th>Email</th>
                        </tr>
                      </thead>
                      <tbody>
                        {userGroups.Groups.map((group, i) => (
                          <tr key={i}>
                            <td>
                              <button 
                                className="link-button"
                                onClick={() => handleViewGroupMembers(group)}
                                title="Click to view group members"
                              >
                                {group.Name}
                              </button>
                            </td>
                            <td>{group.Type}</td>
                            <td>{group.Scope}</td>
                            <td>{group.Domain || '-'}</td>
                            <td>{group.Email || '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </>
              )}
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setShowGroups(false)}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Group Members Modal */}
      {showGroupMembers && (
        <div className="modal-overlay" onClick={closeGroupMembers}>
          <div className="modal modal-large" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Members of: {selectedGroup?.Name}</h3>
              <button className="modal-close" onClick={closeGroupMembers} aria-label="Close group members modal">×</button>
            </div>
            <div className="modal-body">
              {loadingGroupMembers ? (
                <div className="loading-spinner" style={{ textAlign: 'center', padding: '40px' }}>
                  <p>Loading group members...</p>
                </div>
              ) : groupMembers?.Error ? (
                <div className="error-message">{groupMembers.Error}</div>
              ) : (
                <>
                  <div className="group-info" style={{ marginBottom: '15px', padding: '10px', backgroundColor: 'var(--bg-secondary)', borderRadius: '5px' }}>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '10px', fontSize: '0.9rem' }}>
                      <div><strong>Type:</strong> {groupMembers?.GroupCategory || '-'}</div>
                      <div><strong>Scope:</strong> {groupMembers?.GroupScope || '-'}</div>
                      <div><strong>Managed By:</strong> {groupMembers?.ManagedBy || '-'}</div>
                    </div>
                    {groupMembers?.Description && (
                      <div style={{ marginTop: '10px', fontSize: '0.9rem' }}>
                        <strong>Description:</strong> {groupMembers.Description}
                      </div>
                    )}
                  </div>
                  <div className="groups-header">
                    <p className="groups-count">Total Members: {groupMembers?.TotalMembers || 0}</p>
                    {groupMembers?.Members?.length > 0 && (
                      <ExportButton 
                        data={groupMembers.Members} 
                        filename={`group_members_${selectedGroup?.Name || 'export'}`} 
                        title={`Members of ${selectedGroup?.Name}`} 
                      />
                    )}
                  </div>
                  {(!groupMembers?.Members || groupMembers.Members.length === 0) ? (
                    <p style={{ color: 'var(--text-secondary)', textAlign: 'center', padding: '20px' }}>
                      No members found in this group.
                    </p>
                  ) : (
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Name</th>
                          <th>Type</th>
                          <th>Account Name</th>
                          <th>Title</th>
                          <th>Department</th>
                          <th>Email</th>
                          <th>Enabled</th>
                        </tr>
                      </thead>
                      <tbody>
                        {groupMembers.Members.map((member, i) => (
                          <tr key={i}>
                            <td>{member.Name}</td>
                            <td>
                              <span className={`type-badge type-${member.Type?.toLowerCase()}`}>
                                {member.Type}
                              </span>
                            </td>
                            <td>{member.SamAccountName || '-'}</td>
                            <td>{member.Title || '-'}</td>
                            <td>{member.Department || '-'}</td>
                            <td>{member.Email || '-'}</td>
                            <td>
                              {member.Enabled === null ? '-' : (
                                <span className={`status-badge ${member.Enabled ? 'enabled' : 'disabled'}`}>
                                  {member.Enabled ? 'Yes' : 'No'}
                                </span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </>
              )}
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={closeGroupMembers}>
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Users;

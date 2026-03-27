import React, { useState, useMemo } from 'react';
import axios from 'axios';
import './Groups.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { formatDate } from '../../utils/dateUtils';
import { useApp } from '../../context/AppContext';

const Groups = () => {
  const { selectedDomain, forestData } = useApp();
  const [searchValue, setSearchValue] = useState('');
  const [dlGroupInput, setDlGroupInput] = useState('');
  const [reportDomainInput, setReportDomainInput] = useState('');
  const [data, setData] = useState([]);
  const [dlMeta, setDlMeta] = useState(null);
  const [dlCandidates, setDlCandidates] = useState([]);
  const [selectedGroup, setSelectedGroup] = useState(null);
  const [members, setMembers] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  // Separate domain inputs for each privileged query
  const [domainAdminsDomain, setDomainAdminsDomain] = useState('');
  const [builtinAdminsDomain, setBuiltinAdminsDomain] = useState('');
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  
  // Sorting state
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  // Sorting logic
  const sortedData = useMemo(() => {
    if (!data || !Array.isArray(data)) return [];
    let sortableItems = [...data];
    if (sortConfig.key !== null) {
      sortableItems.sort((a, b) => {
        let aValue = a[sortConfig.key];
        let bValue = b[sortConfig.key];
        if (aValue === null || aValue === undefined) aValue = '';
        if (bValue === null || bValue === undefined) bValue = '';
        if (typeof aValue === 'boolean') aValue = aValue ? 1 : 0;
        if (typeof bValue === 'boolean') bValue = bValue ? 1 : 0;
        if (typeof aValue === 'number' && typeof bValue === 'number') {
          return sortConfig.direction === 'asc' ? aValue - bValue : bValue - aValue;
        }
        const aStr = String(aValue).toLowerCase();
        const bStr = String(bValue).toLowerCase();
        if (aStr < bStr) return sortConfig.direction === 'asc' ? -1 : 1;
        if (aStr > bStr) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }
    return sortableItems;
  }, [data, sortConfig]);

  const handleSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
    setCurrentPage(1);
  };

  const getSortIndicator = (key) => {
    if (sortConfig.key !== key) return <span className="sort-indicator">↕</span>;
    return <span className="sort-indicator active">{sortConfig.direction === 'asc' ? '↑' : '↓'}</span>;
  };

  const SortableHeader = ({ columnKey, label }) => (
    <th onClick={() => handleSort(columnKey)} style={{ cursor: 'pointer' }} className="sortable-th">
      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
        {label}
        {getSortIndicator(columnKey)}
      </div>
    </th>
  );

  const predefinedQueries = [
    { id: 'all', label: 'All Groups', description: 'All groups in the forest' },
    { id: 'security', label: 'Security Groups', description: 'Security groups only' },
    { id: 'distribution', label: 'Distribution Groups', description: 'Distribution groups only' },
    { id: 'empty', label: 'Empty Groups', description: 'Groups with no members' },
    { id: 'mail-enabled', label: 'Mail-Enabled Groups', description: 'Groups with email addresses' },
    { id: 'dl-members', label: 'DL Members Lookup', description: 'Get distribution list members', needsGroupName: true },
  ];

  const effectiveDomain = useMemo(() => {
    const typed = String(reportDomainInput || '').trim();
    if (typed) return typed;
    return selectedDomain || '';
  }, [reportDomainInput, selectedDomain]);

  const effectiveDomainLabel = useMemo(() => {
    if (!effectiveDomain) return 'Default domain';
    const found = forestData?.domains?.find(d => d.name === effectiveDomain);
    return found?.label || effectiveDomain;
  }, [effectiveDomain, forestData]);

  const selectedForest = useMemo(() => {
    const domains = forestData?.domains || [];
    const match = domains.find(d => String(d.name).toLowerCase() === String(selectedDomain || '').toLowerCase());
    return match?.forest || '';
  }, [forestData, selectedDomain]);

  const privilegedQueries = [
    { 
      id: 'privileged-groups', 
      label: 'All Privileged Groups', 
      description: selectedForest?.toLowerCase() === 'viacom' 
        ? 'DA + EA + SA + BA + AD-Enterprise Systems Admins (Forest-wide)' 
        : 'DA + EA + SA + BA (Forest-wide)' 
    },
    { id: 'enterprise-admins', label: 'Enterprise Admins', description: 'Enterprise Admins group members (Forest-wide)' },
    { id: 'schema-admins', label: 'Schema Admins', description: 'Schema Admins group members (Forest-wide)' },
    ...(selectedForest?.toLowerCase() === 'viacom'
      ? [{ id: 'enterprise-systems-admins', label: 'AD-Enterprise Systems Admins', description: 'Custom security group in ad.viacom.com' }]
      : []),
    { id: 'domain-admins', label: 'Domain Admins', description: 'Domain Admins group members', needsDomain: true },
    { id: 'builtin-admins', label: 'Built-in Administrators', description: 'Built-in Administrators members', needsDomain: true },
  ];

  const handleSearch = async () => {
    if (!searchValue.trim()) return;
    setLoading(true);
    setError(null);
    setActiveQuery('search');
    setCurrentPage(1);
    setDlMeta(null);
    setDlCandidates([]);
    try {
      const response = await axios.get('/api/groups/search', { params: { q: searchValue, domain: effectiveDomain || '' } });
      
      // Handle nested response structure
      let resultData = response.data?.data;
      if (resultData && typeof resultData === 'object' && !Array.isArray(resultData)) {
        if (resultData.data && Array.isArray(resultData.data)) {
          resultData = resultData.data;
        } else if (resultData.error) {
          setError(resultData.error);
          setData([]);
          return;
        }
      }
      setData(Array.isArray(resultData) ? resultData : []);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  const handleQuery = async (queryId, domainValue = null) => {
    // Check if domain is required for this query
    const needsDomain = ['domain-admins', 'builtin-admins'].includes(queryId);
    if (needsDomain && (!domainValue || !domainValue.trim()) && !effectiveDomain) {
      setError('Please enter a domain name');
      return;
    }
    
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    setSortConfig({ key: null, direction: 'asc' }); // Reset sorting
    setDlMeta(null);
    setDlCandidates([]);
    try {
      const params = { query: queryId };
      if (needsDomain) {
        params.domain = (domainValue && String(domainValue).trim()) ? String(domainValue).trim() : effectiveDomain;
      } else if (effectiveDomain) {
        params.domain = effectiveDomain;
      }

      if (queryId === 'dl-members') {
        if (!dlGroupInput.trim()) {
          setError('Please enter a DL name or email');
          setLoading(false);
          return;
        }
        params.groupName = dlGroupInput.trim();
      }
      
      const response = await axios.get('/api/groups', { params });
      
      // DL Members special handling: script returns a single object wrapped as [obj]
      if (queryId === 'dl-members') {
        const payload = response.data?.data;
        const first = Array.isArray(payload) ? payload[0] : payload;

        if (!first) {
          setData([]);
          return;
        }

        if (first.Error) {
          setError(first.Error);
          setData([]);
          return;
        }

        if (first.MultipleResults) {
          setDlCandidates(Array.isArray(first.Groups) ? first.Groups : []);
          setData([]);
          return;
        }

        setDlMeta({
          GroupName: first.GroupName,
          GroupEmail: first.GroupEmail,
          MemberCount: first.MemberCount
        });
        setData(Array.isArray(first.Members) ? first.Members : []);
        return;
      }

      // Handle nested response structure
      let resultData = response.data?.data;
      if (resultData && typeof resultData === 'object' && !Array.isArray(resultData)) {
        if (resultData.data && Array.isArray(resultData.data)) {
          resultData = resultData.data;
        } else if (resultData.error) {
          setError(resultData.error);
          setData([]);
          return;
        }
      }
      setData(Array.isArray(resultData) ? resultData : []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  // Pagination helpers
  const totalPages = Math.ceil(sortedData.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = sortedData.slice(startIndex, endIndex);

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

  const handleViewMembers = async (group) => {
    setSelectedGroup(group);
    try {
      const response = await axios.get(`/api/groups/${encodeURIComponent(group.Name)}/members`, { params: { domain: effectiveDomain || '' } });
      setMembers(response.data);
    } catch (err) {
      setError('Failed to load members');
    }
  };

  const handleSelectDlCandidate = async (group) => {
    setLoading(true);
    setError(null);
    setActiveQuery('dl-members');
    setCurrentPage(1);
    setSortConfig({ key: null, direction: 'asc' });
    setDlMeta(null);
    setData([]);
    try {
      const params = { query: 'dl-members', groupDN: group.DistinguishedName };
      if (effectiveDomain) params.domain = effectiveDomain;

      const response = await axios.get('/api/groups', { params });
      const payload = response.data?.data;
      const first = Array.isArray(payload) ? payload[0] : payload;

      if (!first) {
        setData([]);
        setDlCandidates([]);
        return;
      }
      if (first.Error) {
        setError(first.Error);
        setData([]);
        setDlCandidates([]);
        return;
      }

      setDlCandidates([]);
      setDlMeta({
        GroupName: first.GroupName,
        GroupEmail: first.GroupEmail,
        MemberCount: first.MemberCount
      });
      setData(Array.isArray(first.Members) ? first.Members : []);
    } catch (err) {
      setError(err.response?.data?.error || 'DL lookup failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  const isPrivilegedQuery = ['privileged-groups', 'enterprise-admins', 'schema-admins', 'enterprise-systems-admins', 'domain-admins', 'builtin-admins'].includes(activeQuery);

  const renderGroupsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="Email" label="Email" />
          <SortableHeader columnKey="Type" label="Type" />
          <SortableHeader columnKey="Scope" label="Scope" />
          <SortableHeader columnKey="MemberCount" label="Members" />
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((group, i) => (
          <tr key={i}>
            <td>{group.Name}</td>
            <td>{group.Email || '-'}</td>
            <td>{group.Type}</td>
            <td>{group.Scope}</td>
            <td>{group.MemberCount}</td>
            <td>
              <button className="btn btn-secondary" onClick={() => handleViewMembers(group)}>
                View Members
              </button>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderDlCandidatesTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Name</th>
          <th>Display Name</th>
          <th>Email</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>
        {dlCandidates.map((g, i) => (
          <tr key={i}>
            <td>{g.Name}</td>
            <td>{g.DisplayName || '-'}</td>
            <td>{g.Email || '-'}</td>
            <td>
              <button className="btn btn-secondary" onClick={() => handleSelectDlCandidate(g)}>
                View Members
              </button>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderDlMembersTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="Type" label="Type" />
          <SortableHeader columnKey="Title" label="Title" />
          <SortableHeader columnKey="Department" label="Department" />
          <SortableHeader columnKey="Email" label="Email" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((m, i) => (
          <tr key={i}>
            <td>{m.Name}</td>
            <td>{m.Type}</td>
            <td>{m.Title || '-'}</td>
            <td>{m.Department || '-'}</td>
            <td>{m.Email || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderPrivilegedTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="GroupName" label="Group Name" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="MemberName" label="Member Name" />
          <SortableHeader columnKey="SamAccountName" label="SAM Account" />
          <SortableHeader columnKey="Title" label="Title" />
          <SortableHeader columnKey="Department" label="Department" />
          <SortableHeader columnKey="Enabled" label="Enabled" />
          <SortableHeader columnKey="GroupModified" label="Group Modified" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((member, i) => (
          <tr key={i}>
            <td>{member.GroupName}</td>
            <td>{member.Domain}</td>
            <td>{member.MemberName || member.DisplayName}</td>
            <td>{member.SamAccountName}</td>
            <td>{member.Title || '-'}</td>
            <td>{member.Department || '-'}</td>
            <td>
              <span className={`status-badge ${member.Enabled ? 'enabled' : 'disabled'}`}>
                {member.Enabled ? 'Yes' : 'No'}
              </span>
            </td>
            <td>{formatDate(member.GroupModified)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderTable = () => {
    if (isPrivilegedQuery) {
      return renderPrivilegedTable();
    }
    if (activeQuery === 'dl-members') {
      if (dlCandidates.length > 0) return renderDlCandidatesTable();
      return renderDlMembersTable();
    }
    return renderGroupsTable();
  };

  return (
    <div className="groups-page">
      <h1>Groups</h1>

      <div className="card report-scope-card">
        <div className="report-scope-row">
          <div className="report-scope-left">
            <div className="report-scope-title">Report domain</div>
            <div className="report-scope-hint">
              Enter a domain name (including child domains). Using: <strong>{effectiveDomainLabel}</strong>
            </div>
          </div>
          <div className="report-scope-right">
            <input
              className="report-scope-input"
              type="text"
              placeholder="e.g. cbs.ad.cbs.net"
              value={reportDomainInput}
              onChange={(e) => setReportDomainInput(e.target.value)}
            />
            <button
              className="btn btn-secondary"
              onClick={() => setReportDomainInput('')}
              style={{ padding: '10px 14px' }}
              disabled={!reportDomainInput.trim()}
            >
              Clear
            </button>
          </div>
        </div>
      </div>

      {/* Standard Group Queries */}
      <div className="query-panel card">
        <h3>Standard Queries</h3>
        <div className="query-cards" style={{ 
          display: 'flex', 
          flexWrap: 'wrap',
          justifyContent: 'center',
          gap: '12px',
          marginBottom: '16px'
        }}>
          {predefinedQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => !q.needsGroupName && handleQuery(q.id)}
              style={{
                flex: '0 1 200px',
                maxWidth: '280px',
                padding: '14px',
                border: `2px solid ${activeQuery === q.id ? 'var(--accent-primary)' : 'transparent'}`,
                borderRadius: '8px',
                backgroundColor: activeQuery === q.id ? 'var(--accent-light)' : 'var(--bg-secondary)',
                cursor: 'pointer',
                transition: 'all 0.2s ease'
              }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)', marginBottom: '4px' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/groups" queryId={q.id} label={q.label} />
              </div>
              <div style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                {q.description}
              </div>

              {q.needsGroupName && (
                <div className="query-input" onClick={(e) => e.stopPropagation()} style={{ display: 'flex', gap: '8px', marginTop: '10px' }}>
                  <input
                    type="text"
                    placeholder="Enter DL name or email"
                    value={dlGroupInput}
                    onChange={(e) => setDlGroupInput(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleQuery(q.id)}
                    style={{
                      flex: 1,
                      padding: '8px 12px',
                      border: '1px solid var(--border-input)',
                      borderRadius: '4px',
                      backgroundColor: 'var(--bg-input)',
                      color: 'var(--text-primary)',
                      fontSize: '13px'
                    }}
                  />
                  <button className="btn btn-primary" onClick={() => handleQuery(q.id)} style={{ padding: '8px 16px' }}>
                    Run
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>
      </div>

      {/* Privileged Groups Queries */}
      <div className="query-panel card">
        <h3>Privileged Groups Reports</h3>
        <div className="query-cards" style={{ 
          display: 'flex', 
          flexWrap: 'wrap',
          justifyContent: 'center',
          gap: '12px'
        }}>
          {privilegedQueries.map((q) => {
            // Use separate domain input for each query type
            const getDomainValue = () => {
              if (q.id === 'domain-admins') return domainAdminsDomain;
              if (q.id === 'builtin-admins') return builtinAdminsDomain;
              return '';
            };
            const setDomainValue = (value) => {
              if (q.id === 'domain-admins') setDomainAdminsDomain(value);
              if (q.id === 'builtin-admins') setBuiltinAdminsDomain(value);
            };
            
            return (
              <div
                key={q.id}
                className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
                onClick={() => !q.needsDomain && handleQuery(q.id)}
                style={{
                  padding: '14px',
                  border: `2px solid ${activeQuery === q.id ? 'var(--accent-primary)' : 'transparent'}`,
                  borderRadius: '8px',
                  backgroundColor: activeQuery === q.id ? 'var(--accent-light)' : 'var(--bg-secondary)',
                  cursor: q.needsDomain ? 'default' : 'pointer',
                  transition: 'all 0.2s ease'
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)', marginBottom: '4px' }}>
                  <span>{q.label}</span>
                  <FavoriteButton page="/groups" queryId={q.id} label={q.label} />
                </div>
                <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: q.needsDomain ? '10px' : '0' }}>
                  {q.description}
                </div>
                
                {q.needsDomain && (
                  <div className="query-input" onClick={(e) => e.stopPropagation()} style={{ display: 'flex', gap: '8px', marginTop: '8px' }}>
                    <input
                      type="text"
                      placeholder="Enter domain (e.g., corp.domain.com)"
                      value={getDomainValue()}
                      onChange={(e) => setDomainValue(e.target.value)}
                      onKeyPress={(e) => e.key === 'Enter' && handleQuery(q.id, getDomainValue())}
                      style={{
                        flex: 1,
                        padding: '8px 12px',
                        border: '1px solid var(--border-input)',
                        borderRadius: '4px',
                        backgroundColor: 'var(--bg-input)',
                        color: 'var(--text-primary)',
                        fontSize: '13px'
                      }}
                    />
                    <button className="btn btn-primary" onClick={() => handleQuery(q.id, getDomainValue())} style={{ padding: '8px 16px' }}>
                      Run
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      <div className="search-panel card">
        <div className="search-input-group">
          <input
            type="text"
            placeholder="Search by group name or email..."
            value={searchValue}
            onChange={(e) => setSearchValue(e.target.value)}
            onKeyPress={(e) => e.key === 'Enter' && handleSearch()}
          />
          <button className="btn btn-primary" onClick={handleSearch}>SEARCH</button>
        </div>
      </div>

      {error && <div className="error-message">{error}</div>}

      {loading && (
        <div className="loading">
          <div className="spinner"></div>
          <span>Loading...</span>
        </div>
      )}

      {!loading && data.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <div>
              <span>
                Found {data.length} total result(s)
                {activeQuery === 'dl-members' && dlMeta?.GroupName ? ` — ${dlMeta.GroupName} (${dlMeta.MemberCount || data.length})` : ''}
              </span>
              <span style={{ marginLeft: '16px', fontSize: '12px', color: 'var(--text-muted)' }}>
                Click column headers to sort ↕
              </span>
            </div>
            <ExportButton data={data} filename={`groups_${activeQuery || 'export'}`} title="Groups Report" />
          </div>
          <div style={{ overflowX: 'auto' }}>
            {renderTable()}
          </div>
          {renderPagination()}
        </div>
      )}

      {members && (
        <div className="modal-overlay" onClick={() => setMembers(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Members of {selectedGroup?.Name}</h3>
              <button className="modal-close" onClick={() => setMembers(null)} aria-label="Close members modal">×</button>
            </div>
            <div className="modal-body">
              <p>Total Members: {members.MemberCount}</p>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Type</th>
                    <th>Title</th>
                    <th>Department</th>
                  </tr>
                </thead>
                <tbody>
                  {members.Members?.map((member, i) => (
                    <tr key={i}>
                      <td>{member.Name}</td>
                      <td>{member.Type}</td>
                      <td>{member.Title || '-'}</td>
                      <td>{member.Department || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setMembers(null)}>Close</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Groups;

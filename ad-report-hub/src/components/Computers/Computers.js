import React, { useState, useMemo } from 'react';
import axios from 'axios';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { formatDate, formatLastLogon } from '../../utils/dateUtils';
import { useApp } from '../../context/AppContext';

const Computers = () => {
  const { selectedDomain } = useApp();
  const [searchValue, setSearchValue] = useState('');
  const [computers, setComputers] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  
  // Domain inputs for reports
  const [allByDomainInput, setAllByDomainInput] = useState('');
  const [osByDomainInput, setOsByDomainInput] = useState('');
  const [selectedOS, setSelectedOS] = useState('Windows Server 2022');
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  
  // Sorting state
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  // Sorting logic
  const sortedData = useMemo(() => {
    if (!computers || !Array.isArray(computers)) return [];
    let sortableItems = [...computers];
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
  }, [computers, sortConfig]);

  const handleSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
    setCurrentPage(1);
  };

  const getSortIndicator = (key) => {
    if (sortConfig.key !== key) return <span className="sort-indicator" style={{ opacity: 0.4, marginLeft: '4px' }}>↕</span>;
    return <span className="sort-indicator" style={{ color: 'var(--accent-primary)', marginLeft: '4px', fontWeight: 'bold' }}>{sortConfig.direction === 'asc' ? '↑' : '↓'}</span>;
  };

  const SortableHeader = ({ columnKey, label }) => (
    <th onClick={() => handleSort(columnKey)} style={{ cursor: 'pointer' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
        {label}
        {getSortIndicator(columnKey)}
      </div>
    </th>
  );

  // Standard queries (quick access)
  const standardQueries = [
    { id: 'enabled-forest', label: 'Enabled Computers', description: 'Forest-wide' },
    { id: 'disabled-forest', label: 'Disabled Computers', description: 'Forest-wide' },
    { id: 'created-30-days', label: 'Created in Last 30 Days', description: 'Forest-wide' },
    { id: 'disabled-30-days', label: 'Disabled in Last 30 Days', description: 'Forest-wide' },
    { id: 'never-logged-on', label: 'Never Logged On', description: 'Forest-wide' },
    { id: 'not-logged-60-days', label: 'Not Logged On for 60 Days', description: 'Forest-wide' },
  ];

  // Report queries (with domain/OS selection)
  const reportQueries = [
    { id: 'all-by-domain', label: 'All Computers by Domain', description: 'All computers in a specific domain', needsDomain: true },
    { id: 'os-by-domain', label: 'Computers by OS', description: 'Filter by Operating System per domain', needsDomain: true, needsOS: true },
  ];

  // Operating system options
  const osOptions = [
    'Windows Server 2022',
    'Windows Server 2019',
    'Windows Server 2016',
    'Windows Server 2012',
    'Windows 11',
    'Windows 10',
    'Windows 7',
  ];

  const handleSearch = async () => {
    if (!searchValue.trim()) return;
    setLoading(true);
    setError(null);
    setCurrentPage(1);
    setActiveQuery('search');
    try {
      const response = await axios.get('/api/computers/search', { params: { q: searchValue } });
      
      let resultData = response.data?.data;
      if (resultData && typeof resultData === 'object' && !Array.isArray(resultData)) {
        if (resultData.data && Array.isArray(resultData.data)) {
          resultData = resultData.data;
        } else if (resultData.error) {
          setError(resultData.error);
          setComputers([]);
          return;
        }
      }
      setComputers(Array.isArray(resultData) ? resultData : []);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
      setComputers([]);
    } finally {
      setLoading(false);
    }
  };

  const handleQuery = async (queryId, options = {}) => {
    const { domain, os } = options;
    
    // Validate required fields
    if (queryId === 'all-by-domain' && (!domain || !domain.trim()) && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    if (queryId === 'os-by-domain' && (!domain || !domain.trim()) && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    
    setLoading(true);
    setError(null);
    setCurrentPage(1);
    setActiveQuery(queryId);
    
    try {
      const params = { query: queryId };
      const effectiveDomain = (domain && domain.trim()) || selectedDomain;
      if (effectiveDomain) {
        params.domain = effectiveDomain;
      }
      if (os) {
        params.os = os;
      }
      if (!params.domain && selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/computers', { params });
      
      let resultData = response.data?.data;
      if (resultData && typeof resultData === 'object' && !Array.isArray(resultData)) {
        if (resultData.data && Array.isArray(resultData.data)) {
          resultData = resultData.data;
        } else if (resultData.error) {
          setError(resultData.error);
          setComputers([]);
          return;
        }
      }
      setComputers(Array.isArray(resultData) ? resultData : []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setComputers([]);
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
    if (computers.length === 0) return null;
    const startItem = startIndex + 1;
    const endItem = Math.min(endIndex, computers.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {computers.length} results
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

  return (
    <div className="computers-page">
      <h1>Computers</h1>

      {/* Standard Queries */}
      <div className="query-panel card">
        <h3>Standard Queries (Forest-Wide)</h3>
        <div className="query-cards" style={{ 
          display: 'flex', 
          flexWrap: 'wrap',
          justifyContent: 'center',
          gap: '16px',
          marginBottom: '16px'
        }}>
          {standardQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handleQuery(q.id)}
              style={{
                flex: '0 1 calc(33.33% - 16px)',
                minWidth: '280px',
                maxWidth: '350px',
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
                <FavoriteButton page="/computers" queryId={q.id} label={q.label} />
              </div>
              <div style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                {q.description}
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Computer Reports */}
      <div className="query-panel card">
        <h3>Computer Reports (Per Domain)</h3>
        <div className="query-cards" style={{ 
          display: 'flex', 
          flexWrap: 'wrap',
          justifyContent: 'center',
          gap: '16px'
        }}>
          {/* All Computers by Domain */}
          <div
            className={`query-card ${activeQuery === 'all-by-domain' ? 'active' : ''}`}
            style={{
              flex: '0 1 calc(33.33% - 16px)',
              minWidth: '320px',
              maxWidth: '400px',
              padding: '16px',
              border: `2px solid ${activeQuery === 'all-by-domain' ? 'var(--accent-primary)' : 'transparent'}`,
              borderRadius: '8px',
              backgroundColor: activeQuery === 'all-by-domain' ? 'var(--accent-light)' : 'var(--bg-secondary)',
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)', marginBottom: '4px' }}>
              <span>All Computers by Domain</span>
              <FavoriteButton page="/computers" queryId="all-by-domain" label="All Computers by Domain" />
            </div>
            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '10px' }}>
              Get all computers from a specific domain
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <input
                type="text"
                placeholder="Enter domain (e.g., corp.domain.com)"
                value={allByDomainInput}
                onChange={(e) => setAllByDomainInput(e.target.value)}
                onKeyPress={(e) => e.key === 'Enter' && handleQuery('all-by-domain', { domain: allByDomainInput })}
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
              <button className="btn btn-primary" onClick={() => handleQuery('all-by-domain', { domain: allByDomainInput })} style={{ padding: '8px 16px' }}>
                Run
              </button>
            </div>
          </div>

          {/* Computers by OS */}
          <div
            className={`query-card ${activeQuery === 'os-by-domain' ? 'active' : ''}`}
            style={{
              flex: '0 1 calc(33.33% - 16px)',
              minWidth: '320px',
              maxWidth: '400px',
              padding: '16px',
              border: `2px solid ${activeQuery === 'os-by-domain' ? 'var(--accent-primary)' : 'transparent'}`,
              borderRadius: '8px',
              backgroundColor: activeQuery === 'os-by-domain' ? 'var(--accent-light)' : 'var(--bg-secondary)',
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)', marginBottom: '4px' }}>
              <span>Computers by Operating System</span>
              <FavoriteButton page="/computers" queryId="os-by-domain" label="Computers by Operating System" />
            </div>
            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '10px' }}>
              Filter computers by OS in a specific domain
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div style={{ display: 'flex', gap: '8px' }}>
                <input
                  type="text"
                  placeholder="Enter domain (e.g., corp.domain.com)"
                  value={osByDomainInput}
                  onChange={(e) => setOsByDomainInput(e.target.value)}
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
              </div>
              <div style={{ display: 'flex', gap: '8px' }}>
                <select
                  value={selectedOS}
                  onChange={(e) => setSelectedOS(e.target.value)}
                  style={{
                    flex: 1,
                    padding: '8px 12px',
                    border: '1px solid var(--border-input)',
                    borderRadius: '4px',
                    backgroundColor: 'var(--bg-input)',
                    color: 'var(--text-primary)',
                    fontSize: '13px'
                  }}
                >
                  {osOptions.map((os) => (
                    <option key={os} value={os}>{os}</option>
                  ))}
                </select>
                <button className="btn btn-primary" onClick={() => handleQuery('os-by-domain', { domain: osByDomainInput, os: selectedOS })} style={{ padding: '8px 16px' }}>
                  Run
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="search-panel card">
        <div className="search-input-group">
          <input
            type="text"
            placeholder="Search by computer name..."
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

      {!loading && computers.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <div>
              <span>Found {computers.length} total computer(s)</span>
              <span style={{ marginLeft: '16px', fontSize: '12px', color: 'var(--text-muted)' }}>
                Click column headers to sort ↕
              </span>
            </div>
            <ExportButton data={computers} filename="computers_export" title="Computers Report" />
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table className="data-table">
              <thead>
                <tr>
                  <SortableHeader columnKey="Name" label="Name" />
                  <SortableHeader columnKey="OperatingSystem" label="Operating System" />
                  <SortableHeader columnKey="LastLogon" label="Last Logon" />
                  <SortableHeader columnKey="Enabled" label="Status" />
                  <SortableHeader columnKey="Created" label="Created" />
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((comp, i) => (
                  <tr key={i}>
                    <td>{comp.Name}</td>
                    <td>{comp.OperatingSystem || '-'}</td>
                    <td>{formatLastLogon(comp.LastLogon)}</td>
                    <td>
                      <span className={`status-badge ${comp.Enabled ? 'enabled' : 'disabled'}`}>
                        {comp.Enabled ? 'Enabled' : 'Disabled'}
                      </span>
                    </td>
                    <td>{formatDate(comp.Created)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {renderPagination()}
        </div>
      )}
    </div>
  );
};

export default Computers;

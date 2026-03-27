import React, { useState } from 'react';
import axios from 'axios';
import ExportButton from '../common/ExportButton';
import { formatDate } from '../../utils/dateUtils';
import { useApp } from '../../context/AppContext';

const GPOs = () => {
  const { selectedDomain } = useApp();
  const [searchValue, setSearchValue] = useState('');
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  
  // Separate domain inputs for each report
  const [gposByDomainInput, setGposByDomainInput] = useState('');
  const [unlinkedDomainInput, setUnlinkedDomainInput] = useState('');
  const [disabledDomainInput, setDisabledDomainInput] = useState('');
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  const predefinedQueries = [
    { id: 'all', label: 'All GPOs', description: 'All Group Policy Objects in Forest' },
    { id: 'password-policies', label: 'Password Policies', description: 'Default & Fine-Grained policies' },
  ];

  const reportQueries = [
    { id: 'gpos-by-domain', label: 'GPOs by Domain', description: 'All GPOs in a specific domain', needsDomain: true, required: true },
    { id: 'unlinked-by-domain', label: 'Unlinked GPOs', description: 'GPOs not linked to any OU', needsDomain: true, required: false, placeholder: 'Domain (optional, blank=forest-wide)' },
    { id: 'disabled-by-domain', label: 'Disabled GPOs', description: 'GPOs that are disabled', needsDomain: true, required: false, placeholder: 'Domain (optional, blank=forest-wide)' },
  ];

  // Get domain input value based on query id
  const getDomainValue = (queryId) => {
    switch(queryId) {
      case 'gpos-by-domain': return gposByDomainInput;
      case 'unlinked-by-domain': return unlinkedDomainInput;
      case 'disabled-by-domain': return disabledDomainInput;
      default: return '';
    }
  };

  // Set domain input value based on query id
  const setDomainValue = (queryId, value) => {
    switch(queryId) {
      case 'gpos-by-domain': setGposByDomainInput(value); break;
      case 'unlinked-by-domain': setUnlinkedDomainInput(value); break;
      case 'disabled-by-domain': setDisabledDomainInput(value); break;
      default: break;
    }
  };

  const handleQuery = async (queryId, domainValue = null) => {
    const query = reportQueries.find(q => q.id === queryId) || predefinedQueries.find(q => q.id === queryId);
    const isRequired = query?.required === true;
    
    if (isRequired && (!domainValue || !domainValue.trim()) && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    try {
      const params = { query: queryId };
      if (domainValue && domainValue.trim()) {
        params.domain = domainValue.trim();
      }
      if (!params.domain && selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/gpos', { params });
      
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

  const handleSearch = async () => {
    if (!searchValue.trim()) return;
    setLoading(true);
    setError(null);
    setActiveQuery('search');
    setCurrentPage(1);
    try {
      const response = await axios.get('/api/gpos/search', { params: { q: searchValue } });
      
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

  // Pagination helpers
  const totalPages = Math.ceil(data.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = data.slice(startIndex, endIndex);

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

  const renderGPOsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Name</th>
          <th>Status</th>
          <th>Owner</th>
          <th>Created</th>
          <th>Modified</th>
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((gpo, i) => (
          <tr key={i}>
            <td>{gpo.Name}</td>
            <td>{gpo.Status}</td>
            <td>{gpo.Owner || '-'}</td>
            <td>{formatDate(gpo.CreationTime)}</td>
            <td>{formatDate(gpo.ModificationTime)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderPoliciesTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Domain</th>
          <th>Policy Name</th>
          <th>Type</th>
          <th>Precedence</th>
          <th>Complexity</th>
          <th>Min Length</th>
          <th>History</th>
          <th>Lockout Threshold</th>
          <th>Max Age (Days)</th>
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((policy, i) => (
          <tr key={i}>
            <td>{policy.Domain}</td>
            <td>{policy.PolicyName}</td>
            <td>
              <span className={`status-badge ${policy.Type === 'Default' ? 'enabled' : 'disabled'}`}>
                {policy.Type}
              </span>
            </td>
            <td>{policy.Precedence || '-'}</td>
            <td>{policy.Complexity ? 'Yes' : 'No'}</td>
            <td>{policy.MinPasswordLength}</td>
            <td>{policy.PasswordHistory}</td>
            <td>{policy.LockoutThreshold}</td>
            <td>{policy.MaxPasswordAgeDays}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderTable = () => {
    if (activeQuery === 'password-policies') {
      return renderPoliciesTable();
    }
    return renderGPOsTable();
  };

  return (
    <div className="gpos-page">
      <h1>Group Policy</h1>

      <div className="query-panel card">
        <h3>Standard Queries</h3>
        <div className="query-cards" style={{ 
          display: 'grid', 
          gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', 
          gap: '12px',
          marginBottom: '16px'
        }}>
          {predefinedQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handleQuery(q.id)}
              style={{
                padding: '16px',
                border: `2px solid ${activeQuery === q.id ? 'var(--accent-primary)' : 'transparent'}`,
                borderRadius: '8px',
                backgroundColor: activeQuery === q.id ? 'var(--accent-light)' : 'var(--bg-secondary)',
                cursor: 'pointer',
                transition: 'all 0.2s ease'
              }}
            >
              <div style={{ fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)', marginBottom: '4px' }}>
                {q.label}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                {q.description}
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* GPO Reports Section */}
      <div className="query-panel card">
        <h3>GPO Reports</h3>
        <div className="query-cards" style={{ 
          display: 'grid', 
          gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', 
          gap: '12px'
        }}>
          {reportQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              style={{
                padding: '16px',
                border: `2px solid ${activeQuery === q.id ? 'var(--accent-primary)' : 'transparent'}`,
                borderRadius: '8px',
                backgroundColor: activeQuery === q.id ? 'var(--accent-light)' : 'var(--bg-secondary)',
                cursor: 'default',
                transition: 'all 0.2s ease'
              }}
            >
              <div style={{ fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)', marginBottom: '4px' }}>
                {q.label}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '10px' }}>
                {q.description}
              </div>
              
              <div className="query-input" onClick={(e) => e.stopPropagation()} style={{ display: 'flex', gap: '8px', marginTop: '8px' }}>
                <input
                  type="text"
                  placeholder={q.placeholder || "Enter domain (e.g., corp.domain.com)"}
                  value={getDomainValue(q.id)}
                  onChange={(e) => setDomainValue(q.id, e.target.value)}
                  onKeyPress={(e) => e.key === 'Enter' && handleQuery(q.id, getDomainValue(q.id))}
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
                <button className="btn btn-primary" onClick={() => handleQuery(q.id, getDomainValue(q.id))} style={{ padding: '8px 16px' }}>
                  Run
                </button>
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="search-panel card">
        <div className="search-input-group">
          <input
            type="text"
            placeholder="Search GPOs..."
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
            <span>Found {data.length} total result(s)</span>
            <ExportButton data={data} filename={`gpo_${activeQuery || 'export'}`} title="GPO Report" />
          </div>
          <div style={{ overflowX: 'auto' }}>
            {renderTable()}
          </div>
          {renderPagination()}
        </div>
      )}
    </div>
  );
};

export default GPOs;

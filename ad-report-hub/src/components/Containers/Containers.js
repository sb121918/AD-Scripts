import React, { useState } from 'react';
import axios from 'axios';
import ExportButton from '../common/ExportButton';
import OUTree from './OUTree';
import { useApp } from '../../context/AppContext';

const Containers = () => {
  const { selectedDomain } = useApp();
  const [searchValue, setSearchValue] = useState('');
  const [domain, setDomain] = useState('');
  const [ous, setOUs] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [viewMode, setViewMode] = useState('list'); // 'list' | 'tree'
  const [treeData, setTreeData] = useState(null);
  const [treeLoading, setTreeLoading] = useState(false);
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  const predefinedQueries = [
    { id: 'all-ous', label: 'All Organizational Units', description: 'All OUs in forest or specified domain' },
    { id: 'empty', label: 'Empty OUs', description: 'OUs with no child objects' },
    { id: 'protected', label: 'Protected OUs', description: 'OUs with deletion protection enabled' },
  ];

  const handleSearch = async () => {
    if (!searchValue.trim()) return;
    setLoading(true);
    setCurrentPage(1); // Reset pagination
    try {
      const response = await axios.get('/api/containers/search', { params: { q: searchValue } });
      setOUs(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
    } finally {
      setLoading(false);
    }
  };

  const handlePredefinedQuery = async (queryId) => {
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1); // Reset pagination
    try {
      const params = { query: queryId };
      const effectiveDomain = domain.trim() || selectedDomain;
      if (effectiveDomain) {
        params.domain = effectiveDomain;
      }
      const response = await axios.get('/api/containers', { params });
      setOUs(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
    } finally {
      setLoading(false);
    }
  };

  const handleLoadTree = async () => {
    setTreeLoading(true);
    setError(null);
    try {
      const params = { query: 'ou-tree' };
      const effectiveDomain = domain.trim() || selectedDomain;
      if (effectiveDomain) {
        params.domain = effectiveDomain;
      }
      const response = await axios.get('/api/containers', { params });
      // Backend may return either a root object or [rootObject]
      const raw = response.data.data;
      setTreeData(Array.isArray(raw) ? (raw[0] || null) : raw);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to load OU tree');
    } finally {
      setTreeLoading(false);
    }
  };

  const handleViewModeChange = (mode) => {
    setViewMode(mode);
    if (mode === 'tree' && !treeData) {
      handleLoadTree();
    }
  };

  // Pagination helpers
  const totalPages = Math.ceil(ous.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = ous.slice(startIndex, endIndex);

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
    if (ous.length === 0) return null;
    const startItem = startIndex + 1;
    const endItem = Math.min(endIndex, ous.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {ous.length} results
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
    <div className="containers-page">
      <h1>Containers & OUs</h1>

      {/* View Mode Toggle */}
      <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
        <button
          className={`btn ${viewMode === 'list' ? 'btn-primary' : 'btn-secondary'}`}
          onClick={() => handleViewModeChange('list')}
        >
          &#x1F4CB; List View
        </button>
        <button
          className={`btn ${viewMode === 'tree' ? 'btn-primary' : 'btn-secondary'}`}
          onClick={() => handleViewModeChange('tree')}
        >
          &#x1F333; Tree View
        </button>
      </div>

      {/* Tree View */}
      {viewMode === 'tree' && (
        <div className="card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>OU Hierarchy</span>
            <button className="btn btn-secondary" onClick={handleLoadTree} style={{ fontSize: '13px' }}>
              Refresh Tree
            </button>
          </div>
          {treeLoading ? (
            <div className="loading" style={{ padding: '40px' }}>
              <div className="spinner"></div>
              <span>Loading OU tree...</span>
            </div>
          ) : treeData ? (
            <OUTree treeData={treeData} domain={domain.trim() || selectedDomain} />
          ) : (
            <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-secondary)' }}>
              Click "Refresh Tree" or switch to tree view to load the OU hierarchy.
            </div>
          )}
        </div>
      )}

      {/* List View */}
      {viewMode === 'list' && (
      <>
      <div className="query-panel card">
        <h3>Predefined Queries</h3>
        
        {/* Domain Selection */}
        <div className="domain-input-section" style={{ marginBottom: '16px' }}>
          <label style={{ 
            display: 'block', 
            marginBottom: '8px', 
            color: 'var(--text-secondary)', 
            fontSize: '13px' 
          }}>
            Target Domain (optional - leave empty for forest-wide query)
          </label>
          <input
            type="text"
            placeholder="Enter domain (e.g., corp.domain.com)"
            value={domain}
            onChange={(e) => setDomain(e.target.value)}
            style={{
              width: '100%',
              maxWidth: '400px',
              padding: '10px 14px',
              border: '1px solid var(--border-input)',
              borderRadius: '6px',
              backgroundColor: 'var(--bg-input)',
              color: 'var(--text-primary)',
              fontSize: '14px'
            }}
          />
        </div>

        <div className="query-cards" style={{ 
          display: 'grid', 
          gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', 
          gap: '12px' 
        }}>
          {predefinedQueries.map((q) => (
            <div 
              key={q.id} 
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handlePredefinedQuery(q.id)}
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

      <div className="search-panel card">
        <div className="search-input-group">
          <input
            type="text"
            placeholder="Search OUs..."
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

      {!loading && ous.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header">
            <span>Found {ous.length} total OU(s)</span>
            <ExportButton data={ous} filename={`containers_${activeQuery || 'export'}`} title="Containers & OUs Report" />
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Description</th>
                  <th>Protected</th>
                  <th>Created</th>
                  <th>Distinguished Name</th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((ou, i) => (
                  <tr key={i}>
                    <td>{ou.Name}</td>
                    <td>{ou.Description || '-'}</td>
                    <td>{ou.Protected ? 'Yes' : 'No'}</td>
                    <td>{ou.Created ? new Date(ou.Created).toLocaleDateString() : '-'}</td>
                    <td style={{ fontSize: '12px', maxWidth: '300px', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                      {ou.DistinguishedName}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {renderPagination()}
        </div>
      )}
      </>
      )}

      {error && viewMode === 'tree' && <div className="error-message">{error}</div>}
    </div>
  );
};

export default Containers;

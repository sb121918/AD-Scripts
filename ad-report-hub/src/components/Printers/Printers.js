import React, { useState } from 'react';
import axios from 'axios';
import ExportButton from '../common/ExportButton';
import { useSortableData } from '../../utils/useSortableData';
import { useApp } from '../../context/AppContext';

const Printers = () => {
  const { selectedDomain } = useApp();
  const [searchValue, setSearchValue] = useState('');
  const [domainInput, setDomainInput] = useState('');
  const [printers, setPrinters] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  // Sorting
  const { sortedItems, requestSort, getSortIndicator, getSortClass } = useSortableData(printers);

  const predefinedQueries = [
    { id: 'by-domain', label: 'Printers by Domain', description: 'Get all printers from a specific domain', needsDomain: true },
    { id: 'search', label: 'Search Printers', description: 'Search printers by name or location', needsSearch: true },
  ];

  const handleQuery = async (queryId) => {
    if (queryId === 'by-domain' && !domainInput.trim() && !selectedDomain) {
      setError('Please enter a domain name');
      return;
    }
    
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    
    try {
      const params = { query: queryId };
      if (queryId === 'by-domain') params.domain = domainInput || selectedDomain;
      if (!params.domain && selectedDomain) params.domain = selectedDomain;
      
      const response = await axios.get('/api/printers', { params });
      setPrinters(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setPrinters([]);
    } finally {
      setLoading(false);
    }
  };

  const handleSearch = async () => {
    if (!searchValue.trim()) {
      setError('Please enter a search term');
      return;
    }
    setLoading(true);
    setError(null);
    setActiveQuery('search');
    setCurrentPage(1);
    try {
      const response = await axios.get('/api/printers/search', { params: { q: searchValue } });
      setPrinters(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
      setPrinters([]);
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

  return (
    <div className="printers-page">
      <h1>Printers</h1>

      <div className="query-panel card">
        <h3>Predefined Queries</h3>
        <div className="query-cards">
          {predefinedQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => !q.needsDomain && !q.needsSearch && handleQuery(q.id)}
            >
              <div className="query-card-title">{q.label}</div>
              <div className="query-card-desc">{q.description}</div>
              
              {q.needsDomain && (
                <div className="query-input" onClick={(e) => e.stopPropagation()}>
                  <input
                    type="text"
                    placeholder="Enter domain (e.g., corp.domain.com)"
                    value={domainInput}
                    onChange={(e) => setDomainInput(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleQuery(q.id)}
                  />
                  <button className="btn btn-primary btn-sm" onClick={() => handleQuery(q.id)}>
                    Run
                  </button>
                </div>
              )}
              
              {q.needsSearch && (
                <div className="query-input" onClick={(e) => e.stopPropagation()}>
                  <input
                    type="text"
                    placeholder="Search printers by name or location..."
                    value={searchValue}
                    onChange={(e) => setSearchValue(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleSearch()}
                  />
                  <button className="btn btn-primary btn-sm" onClick={handleSearch}>
                    Search
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
          <span>Loading printers...</span>
        </div>
      )}

      {!loading && printers.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <span>Found {printers.length} total printer(s)</span>
            <ExportButton data={printers} filename={`printers_${activeQuery || 'export'}`} title="Printers Report" />
          </div>
          <p style={{ padding: '8px 20px', margin: 0, fontSize: '12px', color: 'var(--text-secondary)' }}>
            Click on column headers to sort A-Z or Z-A
          </p>
          <div style={{ overflowX: 'auto' }}>
            <table className="data-table">
              <thead>
                <tr>
                  {renderSortableHeader('Name', 'Name')}
                  {renderSortableHeader('Server', 'Server')}
                  {renderSortableHeader('Location', 'Location')}
                  {renderSortableHeader('Domain', 'Domain')}
                  {renderSortableHeader('Color', 'Color')}
                  {renderSortableHeader('Duplex', 'Duplex')}
                  {renderSortableHeader('Published', 'Published')}
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((printer, i) => (
                  <tr key={i}>
                    <td>{printer.Name}</td>
                    <td>{printer.Server || '-'}</td>
                    <td>{printer.Location || '-'}</td>
                    <td>{printer.Domain || '-'}</td>
                    <td>
                      <span className={`status-badge ${printer.Color ? 'enabled' : 'disabled'}`}>
                        {printer.Color ? 'Yes' : 'No'}
                      </span>
                    </td>
                    <td>
                      <span className={`status-badge ${printer.Duplex ? 'enabled' : 'disabled'}`}>
                        {printer.Duplex ? 'Yes' : 'No'}
                      </span>
                    </td>
                    <td>
                      <span className={`status-badge ${printer.Published ? 'enabled' : 'disabled'}`}>
                        {printer.Published ? 'Yes' : 'No'}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {renderPagination()}
        </div>
      )}

      {!loading && activeQuery && printers.length === 0 && !error && (
        <div className="no-results card">
          <p>No printers found matching the criteria.</p>
        </div>
      )}
    </div>
  );
};

export default Printers;

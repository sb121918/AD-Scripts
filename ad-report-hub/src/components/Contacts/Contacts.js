import React, { useState } from 'react';
import axios from 'axios';
import ExportButton from '../common/ExportButton';

const Contacts = () => {
  const [searchValue, setSearchValue] = useState('');
  const [contacts, setContacts] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  const handleSearch = async () => {
    if (!searchValue.trim()) return;
    setLoading(true);
    setCurrentPage(1); // Reset pagination
    try {
      const response = await axios.get('/api/contacts/search', { params: { q: searchValue } });
      setContacts(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
    } finally {
      setLoading(false);
    }
  };

  const handleLoadAll = async () => {
    setLoading(true);
    setCurrentPage(1); // Reset pagination
    try {
      const response = await axios.get('/api/contacts', { params: { query: 'all' } });
      setContacts(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to load contacts');
    } finally {
      setLoading(false);
    }
  };

  // Pagination helpers
  const totalPages = Math.ceil(contacts.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = contacts.slice(startIndex, endIndex);

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
    if (contacts.length === 0) return null;
    const startItem = startIndex + 1;
    const endItem = Math.min(endIndex, contacts.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {contacts.length} results
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
    <div className="contacts-page">
      <h1>Contacts</h1>

      <div className="action-bar">
        <button className="btn btn-primary" onClick={handleLoadAll}>Load All Contacts</button>
      </div>

      <div className="search-panel card">
        <div className="search-input-group">
          <input
            type="text"
            placeholder="Search contacts..."
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

      {!loading && contacts.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header">
            <span>Found {contacts.length} total contact(s)</span>
            <ExportButton data={contacts} filename="contacts_export" title="Contacts Report" />
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table className="data-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Email</th>
                  <th>Company</th>
                  <th>Department</th>
                  <th>Phone</th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((contact, i) => (
                  <tr key={i}>
                    <td>{contact.DisplayName || contact.Name}</td>
                    <td>{contact.Email || '-'}</td>
                    <td>{contact.Company || '-'}</td>
                    <td>{contact.Department || '-'}</td>
                    <td>{contact.Phone || '-'}</td>
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

export default Contacts;

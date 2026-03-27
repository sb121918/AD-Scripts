import React, { useState, useCallback } from 'react';
import axios from 'axios';
import QueryBuilder, { buildLdapFilter } from './QueryBuilder';
import SavedSearches from './SavedSearches';
import SearchHistory, { addToHistory } from './SearchHistory';
import ExportButton from '../common/ExportButton';
import './Search.css';

const Search = () => {
  const [searchMode, setSearchMode] = useState('simple'); // 'simple' | 'query-builder' | 'raw-ldap'
  const [searchValue, setSearchValue] = useState('');
  const [rawLdapFilter, setRawLdapFilter] = useState('');
  const [rawLdapObjectType, setRawLdapObjectType] = useState('user');
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [objectType, setObjectType] = useState('user');
  const [objectTypes, setObjectTypes] = useState({
    user: true,
    computer: true,
    group: true,
    contact: false,
    ou: false,
    gpo: false
  });
  const [conditions, setConditions] = useState([
    { attribute: 'samAccountName', operator: 'contains', value: '', logic: 'AND' }
  ]);
  const [logicBetween, setLogicBetween] = useState('AND');
  const [ldapFilter, setLdapFilter] = useState('');
  const [saveModal, setSaveModal] = useState({ open: false, data: null });
  const [searchHistory, setSearchHistory] = useState(() => {
    try {
      const saved = localStorage.getItem('ad-search-history');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  const [pagination, setPagination] = useState(null);
  const [lastSearchParams, setLastSearchParams] = useState(null);

  const persistHistory = useCallback((history) => {
    setSearchHistory(history);
    localStorage.setItem('ad-search-history', JSON.stringify(history));
  }, []);

  const executeSearch = async (page = 1, size = pageSize) => {
    if (searchMode === 'simple') {
      if (!searchValue.trim()) return;
    } else if (searchMode === 'query-builder') {
      const filter = buildLdapFilter(conditions, objectType, logicBetween);
      if (!filter) return;
    } else if (searchMode === 'raw-ldap') {
      if (!rawLdapFilter.trim()) return;
    }

    setLoading(true);
    setError(null);
    try {
      let params = { pageSize: size, page };
      let searchParamsForHistory = {};

      if (searchMode === 'simple') {
        const types = Object.entries(objectTypes)
          .filter(([_, enabled]) => enabled)
          .map(([type]) => type)
          .join(',');
        params = { ...params, q: searchValue, types };
        searchParamsForHistory = { mode: 'simple', query: searchValue, types };
      } else if (searchMode === 'query-builder') {
        const filter = buildLdapFilter(conditions, objectType, logicBetween);
        params = { ...params, ldapFilter: filter, objectType };
        searchParamsForHistory = { mode: 'ldap', ldapFilter: filter, types: objectType };
      } else {
        params = { ...params, ldapFilter: rawLdapFilter, objectType: rawLdapObjectType };
        searchParamsForHistory = { mode: 'ldap', ldapFilter: rawLdapFilter, types: rawLdapObjectType };
      }

      const response = await axios.get('/api/search', { params });
      setResults(response.data.data);
      setPagination(response.data.data?.Pagination || null);
      setCurrentPage(page);
      setLastSearchParams(params);

      const newHistory = addToHistory(searchHistory, searchParamsForHistory);
      persistHistory(newHistory);
    } catch (err) {
      setError(err.response?.data?.error || 'Search failed');
    } finally {
      setLoading(false);
    }
  };

  const handleSearch = () => {
    setCurrentPage(1);
    executeSearch(1, pageSize);
  };

  const handlePageSizeChange = (newSize) => {
    setPageSize(newSize);
    setCurrentPage(1);
    if (lastSearchParams) {
      executeSearch(1, newSize);
    }
  };

  const handleNextPage = () => {
    executeSearch(currentPage + 1, pageSize);
  };

  const handlePrevPage = () => {
    if (currentPage > 1) {
      executeSearch(currentPage - 1, pageSize);
    }
  };

  const handleLoadSavedSearch = (saved) => {
    if ((saved.mode === 'ldap' || saved.ldapFilter) && saved.conditions) {
      setSearchMode('query-builder');
      setConditions(saved.conditions);
      setObjectType(saved.objectType || 'user');
      setLogicBetween(saved.logicBetween || 'AND');
    } else if (saved.ldapFilter && !saved.conditions) {
      setSearchMode('raw-ldap');
      setRawLdapFilter(saved.ldapFilter);
      setRawLdapObjectType(saved.objectType || 'user');
    } else {
      setSearchMode('simple');
      setSearchValue(saved.query || '');
      setObjectTypes(saved.objectTypes || { user: true, computer: true, group: true });
    }
  };

  const handleSaveSearch = () => {
    let data = {};
    if (searchMode === 'simple') {
      data = { mode: 'simple', query: searchValue, objectTypes: { ...objectTypes } };
    } else if (searchMode === 'query-builder') {
      data = { mode: 'ldap', conditions: [...conditions], objectType, logicBetween, ldapFilter: buildLdapFilter(conditions, objectType, logicBetween) };
    } else {
      data = { mode: 'ldap', ldapFilter: rawLdapFilter };
    }
    setSaveModal({ open: true, data });
  };

  const getTotalPages = (totalCount) => {
    return Math.ceil((totalCount || 0) / pageSize) || 1;
  };

  const renderPagination = (totalCount, typeName) => {
    if (!totalCount || totalCount === 0) return null;
    const totalPages = getTotalPages(totalCount);
    const startItem = ((currentPage - 1) * pageSize) + 1;
    const endItem = Math.min(currentPage * pageSize, totalCount);

    return (
      <div className="pagination" style={{ borderTop: '1px solid var(--border-color)', padding: '12px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
        <div className="pagination-info" style={{ color: 'var(--text-secondary)', fontSize: '14px' }}>
          Showing {startItem}-{endItem} of {totalCount} {typeName}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <label style={{ color: 'var(--text-secondary)', fontSize: '13px' }}>Per page:</label>
            <select
              value={pageSize}
              onChange={(e) => handlePageSizeChange(parseInt(e.target.value))}
              style={{ padding: '4px 8px', border: '1px solid var(--border-input)', borderRadius: '4px', backgroundColor: 'var(--bg-input)', color: 'var(--text-primary)', fontSize: '13px' }}
            >
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={200}>200</option>
              <option value={500}>500</option>
              <option value={1000}>1000</option>
            </select>
          </div>
          <div className="pagination-buttons" style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
            <button
              onClick={handlePrevPage}
              disabled={currentPage <= 1 || loading}
              className="btn btn-secondary"
              style={{ padding: '6px 12px' }}
            >
              Back
            </button>
            <span style={{ padding: '0 12px', color: 'var(--text-secondary)' }}>
              Page {currentPage} of {totalPages}
            </span>
            <button
              onClick={handleNextPage}
              disabled={currentPage >= totalPages || loading}
              className="btn btn-secondary"
              style={{ padding: '6px 12px' }}
            >
              Next
            </button>
          </div>
        </div>
      </div>
    );
  };

  const getActiveTypes = () => {
    if (searchMode === 'query-builder') return [objectType];
    return Object.entries(objectTypes).filter(([_, v]) => v).map(([k]) => k);
  };

  const activeTypes = getActiveTypes();

  return (
    <div className="search-page">
      <h1>Global Search</h1>

      <div className="search-mode-tabs">
        <button
          type="button"
          className={`mode-tab ${searchMode === 'simple' ? 'active' : ''}`}
          onClick={() => setSearchMode('simple')}
        >
          Simple Search
        </button>
        <button
          type="button"
          className={`mode-tab ${searchMode === 'query-builder' ? 'active' : ''}`}
          onClick={() => setSearchMode('query-builder')}
        >
          Query Builder
        </button>
        <button
          type="button"
          className={`mode-tab ${searchMode === 'raw-ldap' ? 'active' : ''}`}
          onClick={() => setSearchMode('raw-ldap')}
        >
          Raw LDAP
        </button>
      </div>

      <div className="search-layout">
        <div className="search-main">
          <div className="search-panel card">
            {searchMode === 'simple' && (
              <>
                <div className="search-input-group" style={{ marginBottom: '16px', display: 'flex', gap: '12px' }}>
                  <input
                    type="text"
                    placeholder="Enter search term..."
                    value={searchValue}
                    onChange={(e) => setSearchValue(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleSearch()}
                    style={{ flex: 1, padding: '10px 14px', border: '1px solid var(--border-input)', borderRadius: '4px', backgroundColor: 'var(--bg-input)', color: 'var(--text-primary)' }}
                  />
                  <button className="btn btn-primary" onClick={handleSearch}>SEARCH</button>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
                  <div className="object-types" style={{ display: 'flex', gap: '20px', flexWrap: 'wrap' }}>
                    {['user', 'computer', 'group', 'contact', 'ou', 'gpo'].map((t) => (
                      <label key={t} style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-primary)' }}>
                        <input
                          type="checkbox"
                          checked={objectTypes[t]}
                          onChange={(e) => setObjectTypes({ ...objectTypes, [t]: e.target.checked })}
                        />
                        {t.charAt(0).toUpperCase() + t.slice(1)}s
                      </label>
                    ))}
                  </div>
                </div>
              </>
            )}

            {searchMode === 'query-builder' && (
              <>
                <div className="query-builder-type" style={{ marginBottom: '16px' }}>
                  <label style={{ color: 'var(--text-secondary)', marginRight: '8px' }}>Search in:</label>
                  <select
                    value={objectType}
                    onChange={(e) => {
                      setObjectType(e.target.value);
                      setConditions([{ attribute: 'samAccountName', operator: 'contains', value: '', logic: 'AND' }]);
                    }}
                    style={{ padding: '8px 12px', border: '1px solid var(--border-input)', borderRadius: '4px', backgroundColor: 'var(--bg-input)', color: 'var(--text-primary)' }}
                  >
                    <option value="user">Users</option>
                    <option value="computer">Computers</option>
                    <option value="group">Groups</option>
                    <option value="contact">Contacts</option>
                    <option value="ou">Organizational Units</option>
                    <option value="gpo">Group Policy Objects</option>
                  </select>
                </div>
                <QueryBuilder
                  objectType={objectType}
                  conditions={conditions}
                  onConditionsChange={setConditions}
                  onLdapFilterChange={setLdapFilter}
                  logicBetween={logicBetween}
                  onLogicBetweenChange={setLogicBetween}
                />
                <button className="btn btn-primary" onClick={handleSearch} disabled={!buildLdapFilter(conditions, objectType, logicBetween)}>
                  SEARCH
                </button>
              </>
            )}

            {searchMode === 'raw-ldap' && (
              <>
                <div className="search-input-group" style={{ marginBottom: '16px' }}>
                  <label style={{ display: 'block', marginBottom: '8px', color: 'var(--text-secondary)' }}>
                    LDAP Filter (e.g. (samAccountName=*john*))
                  </label>
                  <input
                    type="text"
                    placeholder="(samAccountName=*value*) or (&(objectClass=user)(department=IT))"
                    value={rawLdapFilter}
                    onChange={(e) => setRawLdapFilter(e.target.value)}
                    onKeyPress={(e) => e.key === 'Enter' && handleSearch()}
                    style={{ width: '100%', padding: '10px 14px', border: '1px solid var(--border-input)', borderRadius: '4px', backgroundColor: 'var(--bg-input)', color: 'var(--text-primary)', fontFamily: 'monospace' }}
                  />
                </div>
                <div style={{ marginBottom: '16px' }}>
                  <label style={{ color: 'var(--text-secondary)', marginRight: '8px' }}>Object type:</label>
                  <select
                    value={rawLdapObjectType}
                    onChange={(e) => setRawLdapObjectType(e.target.value)}
                    style={{ padding: '8px 12px', border: '1px solid var(--border-input)', borderRadius: '4px', backgroundColor: 'var(--bg-input)', color: 'var(--text-primary)' }}
                  >
                    <option value="user">Users</option>
                    <option value="computer">Computers</option>
                    <option value="group">Groups</option>
                    <option value="contact">Contacts</option>
                    <option value="ou">Organizational Units</option>
                    <option value="gpo">Group Policy Objects</option>
                  </select>
                </div>
                <button className="btn btn-primary" onClick={handleSearch} disabled={!rawLdapFilter.trim()}>
                  SEARCH
                </button>
              </>
            )}

            <div style={{ marginTop: '16px', display: 'flex', gap: '12px', alignItems: 'center' }}>
              <label style={{ color: 'var(--text-secondary)', fontSize: '14px' }}>Results per page:</label>
              <select
                value={pageSize}
                onChange={(e) => handlePageSizeChange(parseInt(e.target.value))}
                style={{ padding: '6px 12px', border: '1px solid var(--border-input)', borderRadius: '4px', backgroundColor: 'var(--bg-input)', color: 'var(--text-primary)' }}
              >
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
                <option value={200}>200</option>
                <option value={500}>500</option>
                <option value={1000}>1000</option>
              </select>
              <button className="btn btn-secondary" onClick={handleSaveSearch} style={{ marginLeft: 'auto' }}>
                Save Search
              </button>
            </div>
          </div>

          {error && <div className="error-message">{error}</div>}

          {loading && (
            <div className="loading">
              <div className="spinner"></div>
              <span>Searching...</span>
            </div>
          )}

          {results && (
            <div className="search-results">
              {activeTypes.includes('user') && (pagination?.TotalUsers > 0 || results.Users?.length > 0) && (
                <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
                  <div className="results-header" style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span>Users ({pagination?.TotalUsers || results.Users?.length || 0} total)</span>
                    <ExportButton data={results.Users} filename="search_users" title="Search Results - Users" />
                  </div>
                  {results.Users?.length > 0 ? (
                    <>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>SAM Account</th>
                              <th>Email</th>
                              <th>Department</th>
                            </tr>
                          </thead>
                          <tbody>
                            {results.Users.map((user, i) => (
                              <tr key={i}>
                                <td>{user.Name}</td>
                                <td>{user.SamAccountName}</td>
                                <td>{user.Email || '-'}</td>
                                <td>{user.Department || '-'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {renderPagination(pagination?.TotalUsers, 'users')}
                    </>
                  ) : (
                    <div className="no-results">No users found on this page</div>
                  )}
                </div>
              )}

              {activeTypes.includes('computer') && (pagination?.TotalComputers > 0 || results.Computers?.length > 0) && (
                <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
                  <div className="results-header" style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span>Computers ({pagination?.TotalComputers || results.Computers?.length || 0} total)</span>
                    <ExportButton data={results.Computers} filename="search_computers" title="Search Results - Computers" />
                  </div>
                  {results.Computers?.length > 0 ? (
                    <>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>Operating System</th>
                              <th>Status</th>
                            </tr>
                          </thead>
                          <tbody>
                            {results.Computers.map((comp, i) => (
                              <tr key={i}>
                                <td>{comp.Name}</td>
                                <td>{comp.OperatingSystem || '-'}</td>
                                <td>{comp.Enabled ? 'Enabled' : 'Disabled'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {renderPagination(pagination?.TotalComputers, 'computers')}
                    </>
                  ) : (
                    <div className="no-results">No computers found on this page</div>
                  )}
                </div>
              )}

              {activeTypes.includes('group') && (pagination?.TotalGroups > 0 || results.Groups?.length > 0) && (
                <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
                  <div className="results-header" style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span>Groups ({pagination?.TotalGroups || results.Groups?.length || 0} total)</span>
                    <ExportButton data={results.Groups} filename="search_groups" title="Search Results - Groups" />
                  </div>
                  {results.Groups?.length > 0 ? (
                    <>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>Type</th>
                              <th>Email</th>
                            </tr>
                          </thead>
                          <tbody>
                            {results.Groups.map((group, i) => (
                              <tr key={i}>
                                <td>{group.Name}</td>
                                <td>{group.Type}</td>
                                <td>{group.Email || '-'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {renderPagination(pagination?.TotalGroups, 'groups')}
                    </>
                  ) : (
                    <div className="no-results">No groups found on this page</div>
                  )}
                </div>
              )}

              {activeTypes.includes('contact') && (pagination?.TotalContacts > 0 || results.Contacts?.length > 0) && (
                <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
                  <div className="results-header" style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span>Contacts ({pagination?.TotalContacts || results.Contacts?.length || 0} total)</span>
                    <ExportButton data={results.Contacts} filename="search_contacts" title="Search Results - Contacts" />
                  </div>
                  {results.Contacts?.length > 0 ? (
                    <>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>Email</th>
                              <th>Company</th>
                              <th>Department</th>
                            </tr>
                          </thead>
                          <tbody>
                            {results.Contacts.map((contact, i) => (
                              <tr key={i}>
                                <td>{contact.Name}</td>
                                <td>{contact.Email || '-'}</td>
                                <td>{contact.Company || '-'}</td>
                                <td>{contact.Department || '-'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {renderPagination(pagination?.TotalContacts, 'contacts')}
                    </>
                  ) : (
                    <div className="no-results">No contacts found on this page</div>
                  )}
                </div>
              )}

              {activeTypes.includes('ou') && (pagination?.TotalOUs > 0 || results.OUs?.length > 0) && (
                <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
                  <div className="results-header" style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span>Organizational Units ({pagination?.TotalOUs || results.OUs?.length || 0} total)</span>
                    <ExportButton data={results.OUs} filename="search_ous" title="Search Results - OUs" />
                  </div>
                  {results.OUs?.length > 0 ? (
                    <>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>Description</th>
                              <th>Distinguished Name</th>
                            </tr>
                          </thead>
                          <tbody>
                            {results.OUs.map((ou, i) => (
                              <tr key={i}>
                                <td>{ou.Name}</td>
                                <td>{ou.Description || '-'}</td>
                                <td style={{ fontSize: '12px', maxWidth: '300px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{ou.DistinguishedName || '-'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {renderPagination(pagination?.TotalOUs, 'OUs')}
                    </>
                  ) : (
                    <div className="no-results">No OUs found on this page</div>
                  )}
                </div>
              )}

              {activeTypes.includes('gpo') && (pagination?.TotalGPOs > 0 || results.GPOs?.length > 0) && (
                <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
                  <div className="results-header" style={{ padding: '16px 20px', borderBottom: '1px solid var(--border-color)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span>Group Policy Objects ({pagination?.TotalGPOs || results.GPOs?.length || 0} total)</span>
                    <ExportButton data={results.GPOs} filename="search_gpos" title="Search Results - GPOs" />
                  </div>
                  {results.GPOs?.length > 0 ? (
                    <>
                      <div style={{ overflowX: 'auto' }}>
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Name</th>
                              <th>Status</th>
                              <th>Created</th>
                            </tr>
                          </thead>
                          <tbody>
                            {results.GPOs.map((gpo, i) => (
                              <tr key={i}>
                                <td>{gpo.Name}</td>
                                <td>{gpo.Status || '-'}</td>
                                <td>{gpo.Created || '-'}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      {renderPagination(pagination?.TotalGPOs, 'GPOs')}
                    </>
                  ) : (
                    <div className="no-results">No GPOs found on this page</div>
                  )}
                </div>
              )}

              {results.Users?.length === 0 && results.Computers?.length === 0 && results.Groups?.length === 0 && !results.Contacts?.length && !results.OUs?.length && !results.GPOs?.length && (
                <div className="no-results card" style={{ textAlign: 'center', padding: '40px' }}>
                  No results found
                </div>
              )}
            </div>
          )}
        </div>

        <aside className="search-sidebar">
          <SavedSearches
            onLoadSearch={handleLoadSavedSearch}
            saveModal={saveModal}
            onSaveComplete={() => setSaveModal({ open: false, data: null })}
          />
          <SearchHistory
            history={searchHistory}
            onSelect={(item) => {
              if (item.mode === 'ldap' && item.ldapFilter) {
                setSearchMode('raw-ldap');
                setRawLdapFilter(item.ldapFilter);
                setRawLdapObjectType(item.types || 'user');
              } else if (item.query) {
                setSearchMode('simple');
                setSearchValue(item.query);
              }
            }}
          />
        </aside>
      </div>
    </div>
  );
};

export default Search;

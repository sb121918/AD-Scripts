import React, { useState, useMemo } from 'react';
import axios from 'axios';
import './Topology.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import TopologyDiagram from './TopologyDiagram';
import { useApp } from '../../context/AppContext';
import { formatDate } from '../../utils/dateUtils';

const Topology = () => {
  const { selectedDomain } = useApp();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [viewMode, setViewMode] = useState('table'); // 'table' | 'diagram'
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  
  // Sorting state
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  // Trusts view filters
  const [trustSidFilter, setTrustSidFilter] = useState('all'); // 'all' | 'enabled' | 'disabled'

  const predefinedQueries = [
    { id: 'forest-info', label: 'Forest Information', description: 'AD Forest details and functional levels' },
    { id: 'domains', label: 'All Domains', description: 'All domains in the forest' },
    { id: 'trusts', label: 'AD Trusts', description: 'Forest and domain trust relationships' },
    { id: 'naming-contexts', label: 'Naming Contexts', description: 'Directory partitions and naming contexts' },
    { id: 'fsmo-roles', label: 'FSMO Role Holders', description: 'Flexible Single Master Operations roles' },
    { id: 'ad-sites', label: 'AD Sites Summary', description: 'Sites with DC/subnet counts and site link connectivity' },
    { id: 'site-links', label: 'Site-to-Site Replication', description: 'AD Site link replication topology' },
    { id: 'dc-replication', label: 'DC Replication Status', description: 'DC-to-DC replication connections and status' },
  ];

  const handleQuery = async (queryId) => {
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    setSortConfig({ key: null, direction: 'asc' });
    if (queryId === 'trusts') {
      setTrustSidFilter('all');
    }
    try {
      const params = { query: queryId };
      if (selectedDomain) {
        params.domain = selectedDomain;
      }
      const response = await axios.get('/api/topology', { params });
      setData(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

  // Optional filter/augmentation (used by AD Trusts view)
  const filteredData = useMemo(() => {
    if (!data || !Array.isArray(data)) return [];
    if (activeQuery !== 'trusts') return data;

    const augmented = data
      .filter(t => t && typeof t === 'object')
      .map((t) => {
        const attrs = String(t.TrustAttributes || '');
        const sidFiltered = attrs.includes('SID Filtered');
        return { ...t, SidFilteringEnabled: sidFiltered };
      });

    if (trustSidFilter === 'enabled') return augmented.filter(t => t.SidFilteringEnabled);
    if (trustSidFilter === 'disabled') return augmented.filter(t => !t.SidFilteringEnabled);
    return augmented;
  }, [data, activeQuery, trustSidFilter]);

  // Sorting logic
  const sortedData = useMemo(() => {
    if (!filteredData || !Array.isArray(filteredData)) return [];
    let sortableItems = [...filteredData];
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
  }, [filteredData, sortConfig]);

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
    if (currentPage < totalPages) setCurrentPage(currentPage + 1);
  };

  const handlePrevPage = () => {
    if (currentPage > 1) setCurrentPage(currentPage - 1);
  };

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

  const renderForestInfoTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Property" label="Property" />
          <SortableHeader columnKey="Value" label="Value" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.Property}</strong></td>
            <td>{item.Value || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderDomainsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="DomainName" label="Domain Name" />
          <SortableHeader columnKey="NetBIOSName" label="NetBIOS Name" />
          <SortableHeader columnKey="DomainMode" label="Domain Mode" />
          <SortableHeader columnKey="ParentDomain" label="Parent Domain" />
          <SortableHeader columnKey="ChildDomains" label="Child Domains" />
          <SortableHeader columnKey="DCCount" label="DC Count" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((domain, i) => (
          <tr key={i}>
            <td>{domain.DomainName}</td>
            <td>{domain.NetBIOSName || '-'}</td>
            <td><span className="mode-badge">{domain.DomainMode || '-'}</span></td>
            <td>{domain.ParentDomain || 'Root'}</td>
            <td>{domain.ChildDomains || '-'}</td>
            <td>{domain.DCCount || 0}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderTrustsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SourceDomain" label="Source Domain" />
          <SortableHeader columnKey="TargetDomain" label="Target Domain" />
          <SortableHeader columnKey="TrustType" label="Trust Type" />
          <SortableHeader columnKey="TrustDirection" label="Direction" />
          <SortableHeader columnKey="TrustAttributes" label="Attributes" />
          <SortableHeader columnKey="SidFilteringEnabled" label="SID Filtering" />
          <SortableHeader columnKey="IsTransitive" label="Transitive" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((trust, i) => (
          <tr key={i}>
            <td>{trust.SourceDomain}</td>
            <td>{trust.TargetDomain}</td>
            <td><span className="trust-type-badge">{trust.TrustType}</span></td>
            <td>
              <span className={`direction-badge ${trust.TrustDirection?.toLowerCase()}`}>
                {trust.TrustDirection}
              </span>
            </td>
            <td>{trust.TrustAttributes || '-'}</td>
            <td>
              <span className={`status-badge ${trust.SidFilteringEnabled ? 'enabled' : 'disabled'}`}>
                {trust.SidFilteringEnabled ? 'Yes' : 'No'}
              </span>
            </td>
            <td>
              <span className={`status-badge ${trust.IsTransitive ? 'enabled' : 'disabled'}`}>
                {trust.IsTransitive ? 'Yes' : 'No'}
              </span>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderNamingContextsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="Type" label="Type" />
          <SortableHeader columnKey="DistinguishedName" label="Distinguished Name" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((nc, i) => (
          <tr key={i}>
            <td><strong>{nc.Name}</strong></td>
            <td><span className="type-badge">{nc.Type}</span></td>
            <td className="dn-cell">{nc.DistinguishedName}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderSiteLinksTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SiteLinkName" label="Site Link Name" />
          <SortableHeader columnKey="Sites" label="Connected Sites" />
          <SortableHeader columnKey="Cost" label="Cost" />
          <SortableHeader columnKey="ReplicationInterval" label="Replication Interval" />
          <SortableHeader columnKey="Schedule" label="Schedule" />
          <SortableHeader columnKey="ChangeNotification" label="Change Notify" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((link, i) => (
          <tr key={i}>
            <td><strong>{link.SiteLinkName}</strong></td>
            <td className="sites-cell">{link.Sites || '-'}</td>
            <td>{link.Cost}</td>
            <td>{link.ReplicationInterval} min</td>
            <td>{link.Schedule || '24x7'}</td>
            <td>
              <span className={`status-badge ${link.ChangeNotification ? 'enabled' : 'disabled'}`}>
                {link.ChangeNotification ? 'Yes' : 'No'}
              </span>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderSitesSummaryTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SiteName" label="Site Name" />
          <SortableHeader columnKey="DCCount" label="DCs" />
          <SortableHeader columnKey="SubnetCount" label="Subnets" />
          <SortableHeader columnKey="SiteLinksCount" label="Site Links" />
          <SortableHeader columnKey="LinkedSitesCount" label="Linked Sites" />
          <SortableHeader columnKey="LinkedSites" label="Linked Site Names" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((site, i) => (
          <tr key={i}>
            <td><strong>{site.SiteName}</strong></td>
            <td>{site.DCCount || 0}</td>
            <td>{site.SubnetCount || 0}</td>
            <td>{site.SiteLinksCount || 0}</td>
            <td>{site.LinkedSitesCount || 0}</td>
            <td className="sites-cell">{site.LinkedSites || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderDCReplicationTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SourceDC" label="Source DC" />
          <SortableHeader columnKey="DestinationDC" label="Destination DC" />
          <SortableHeader columnKey="NamingContext" label="Naming Context" />
          <SortableHeader columnKey="LastReplication" label="Last Replication" />
          <SortableHeader columnKey="ReplicationStatus" label="Status" />
          <SortableHeader columnKey="FailureCount" label="Failures" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((repl, i) => (
          <tr key={i}>
            <td>{repl.SourceDC}</td>
            <td>{repl.DestinationDC}</td>
            <td className="dn-cell">{repl.NamingContext || '-'}</td>
            <td>{repl.LastReplication || 'Never'}</td>
            <td>
              <span className={`status-badge ${repl.ReplicationStatus === 'Success' ? 'enabled' : repl.ReplicationStatus === 'In Progress' ? 'warning' : 'disabled'}`}>
                {repl.ReplicationStatus || 'Unknown'}
              </span>
            </td>
            <td>
              <span className={`failure-badge ${repl.FailureCount > 0 ? 'has-failures' : 'no-failures'}`}>
                {repl.FailureCount || 0}
              </span>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderFSMOTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Role" label="FSMO Role" />
          <SortableHeader columnKey="Holder" label="Role Holder" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Scope" label="Scope" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((role, i) => (
          <tr key={i}>
            <td><strong>{role.Role}</strong></td>
            <td>{role.Holder}</td>
            <td>{role.Domain || '-'}</td>
            <td><span className="scope-badge">{role.Scope}</span></td>
          </tr>
        ))}
      </tbody>
    </table>
  );


  const renderPagination = () => {
    if (sortedData.length === 0) return null;
    const startItem = startIndex + 1;
    const endItem = Math.min(endIndex, sortedData.length);

    return (
      <div className="pagination-controls">
        <div className="pagination-info">
          Showing {startItem}-{endItem} of {sortedData.length} results
        </div>
        <div className="pagination-actions">
          <div className="page-size-select">
            <label>Per page:</label>
            <select value={pageSize} onChange={(e) => handlePageSizeChange(parseInt(e.target.value))}>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
            </select>
          </div>
          <div className="page-nav">
            <button onClick={handlePrevPage} disabled={currentPage <= 1} className="btn btn-secondary">
              Back
            </button>
            <span className="page-indicator">Page {currentPage} of {totalPages}</span>
            <button onClick={handleNextPage} disabled={currentPage >= totalPages} className="btn btn-secondary">
              Next
            </button>
          </div>
        </div>
      </div>
    );
  };

  const renderTable = () => {
    switch (activeQuery) {
      case 'forest-info':
        return renderForestInfoTable();
      case 'domains':
        return renderDomainsTable();
      case 'trusts':
        return renderTrustsTable();
      case 'naming-contexts':
        return renderNamingContextsTable();
      case 'fsmo-roles':
        return renderFSMOTable();
      case 'ad-sites':
        return renderSitesSummaryTable();
      case 'site-links':
        return renderSiteLinksTable();
      case 'dc-replication':
        return renderDCReplicationTable();
      default:
        return null;
    }
  };

  return (
    <div className="topology-page">
      <div className="page-header">
        <div className="page-title">
          <span className="page-icon">🔗</span>
          <h1>AD Topology</h1>
        </div>
        <p className="page-description">View Active Directory forest structure, domains, trusts, and replication topology</p>
      </div>

      <div className="query-panel card">
        <h3>Topology Reports</h3>
        <div className="query-cards">
          {predefinedQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handleQuery(q.id)}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/topology" queryId={q.id} label={q.label} />
              </div>
              <div className="query-card-desc">{q.description}</div>
            </div>
          ))}
        </div>
      </div>

      {error && <div className="error-message">{error}</div>}

      {loading && (
        <div className="loading">
          <div className="spinner"></div>
          <span>Loading topology data...</span>
        </div>
      )}

      {!loading && data.length > 0 && (
        <div className="results-panel card">
          <div className="results-header">
            <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
              <span>Found {sortedData.length} result(s)</span>

              {/* Trusts quick filter tabs */}
              {activeQuery === 'trusts' && viewMode === 'table' && (
                <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                  <button
                    className="btn btn-secondary"
                    onClick={() => setTrustSidFilter('all')}
                    style={{
                      padding: '4px 10px',
                      fontSize: '12px',
                      borderColor: trustSidFilter === 'all' ? 'var(--accent-color)' : 'var(--border-color)',
                      color: trustSidFilter === 'all' ? 'var(--text-primary)' : 'var(--text-secondary)',
                    }}
                    title="Show all trusts"
                  >
                    All
                  </button>
                  <button
                    className="btn btn-secondary"
                    onClick={() => setTrustSidFilter('enabled')}
                    style={{
                      padding: '4px 10px',
                      fontSize: '12px',
                      borderColor: trustSidFilter === 'enabled' ? 'var(--accent-color)' : 'var(--border-color)',
                      color: trustSidFilter === 'enabled' ? 'var(--text-primary)' : 'var(--text-secondary)',
                    }}
                    title="Show trusts with SID Filtering enabled"
                  >
                    SID Filtering: Yes
                  </button>
                  <button
                    className="btn btn-secondary"
                    onClick={() => setTrustSidFilter('disabled')}
                    style={{
                      padding: '4px 10px',
                      fontSize: '12px',
                      borderColor: trustSidFilter === 'disabled' ? 'var(--accent-color)' : 'var(--border-color)',
                      color: trustSidFilter === 'disabled' ? 'var(--text-primary)' : 'var(--text-secondary)',
                    }}
                    title="Show trusts with SID Filtering disabled"
                  >
                    SID Filtering: No
                  </button>
                </div>
              )}

              <div className="view-toggle" style={{ display: 'flex', gap: '4px' }}>
                <button
                  className={`toggle-btn ${viewMode === 'table' ? 'active' : ''}`}
                  onClick={() => setViewMode('table')}
                  style={{
                    padding: '4px 12px',
                    fontSize: '12px',
                    border: '1px solid var(--border-color)',
                    borderRadius: '4px 0 0 4px',
                    background: viewMode === 'table' ? 'var(--accent-color)' : 'var(--bg-secondary)',
                    color: viewMode === 'table' ? '#fff' : 'var(--text-secondary)',
                    cursor: 'pointer',
                  }}
                >
                  Table
                </button>
                <button
                  className={`toggle-btn ${viewMode === 'diagram' ? 'active' : ''}`}
                  onClick={() => setViewMode('diagram')}
                  style={{
                    padding: '4px 12px',
                    fontSize: '12px',
                    border: '1px solid var(--border-color)',
                    borderRadius: '0 4px 4px 0',
                    background: viewMode === 'diagram' ? 'var(--accent-color)' : 'var(--bg-secondary)',
                    color: viewMode === 'diagram' ? '#fff' : 'var(--text-secondary)',
                    cursor: 'pointer',
                  }}
                >
                  Diagram
                </button>
              </div>
              {viewMode === 'table' && <span className="sort-hint">Click column headers to sort</span>}
            </div>
            <ExportButton data={sortedData} filename={`topology_${activeQuery || 'export'}`} title="AD Topology Report" />
          </div>

          {viewMode === 'table' ? (
            <>
              <div className="table-scroll">
                {renderTable()}
              </div>
              {renderPagination()}
            </>
          ) : (
            <TopologyDiagram data={sortedData} activeQuery={activeQuery} />
          )}
        </div>
      )}

      {!loading && data.length === 0 && activeQuery && (
        <div className="empty-state card">
          <div className="empty-icon">📭</div>
          <p>No data found for this query.</p>
        </div>
      )}
    </div>
  );
};

export default Topology;

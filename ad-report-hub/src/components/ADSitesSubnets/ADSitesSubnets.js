import React, { useState, useMemo } from 'react';
import axios from 'axios';
import './ADSitesSubnets.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { useApp } from '../../context/AppContext';
import { formatDate } from '../../utils/dateUtils';

const ADSitesSubnets = () => {
  const { selectedDomain } = useApp();
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  
  // Sorting state
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  const predefinedQueries = [
    { id: 'site-links', label: 'AD Site Links', description: 'Replication topology and costs' },
    { id: 'subnets', label: 'All Subnets', description: 'Subnet to site assignments' },
    { id: 'subnets-detailed', label: 'Subnets Detailed', description: 'Subnets with DC count and metadata' },
    { id: 'sites-no-dc', label: 'Sites with No DC', description: 'AD Sites with no Domain Controller' },
    { id: 'unassigned-subnets', label: 'Unassigned Subnets', description: 'Subnets not associated with any AD Site' },
  ];

  const handleQuery = async (queryId) => {
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCurrentPage(1);
    setSortConfig({ key: null, direction: 'asc' }); // Reset sorting
    try {
      const params = { query: queryId };
      if (selectedDomain) params.domain = selectedDomain;
      const response = await axios.get('/api/sites-subnets', { params });
      const result = response.data.data;
      // Ensure data is always an array of objects (not a raw string)
      setData(Array.isArray(result) ? result : []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
    } finally {
      setLoading(false);
    }
  };

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

  // Pagination helpers - use sortedData
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
    setCurrentPage(1); // Reset to first page when sorting
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

  const renderSiteLinksTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SiteLinkName" label="Site Link Name" />
          <SortableHeader columnKey="Description" label="Description" />
          <SortableHeader columnKey="Cost" label="Cost" />
          <SortableHeader columnKey="ReplicationIntervalMin" label="Replication Interval (Min)" />
          <SortableHeader columnKey="ReplicationSchedule" label="Schedule" />
          <SortableHeader columnKey="ChangeNotification" label="Change Notify" />
          <SortableHeader columnKey="SiteCount" label="Site Count" />
          <SortableHeader columnKey="AssociatedSites" label="Associated Sites" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((link, i) => (
          <tr key={i}>
            <td>{link.SiteLinkName || link['Site Link Name']}</td>
            <td>{link.Description || '-'}</td>
            <td>{link.Cost}</td>
            <td>{link.ReplicationIntervalMin || link.DurationMin || '-'}</td>
            <td>{link.ReplicationSchedule || '-'}</td>
            <td>
              <span className={`status-badge ${link.ChangeNotification ? 'enabled' : 'disabled'}`}>
                {link.ChangeNotification ? 'Yes' : 'No'}
              </span>
            </td>
            <td>{link.SiteCount || link['Site Count']}</td>
            <td className="sites-list">{link.AssociatedSites || link['Associated Sites'] || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderSubnetsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SubnetName" label="Subnet Name" />
          <SortableHeader columnKey="ADSite" label="AD Site" />
          <SortableHeader columnKey="DCCount" label="DC Count" />
          <SortableHeader columnKey="DCNames" label="DC Names" />
          <SortableHeader columnKey="Description" label="Description" />
          <SortableHeader columnKey="Location" label="Location" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((subnet, i) => (
          <tr key={i}>
            <td>{subnet['Subnet Name'] || subnet.SubnetName || subnet.Subnet}</td>
            <td>{subnet['AD Site'] || subnet.ADSite}</td>
            <td>{subnet['DC Count'] || subnet.DCCount || 0}</td>
            <td className="dc-names">{subnet['DC Names'] || subnet.DCNames || '-'}</td>
            <td>{subnet['Subnet Description'] || subnet.Description || '-'}</td>
            <td>{subnet['Subnet Location'] || subnet.Location || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderSubnetsDetailedTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Subnet" label="Subnet" />
          <SortableHeader columnKey="ADSite" label="AD Site" />
          <SortableHeader columnKey="DCCount" label="DC Count" />
          <SortableHeader columnKey="Description" label="Description" />
          <SortableHeader columnKey="PhysicalLocation" label="Physical Location" />
          <SortableHeader columnKey="CreatedDate" label="Created" />
          <SortableHeader columnKey="LastModified" label="Last Modified" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((subnet, i) => (
          <tr key={i}>
            <td>{subnet.Subnet || subnet['Subnet Name']}</td>
            <td>{subnet['AD Site'] || subnet.ADSite}</td>
            <td>{subnet['DC Count'] || subnet.DCCount || 0}</td>
            <td>{subnet.Description || '-'}</td>
            <td>{subnet['Physical Location'] || subnet.PhysicalLocation || '-'}</td>
            <td>{formatDate(subnet.CreatedDate || subnet['Created Date'])}</td>
            <td>{formatDate(subnet.LastModified || subnet['Last Modified'])}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

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

  const renderSitesNoDCTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="SiteName" label="Site Name" />
          <SortableHeader columnKey="Description" label="Description" />
          <SortableHeader columnKey="DCCount" label="DC Count" />
          <SortableHeader columnKey="SubnetCount" label="Subnet Count" />
          <SortableHeader columnKey="SiteLinks" label="Site Links" />
          <SortableHeader columnKey="Created" label="Created" />
          <SortableHeader columnKey="LastModified" label="Last Modified" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((site, i) => (
          <tr key={i}>
            <td>{site.SiteName}</td>
            <td>{site.Description || '-'}</td>
            <td>{site.DCCount}</td>
            <td>{site.SubnetCount}</td>
            <td>{site.SiteLinks || '-'}</td>
            <td>{formatDate(site.Created)}</td>
            <td>{formatDate(site.LastModified)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderUnassignedSubnetsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Subnet" label="Subnet" />
          <SortableHeader columnKey="ADSite" label="AD Site" />
          <SortableHeader columnKey="Description" label="Description" />
          <SortableHeader columnKey="PhysicalLocation" label="Physical Location" />
          <SortableHeader columnKey="Created" label="Created" />
          <SortableHeader columnKey="LastModified" label="Last Modified" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((subnet, i) => (
          <tr key={i}>
            <td>{subnet.Subnet}</td>
            <td><span style={{ color: 'var(--error-text)' }}>{subnet.ADSite}</span></td>
            <td>{subnet.Description || '-'}</td>
            <td>{subnet.PhysicalLocation || '-'}</td>
            <td>{formatDate(subnet.Created)}</td>
            <td>{formatDate(subnet.LastModified)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderTable = () => {
    switch (activeQuery) {
      case 'site-links':
        return renderSiteLinksTable();
      case 'subnets':
        return renderSubnetsTable();
      case 'subnets-detailed':
        return renderSubnetsDetailedTable();
      case 'sites-no-dc':
        return renderSitesNoDCTable();
      case 'unassigned-subnets':
        return renderUnassignedSubnetsTable();
      default:
        return null;
    }
  };

  return (
    <div className="sites-page">
      <h1>AD Sites & Subnets</h1>

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
              onClick={() => handleQuery(q.id)}
              style={{
                flex: '0 0 calc(33.33% - 16px)',
                maxWidth: 'calc(33.33% - 16px)',
                minWidth: '280px'
              }}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/sites-subnets" queryId={q.id} label={q.label} />
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
          <span>Querying AD Sites & Subnets...</span>
        </div>
      )}

      {!loading && data.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <div>
              <span>Found {data.length} total result(s)</span>
              <span style={{ marginLeft: '16px', fontSize: '12px', color: 'var(--text-muted)' }}>
                Click column headers to sort ↕
              </span>
            </div>
            <ExportButton data={data} filename={`ad_${activeQuery || 'export'}`} title="Sites & Subnets Report" />
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

export default ADSitesSubnets;

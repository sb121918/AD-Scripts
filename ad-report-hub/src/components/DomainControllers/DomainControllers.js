import React, { useState } from 'react';
import axios from 'axios';
import './DomainControllers.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { formatDate } from '../../utils/dateUtils';
import { useSortableData } from '../../utils/useSortableData';
import { useApp } from '../../context/AppContext';

const DomainControllers = () => {
  const { startReport, endReport, selectedDomain, isLoading: globalIsLoading } = useApp();
  const [perDomainInput, setPerDomainInput] = useState('');
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [activeMode, setActiveMode] = useState(null); // 'forest-wide' | 'per-domain'
  const [activeTargetDomain, setActiveTargetDomain] = useState('');
  
  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  // Sorting
  const { sortedItems, requestSort, getSortIndicator, getSortClass } = useSortableData(data);

  const forestWideQueries = [
    { id: 'all', label: 'All Domain Controllers (Forest-wide)', description: 'Complete list of all DCs in the forest' },
    { id: 'health', label: 'DC Resource Health (Forest-wide)', description: 'CPU, RAM, Disk usage per DC' },
    { id: 'services', label: 'DC Services Status (Forest-wide)', description: 'Service matrix across all DCs' },
  ];

  const perDomainQueries = [
    { id: 'all', label: 'All DCs (Per Domain)', description: 'Domain Controllers for the specified domain' },
    { id: 'health', label: 'DC Resource Health (Per Domain)', description: 'CPU, RAM, Disk usage per DC for the specified domain' },
    { id: 'services', label: 'DC Service Status (Per Domain)', description: 'Service matrix across DCs for the specified domain' },
    { id: 'fsmo', label: 'FSMO Role Holders (Per Domain)', description: 'FSMO roles and their holders for the specified domain' },
  ];

  const getQueryLabel = (queryId, mode) => {
    const list = mode === 'per-domain' ? perDomainQueries : forestWideQueries;
    const query = list.find(q => q.id === queryId);
    return query ? query.label : 'Report';
  };

  const runQuery = async (queryId, mode, targetDomain) => {
    // Enforce: only one report at a time (across app)
    if (loading || globalIsLoading) {
      setError('A report is already running. Please wait for it to finish before starting another.');
      return;
    }

    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setActiveMode(mode);
    setActiveTargetDomain(targetDomain || '');
    setCurrentPage(1);
    const scopeLabel = mode === 'per-domain' ? `Per-domain: ${targetDomain}` : 'Forest-wide';
    startReport(`${getQueryLabel(queryId, mode)} (${scopeLabel})`, '/domain-controllers');
    
    try {
      const params = { query: queryId };
      // For per-domain reports, anchor to the typed domain so it runs in that forest
      if (mode === 'per-domain') {
        params.domain = targetDomain;
        params.targetDomain = targetDomain;
      } else {
        // Forest-wide reports anchor to the top selector
        if (selectedDomain) params.domain = selectedDomain;
      }
      const response = await axios.get('/api/domain-controllers', { params });
      setData(response.data.data || []);
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
    } finally {
      setLoading(false);
      endReport();
    }
  };

  const handleForestWideQuery = (queryId) => runQuery(queryId, 'forest-wide', undefined);

  const handlePerDomainQuery = (queryId) => {
    const td = (perDomainInput || '').trim();
    if (!td) {
      setError('Please enter a domain for per-domain reports (e.g. cbs.ad.cbs.net)');
      return;
    }
    runQuery(queryId, 'per-domain', td);
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

  // Pagination helpers - use sortedItems for sorted data
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


  const renderAllDCsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          {renderSortableHeader('Domain', 'Domain')}
          {renderSortableHeader('DCName', 'DC Name')}
          {renderSortableHeader('HostName', 'Hostname')}
          {renderSortableHeader('IPv4Address', 'IP Address')}
          {renderSortableHeader('Site', 'Site')}
          {renderSortableHeader('OperatingSystem', 'Operating System')}
          {renderSortableHeader('Uptime', 'Uptime')}
          {renderSortableHeader('IsGlobalCatalog', 'Global Catalog')}
          {renderSortableHeader('IsReadOnly', 'Read Only')}
          {renderSortableHeader('FSMORoles', 'FSMO Roles')}
          {renderSortableHeader('Created', 'Created')}
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((dc, i) => (
          <tr key={i}>
            <td>{dc.Domain}</td>
            <td>{dc.DCName}</td>
            <td>{dc.HostName || '-'}</td>
            <td>{dc.IPv4Address || '-'}</td>
            <td>{dc.Site || '-'}</td>
            <td>{dc.OperatingSystem || '-'}</td>
            <td>{dc.Uptime || '-'}</td>
            <td>
              <span className={`status-badge ${dc.IsGlobalCatalog ? 'enabled' : 'disabled'}`}>
                {dc.IsGlobalCatalog ? 'Yes' : 'No'}
              </span>
            </td>
            <td>
              <span className={`status-badge ${dc.IsReadOnly ? 'enabled' : 'disabled'}`}>
                {dc.IsReadOnly ? 'Yes' : 'No'}
              </span>
            </td>
            <td>{dc.FSMORoles || '-'}</td>
            <td>{formatDate(dc.Created)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderHealthTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Domain</th>
          <th>DC Name</th>
          <th>Status</th>
          <th>CPU %</th>
          <th>RAM Used GB</th>
          <th>RAM %</th>
          <th>Disk Used %</th>
          <th>IP Address</th>
          <th>Site</th>
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((dc, i) => (
          <tr key={i}>
            <td>{dc.Domain}</td>
            <td>{dc.DCName}</td>
            <td>
              <span className={`status-badge ${dc.Status === 'Online' ? 'enabled' : 'disabled'}`}>
                {dc.Status}
              </span>
            </td>
            <td>{dc['CPU_Load_%'] || '-'}</td>
            <td>{dc['RAM_Used_GB'] || '-'}</td>
            <td>{dc['RAM_Used_%'] || '-'}</td>
            <td>{dc['C_Used_%'] || '-'}</td>
            <td>{dc.IPAddress || '-'}</td>
            <td>{dc.Site || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  // Services to display in the table
  const serviceColumns = [
    { key: 'NTDS', label: 'NTDS' },
    { key: 'ADWS', label: 'ADWS' },
    { key: 'DNS', label: 'DNS' },
    { key: 'Kdc', label: 'Kdc' },
    { key: 'Netlogon', label: 'Netlogon' },
    { key: 'RpcSs', label: 'RpcSs' },
    { key: 'DFSR', label: 'DFSR' },
    { key: 'AATPSensor', label: 'ATP Sensor' },
    { key: 'AATPSensorUpdater', label: 'ATP Updater' },
    { key: 'AzureADConnectHealthAgent', label: 'Entra Health' },
    { key: 'CSFalconService', label: 'CrowdStrike' },
    { key: 'NPSrvHost', label: 'Quest Auditor' },
    { key: 'Tanium Client', label: 'Tanium' },
    { key: 'W32Time', label: 'W32Time' },
    { key: 'Cribl', label: 'Cribl' }
  ];

  const getServiceStatusClass = (status) => {
    if (!status) return 'status-unknown';
    if (status === 'Running') return 'status-running';
    if (status === 'Stopped') return 'status-stopped';
    if (status === 'Not Installed') return 'status-notinstalled';
    if (status === 'Requires Elevated Access') return 'status-noaccess';
    if (status === 'UNREACHABLE' || status === 'ERROR') return 'status-error';
    return 'status-unknown';
  };

  const getOverallStatusClass = (status) => {
    if (status === 'Healthy') return 'overall-healthy';
    if (status === 'Warning') return 'overall-warning';
    if (status === 'Critical') return 'overall-critical';
    if (status === 'Offline') return 'overall-offline';
    return 'overall-unknown';
  };

  const renderServicesTable = () => (
    <table className="data-table services-matrix">
      <thead>
        <tr>
          <th>Domain</th>
          <th>DC Name</th>
          <th>Site</th>
          <th>Status</th>
          {serviceColumns.map(svc => (
            <th key={svc.key} title={svc.key}>{svc.label}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((dc, i) => (
          <tr key={i}>
            <td>{dc.Domain}</td>
            <td>{dc.DCName}</td>
            <td>{dc.Site || '-'}</td>
            <td>
              <span className={`overall-status ${getOverallStatusClass(dc.OverallStatus)}`}>
                {dc.OverallStatus || '-'}
              </span>
            </td>
            {serviceColumns.map(svc => (
              <td key={svc.key}>
                <span className={`service-status ${getServiceStatusClass(dc[svc.key])}`} title={dc[svc.key] || 'Unknown'}>
                  {dc[svc.key] === 'Running' ? '✓' : 
                   dc[svc.key] === 'Stopped' ? '✗' : 
                   dc[svc.key] === 'Not Installed' ? 'N/A' : 
                   dc[svc.key] === 'Requires Elevated Access' ? '🔒' :
                   dc[svc.key] === 'UNREACHABLE' ? '?' :
                   dc[svc.key] === 'ERROR' ? '!' : '-'}
                </span>
              </td>
            ))}
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
              <span className={`type-badge ${policy.Type === 'Default' ? 'default' : 'fgpp'}`}>
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

  const getTrustDirectionClass = (direction) => {
    switch (direction) {
      case 'Bidirectional': return 'trust-bidirectional';
      case 'Inbound': return 'trust-inbound';
      case 'Outbound': return 'trust-outbound';
      case 'Disabled': return 'trust-disabled';
      default: return '';
    }
  };

  const renderTrustsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Source Domain</th>
          <th>Level</th>
          <th>Target Domain</th>
          <th>Trust Type</th>
          <th>Direction</th>
          <th>Forest Transitive</th>
          <th>Intra-Forest</th>
          <th>Selective Auth</th>
          <th>SID Filtering</th>
          <th>Uses AES</th>
          <th>Created</th>
          <th>Modified</th>
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((trust, i) => (
          <tr key={i}>
            <td>{trust.SourceDomain}</td>
            <td>
              <span className={`level-badge ${trust.SourceLevel === 'Forest Root' ? 'level-root' : 'level-child'}`}>
                {trust.SourceLevel}
              </span>
            </td>
            <td>{trust.TargetDomain}</td>
            <td>{trust.TrustType}</td>
            <td>
              <span className={`trust-direction ${getTrustDirectionClass(trust.TrustDirection)}`}>
                {trust.TrustDirection === 'Bidirectional' && '↔ '}
                {trust.TrustDirection === 'Inbound' && '← '}
                {trust.TrustDirection === 'Outbound' && '→ '}
                {trust.TrustDirection}
              </span>
            </td>
            <td>
              <span className={`status-badge ${trust.ForestTransitive ? 'enabled' : 'disabled'}`}>
                {trust.ForestTransitive ? 'Yes' : 'No'}
              </span>
            </td>
            <td>
              <span className={`status-badge ${trust.IntraForest ? 'enabled' : 'disabled'}`}>
                {trust.IntraForest ? 'Yes' : 'No'}
              </span>
            </td>
            <td>
              <span className={`status-badge ${trust.SelectiveAuth ? 'enabled' : 'disabled'}`}>
                {trust.SelectiveAuth ? 'Yes' : 'No'}
              </span>
            </td>
            <td>
              <span className={`status-badge ${trust.SIDFilteringQuarantined ? 'enabled' : 'disabled'}`}>
                {trust.SIDFilteringQuarantined ? 'Enabled' : 'Disabled'}
              </span>
            </td>
            <td>
              <span className={`status-badge ${trust.UsesAESKeys ? 'enabled' : 'disabled'}`}>
                {trust.UsesAESKeys ? 'Yes' : 'No'}
              </span>
            </td>
            <td>{formatDate(trust.Created)}</td>
            <td>{formatDate(trust.Modified)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderFSMOTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Scope</th>
          <th>Domain</th>
          <th>Role</th>
          <th>Holder</th>
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((row, i) => (
          <tr key={i}>
            <td>{row.Scope || '-'}</td>
            <td>{row.Domain || '-'}</td>
            <td>{row.Role || '-'}</td>
            <td>{row.Holder || '-'}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderTable = () => {
    switch (activeQuery) {
      case 'all':
        return renderAllDCsTable();
      case 'health':
        return renderHealthTable();
      case 'services':
        return renderServicesTable();
      case 'trusts':
        return renderTrustsTable();
      case 'fsmo':
        return renderFSMOTable();
      default:
        return null;
    }
  };

  return (
    <div className="dc-page">
      <h1>Active Directory Reports</h1>

      {/* Forest-wide section */}
      <div className="query-panel card">
        <h3>All Domain Controllers (Forest-wide)</h3>
        <div className="query-cards">
          {forestWideQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeMode === 'forest-wide' && activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handleForestWideQuery(q.id)}
              style={{ pointerEvents: (loading || globalIsLoading) ? 'none' : 'auto', opacity: (loading || globalIsLoading) ? 0.6 : 1 }}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/domain-controllers" queryId={q.id} label={q.label} />
              </div>
              <div className="query-card-desc">{q.description}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Per-domain section */}
      <div className="query-panel card" style={{ marginTop: '16px' }}>
        <h3>Per-domain Reports</h3>

        <div style={{ marginBottom: '14px', display: 'flex', gap: '10px', alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <div style={{ minWidth: '320px', flex: 1 }}>
            <label style={{ display: 'block', marginBottom: '6px', color: 'var(--text-secondary)', fontSize: '13px' }}>
              Target domain (required for per-domain reports)
            </label>
            <input
              type="text"
              placeholder="e.g. cbs.ad.cbs.net"
              value={perDomainInput}
              onChange={(e) => setPerDomainInput(e.target.value)}
              style={{
                width: '100%',
                maxWidth: '520px',
                padding: '10px 12px',
                border: '1px solid var(--border-input)',
                borderRadius: '6px',
                backgroundColor: 'var(--bg-input)',
                color: 'var(--text-primary)',
                fontSize: '14px',
              }}
            />
            <div style={{ marginTop: '6px', fontSize: '12px', color: 'var(--text-muted)' }}>
              Uses the top selector as the forest anchor, and filters down to this specific domain.
            </div>
          </div>
          <button
            className="btn btn-secondary"
            onClick={() => setPerDomainInput('')}
            disabled={!perDomainInput}
            style={{ padding: '10px 12px' }}
          >
            Clear
          </button>
        </div>

        <div className="query-cards">
          {perDomainQueries.map((q) => (
            <div
              key={`per-${q.id}`}
              className={`query-card ${activeMode === 'per-domain' && activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handlePerDomainQuery(q.id)}
              style={{ pointerEvents: (loading || globalIsLoading) ? 'none' : 'auto', opacity: (loading || globalIsLoading) ? 0.6 : 1 }}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/domain-controllers" queryId={`${q.id}-per-domain`} label={q.label} params={{ targetDomain: perDomainInput }} />
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
          <span>Querying Domain Controllers...</span>
        </div>
      )}

      {!loading && data.length > 0 && (
        <div className="results-panel card" style={{ padding: 0, overflow: 'hidden' }}>
          <div className="results-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '16px 20px', borderBottom: '1px solid var(--border-color)' }}>
            <div>
              <span>Found {data.length} total result(s)</span>
              <span style={{ marginLeft: '16px', fontSize: '12px', color: 'var(--text-muted)' }}>
                Scope: {activeMode === 'per-domain' ? `Per-domain (${activeTargetDomain})` : 'Forest-wide'} • Click column headers to sort
              </span>
            </div>
            <ExportButton data={data} filename={`dc_${activeQuery || 'export'}`} title="Domain Controllers Report" />
          </div>
          {activeQuery === 'services' && (
            <div className="services-legend">
              <span style={{ fontWeight: 500, color: 'var(--text-primary)', marginRight: '10px' }}>Legend:</span>
              <div className="legend-item">
                <span className="legend-icon" style={{ backgroundColor: 'rgba(46, 204, 113, 0.2)', color: '#27ae60' }}>✓</span>
                <span>Running</span>
              </div>
              <div className="legend-item">
                <span className="legend-icon" style={{ backgroundColor: 'rgba(231, 76, 60, 0.2)', color: '#e74c3c' }}>✗</span>
                <span>Stopped</span>
              </div>
              <div className="legend-item">
                <span className="legend-icon" style={{ backgroundColor: 'rgba(149, 165, 166, 0.15)', color: '#95a5a6', fontSize: '9px' }}>N/A</span>
                <span>Not Installed</span>
              </div>
              <div className="legend-item">
                <span className="legend-icon" style={{ backgroundColor: 'rgba(155, 89, 182, 0.2)', color: '#9b59b6' }}>🔒</span>
                <span>Requires Elevated Access</span>
              </div>
              <div className="legend-item">
                <span className="legend-icon" style={{ backgroundColor: 'rgba(243, 156, 18, 0.2)', color: '#f39c12' }}>?</span>
                <span>Unreachable</span>
              </div>
            </div>
          )}
          <div style={{ overflowX: 'auto' }}>
            {renderTable()}
          </div>
          {renderPagination()}
        </div>
      )}
    </div>
  );
};

export default DomainControllers;

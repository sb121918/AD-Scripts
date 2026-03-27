import React, { useState, useMemo } from 'react';
import axios from 'axios';
import './Compliance.css';
import ExportButton from '../common/ExportButton';
import FavoriteButton from '../common/FavoriteButton';
import { useApp } from '../../context/AppContext';

const Compliance = () => {
  const { selectedDomain } = useApp();
  const [domainOverride, setDomainOverride] = useState('');
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [activeQuery, setActiveQuery] = useState(null);
  const [daysInput, setDaysInput] = useState(90);
  /** When true, Privileged Group Changes uses linked-value replication metadata (per-member add/remove). */
  const [memberDetailInPrivilegedChanges, setMemberDetailInPrivilegedChanges] = useState(true);
  /** Combined built-in Administrators: membership changes + Security log audit (same API round-trip). */
  const [combinedBuiltin, setCombinedBuiltin] = useState(null);

  // Pagination state
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);

  // Sorting state
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  /** Per-group change reports reuse query `privileged-changes` with a server-side groupFilter. */
  const predefinedQueries = [
    { id: 'privileged-changes', label: 'Privileged Group Changes', description: 'Membership changes in all privileged groups (DA, EA, SA, BA, AD-Enterprise Systems Admins — replication metadata)', hasDays: true },
    { id: 'privileged-membership-audit', label: 'Privileged Group Membership Audit', description: 'Who added/removed whom, when, and from which privileged group (DC Security Logs)', hasDays: true },
    {
      id: 'privileged-changes-ea',
      apiQuery: 'privileged-changes',
      groupFilter: 'Enterprise Admins',
      label: 'Enterprise Admins Changes',
      description: 'Membership changes to Enterprise Admins only (forest-wide)',
      hasDays: true,
    },
    {
      id: 'privileged-changes-sa',
      apiQuery: 'privileged-changes',
      groupFilter: 'Schema Admins',
      label: 'Schema Admins Changes',
      description: 'Membership changes to Schema Admins only (forest-wide)',
      hasDays: true,
    },
    {
      id: 'privileged-changes-da',
      apiQuery: 'privileged-changes',
      groupFilter: 'Domain Admins',
      label: 'Domain Admins Changes',
      description: 'Membership changes to Domain Admins across domains in scope',
      hasDays: true,
    },
    {
      id: 'privileged-changes-esa',
      apiQuery: 'privileged-changes',
      groupFilter: 'AD-Enterprise Systems Admins',
      label: 'AD-Enterprise Systems Admins Changes',
      description: 'Membership changes to AD-Enterprise Systems Admins (custom group; Viacom forest)',
      hasDays: true,
    },
    {
      id: 'builtin-administrators-combined',
      label: 'Built-in Administrators (Changes & membership audit)',
      description: 'Single report: membership changes (replication + Security event fallback) and Security log audit (4728–4757) for Builtin\\Administrators',
      hasDays: true,
    },
    { id: 'stale-admins', label: 'Stale Admin Accounts', description: 'Privileged users who haven\'t logged on recently', hasDays: true },
    { id: 'pwd-never-expires-priv', label: 'Password Never Expires', description: 'Privileged accounts with non-expiring passwords', hasDays: false },
    { id: 'kerberos-delegation', label: 'Kerberos Delegation', description: 'Accounts with delegation settings (unconstrained/constrained/RBCD)', hasDays: false },
    { id: 'sid-history', label: 'SID History', description: 'Accounts with SID history attributes from migrations', hasDays: false },
    { id: 'adminsdholder', label: 'AdminSDHolder Protected', description: 'Accounts with adminCount=1 (AdminSDHolder protection)', hasDays: false },
  ];

  const handleQuery = async (queryId) => {
    setLoading(true);
    setError(null);
    setActiveQuery(queryId);
    setCombinedBuiltin(null);
    setCurrentPage(1);
    setSortConfig({ key: null, direction: 'asc' });
    try {
      const queryDef = predefinedQueries.find(q => q.id === queryId);
      const params = { query: queryDef?.apiQuery || queryId };
      if (queryDef?.hasDays && daysInput) {
        params.days = daysInput;
      }
      if (queryDef?.groupFilter) {
        params.groupFilter = queryDef.groupFilter;
      }
      if (
        (queryDef?.apiQuery || queryId) === 'privileged-changes' ||
        queryId === 'builtin-administrators-combined'
      ) {
        params.memberDetails = memberDetailInPrivilegedChanges;
      }
      // domain = forest anchor (selected in navbar)
      if (selectedDomain) params.domain = selectedDomain;
      // targetDomain = optional per-domain override
      if (domainOverride && domainOverride.trim()) {
        params.targetDomain = domainOverride.trim();
      }
      const response = await axios.get('/api/compliance', { params });
      const result = response.data.data;
      if (result && typeof result === 'object' && result.builtinCombined) {
        setCombinedBuiltin({
          changes: Array.isArray(result.changes) ? result.changes : [],
          audit: Array.isArray(result.membershipAudit) ? result.membershipAudit : [],
        });
        setData([]);
      } else {
        setCombinedBuiltin(null);
        setData(Array.isArray(result) ? result : [result]);
      }
    } catch (err) {
      setError(err.response?.data?.error || 'Query failed');
      setData([]);
      setCombinedBuiltin(null);
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

  const complianceExportData = useMemo(() => {
    if (activeQuery === 'builtin-administrators-combined' && combinedBuiltin) {
      return [
        { ReportSection: 'Built-in Administrators — membership changes (replication / events)' },
        ...combinedBuiltin.changes,
        { ReportSection: 'Built-in Administrators — Security log audit' },
        ...combinedBuiltin.audit,
      ];
    }
    return data;
  }, [activeQuery, combinedBuiltin, data]);

  const isActivePrivilegedChangesReport =
    activeQuery &&
    (activeQuery === 'privileged-changes' ||
      activeQuery.startsWith('privileged-changes-') ||
      activeQuery === 'builtin-administrators-combined');

  const privilegedChangesShowMemberDetail = Boolean(
    isActivePrivilegedChangesReport && memberDetailInPrivilegedChanges
  );

  const privilegedChangesHasDetailNotes = useMemo(
    () => data.some((r) => typeof r.DetailNote === 'string' && r.DetailNote.trim() !== ''),
    [data]
  );

  const privilegedChangesHint = useMemo(() => {
    if (activeQuery === 'builtin-administrators-combined') return null;
    if (!isActivePrivilegedChangesReport || !data.length || data[0]?.GroupName === 'No changes found') {
      return null;
    }
    const hasMemberDn = data.some((r) => typeof r.MemberDN === 'string' && r.MemberDN.trim() !== '');
    if (hasMemberDn) return null;
    if (data.some((r) => typeof r.DetailNote === 'string' && r.DetailNote.trim())) return null;

    if (!memberDetailInPrivilegedChanges) {
      return 'To see which members were added or removed, enable “show who was added or removed” under the lookback period, then run this report again. For actor and member from Security events, use Privileged Group Membership Audit.';
    }
    return 'Per-member linked replication metadata was not returned. Try a longer lookback, confirm the service account can read linked-value replication metadata on each domain PDC, or use Privileged Group Membership Audit for Security log detail.';
  }, [isActivePrivilegedChangesReport, memberDetailInPrivilegedChanges, data]);

  // Pagination helpers
  const totalPages = Math.ceil(sortedData.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = sortedData.slice(startIndex, endIndex);

  const handlePageSizeChange = (newSize) => {
    setPageSize(newSize);
    setCurrentPage(1);
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
    if (sortConfig.key !== key) return <span className="sort-indicator">&#8597;</span>;
    return <span className="sort-indicator active">{sortConfig.direction === 'asc' ? '\u2191' : '\u2193'}</span>;
  };

  const SortableHeader = ({ columnKey, label }) => (
    <th onClick={() => handleSort(columnKey)} style={{ cursor: 'pointer' }} className="sortable-th">
      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
        {label}
        {getSortIndicator(columnKey)}
      </div>
    </th>
  );

  const getRiskBadge = (risk) => {
    const r = (risk || '').toLowerCase();
    return <span className={`risk-badge ${r}`}>{risk || 'Unknown'}</span>;
  };

  const getDelegationBadge = (type) => {
    const t = (type || '').toLowerCase();
    let cls = 'constrained';
    if (t.includes('unconstrained')) cls = 'unconstrained';
    else if (t.includes('rbcd') || t.includes('resource')) cls = 'rbcd';
    return <span className={`delegation-badge ${cls}`}>{type}</span>;
  };

  const renderPrivilegedChangesTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="GroupName" label="Group Name" />
          <SortableHeader columnKey="Domain" label="Domain" />
          {privilegedChangesShowMemberDetail && (
            <SortableHeader columnKey="ChangeType" label="Change" />
          )}
          {privilegedChangesShowMemberDetail && (
            <SortableHeader columnKey="MemberCN" label="Member" />
          )}
          {privilegedChangesShowMemberDetail && <th>Member DN</th>}
          {privilegedChangesShowMemberDetail && privilegedChangesHasDetailNotes && <th>Note</th>}
          <SortableHeader columnKey="AttributeChanged" label="Attribute" />
          <SortableHeader columnKey="ChangeTime" label="Change Time" />
          <SortableHeader columnKey="OriginatingDC" label="Originating DC" />
          <SortableHeader columnKey="ChangeVersion" label="Version" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.GroupName}</strong></td>
            <td>{item.Domain}</td>
            {privilegedChangesShowMemberDetail && (
              <td>
                {item.ChangeType ? (
                  <span className={`status-badge ${String(item.ChangeType).toLowerCase() === 'removed' ? 'disabled' : 'enabled'}`}>
                    {item.ChangeType}
                  </span>
                ) : '—'}
              </td>
            )}
            {privilegedChangesShowMemberDetail && (
              <td>{item.MemberCN || '—'}</td>
            )}
            {privilegedChangesShowMemberDetail && (
              <td
                style={{ fontSize: '12px', maxWidth: '280px', overflow: 'hidden', textOverflow: 'ellipsis', wordBreak: 'break-all' }}
                title={item.MemberDN || ''}
              >
                {item.MemberDN || '—'}
              </td>
            )}
            {privilegedChangesShowMemberDetail && privilegedChangesHasDetailNotes && (
              <td style={{ fontSize: '12px', maxWidth: '400px', color: 'var(--text-muted)' }}>
                {item.DetailNote || '—'}
              </td>
            )}
            <td>{item.AttributeChanged}</td>
            <td>{item.ChangeTime}</td>
            <td style={{ fontSize: '12px', maxWidth: '250px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.OriginatingDC}</td>
            <td>{item.ChangeVersion}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderStaleAdminsTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="SamAccountName" label="SAM Account" />
          <SortableHeader columnKey="PrivilegedGroup" label="Privileged Group" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Enabled" label="Enabled" />
          <SortableHeader columnKey="LastLogonDate" label="Last Logon" />
          <SortableHeader columnKey="DaysSinceLogon" label="Days Since Logon" />
          <SortableHeader columnKey="PasswordLastSet" label="Password Last Set" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.Name}</strong></td>
            <td>{item.SamAccountName}</td>
            <td>{item.PrivilegedGroup}</td>
            <td>{item.Domain}</td>
            <td>
              <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                {item.Enabled ? 'Yes' : 'No'}
              </span>
            </td>
            <td>{item.LastLogonDate}</td>
            <td>{item.DaysSinceLogon}</td>
            <td>{item.PasswordLastSet}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderPwdNeverExpiresTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="SamAccountName" label="SAM Account" />
          <SortableHeader columnKey="PrivilegedGroup" label="Privileged Group" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Enabled" label="Enabled" />
          <SortableHeader columnKey="PasswordLastSet" label="Password Last Set" />
          <SortableHeader columnKey="PasswordAgeDays" label="Password Age (Days)" />
          <SortableHeader columnKey="Risk" label="Risk" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.Name}</strong></td>
            <td>{item.SamAccountName}</td>
            <td>{item.PrivilegedGroup}</td>
            <td>{item.Domain}</td>
            <td>
              <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                {item.Enabled ? 'Yes' : 'No'}
              </span>
            </td>
            <td>{item.PasswordLastSet}</td>
            <td>{item.PasswordAgeDays}</td>
            <td>{getRiskBadge(item.Risk)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderKerberosDelegationTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="SamAccountName" label="SAM Account" />
          <SortableHeader columnKey="ObjectType" label="Object Type" />
          <SortableHeader columnKey="DelegationType" label="Delegation Type" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Enabled" label="Enabled" />
          <SortableHeader columnKey="DelegationTarget" label="Delegation Target" />
          <SortableHeader columnKey="Risk" label="Risk" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.Name}</strong></td>
            <td>{item.SamAccountName}</td>
            <td>{item.ObjectType}</td>
            <td>{getDelegationBadge(item.DelegationType)}</td>
            <td>{item.Domain}</td>
            <td>
              <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                {item.Enabled ? 'Yes' : 'No'}
              </span>
            </td>
            <td style={{ fontSize: '12px', maxWidth: '300px', wordBreak: 'break-all' }}>{item.DelegationTarget}</td>
            <td>{getRiskBadge(item.Risk)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderSIDHistoryTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="SamAccountName" label="SAM Account" />
          <SortableHeader columnKey="ObjectType" label="Object Type" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Enabled" label="Enabled" />
          <SortableHeader columnKey="SIDHistoryCount" label="SID Count" />
          <SortableHeader columnKey="LastLogonDate" label="Last Logon" />
          <SortableHeader columnKey="Risk" label="Risk" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.Name}</strong></td>
            <td>{item.SamAccountName}</td>
            <td>{item.ObjectType}</td>
            <td>{item.Domain}</td>
            <td>{String(item.Enabled)}</td>
            <td>{item.SIDHistoryCount}</td>
            <td>{item.LastLogonDate}</td>
            <td>{getRiskBadge(item.Risk)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderAdminSDHolderTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="Name" label="Name" />
          <SortableHeader columnKey="SamAccountName" label="SAM Account" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Enabled" label="Enabled" />
          <SortableHeader columnKey="IsOrphanedAdminCount" label="Orphaned" />
          <SortableHeader columnKey="PrivilegedGroups" label="Privileged Groups" />
          <SortableHeader columnKey="LastLogonDate" label="Last Logon" />
          <SortableHeader columnKey="PasswordAgeDays" label="Password Age" />
          <SortableHeader columnKey="Risk" label="Risk" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          <tr key={i}>
            <td><strong>{item.Name}</strong></td>
            <td>{item.SamAccountName}</td>
            <td>{item.Domain}</td>
            <td>
              <span className={`status-badge ${item.Enabled ? 'enabled' : 'disabled'}`}>
                {item.Enabled ? 'Yes' : 'No'}
              </span>
            </td>
            <td>
              {item.IsOrphanedAdminCount ? <span className="orphaned-badge">Orphaned</span> : 'No'}
            </td>
            <td style={{ fontSize: '12px', maxWidth: '250px', wordBreak: 'break-all' }}>{item.PrivilegedGroups}</td>
            <td>{item.LastLogonDate}</td>
            <td>{item.PasswordAgeDays}</td>
            <td>{getRiskBadge(item.Risk)}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );

  const renderPrivilegedMembershipAuditTable = () => (
    <table className="data-table">
      <thead>
        <tr>
          <SortableHeader columnKey="GroupName" label="Group Name" />
          <SortableHeader columnKey="Domain" label="Domain" />
          <SortableHeader columnKey="Action" label="Action" />
          <SortableHeader columnKey="Member" label="Member" />
          <SortableHeader columnKey="ChangedBy" label="Changed By" />
          <SortableHeader columnKey="ChangeTime" label="Change Time" />
          <SortableHeader columnKey="DomainController" label="DC" />
          <SortableHeader columnKey="EventId" label="Event ID" />
        </tr>
      </thead>
      <tbody>
        {paginatedData.map((item, i) => (
          item.Status === 'ERROR' ? (
            <tr key={i}>
              <td><strong>{item.Domain || '-'}</strong></td>
              <td style={{ fontSize: '12px' }}>{item.DomainController || '-'}</td>
              <td colSpan={6} style={{ color: '#fca5a5' }}>
                {item.Error || 'Unable to read DC Security logs'}
              </td>
            </tr>
          ) : (
            <tr key={i}>
              <td><strong>{item.GroupName}</strong></td>
              <td>{item.Domain}</td>
              <td>
                <span className={`status-badge ${String(item.Action).toLowerCase() === 'removed' ? 'disabled' : 'enabled'}`}>
                  {item.Action}
                </span>
              </td>
              <td style={{ fontSize: '12px', maxWidth: '280px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={item.MemberRaw || item.Member}>
                {item.Member || '-'}
              </td>
              <td>{item.ChangedBy || '-'}</td>
              <td>{item.ChangeTime}</td>
              <td style={{ fontSize: '12px', maxWidth: '240px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.DomainController}</td>
              <td>{item.EventId}</td>
            </tr>
          )
        ))}
      </tbody>
    </table>
  );

  const renderCombinedPrivilegedChangesTable = (rows, hasNotes) => {
    const showMember = memberDetailInPrivilegedChanges;
    return (
      <table className="data-table">
        <thead>
          <tr>
            <th>Group Name</th>
            <th>Domain</th>
            {showMember && <th>Change</th>}
            {showMember && <th>Member</th>}
            {showMember && <th>Member DN</th>}
            {showMember && hasNotes && <th>Note</th>}
            <th>Attribute</th>
            <th>Change Time</th>
            <th>Originating DC</th>
            <th>Version</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((item, i) => (
            <tr key={i}>
              <td><strong>{item.GroupName}</strong></td>
              <td>{item.Domain}</td>
              {showMember && (
                <td>
                  {item.ChangeType ? (
                    <span className={`status-badge ${String(item.ChangeType).toLowerCase() === 'removed' ? 'disabled' : 'enabled'}`}>
                      {item.ChangeType}
                    </span>
                  ) : '—'}
                </td>
              )}
              {showMember && <td>{item.MemberCN || '—'}</td>}
              {showMember && (
                <td
                  style={{ fontSize: '12px', maxWidth: '280px', overflow: 'hidden', textOverflow: 'ellipsis', wordBreak: 'break-all' }}
                  title={item.MemberDN || ''}
                >
                  {item.MemberDN || '—'}
                </td>
              )}
              {showMember && hasNotes && (
                <td style={{ fontSize: '12px', maxWidth: '400px', color: 'var(--text-muted)' }}>
                  {item.DetailNote || '—'}
                </td>
              )}
              <td>{item.AttributeChanged}</td>
              <td>{item.ChangeTime}</td>
              <td style={{ fontSize: '12px', maxWidth: '250px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.OriginatingDC}</td>
              <td>{item.ChangeVersion}</td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  const renderCombinedMembershipAuditTable = (rows) => (
    <table className="data-table">
      <thead>
        <tr>
          <th>Group Name</th>
          <th>Domain</th>
          <th>Action</th>
          <th>Member</th>
          <th>Changed By</th>
          <th>Change Time</th>
          <th>DC</th>
          <th>Event ID</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((item, i) => (
          item.Status === 'ERROR' ? (
            <tr key={i}>
              <td><strong>{item.Domain || '-'}</strong></td>
              <td style={{ fontSize: '12px' }}>{item.DomainController || '-'}</td>
              <td colSpan={6} style={{ color: '#fca5a5' }}>
                {item.Error || 'Unable to read DC Security logs'}
              </td>
            </tr>
          ) : (
            <tr key={i}>
              <td><strong>{item.GroupName}</strong></td>
              <td>{item.Domain}</td>
              <td>
                <span className={`status-badge ${String(item.Action).toLowerCase() === 'removed' ? 'disabled' : 'enabled'}`}>
                  {item.Action}
                </span>
              </td>
              <td style={{ fontSize: '12px', maxWidth: '280px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={item.MemberRaw || item.Member}>
                {item.Member || '-'}
              </td>
              <td>{item.ChangedBy || '-'}</td>
              <td>{item.ChangeTime}</td>
              <td style={{ fontSize: '12px', maxWidth: '240px', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.DomainController}</td>
              <td>{item.EventId}</td>
            </tr>
          )
        ))}
      </tbody>
    </table>
  );

  const renderBuiltinCombinedPanel = () => {
    if (!combinedBuiltin) return null;
    const chg = combinedBuiltin.changes || [];
    const aud = combinedBuiltin.audit || [];
    const chgHasNotes = chg.some((r) => typeof r.DetailNote === 'string' && r.DetailNote.trim() !== '');
    let chgHint = null;
    if (chg.length && chg[0]?.GroupName !== 'No changes found') {
      const hasMemberDn = chg.some((r) => typeof r.MemberDN === 'string' && r.MemberDN.trim() !== '');
      const hasDetailNote = chg.some((r) => typeof r.DetailNote === 'string' && r.DetailNote.trim());
      if (!hasMemberDn && !hasDetailNote) {
        chgHint = !memberDetailInPrivilegedChanges
          ? 'To see which members were added or removed, enable “show who was added or removed” under the lookback period, then run this report again.'
          : 'Per-member linked replication metadata was not returned. Try a longer lookback or confirm the service account can read linked-value replication metadata / Security logs on the PDC.';
      }
    }

    return (
      <>
        <h4 className="compliance-combined-section-title">Membership changes (replication metadata and Security event fallback)</h4>
        {chgHint && (
          <p className="compliance-hint-banner" style={{ margin: '0 0 12px', fontSize: '13px', color: 'var(--text-muted)' }}>
            {chgHint}
          </p>
        )}
        <div className="table-scroll">{renderCombinedPrivilegedChangesTable(chg, chgHasNotes)}</div>
        <h4 className="compliance-combined-section-title">Security log membership audit (who performed the change)</h4>
        <div className="table-scroll">{renderCombinedMembershipAuditTable(aud)}</div>
      </>
    );
  };

  const renderTable = () => {
    switch (activeQuery) {
      case 'builtin-administrators-combined':
        return renderBuiltinCombinedPanel();
      case 'privileged-changes':
      case 'privileged-changes-ea':
      case 'privileged-changes-sa':
      case 'privileged-changes-da':
      case 'privileged-changes-esa':
        return renderPrivilegedChangesTable();
      case 'privileged-membership-audit':
        return renderPrivilegedMembershipAuditTable();
      case 'stale-admins': return renderStaleAdminsTable();
      case 'pwd-never-expires-priv': return renderPwdNeverExpiresTable();
      case 'kerberos-delegation': return renderKerberosDelegationTable();
      case 'sid-history': return renderSIDHistoryTable();
      case 'adminsdholder': return renderAdminSDHolderTable();
      default: return null;
    }
  };

  const renderPagination = () => {
    if (activeQuery === 'builtin-administrators-combined') return null;
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
            <button onClick={() => currentPage > 1 && setCurrentPage(currentPage - 1)} disabled={currentPage <= 1} className="btn btn-secondary">
              Back
            </button>
            <span className="page-indicator">Page {currentPage} of {totalPages}</span>
            <button onClick={() => currentPage < totalPages && setCurrentPage(currentPage + 1)} disabled={currentPage >= totalPages} className="btn btn-secondary">
              Next
            </button>
          </div>
        </div>
      </div>
    );
  };

  const hasSingleReportRows = data.length > 0 && activeQuery !== 'builtin-administrators-combined';
  const hasCombinedReportRows =
    activeQuery === 'builtin-administrators-combined' &&
    combinedBuiltin &&
    ((combinedBuiltin.changes && combinedBuiltin.changes.length > 0) ||
      (combinedBuiltin.audit && combinedBuiltin.audit.length > 0));
  const showResultsPanel = !loading && (hasSingleReportRows || hasCombinedReportRows);
  const showEmptyState = !loading && Boolean(activeQuery) && !error && !showResultsPanel;

  return (
    <div className="compliance-page">
      <div className="page-header">
        <div className="page-title">
          <span className="page-icon">&#x1F6E1;&#xFE0F;</span>
          <h1>Compliance & Security Audit</h1>
        </div>
        <p className="page-description">Audit privileged access, delegation settings, SID history, and security compliance across your AD forest</p>
      </div>

      <div className="query-panel card">
        <h3>Security Audit Reports</h3>

        {/* Optional per-domain override */}
        <div className="days-input-section" style={{ marginTop: '10px' }}>
          <label>Optional: Run for a specific domain:</label>
          <input
            type="text"
            placeholder="e.g. cbs.ad.cbs.net"
            value={domainOverride}
            onChange={(e) => setDomainOverride(e.target.value)}
          />
          <button
            className="btn btn-secondary"
            style={{ marginLeft: '8px', padding: '6px 10px' }}
            onClick={() => setDomainOverride('')}
            disabled={!domainOverride}
            type="button"
          >
            Clear
          </button>
          <span style={{ color: 'var(--text-muted)', fontSize: '12px' }}>
            (leave empty to run forest-wide for the selected forest/domain in the top selector)
          </span>
        </div>

        {/* Days input for applicable queries */}
        <div className="days-input-section">
          <label>Lookback period (days):</label>
          <input
            type="number"
            min="1"
            max="365"
            value={daysInput}
            onChange={(e) => setDaysInput(parseInt(e.target.value) || 90)}
          />
          <span style={{ color: 'var(--text-muted)', fontSize: '12px' }}>
            (applies to privileged group / built-in Administrators change and membership-audit reports, and Stale Admin queries)
          </span>
        </div>

        <div className="days-input-section">
          <label style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', cursor: 'pointer' }}>
            <input
              type="checkbox"
              checked={memberDetailInPrivilegedChanges}
              onChange={(e) => setMemberDetailInPrivilegedChanges(e.target.checked)}
              style={{ marginTop: '4px' }}
            />
            <span>
              For <strong>Privileged Group Changes</strong> and <strong>Built-in Administrators (combined)</strong>, show <strong>who</strong> was added or removed (member-level replication metadata where available).
            </span>
          </label>
          <span style={{ color: 'var(--text-muted)', fontSize: '12px', display: 'block', marginTop: '8px', marginLeft: '28px' }}>
            Uncheck for a faster summary (only shows that the <code>member</code> attribute changed). Full-group forest scans are heavier with detail on.
          </span>
        </div>

        <div className="query-cards">
          {predefinedQueries.map((q) => (
            <div
              key={q.id}
              className={`query-card ${activeQuery === q.id ? 'active' : ''}`}
              onClick={() => handleQuery(q.id)}
            >
              <div className="query-card-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span>{q.label}</span>
                <FavoriteButton page="/compliance" queryId={q.id} label={q.label} />
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
          <span>Running security audit...</span>
        </div>
      )}

      {!loading && showResultsPanel && (
        <div className="results-panel card">
          <div className="results-header">
            <div>
              <span>
                {activeQuery === 'builtin-administrators-combined' && combinedBuiltin
                  ? `Built-in Administrators combined: ${combinedBuiltin.changes.length} change row(s), ${combinedBuiltin.audit.length} audit row(s)`
                  : `Found ${data.length} result(s)`}
              </span>
              {activeQuery !== 'builtin-administrators-combined' && (
                <span className="sort-hint">Click column headers to sort &#8597;</span>
              )}
            </div>
            <ExportButton
              data={complianceExportData}
              filename={`compliance_${activeQuery || 'export'}`}
              title="Compliance & Security Audit Report"
            />
          </div>
          {activeQuery !== 'builtin-administrators-combined' && privilegedChangesHint && (
            <p className="compliance-hint-banner" style={{ margin: '0 0 12px', fontSize: '13px', color: 'var(--text-muted)' }}>
              {privilegedChangesHint}
            </p>
          )}
          <div className="table-scroll">
            {renderTable()}
          </div>
          {renderPagination()}
        </div>
      )}

      {showEmptyState && (
        <div className="empty-state card">
          <div className="empty-icon">&#x2705;</div>
          <p>No issues found for this audit. Your environment looks clean!</p>
        </div>
      )}
    </div>
  );
};

export default Compliance;

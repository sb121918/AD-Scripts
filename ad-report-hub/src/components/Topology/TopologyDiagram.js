import React, { useRef, useEffect, useState, useCallback } from 'react';
import './TopologyDiagram.css';

// Color palette
const COLORS = {
  forest: '#16a34a',
  domain: '#2563eb',
  childDomain: '#3b82f6',
  dc: '#7c3aed',
  site: '#0891b2',
  fsmoForest: '#b45309',
  fsmoDomain: '#d97706',
  partition: '#6366f1',
  trustBidirectional: '#2563eb',
  trustInbound: '#ea580c',
  trustOutbound: '#7c3aed',
  trustDisabled: '#9ca3af',
  replicationOk: '#16a34a',
  replicationFail: '#dc2626',
  siteLink: '#0891b2',
  gcServer: '#059669',
};

// ─── Forest Info ────────────────────────────────────────────────────────────
// Data shape: [{ Property: "Forest Name", Value: "AD.viacom.com" }, ...]
const buildForestInfoGraph = (data) => {
  const nodes = [];
  const edges = [];

  const lookup = {};
  data.forEach(row => { lookup[row.Property] = row.Value; });

  const forestName = lookup['Forest Name'] || 'Forest';
  const rootDomain = lookup['Root Domain'] || forestName;
  const domainsStr = lookup['Domains'] || '';
  const sitesStr = lookup['Sites'] || '';
  const gcServersStr = lookup['Global Catalog Servers'] || '';
  const schemaMaster = lookup['Schema Master'] || '';
  const namingMaster = lookup['Domain Naming Master'] || '';
  const forestMode = lookup['Forest Mode'] || '';

  // Forest root node
  nodes.push({
    id: 'forest',
    label: `${forestName}\n(${forestMode})`,
    color: COLORS.forest,
    shape: 'diamond',
    size: 35,
    font: { color: '#fff', size: 13, bold: true, multi: 'md' },
    level: 0,
  });

  // Domains
  const domains = domainsStr.split(',').map(s => s.trim()).filter(Boolean);
  domains.forEach((d, i) => {
    const isRoot = d.toLowerCase() === rootDomain.toLowerCase();
    nodes.push({
      id: `domain-${d}`,
      label: d,
      color: isRoot ? COLORS.domain : COLORS.childDomain,
      shape: 'box',
      font: { color: '#fff', size: 12 },
      level: 1,
      borderWidth: isRoot ? 3 : 1,
    });
    edges.push({
      from: 'forest',
      to: `domain-${d}`,
      color: { color: '#888' },
      width: 2,
      label: isRoot ? 'Root' : 'Child',
      font: { size: 9, color: '#999' },
    });
  });

  // Sites
  const sites = sitesStr.split(',').map(s => s.trim()).filter(Boolean);
  if (sites.length > 0) {
    nodes.push({
      id: 'sites-group',
      label: `Sites (${sites.length})`,
      color: COLORS.site,
      shape: 'ellipse',
      font: { color: '#fff', size: 11 },
      level: 1,
    });
    edges.push({
      from: 'forest',
      to: 'sites-group',
      color: { color: '#888' },
      width: 1,
      dashes: true,
    });
    sites.forEach(site => {
      nodes.push({
        id: `site-${site}`,
        label: site,
        color: '#0e7490',
        shape: 'ellipse',
        size: 15,
        font: { color: '#fff', size: 10 },
        level: 2,
      });
      edges.push({
        from: 'sites-group',
        to: `site-${site}`,
        color: { color: COLORS.site },
        width: 1,
      });
    });
  }

  // GC Servers
  const gcServers = gcServersStr.split(',').map(s => s.trim()).filter(Boolean);
  if (gcServers.length > 0) {
    nodes.push({
      id: 'gc-group',
      label: `GC Servers (${gcServers.length})`,
      color: COLORS.gcServer,
      shape: 'ellipse',
      font: { color: '#fff', size: 11 },
      level: 1,
    });
    edges.push({
      from: 'forest',
      to: 'gc-group',
      color: { color: '#888' },
      width: 1,
      dashes: true,
    });
    gcServers.forEach(gc => {
      const shortName = gc.split('.')[0];
      nodes.push({
        id: `gc-${gc}`,
        label: shortName,
        color: '#047857',
        shape: 'box',
        size: 12,
        font: { color: '#fff', size: 9 },
        title: gc,
        level: 2,
      });
      edges.push({
        from: 'gc-group',
        to: `gc-${gc}`,
        color: { color: COLORS.gcServer },
        width: 1,
      });
    });
  }

  // FSMO roles
  if (schemaMaster) {
    const shortSchema = schemaMaster.split('.')[0];
    nodes.push({
      id: 'fsmo-schema',
      label: `Schema Master\n${shortSchema}`,
      color: COLORS.fsmoForest,
      shape: 'box',
      font: { color: '#fff', size: 10, multi: 'md' },
      title: schemaMaster,
      level: 2,
    });
    edges.push({
      from: 'forest',
      to: 'fsmo-schema',
      color: { color: COLORS.fsmoForest },
      width: 1,
      dashes: [5, 5],
      label: 'FSMO',
      font: { size: 8, color: '#999' },
    });
  }
  if (namingMaster) {
    const shortNaming = namingMaster.split('.')[0];
    nodes.push({
      id: 'fsmo-naming',
      label: `Naming Master\n${shortNaming}`,
      color: COLORS.fsmoForest,
      shape: 'box',
      font: { color: '#fff', size: 10, multi: 'md' },
      title: namingMaster,
      level: 2,
    });
    edges.push({
      from: 'forest',
      to: 'fsmo-naming',
      color: { color: COLORS.fsmoForest },
      width: 1,
      dashes: [5, 5],
      label: 'FSMO',
      font: { size: 8, color: '#999' },
    });
  }

  return { nodes, edges };
};

// ─── All Domains ────────────────────────────────────────────────────────────
// Data shape: [{ DomainName, NetBIOSName, DomainMode, ParentDomain, ChildDomains, DCCount, ... }]
const buildDomainsGraph = (data) => {
  const nodes = [];
  const edges = [];

  // Find forest root (no parent)
  const rootDomains = data.filter(d => !d.ParentDomain);
  const childDomains = data.filter(d => d.ParentDomain);

  data.forEach((item) => {
    const name = item.DomainName || '';
    if (!name || name === '-') return;

    const isRoot = !item.ParentDomain;
    const dcCount = item.DCCount || 0;
    const mode = item.DomainMode || '';
    const shortMode = mode.replace('Windows', 'Win').replace('Domain', '');

    nodes.push({
      id: name,
      label: `${name}\n${item.NetBIOSName || ''}\n${shortMode}\nDCs: ${dcCount}`,
      color: isRoot ? COLORS.forest : COLORS.childDomain,
      shape: isRoot ? 'diamond' : 'box',
      font: { color: '#fff', size: 11, multi: 'md' },
      size: isRoot ? 35 : 25,
      borderWidth: isRoot ? 3 : 1,
      title: [
        `Domain: ${name}`,
        `NetBIOS: ${item.NetBIOSName}`,
        `Mode: ${mode}`,
        `DCs: ${dcCount}`,
        `PDC: ${item.PDCEmulator || 'N/A'}`,
        `RID: ${item.RIDMaster || 'N/A'}`,
        `Infra: ${item.InfrastructureMaster || 'N/A'}`,
      ].join('\n'),
    });

    if (item.ParentDomain) {
      edges.push({
        from: item.ParentDomain,
        to: name,
        arrows: 'to',
        color: { color: '#888' },
        width: 2,
        label: 'Child',
        font: { size: 9, color: '#999' },
      });
    }

    // Show child domains as text if they exist but aren't separate rows
    if (item.ChildDomains) {
      const children = item.ChildDomains.split(',').map(s => s.trim()).filter(Boolean);
      children.forEach(child => {
        // Only add edge if the child isn't already a separate node
        const alreadyExists = data.some(d => d.DomainName === child);
        if (!alreadyExists) {
          nodes.push({
            id: child,
            label: child,
            color: COLORS.childDomain,
            shape: 'box',
            font: { color: '#fff', size: 11 },
            size: 20,
          });
          edges.push({
            from: name,
            to: child,
            arrows: 'to',
            color: { color: '#888' },
            width: 2,
            label: 'Child',
            font: { size: 9, color: '#999' },
          });
        }
      });
    }
  });

  return { nodes, edges };
};

// ─── AD Trusts ──────────────────────────────────────────────────────────────
// Data shape: [{ SourceDomain, TargetDomain, TrustType, TrustDirection, TrustAttributes, IsTransitive, ... }]
const buildTrustsGraph = (data) => {
  const nodes = [];
  const edges = [];
  const domainSet = new Set();

  data.forEach((trust) => {
    const source = trust.SourceDomain || 'Unknown';
    const target = trust.TargetDomain || 'Unknown';
    // Use TrustDirection (the actual field name from PowerShell)
    const direction = (trust.TrustDirection || trust.Direction || '').toLowerCase();
    const trustType = trust.TrustType || '';
    const attributes = trust.TrustAttributes || '';
    const isTransitive = trust.IsTransitive;

    // Skip placeholder rows
    if (target === 'No trusts found' || target === '-') return;

    if (!domainSet.has(source)) {
      domainSet.add(source);
      nodes.push({
        id: source,
        label: source,
        color: COLORS.domain,
        shape: 'box',
        font: { color: '#fff', size: 12 },
        size: 25,
        borderWidth: 2,
      });
    }
    if (!domainSet.has(target)) {
      domainSet.add(target);
      nodes.push({
        id: target,
        label: target,
        color: COLORS.domain,
        shape: 'box',
        font: { color: '#fff', size: 12 },
        size: 25,
        borderWidth: 2,
      });
    }

    let edgeColor = COLORS.trustBidirectional;
    let arrows = 'to,from';
    let dashes = false;

    if (direction.includes('disabled')) {
      edgeColor = COLORS.trustDisabled;
      arrows = '';
      dashes = true;
    } else if (direction.includes('inbound')) {
      edgeColor = COLORS.trustInbound;
      arrows = 'from';
    } else if (direction.includes('outbound')) {
      edgeColor = COLORS.trustOutbound;
      arrows = 'to';
    }

    const labelParts = [trustType];
    if (attributes && attributes !== 'Standard') {
      labelParts.push(`(${attributes})`);
    }

    edges.push({
      from: source,
      to: target,
      label: labelParts.join('\n'),
      color: { color: edgeColor, highlight: edgeColor },
      arrows,
      width: isTransitive ? 3 : 2,
      dashes,
      font: { size: 10, color: '#888', multi: 'md' },
      title: [
        `${source} → ${target}`,
        `Type: ${trustType}`,
        `Direction: ${trust.TrustDirection || direction}`,
        `Attributes: ${attributes}`,
        `Transitive: ${isTransitive ? 'Yes' : 'No'}`,
        trust.WhenCreated ? `Created: ${trust.WhenCreated}` : '',
      ].filter(Boolean).join('\n'),
    });
  });

  return { nodes, edges };
};

// ─── Site Links ─────────────────────────────────────────────────────────────
const buildSiteLinksGraph = (data) => {
  const nodes = [];
  const edges = [];
  const siteSet = new Set();

  data.forEach((link) => {
    const sitesStr = link.Sites || '';
    const sites = sitesStr.split(',').map(s => s.trim()).filter(Boolean);
    const cost = link.Cost || '';
    const interval = link.ReplicationInterval || link.Frequency || '';
    const linkName = link.SiteLinkName || link.Name || '';

    sites.forEach(site => {
      if (!siteSet.has(site)) {
        siteSet.add(site);
        nodes.push({
          id: site,
          label: site,
          color: COLORS.site,
          shape: 'ellipse',
          font: { color: '#fff', size: 12 },
          size: 25,
        });
      }
    });

    for (let i = 0; i < sites.length; i++) {
      for (let j = i + 1; j < sites.length; j++) {
        edges.push({
          from: sites[i],
          to: sites[j],
          label: `${linkName}\nCost: ${cost}${interval ? ` | ${interval}min` : ''}`,
          color: { color: COLORS.siteLink, highlight: COLORS.siteLink },
          width: 2,
          font: { size: 10, color: '#888', multi: 'md' },
          title: [
            `Link: ${linkName}`,
            `Cost: ${cost}`,
            `Interval: ${interval} min`,
            `Schedule: ${link.Schedule || 'N/A'}`,
            `Change Notification: ${link.ChangeNotification ? 'Yes' : 'No'}`,
          ].join('\n'),
        });
      }
    }
  });

  return { nodes, edges };
};

// ─── DC Replication ─────────────────────────────────────────────────────────
const buildReplicationGraph = (data) => {
  const nodes = [];
  const edges = [];
  const dcSet = new Set();

  data.forEach((entry) => {
    const source = entry.SourceDC || entry.Server || entry.Partner || 'Unknown';
    const target = entry.DestinationDC || entry.PartnerServer || entry.Destination || 'Unknown';
    const statusText = (entry.ReplicationStatus || entry.Status || '').toString().toLowerCase();
    const resultCode = entry.ResultCode;
    const isSuccess = statusText.includes('success') || resultCode === 0;
    const failCount = entry.FailureCount || 0;

    if (source === 'No replication data found' || source === 'Unable to query') return;

    if (!dcSet.has(source)) {
      dcSet.add(source);
      nodes.push({
        id: source,
        label: source.split('.')[0],
        color: COLORS.dc,
        shape: 'box',
        font: { color: '#fff', size: 11 },
        title: source,
        size: 20,
      });
    }
    if (!dcSet.has(target)) {
      dcSet.add(target);
      nodes.push({
        id: target,
        label: target.split('.')[0],
        color: COLORS.dc,
        shape: 'box',
        font: { color: '#fff', size: 11 },
        title: target,
        size: 20,
      });
    }

    edges.push({
      from: source,
      to: target,
      arrows: 'to',
      color: { color: isSuccess ? COLORS.replicationOk : COLORS.replicationFail },
      width: isSuccess ? 2 : 3,
      dashes: !isSuccess,
      label: entry.NamingContext || '',
      font: { size: 9, color: '#888' },
      title: [
        `${source} → ${target}`,
        `Status: ${entry.ReplicationStatus || 'Unknown'}`,
        `Last Success: ${entry.LastReplication || 'N/A'}`,
        `Failures: ${failCount}`,
        `NC: ${entry.NamingContext || 'N/A'}`,
      ].join('\n'),
    });
  });

  return { nodes, edges };
};

// ─── FSMO Roles ─────────────────────────────────────────────────────────────
// Data shape: [{ Role, Holder, Domain, Scope }]
const buildFSMOGraph = (data) => {
  const nodes = [];
  const edges = [];
  const holderSet = new Set();
  const domainSet = new Set();

  // Group by scope
  const forestRoles = data.filter(r => r.Scope === 'Forest');
  const domainRoles = data.filter(r => r.Scope === 'Domain');

  // Central "Forest" node
  nodes.push({
    id: 'forest-center',
    label: 'AD Forest',
    color: COLORS.forest,
    shape: 'diamond',
    size: 30,
    font: { color: '#fff', size: 13, bold: true },
    level: 0,
  });

  // Domain nodes
  const uniqueDomains = [...new Set(domainRoles.map(r => r.Domain))];
  uniqueDomains.forEach(domain => {
    if (!domainSet.has(domain)) {
      domainSet.add(domain);
      nodes.push({
        id: `domain-${domain}`,
        label: domain,
        color: COLORS.domain,
        shape: 'box',
        font: { color: '#fff', size: 12 },
        level: 1,
      });
      edges.push({
        from: 'forest-center',
        to: `domain-${domain}`,
        color: { color: '#888' },
        width: 2,
      });
    }
  });

  // Forest FSMO roles
  forestRoles.forEach(role => {
    const shortHolder = (role.Holder || '').split('.')[0];
    const nodeId = `fsmo-${role.Role}`;
    nodes.push({
      id: nodeId,
      label: `${role.Role}\n${shortHolder}`,
      color: COLORS.fsmoForest,
      shape: 'box',
      font: { color: '#fff', size: 10, multi: 'md' },
      title: `${role.Role}: ${role.Holder}`,
      level: 1,
    });
    edges.push({
      from: 'forest-center',
      to: nodeId,
      color: { color: COLORS.fsmoForest },
      width: 2,
      dashes: [5, 5],
      label: 'Forest FSMO',
      font: { size: 8, color: '#999' },
    });
  });

  // Domain FSMO roles
  domainRoles.forEach(role => {
    const shortHolder = (role.Holder || '').split('.')[0];
    const nodeId = `fsmo-${role.Domain}-${role.Role}`;
    nodes.push({
      id: nodeId,
      label: `${role.Role}\n${shortHolder}`,
      color: COLORS.fsmoDomain,
      shape: 'box',
      font: { color: '#fff', size: 10, multi: 'md' },
      title: `${role.Role}: ${role.Holder}\nDomain: ${role.Domain}`,
      level: 2,
    });
    edges.push({
      from: `domain-${role.Domain}`,
      to: nodeId,
      color: { color: COLORS.fsmoDomain },
      width: 1,
      dashes: [5, 5],
      label: 'Domain FSMO',
      font: { size: 8, color: '#999' },
    });
  });

  return { nodes, edges };
};

// ─── Naming Contexts ────────────────────────────────────────────────────────
// Data shape: [{ Name, Type, DistinguishedName }]
const buildNamingContextsGraph = (data) => {
  const nodes = [];
  const edges = [];

  nodes.push({
    id: 'root',
    label: 'Directory Partitions',
    color: COLORS.forest,
    shape: 'diamond',
    size: 30,
    font: { color: '#fff', size: 13, bold: true },
    level: 0,
  });

  const typeColors = {
    'Schema Partition': '#7c3aed',
    'Configuration Partition': '#b45309',
    'Domain Partition': COLORS.domain,
    'Application Partition': '#6366f1',
  };

  data.forEach((item, i) => {
    const color = typeColors[item.Type] || '#6b7280';
    const nodeId = `nc-${i}`;
    nodes.push({
      id: nodeId,
      label: `${item.Name}\n(${item.Type})`,
      color,
      shape: item.Type === 'Domain Partition' ? 'box' : 'ellipse',
      font: { color: '#fff', size: 11, multi: 'md' },
      title: `DN: ${item.DistinguishedName}`,
      level: 1,
    });
    edges.push({
      from: 'root',
      to: nodeId,
      color: { color },
      width: 2,
    });
  });

  return { nodes, edges };
};


// ═══════════════════════════════════════════════════════════════════════════
// Main Component
// ═══════════════════════════════════════════════════════════════════════════

const SUPPORTED_QUERIES = [
  'forest-info', 'domains', 'trusts', 'site-links',
  'dc-replication', 'fsmo-roles', 'naming-contexts',
];

const TopologyDiagram = ({ data, activeQuery }) => {
  const containerRef = useRef(null);
  const networkRef = useRef(null);
  const [selectedNode, setSelectedNode] = useState(null);

  const getGraphData = useCallback(() => {
    if (!data || data.length === 0) return null;

    switch (activeQuery) {
      case 'forest-info':
        return buildForestInfoGraph(data);
      case 'domains':
        return buildDomainsGraph(data);
      case 'trusts':
        return buildTrustsGraph(data);
      case 'site-links':
        return buildSiteLinksGraph(data);
      case 'dc-replication':
        return buildReplicationGraph(data);
      case 'fsmo-roles':
        return buildFSMOGraph(data);
      case 'naming-contexts':
        return buildNamingContextsGraph(data);
      default:
        return null;
    }
  }, [data, activeQuery]);

  const getLegend = () => {
    switch (activeQuery) {
      case 'forest-info':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.forest }} /> Forest</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.domain }} /> Root Domain</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.childDomain }} /> Child Domain</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.site }} /> Sites</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.gcServer }} /> GC Servers</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.fsmoForest }} /> FSMO Roles</div>
          </>
        );
      case 'domains':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.forest }} /> Forest Root</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.childDomain }} /> Child Domain</div>
          </>
        );
      case 'trusts':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.domain }} /> Domain</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.trustBidirectional }} /> Bidirectional</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.trustInbound }} /> Inbound</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.trustOutbound }} /> Outbound</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.trustDisabled, borderStyle: 'dashed' }} /> Disabled</div>
          </>
        );
      case 'site-links':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.site }} /> AD Site</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.siteLink }} /> Site Link</div>
          </>
        );
      case 'dc-replication':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.dc }} /> Domain Controller</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.replicationOk }} /> Replication OK</div>
            <div className="legend-item"><span className="legend-line" style={{ backgroundColor: COLORS.replicationFail }} /> Replication Failed</div>
          </>
        );
      case 'fsmo-roles':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.forest }} /> Forest</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.domain }} /> Domain</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.fsmoForest }} /> Forest FSMO</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.fsmoDomain }} /> Domain FSMO</div>
          </>
        );
      case 'naming-contexts':
        return (
          <>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.forest }} /> Root</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: '#7c3aed' }} /> Schema</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: '#b45309' }} /> Configuration</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: COLORS.domain }} /> Domain</div>
            <div className="legend-item"><span className="legend-dot" style={{ backgroundColor: '#6366f1' }} /> Application</div>
          </>
        );
      default:
        return null;
    }
  };

  // Determine if this query uses hierarchical layout
  const isHierarchical = ['forest-info', 'domains', 'fsmo-roles', 'naming-contexts'].includes(activeQuery);

  useEffect(() => {
    const graphData = getGraphData();
    if (!graphData || !containerRef.current) {
      if (networkRef.current) {
        networkRef.current.destroy();
        networkRef.current = null;
      }
      return;
    }

    const initNetwork = async () => {
      try {
        const { Network } = await import('vis-network/standalone');

        const options = {
          nodes: {
            borderWidth: 2,
            shadow: { enabled: true, size: 6 },
            font: { size: 12 },
            margin: { top: 8, bottom: 8, left: 10, right: 10 },
          },
          edges: {
            smooth: isHierarchical
              ? { type: 'cubicBezier', forceDirection: 'vertical', roundness: 0.4 }
              : { type: 'cubicBezier', forceDirection: 'horizontal', roundness: 0.4 },
            shadow: { enabled: true, size: 3 },
          },
          layout: {
            hierarchical: isHierarchical
              ? {
                  direction: 'UD',
                  sortMethod: 'directed',
                  levelSeparation: 140,
                  nodeSpacing: 180,
                  treeSpacing: 200,
                  blockShifting: true,
                  edgeMinimization: true,
                }
              : false,
          },
          physics: {
            enabled: !isHierarchical,
            barnesHut: {
              gravitationalConstant: -4000,
              centralGravity: 0.3,
              springLength: 250,
              springConstant: 0.04,
              damping: 0.09,
            },
            stabilization: { iterations: 200 },
          },
          interaction: {
            hover: true,
            tooltipDelay: 200,
            navigationButtons: true,
            keyboard: true,
            zoomView: true,
          },
        };

        if (networkRef.current) {
          networkRef.current.destroy();
        }

        const network = new Network(containerRef.current, graphData, options);
        networkRef.current = network;

        // Fit after stabilization
        network.once('stabilizationIterationsDone', () => {
          network.fit({ animation: { duration: 500, easingFunction: 'easeInOutQuad' } });
        });

        network.on('click', (params) => {
          if (params.nodes.length > 0) {
            const nodeId = params.nodes[0];
            // Find matching data row
            const nodeData = data.find(d =>
              (d.SourceDomain) === nodeId ||
              (d.TargetDomain) === nodeId ||
              (d.DomainName) === nodeId ||
              (d.Name) === nodeId ||
              (d.Value) === nodeId ||
              (d.Holder) === nodeId ||
              (d.SourceDC) === nodeId ||
              (d.DestinationDC) === nodeId
            );
            setSelectedNode({ id: nodeId, data: nodeData });
          } else {
            setSelectedNode(null);
          }
        });
      } catch (err) {
        console.error('Failed to load vis-network:', err);
      }
    };

    initNetwork();

    return () => {
      if (networkRef.current) {
        networkRef.current.destroy();
        networkRef.current = null;
      }
    };
  }, [data, activeQuery, getGraphData, isHierarchical]);

  const supportsDiagram = SUPPORTED_QUERIES.includes(activeQuery);

  if (!supportsDiagram) {
    return (
      <div className="topology-diagram-empty">
        <div className="empty-diagram-content">
          <span className="empty-diagram-icon">📊</span>
          <p>Diagram view is available for:</p>
          <ul>
            <li>Forest Information</li>
            <li>All Domains</li>
            <li>AD Trusts</li>
            <li>FSMO Role Holders</li>
            <li>Naming Contexts</li>
            <li>Site-to-Site Replication</li>
            <li>DC Replication Status</li>
          </ul>
        </div>
      </div>
    );
  }

  if (!data || data.length === 0) {
    return (
      <div className="topology-diagram-empty">
        Run a query to see the diagram visualization.
      </div>
    );
  }

  const graphData = getGraphData();
  if (!graphData || graphData.nodes.length === 0) {
    return (
      <div className="topology-diagram-empty">
        No graph data could be extracted from the results. Try the Table view instead.
      </div>
    );
  }

  return (
    <div className="topology-diagram-container">
      <div ref={containerRef} className="topology-diagram-canvas" />

      <div className="topology-diagram-legend">
        <h4>Legend</h4>
        {getLegend()}
        <div className="legend-hint">
          Scroll to zoom. Drag to pan. Click nodes for details.
        </div>
      </div>

      {selectedNode && (
        <div className="topology-node-detail">
          <div className="detail-header">
            <h4>{selectedNode.id}</h4>
            <button
              className="detail-close"
              onClick={() => setSelectedNode(null)}
              title="Close"
            >
              ✕
            </button>
          </div>
          {selectedNode.data ? (
            Object.entries(selectedNode.data)
              .filter(([key]) => key !== 'DistinguishedName') // Skip long DN by default
              .map(([key, value]) => (
                <div key={key} className="detail-row">
                  <span className="detail-label">{key}:</span>
                  <span className="detail-value">{String(value ?? 'N/A')}</span>
                </div>
              ))
          ) : (
            <div style={{ color: 'var(--text-muted)', fontSize: '12px' }}>
              No additional details for this node.
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default TopologyDiagram;

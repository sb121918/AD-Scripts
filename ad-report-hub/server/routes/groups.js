const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');
const { getForestNameForDomain } = require('../utils/domainConfig');
const { getPrivilegedGroupNamesForForest } = require('../utils/privilegedGroups');

// Predefined queries for Groups tab
const predefinedQueries = {
  'all': 'Get-AllGroups.ps1',
  'security': 'Get-SecurityGroups.ps1',
  'distribution': 'Get-DistributionGroups.ps1',
  'by-name': 'Get-GroupByName.ps1',
  'created-30-days': 'Get-GroupsCreatedLastXDays.ps1',
  'modified-30-days': 'Get-GroupsModifiedLastXDays.ps1',
  'empty': 'Get-EmptyGroups.ps1',
  'with-manager': 'Get-GroupsWithManager.ps1',
  'domain-local': 'Get-DomainLocalGroups.ps1',
  'global': 'Get-GlobalGroups.ps1',
  'universal': 'Get-UniversalGroups.ps1',
  'mail-enabled': 'Get-MailEnabledGroups.ps1',
  'dl-members': 'Get-DLMembers.ps1',
  // Privileged Groups
  'domain-admins': 'Get-DomainAdminsByDomain.ps1',
  'builtin-admins': 'Get-BuiltinAdminsByDomain.ps1',
  'enterprise-admins': 'Get-EnterpriseAdmins.ps1',
  'schema-admins': 'Get-SchemaAdmins.ps1',
  'enterprise-systems-admins': 'Get-EnterpriseSystemsAdmins.ps1',
  'privileged-groups': 'Get-PrivilegedGroups.ps1'
};

// GET /api/groups - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, value, domain, groupName, groupDN, limit } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const forestName = getForestNameForDomain(domain);
    const args = {};
    if (value) {
      args.SearchValue = value;
    }

    // DL members lookup (Distribution List)
    if (queryType === 'dl-members') {
      if (!groupName && !groupDN) {
        return res.status(400).json({ error: 'groupName or groupDN is required for DL members lookup' });
      }
      if (groupName) args.GroupName = groupName;
      if (groupDN) args.GroupDN = groupDN;
      if (limit !== undefined && limit !== null && limit !== '') {
        const n = parseInt(limit, 10);
        if (!Number.isNaN(n)) args.Limit = n;
      }
    }
    
    // Handle domain-specific queries (required for these types)
    if (['domain-admins', 'builtin-admins'].includes(queryType)) {
      if (!domain) {
        return res.status(400).json({ error: 'Domain parameter is required for this query' });
      }
      args.TargetDomain = domain;
    } else if (domain) {
      // Pass domain as optional target for scripts that support it
      args.TargetDomain = domain;
    }

    // Enforce forest-specific visibility for Viacom-only group report
    if (queryType === 'enterprise-systems-admins' && forestName && forestName.toLowerCase() !== 'viacom') {
      return res.json({ data: [], duration: 0 });
    }

    // Provide forest context + privileged group set to scripts that use it
    if (domain) {
      if (queryType === 'privileged-groups') {
        args.ForestDomain = domain;
        args.ForestName = forestName || '';
        args.PrivilegedGroupsCsv = getPrivilegedGroupNamesForForest(forestName || '').join(',');
      }
      if (queryType === 'enterprise-systems-admins') {
        args.ForestName = forestName || '';
        args.ForestDomain = domain;
      }
    }

    const result = await runPowerShell(scriptName, args, 'groups');
    res.json(result);
  } catch (error) {
    console.error('Group query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/groups/search - Search groups
router.get('/search', async (req, res) => {
  try {
    const { q, domain } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-Group.ps1', {
      SearchValue: q,
      TargetDomain: domain || ''
    }, 'groups');
    res.json(result);
  } catch (error) {
    console.error('Group search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/groups/:name/details - Get group details
router.get('/:name/details', async (req, res) => {
  try {
    const { name } = req.params;
    const { domain } = req.query;
    const result = await runPowerShell('Get-GroupDetails.ps1', {
      GroupName: name,
      TargetDomain: domain || ''
    }, 'groups');
    res.json(result);
  } catch (error) {
    console.error('Group details error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/groups/:name/members - Get group members
router.get('/:name/members', async (req, res) => {
  try {
    const { name } = req.params;
    const { domain } = req.query;
    const result = await runPowerShell('Get-GroupMembers.ps1', {
      GroupName: name,
      TargetDomain: domain || ''
    }, 'groups');
    res.json(result);
  } catch (error) {
    console.error('Group members error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/groups/export - Export groups
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json', includeMembers = false } = req.body;
    const result = await runPowerShell('Export-Groups.ps1', {
      ...filters,
      Format: format,
      IncludeMembers: includeMembers
    }, 'groups');
    res.json(result);
  } catch (error) {
    console.error('Group export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

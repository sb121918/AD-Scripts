const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Security reports
const predefinedQueries = {
  'adminsdholder': 'Get-AdminSDHolderUsers.ps1',
  'unconstrained-delegation': 'Get-UnconstrainedDelegation.ps1',
  'kerberos-preauth-disabled': 'Get-KerberosPreAuthDisabled.ps1',
  'reversible-encryption': 'Get-UsersReversibleEncryption.ps1',
  'nested-privileged': 'Get-NestedPrivilegedMembers.ps1',
  'stale-admins': 'Get-StaleAdminAccounts.ps1',
  'disabled-in-groups': 'Get-DisabledUsersInGroups.ps1'
};

// GET /api/security - Run security report query
router.get('/', async (req, res) => {
  try {
    const { query, domain, days } = req.query;
    
    if (!query) {
      return res.status(400).json({ error: 'Query type is required' });
    }
    
    const scriptName = predefinedQueries[query];
    
    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }
    
    const args = {};
    
    // Pass domain if provided
    if (domain && domain.trim()) {
      args.TargetDomain = domain.trim();
    }
    
    // Pass days parameter for stale-admins
    if (query === 'stale-admins' && days) {
      args.InactiveDays = parseInt(days, 10);
    }
    
    const result = await runPowerShell(scriptName, args, 'security');
    res.json(result);
  } catch (error) {
    console.error('Security query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/security/summary - Get security summary/dashboard data
router.get('/summary', async (req, res) => {
  try {
    const { domain } = req.query;
    const args = {};
    
    if (domain && domain.trim()) {
      args.TargetDomain = domain.trim();
    }
    
    // Run multiple queries to build summary
    const [
      adminSDHolder,
      unconstrainedDelegation,
      kerberosPreAuth,
      staleAdmins,
      disabledInGroups
    ] = await Promise.all([
      runPowerShell('Get-AdminSDHolderUsers.ps1', args, 'security').catch(() => ({ data: [] })),
      runPowerShell('Get-UnconstrainedDelegation.ps1', args, 'security').catch(() => ({ data: [] })),
      runPowerShell('Get-KerberosPreAuthDisabled.ps1', args, 'security').catch(() => ({ data: [] })),
      runPowerShell('Get-StaleAdminAccounts.ps1', { ...args, InactiveDays: 90 }, 'security').catch(() => ({ data: [] })),
      runPowerShell('Get-DisabledUsersInGroups.ps1', args, 'security').catch(() => ({ data: [] }))
    ]);
    
    const summary = {
      adminSDHolderCount: Array.isArray(adminSDHolder.data) ? adminSDHolder.data.length : 0,
      unconstrainedDelegationCount: Array.isArray(unconstrainedDelegation.data) ? unconstrainedDelegation.data.length : 0,
      kerberosPreAuthDisabledCount: Array.isArray(kerberosPreAuth.data) ? kerberosPreAuth.data.length : 0,
      staleAdminCount: Array.isArray(staleAdmins.data) ? staleAdmins.data.length : 0,
      disabledInGroupsCount: Array.isArray(disabledInGroups.data) ? disabledInGroups.data.length : 0,
      criticalIssues: 0,
      highIssues: 0,
      mediumIssues: 0
    };
    
    // Calculate risk levels
    const allData = [
      ...(adminSDHolder.data || []),
      ...(unconstrainedDelegation.data || []),
      ...(kerberosPreAuth.data || []),
      ...(staleAdmins.data || []),
      ...(disabledInGroups.data || [])
    ];
    
    allData.forEach(item => {
      if (item.RiskLevel === 'Critical') summary.criticalIssues++;
      else if (item.RiskLevel === 'High') summary.highIssues++;
      else if (item.RiskLevel === 'Medium') summary.mediumIssues++;
    });
    
    res.json({ data: summary });
  } catch (error) {
    console.error('Security summary error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

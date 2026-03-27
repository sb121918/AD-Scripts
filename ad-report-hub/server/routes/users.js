const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Users tab
const predefinedQueries = {
  'all': 'Get-AllUsers.ps1',
  'by-name': 'Get-UserByName.ps1',
  'by-email': 'Get-UserByEmail.ps1',
  'by-employeeid': 'Get-UserByEmployeeId.ps1',
  'created-30-days': 'Get-UsersCreatedLastXDays.ps1',
  'created-x-days': 'Get-UsersCreatedLastXDays.ps1',
  'modified-30-days': 'Get-UsersModifiedLastXDays.ps1',
  'never-logged-on': 'Get-UsersNeverLoggedOn.ps1',
  'not-logged-60-days': 'Get-UsersNotLoggedOnXDays.ps1',
  'not-logged-x-days': 'Get-UsersNotLoggedOnXDays.ps1',
  'enabled': 'Get-EnabledUsers.ps1',
  'disabled': 'Get-DisabledUsers.ps1',
  'disabled-forest': 'Get-AllDisabledUsersForest.ps1',
  'disabled-by-domain': 'Get-DisabledUsersByDomain.ps1',
  'locked-out': 'Get-UserLockoutDiagnostics.ps1',
  'password-expired': 'Get-PasswordExpiredUsers.ps1',
  'password-never-expires': 'Get-PasswordNeverExpires.ps1',
  'with-manager': 'Get-UsersWithManager.ps1',
  'without-manager': 'Get-UsersWithoutManager.ps1',
  'by-department': 'Get-UsersByDepartment.ps1',
  // New queries
  'empty-ea6': 'Get-UsersEmptyEA6.ps1',
  'enabled-forest': 'Get-AllEnabledUsersForest.ps1',
  'enabled-human-forest': 'Get-EnabledHumanAccountsForest.ps1',
  'account-type-count': 'Get-AccountTypeCount.ps1',
  'employee-type-count': 'Get-EmployeeTypeCount.ps1',
  'ea6-value-count': 'Get-EA6ValueCount.ps1',
  'enabled-contractors': 'Get-EnabledContractors.ps1',
  'enabled-human': 'Get-EnabledHumanAccounts.ps1',
  'enabled-by-domain': 'Get-EnabledUsersByDomain.ps1',
  // Lifecycle queries
  'modified-x-days': 'Get-UsersModifiedLastXDays.ps1',
  'missing-manager': 'Get-UsersWithMissingManager.ps1',
  'expiring-soon': 'Get-UsersExpiringSoon.ps1'
};

// GET /api/users - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, days, value, department, domain, limit, lookbackHours } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const args = {};
    
    // Handle days parameter
    if (queryType.includes('days') && days) {
      args.Days = days;
    }
    if (value) {
      args.SearchValue = value;
    }
    if (lookbackHours !== undefined && lookbackHours !== null && lookbackHours !== '') {
      const n = parseInt(lookbackHours, 10);
      if (!Number.isNaN(n)) args.LookbackHours = n;
    }
    if (department) {
      args.Department = department;
    }

    // Optional limit parameter (supported by some scripts)
    if (limit !== undefined && limit !== null && limit !== '') {
      const n = parseInt(limit, 10);
      if (!Number.isNaN(n)) args.Limit = n;
    }
    
    // Handle domain-specific queries (required for these types)
    if (['account-type-count', 'employee-type-count', 'ea6-value-count', 'enabled-contractors', 'enabled-human', 'enabled-by-domain', 'disabled-by-domain'].includes(queryType)) {
      if (!domain) {
        return res.status(400).json({ error: 'Domain parameter is required for this query' });
      }
      args.TargetDomain = domain;
    } else if (domain) {
      // Pass domain as optional target for all other queries (from global selector)
      args.TargetDomain = domain;
    }

    // Locked out user lookup requires a target identity
    if (queryType === 'locked-out' && (!value || !String(value).trim())) {
      return res.status(400).json({ error: 'User identity is required for locked-out lookup (samAccountName/UPN/email)' });
    }
    
    const result = await runPowerShell(scriptName, args, 'users');
    res.json(result);
  } catch (error) {
    console.error('User query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/users/search - Search users
router.get('/search', async (req, res) => {
  try {
    const { q, type = 'wildcard' } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-User.ps1', {
      SearchValue: q,
      SearchType: type
    }, 'users');
    res.json(result);
  } catch (error) {
    console.error('User search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/users/:sam/details - Get full user details
router.get('/:sam/details', async (req, res) => {
  try {
    const { sam } = req.params;
    const result = await runPowerShell('Get-UserDetails.ps1', {
      SamAccountName: sam
    }, 'users');
    res.json(result);
  } catch (error) {
    console.error('User details error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/users/:sam/groups - Get user's group memberships
router.get('/:sam/groups', async (req, res) => {
  try {
    const { sam } = req.params;
    const result = await runPowerShell('Get-UserGroups.ps1', {
      SamAccountName: sam
    }, 'users');
    res.json(result);
  } catch (error) {
    console.error('User groups error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/users/group-members - Get members of a specific group
router.get('/group-members', async (req, res) => {
  try {
    const { groupDN, domain } = req.query;
    if (!groupDN) {
      return res.status(400).json({ error: 'Group DN is required' });
    }
    const result = await runPowerShell('Get-GroupMembers.ps1', {
      GroupDN: groupDN,
      GroupDomain: domain || ''
    }, 'users');
    res.json(result);
  } catch (error) {
    console.error('Group members error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/users/export - Export users
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json' } = req.body;
    const result = await runPowerShell('Export-Users.ps1', {
      ...filters,
      Format: format
    }, 'users');
    res.json(result);
  } catch (error) {
    console.error('User export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

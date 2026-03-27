const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Containers & OUs tab
const predefinedQueries = {
  'all-ous': 'Get-AllOUs.ps1',
  'all-containers': 'Get-AllContainers.ps1',
  'by-name': 'Get-OUByName.ps1',
  'created-30-days': 'Get-OUsCreatedLastXDays.ps1',
  'with-gpo': 'Get-OUsWithGPO.ps1',
  'without-gpo': 'Get-OUsWithoutGPO.ps1',
  'empty': 'Get-EmptyOUs.ps1',
  'protected': 'Get-ProtectedOUs.ps1',
  'ou-tree': 'Get-OUTree.ps1'
};

// GET /api/containers - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, value, domain } = req.query;
    const queryType = query || 'all-ous';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const args = {};
    if (value) {
      args.SearchValue = value;
    }
    // Add domain/server parameter if specified
    if (domain && domain.trim()) {
      if (queryType === 'ou-tree') {
        // Get-OUTree.ps1 expects -Domain (not -Server)
        args.Domain = domain.trim();
      } else {
        args.Server = domain.trim();
      }
    }

    const result = await runPowerShell(scriptName, args, 'containers');
    // For OU tree we expect a single root object
    if (queryType === 'ou-tree') {
      return res.json({ data: result.data?.[0] || null, duration: result.duration });
    }
    res.json(result);
  } catch (error) {
    console.error('Container query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/containers/children - Get child objects of a specific OU
router.get('/children', async (req, res) => {
  try {
    const { dn, domain } = req.query;
    if (!dn) {
      return res.status(400).json({ error: 'Distinguished Name (dn) parameter required' });
    }

    const args = { DN: dn };
    if (domain && domain.trim()) {
      // Get-OUChildren.ps1 accepts optional -Server
      args.Server = domain.trim();
    }
    const result = await runPowerShell('Get-OUChildren.ps1', args, 'containers');
    // children endpoint should return a single object
    res.json({ data: result.data?.[0] || null, duration: result.duration });
  } catch (error) {
    console.error('OU children query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/containers/search - Search OUs
router.get('/search', async (req, res) => {
  try {
    const { q } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-OU.ps1', {
      SearchValue: q
    }, 'containers');
    res.json(result);
  } catch (error) {
    console.error('OU search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/containers/export - Export OUs
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json' } = req.body;
    const result = await runPowerShell('Export-OUs.ps1', {
      ...filters,
      Format: format
    }, 'containers');
    res.json(result);
  } catch (error) {
    console.error('OU export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

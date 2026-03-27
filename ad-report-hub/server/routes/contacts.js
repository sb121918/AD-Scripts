const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Contacts tab
const predefinedQueries = {
  'all': 'Get-AllContacts.ps1',
  'by-name': 'Get-ContactByName.ps1',
  'created-30-days': 'Get-ContactsCreatedLastXDays.ps1',
  'modified-30-days': 'Get-ContactsModifiedLastXDays.ps1',
  'with-email': 'Get-ContactsWithEmail.ps1',
  'without-email': 'Get-ContactsWithoutEmail.ps1'
};

// GET /api/contacts - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, value, domain } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const args = {};
    if (value) {
      args.SearchValue = value;
    }
    // Pass domain as optional target (from global selector)
    if (domain && domain.trim()) {
      args.TargetDomain = domain.trim();
    }

    const result = await runPowerShell(scriptName, args, 'contacts');
    res.json(result);
  } catch (error) {
    console.error('Contact query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/contacts/search - Search contacts
router.get('/search', async (req, res) => {
  try {
    const { q } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-Contact.ps1', {
      SearchValue: q
    }, 'contacts');
    res.json(result);
  } catch (error) {
    console.error('Contact search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/contacts/export - Export contacts
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json' } = req.body;
    const result = await runPowerShell('Export-Contacts.ps1', {
      ...filters,
      Format: format
    }, 'contacts');
    res.json(result);
  } catch (error) {
    console.error('Contact export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

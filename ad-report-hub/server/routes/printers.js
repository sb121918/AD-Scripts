const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');

// Predefined queries for Printers tab
const predefinedQueries = {
  'all': 'Get-AllPrinters.ps1',
  'by-domain': 'Get-PrintersByDomain.ps1',
  'by-name': 'Get-PrinterByName.ps1',
  'by-location': 'Get-PrintersByLocation.ps1',
  'by-server': 'Get-PrintersByServer.ps1',
  'color': 'Get-ColorPrinters.ps1',
  'duplex': 'Get-DuplexPrinters.ps1'
};

// GET /api/printers - Run predefined query
router.get('/', async (req, res) => {
  try {
    const { query, value, location, server, domain } = req.query;
    const queryType = query || 'all';
    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type' });
    }

    const args = {};
    if (value) args.SearchValue = value;
    if (location) args.Location = location;
    if (server) args.Server = server;
    if (domain) args.TargetDomain = domain;

    const result = await runPowerShell(scriptName, args, 'printers');
    res.json(result);
  } catch (error) {
    console.error('Printer query error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/printers/search - Search printers
router.get('/search', async (req, res) => {
  try {
    const { q } = req.query;
    if (!q) {
      return res.status(400).json({ error: 'Search query required' });
    }

    const result = await runPowerShell('Search-Printer.ps1', {
      SearchValue: q
    }, 'printers');
    res.json(result);
  } catch (error) {
    console.error('Printer search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/printers/export - Export printers
router.post('/export', async (req, res) => {
  try {
    const { filters, format = 'json' } = req.body;
    const result = await runPowerShell('Export-Printers.ps1', {
      ...filters,
      Format: format
    }, 'printers');
    res.json(result);
  } catch (error) {
    console.error('Printer export error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

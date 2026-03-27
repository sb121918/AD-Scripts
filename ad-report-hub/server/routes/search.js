const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const { runPowerShell } = require('../utils/powershell');

const SAVED_SEARCHES_PATH = path.join(__dirname, '../data/saved-searches.json');

const ensureSavedSearchesFile = () => {
  const dir = path.dirname(SAVED_SEARCHES_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(SAVED_SEARCHES_PATH)) {
    fs.writeFileSync(SAVED_SEARCHES_PATH, JSON.stringify([], null, 2));
  }
};

const readSavedSearches = () => {
  ensureSavedSearchesFile();
  try {
    const data = fs.readFileSync(SAVED_SEARCHES_PATH, 'utf8');
    return JSON.parse(data);
  } catch {
    return [];
  }
};

const writeSavedSearches = (searches) => {
  ensureSavedSearchesFile();
  fs.writeFileSync(SAVED_SEARCHES_PATH, JSON.stringify(searches, null, 2));
};

// GET /api/search - Global search (simple or LDAP)
router.get('/', async (req, res) => {
  try {
    const { q, ldapFilter, objectType, types = 'user,computer,group', pageSize = 25, page = 1 } = req.query;

    if (ldapFilter) {
      // LDAP filter mode - use objectType for single type or types for multiples
      const searchTypes = objectType || types;
      const result = await runPowerShell('Global-Search-LDAP.ps1', {
        LDAPFilter: ldapFilter,
        ObjectTypes: searchTypes,
        PageSize: parseInt(pageSize),
        Page: parseInt(page)
      }, '');
      return res.json(result);
    }

    if (!q) {
      return res.status(400).json({ error: 'Search query or LDAP filter required' });
    }

    const result = await runPowerShell('Global-Search.ps1', {
      SearchValue: q,
      ObjectTypes: types,
      PageSize: parseInt(pageSize),
      Page: parseInt(page)
    }, '');

    res.json(result);
  } catch (error) {
    console.error('Global search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/search/saved - List saved searches
router.get('/saved', async (req, res) => {
  try {
    const searches = readSavedSearches();
    res.json({ data: searches });
  } catch (error) {
    console.error('List saved searches error:', error);
    res.status(500).json({ error: error.message });
  }
});

// POST /api/search/saved - Save a search
router.post('/saved', async (req, res) => {
  try {
    const { name, ...searchData } = req.body;
    if (!name || !name.trim()) {
      return res.status(400).json({ error: 'Search name required' });
    }

    const searches = readSavedSearches();
    const id = `saved-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
    const newSearch = { id, name: name.trim(), ...searchData };
    searches.push(newSearch);
    writeSavedSearches(searches);

    res.json({ data: newSearch });
  } catch (error) {
    console.error('Save search error:', error);
    res.status(500).json({ error: error.message });
  }
});

// DELETE /api/search/saved/:id - Delete a saved search
router.delete('/saved/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const searches = readSavedSearches().filter((s) => s.id !== id);
    writeSavedSearches(searches);
    res.json({ success: true });
  } catch (error) {
    console.error('Delete search error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

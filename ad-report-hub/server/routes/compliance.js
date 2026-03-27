const express = require('express');
const router = express.Router();
const { runPowerShell } = require('../utils/powershell');
const { getForestNameForDomain } = require('../utils/domainConfig');
const { getPrivilegedGroupNamesForForest } = require('../utils/privilegedGroups');

/** Allowed single-group filter values for privileged change/membership reports (canonical names as in AD). */
const ALLOWED_PRIVILEGED_GROUP_FILTER_LIST = [
  'Enterprise Admins',
  'Schema Admins',
  'Domain Admins',
  'Administrators',
  'AD-Enterprise Systems Admins',
];
const ALLOWED_PRIVILEGED_GROUP_FILTERS = new Set(ALLOWED_PRIVILEGED_GROUP_FILTER_LIST);

/**
 * Normalize query param (handles duplicate keys as array, trimming, case, and common aliases).
 * Builtin Administrators must resolve to "Administrators" for Get-ADGroup -Identity expectations.
 */
function normalizePrivilegedGroupFilter(raw) {
  if (raw == null) return '';
  const first = Array.isArray(raw) ? raw[0] : raw;
  if (typeof first !== 'string') return '';
  const trimmed = first.trim();
  if (!trimmed) return '';
  if (ALLOWED_PRIVILEGED_GROUP_FILTERS.has(trimmed)) return trimmed;
  const key = trimmed.toLowerCase().replace(/\\/g, ' ');
  const aliases = {
    'enterprise admins': 'Enterprise Admins',
    'schema admins': 'Schema Admins',
    'domain admins': 'Domain Admins',
    administrators: 'Administrators',
    'built-in administrators': 'Administrators',
    'built in administrators': 'Administrators',
    'builtin administrators': 'Administrators',
    'builtin domain administrators': 'Administrators',
    'ad-enterprise systems admins': 'AD-Enterprise Systems Admins',
  };
  return aliases[key] || trimmed;
}

/** For Get-PrivilegedGroupChanges: include linked-value replication rows (per-member add/remove). Default true. */
function parseMemberDetailsFlag(raw) {
  if (raw === undefined || raw === null) return true;
  const s = String(Array.isArray(raw) ? raw[0] : raw)
    .trim()
    .toLowerCase();
  if (s === '') return true;
  if (['0', 'false', 'no', 'off'].includes(s)) return false;
  return true;
}

// Predefined queries for Compliance & Security Auditing
const predefinedQueries = {
  'privileged-changes': 'Get-PrivilegedGroupChanges.ps1',
  'privileged-membership-audit': 'Get-PrivilegedGroupMembershipAudit.ps1',
  'stale-admins': 'Get-StaleAdminAccounts.ps1',
  'pwd-never-expires-priv': 'Get-PasswordNeverExpiresPrivileged.ps1',
  'kerberos-delegation': 'Get-KerberosDelegation.ps1',
  'sid-history': 'Get-SIDHistory.ps1',
  'adminsdholder': 'Get-AdminSDHolderProtected.ps1',
};

// GET /api/compliance - Run predefined compliance query
router.get('/', async (req, res) => {
  try {
    const { query, days, domain, targetDomain, groupFilter, memberDetails } = req.query;
    const queryType = query;

    if (queryType === 'builtin-administrators-combined') {
      if (!domain) {
        return res.status(400).json({ error: 'domain (forest anchor) is required' });
      }
      const daysNum = Math.min(365, Math.max(1, parseInt(days, 10) || 90));
      const td = (targetDomain || domain || '').trim();
      const baseArgs = {
        Days: daysNum,
        ForestDomain: domain,
        PrivilegedGroupsCsv: 'Administrators',
      };
      if (td) baseArgs.TargetDomain = td;

      const changeArgs = { ...baseArgs };
      if (parseMemberDetailsFlag(memberDetails)) {
        changeArgs.IncludeMemberDetails = true;
      }

      const auditArgs = {
        ...baseArgs,
        MaxEventsPerDomain: 2000,
        MaxEventsPerDC: 250,
      };

      const [chg, aud] = await Promise.all([
        runPowerShell('Get-PrivilegedGroupChanges.ps1', changeArgs, 'compliance'),
        runPowerShell('Get-PrivilegedGroupMembershipAudit.ps1', auditArgs, 'compliance'),
      ]);

      const asArray = (r) => (r && Array.isArray(r.data) ? r.data : []);

      return res.json({
        data: {
          builtinCombined: true,
          changes: asArray(chg),
          membershipAudit: asArray(aud),
        },
        duration: (chg.duration || 0) + (aud.duration || 0),
      });
    }

    const scriptName = predefinedQueries[queryType];

    if (!scriptName) {
      return res.status(400).json({ error: 'Invalid query type. Valid types: ' + Object.keys(predefinedQueries).join(', ') + ', builtin-administrators-combined' });
    }

    const forestName = getForestNameForDomain(domain);
    const args = {};
    if (days && (queryType === 'privileged-changes' || queryType === 'privileged-membership-audit' || queryType === 'stale-admins')) {
      args.Days = parseInt(days);
    }
    // domain = forest anchor (selected in navbar)
    if (domain) {
      args.ForestDomain = domain;
    }
    // targetDomain = optional per-domain filter override
    // Default behavior: scope to selected domain unless an override is provided.
    // This keeps Compliance consistent with other pages that treat the top selector as the target domain.
    args.TargetDomain = (targetDomain || domain || '').trim();
    if (!args.TargetDomain) delete args.TargetDomain;

    // Ensure "privileged" compliance reports audit the correct privileged set per forest
    const privilegedCsvTypes = ['privileged-changes', 'privileged-membership-audit', 'stale-admins', 'pwd-never-expires-priv', 'adminsdholder'];
    const groupFilterResolved = normalizePrivilegedGroupFilter(groupFilter);
    const canUseGroupFilter = ['privileged-changes', 'privileged-membership-audit'].includes(queryType);

    if (domain && privilegedCsvTypes.includes(queryType)) {
      if (groupFilterResolved && canUseGroupFilter) {
        if (!ALLOWED_PRIVILEGED_GROUP_FILTERS.has(groupFilterResolved)) {
          return res.status(400).json({
            error:
              'Invalid groupFilter. Allowed values: ' +
              ALLOWED_PRIVILEGED_GROUP_FILTER_LIST.join(', '),
          });
        }
        args.PrivilegedGroupsCsv = groupFilterResolved;
      } else {
        args.PrivilegedGroupsCsv = getPrivilegedGroupNamesForForest(forestName || '').join(',');
      }
    }

    if (queryType === 'privileged-changes' && parseMemberDetailsFlag(memberDetails)) {
      args.IncludeMemberDetails = true;
    }

    const result = await runPowerShell(scriptName, args, 'compliance');
    res.json(result);
  } catch (error) {
    console.error('Compliance query error:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;

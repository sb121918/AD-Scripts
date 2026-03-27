function getPrivilegedGroupNamesForForest(forestName) {
  const base = [
    'Domain Admins',
    'Enterprise Admins',
    'Schema Admins',
    'Administrators',
  ];

  if (String(forestName || '').toLowerCase() === 'viacom') {
    // Custom group that exists only in Viacom forest
    base.push('AD-Enterprise Systems Admins');
  }

  // Deduplicate while preserving order
  const seen = new Set();
  return base.filter((g) => {
    const key = String(g).toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

module.exports = {
  getPrivilegedGroupNamesForForest,
};


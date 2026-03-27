const fs = require('fs');
const path = require('path');

function readDomainsConfig() {
  const configPath = path.join(__dirname, '../config/domains.json');
  const raw = fs.readFileSync(configPath, 'utf8');
  return JSON.parse(raw);
}

function getDomainEntry(domainName) {
  if (!domainName) return null;
  const cfg = readDomainsConfig();
  const domains = cfg?.domains || [];
  return domains.find((d) => String(d.name).toLowerCase() === String(domainName).toLowerCase()) || null;
}

function getForestNameForDomain(domainName) {
  return getDomainEntry(domainName)?.forest || null;
}

module.exports = {
  readDomainsConfig,
  getDomainEntry,
  getForestNameForDomain,
};


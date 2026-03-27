import React from 'react';
import { Link, useLocation } from 'react-router-dom';
import './Breadcrumb.css';

const routeMap = {
  '/dashboard': { label: 'Home', group: null },
  '/users': { label: 'Users', group: 'Identity' },
  '/groups': { label: 'Groups', group: 'Identity' },
  '/service-accounts': { label: 'Service Accounts', group: 'Identity' },
  '/contacts': { label: 'Contacts', group: 'Identity' },
  '/domain-controllers': { label: 'Domain Controllers', group: 'Infrastructure' },
  '/computers': { label: 'Computers', group: 'Infrastructure' },
  '/sites-subnets': { label: 'Sites & Subnets', group: 'Infrastructure' },
  '/topology': { label: 'AD Topology', group: 'Infrastructure' },
  '/gpos': { label: 'Group Policy', group: 'Infrastructure' },
  '/containers': { label: 'Containers', group: 'Infrastructure' },
  '/security': { label: 'Security', group: 'Security & Governance' },
  '/compliance': { label: 'Compliance', group: 'Security & Governance' },
  '/search': { label: 'Search', group: 'Quick Links' },
  '/activity-logs': { label: 'Activity Logs', group: 'Quick Links' },
  '/help': { label: 'Help & Release Notes', group: 'Quick Links' },
  '/license': { label: 'License', group: 'Quick Links' },
};

const Breadcrumb = ({ activeReport }) => {
  const location = useLocation();
  const path = location.pathname;
  const route = routeMap[path];

  if (!route || path === '/dashboard') return null;

  return (
    <nav className="breadcrumb" aria-label="Breadcrumb">
      <ol>
        <li>
          <Link to="/dashboard" className="breadcrumb-link">Home</Link>
          <span className="breadcrumb-sep" aria-hidden="true">/</span>
        </li>
        {route.group && (
          <li>
            <span className="breadcrumb-group">{route.group}</span>
            <span className="breadcrumb-sep" aria-hidden="true">/</span>
          </li>
        )}
        <li>
          {activeReport ? (
            <>
              <Link to={path} className="breadcrumb-link">{route.label}</Link>
              <span className="breadcrumb-sep" aria-hidden="true">/</span>
            </>
          ) : (
            <span className="breadcrumb-current" aria-current="page">{route.label}</span>
          )}
        </li>
        {activeReport && (
          <li>
            <span className="breadcrumb-current" aria-current="page">{activeReport}</span>
          </li>
        )}
      </ol>
    </nav>
  );
};

export default Breadcrumb;

import React, { useState } from 'react';
import './Help.css';

const Help = () => {
  const [activeTab, setActiveTab] = useState('release-notes');
  const [expandedFaq, setExpandedFaq] = useState(null);

  const toggleFaq = (id) => {
    setExpandedFaq(expandedFaq === id ? null : id);
  };

  const tabs = [
    { id: 'release-notes', label: 'Release Notes' },
    { id: 'faq', label: 'FAQ' },
    { id: 'getting-started', label: 'Getting Started' },
  ];

  const releases = [
    {
      version: '2.1.0',
      date: 'February 2026',
      codename: 'Sentinel',
      summary: 'Enterprise security reporting, lifecycle management, and professional UI overhaul.',
      sections: [
        {
          title: 'Security & Compliance',
          type: 'new',
          items: [
            'New Security page with 7 enterprise-grade security assessment reports',
            'AdminSDHolder protected users detection — identify accounts with elevated adminCount flags',
            'Unconstrained Kerberos delegation analysis — flag users and computers trusted for delegation (excluding DCs)',
            'Kerberos Pre-Authentication disabled report — detect accounts vulnerable to AS-REP Roasting',
            'Reversible encryption detection — identify accounts storing passwords with reversible encryption',
            'Nested privileged group membership analysis — trace indirect admin paths through group nesting',
            'Stale admin accounts report — surface inactive privileged accounts (configurable threshold)',
            'Disabled users still in groups — compliance report for accounts retaining group memberships after disablement',
          ]
        },
        {
          title: 'User Lifecycle Reports',
          type: 'new',
          items: [
            'Users Modified in X Days — track recent attribute changes across the directory',
            'Users with Missing Manager — identify enabled accounts with no manager assigned',
            'Users Expiring Soon — surface accounts approaching or past their expiration date',
            'Employee Type Count — standalone report for EmployeeType attribute distribution per domain',
            'EA6 Value Count — standalone report for extensionAttribute6 distribution per domain',
          ]
        },
        {
          title: 'User Interface',
          type: 'improved',
          items: [
            'Users page reorganized into Forest-Wide and Domain-Specific report sections',
            'Professional UI refresh — refined typography, spacing, and visual hierarchy across all pages',
            'Consistent pagination controls migrated from inline styles to CSS classes across all 10 report pages',
            'Striped table rows, uppercase column headers, and sticky header support for better data readability',
            'Refined button styles with active press states, improved disabled appearance, and subtle shadows',
            'Dashboard hero section redesigned with cleaner layout and accent gradient top border',
            'Query cards updated with grid layout, refined hover states, and subtle border animations',
          ]
        },
        {
          title: 'Accessibility & Standards',
          type: 'improved',
          items: [
            'ARIA labels added to all modal close buttons, theme toggle, and navigation controls',
            'Keyboard navigation support for query cards (Enter/Space activation, role and tabIndex attributes)',
            'prefers-reduced-motion media query — disables animations for users with motion sensitivity',
            'Focus-visible outlines on buttons and interactive elements for keyboard-only navigation',
            'Scrollbar styling updated to respect theme variables in both light and dark modes',
          ]
        },
        {
          title: 'Theme & Dark Mode',
          type: 'improved',
          items: [
            'Footer component migrated from hardcoded colors to CSS variables — full dark mode support',
            'License page migrated from fixed hex values to theme-aware variables',
            'Missing --primary-color CSS variable added (used by profile header gradient)',
            'Font stack consolidated to Roboto — removed conflicting Inter declaration',
            'Theme-color meta tag aligned with accent-primary (#0055d4)',
          ]
        },
        {
          title: 'Bug Fixes',
          type: 'fixed',
          items: [
            'Account Type Count report now correctly extracts summary data from server response (array unwrapping)',
            'Domain input fields are now independent per query card — typing in one no longer affects others',
            'Created missing PowerShell scripts referenced by contacts routes (Get-ContactByName, Get-ContactsCreatedLastXDays, and 3 others)',
          ]
        },
      ]
    },
    {
      version: '2.0.0',
      date: 'January 2026',
      codename: 'Horizon',
      summary: 'Global search, LDAP query builder, compliance reporting, and favorites system.',
      sections: [
        {
          title: 'Core Features',
          type: 'new',
          items: [
            'Global search across all AD object types — users, groups, computers, contacts, OUs',
            'Visual LDAP query builder with AND/OR logic, attribute selection, and raw LDAP mode',
            'Saved searches — persist and recall favorite queries with custom names',
            'Compliance reports — delegation analysis, orphaned accounts, and risk assessment',
            'Favorites system — bookmark frequently used reports with star toggle on every query card',
            'Keyboard shortcuts — Ctrl+K for search, ? for help overlay, Escape to close modals',
            'Activity logging — all API requests tracked with user, action, target, status, and duration',
          ]
        },
        {
          title: 'Infrastructure',
          type: 'new',
          items: [
            'Forest/domain selector in the navigation bar for multi-domain targeting',
            'AD Topology page — forest info, domains, FSMO roles, trusts, replication status',
            'DC Resource Health — CPU, RAM, disk usage monitoring across domain controllers',
            'DC Services Status — service matrix (NTDS, DNS, Kdc, Netlogon, etc.) per DC',
            'Sites & Subnets — site links, subnet assignments, unassigned subnets, sites with no DC',
          ]
        },
        {
          title: 'Export & Data',
          type: 'improved',
          items: [
            'Export to Excel (.xlsx), CSV, JSON, and PDF (print) formats from any report',
            'Sortable table columns — click headers to sort ascending/descending',
            'Configurable pagination — 25, 50, or 100 results per page',
          ]
        },
      ]
    },
    {
      version: '1.0.0',
      date: 'December 2025',
      codename: 'Foundation',
      summary: 'Initial release — core identity and infrastructure reporting platform.',
      sections: [
        {
          title: 'Identity Reports',
          type: 'new',
          items: [
            'User reports — enabled/disabled, password status, creation date, lockout diagnostics',
            'Group reports — security, distribution, mail-enabled, empty groups, privileged group analysis',
            'Service account inventory — all accounts, orphaned, inactive, per-domain views',
            'Contact management — full contact listing and search',
          ]
        },
        {
          title: 'Infrastructure',
          type: 'new',
          items: [
            'Domain controller listing with site, IP, OS, and FSMO role information',
            'Computer inventory — enabled, disabled, never logged on, OS-based filtering',
            'Group Policy — all GPOs, unlinked, disabled, password policies',
            'OU container management with hierarchical tree view',
            'Printer inventory by domain with search',
          ]
        },
        {
          title: 'Platform',
          type: 'new',
          items: [
            'React frontend with Node.js/Express backend and PowerShell AD integration',
            'Windows Integrated Authentication — no separate login required',
            'Light and dark theme support with localStorage persistence',
            'Real-time queries against Active Directory — no caching or stale data',
          ]
        },
      ]
    },
  ];

  const faqs = [
    {
      category: 'General',
      items: [
        {
          id: 'what-is',
          q: 'What is AD Report Hub?',
          a: 'AD Report Hub is an enterprise-grade Active Directory reporting and visibility platform. It provides real-time querying across identity objects (users, groups, service accounts), infrastructure components (domain controllers, sites, GPOs), and security posture (privileged access, delegation, compliance) — all through a unified web interface.'
        },
        {
          id: 'auth',
          q: 'How does authentication work?',
          a: 'AD Report Hub uses Windows Integrated Authentication (Kerberos/NTLM). It automatically authenticates using the credentials of the currently logged-in Windows user. No separate login is required. The application inherits the AD read permissions of the user running it.'
        },
        {
          id: 'permissions',
          q: 'What permissions are required to run reports?',
          a: 'Most reports require standard Active Directory read permissions (typically granted to all authenticated domain users). Certain reports — such as DC Resource Health, replication status, and security event correlation — may require elevated permissions (e.g., Domain Admins or delegated read access to specific OUs and event logs). The START_AS_ADMIN.bat launcher is provided for these scenarios.'
        },
        {
          id: 'real-time',
          q: 'Is the data real-time or cached?',
          a: 'All queries execute in real-time against Active Directory via PowerShell. There is no caching layer — every report reflects the current state of the directory at the time of execution. This ensures accuracy but means large forest-wide queries may take longer depending on the size of your environment.'
        },
      ]
    },
    {
      category: 'Reports & Data',
      items: [
        {
          id: 'export',
          q: 'What export formats are supported?',
          a: 'Reports can be exported in four formats: Excel (.xlsx) for spreadsheet analysis, CSV for data processing and imports, JSON for programmatic consumption, and PDF via the browser print dialog for formal documentation. The export button appears on any report that returns results.'
        },
        {
          id: 'forest-vs-domain',
          q: 'What is the difference between Forest-Wide and Domain-Specific reports?',
          a: 'Forest-Wide reports query across all domains in the AD forest simultaneously. They do not require a domain input and return results from every domain. Domain-Specific reports target a single domain — you must enter the domain FQDN (e.g., corp.example.com) before running the query. Use the global domain selector in the navigation bar to set a default domain context.'
        },
        {
          id: 'large-results',
          q: 'How are large result sets handled?',
          a: 'Results are paginated client-side with configurable page sizes (25, 50, or 100 rows). All data is loaded into the browser and paginated locally, which means sorting and page navigation are instant. For very large environments (50,000+ objects), forest-wide queries may take 30-60 seconds to complete — a loading indicator is shown during execution.'
        },
        {
          id: 'sort',
          q: 'Can I sort report results?',
          a: 'Yes. Click any column header to sort ascending. Click again to sort descending. The sort indicator arrow shows the current sort direction. Sorting is performed client-side on the full result set, not just the visible page.'
        },
      ]
    },
    {
      category: 'Security Reports',
      items: [
        {
          id: 'adminsdholder',
          q: 'What does the AdminSDHolder report show?',
          a: 'The AdminSDHolder report identifies user accounts with adminCount=1. These accounts have been (or currently are) members of protected groups and have their ACLs overwritten by the AdminSDHolder process every 60 minutes. This includes current admins and potentially former admins whose adminCount was never cleared — a common security hygiene issue.'
        },
        {
          id: 'delegation',
          q: 'Why is unconstrained delegation a security risk?',
          a: 'Unconstrained delegation allows a service to impersonate any user who authenticates to it, by caching their Kerberos TGT. If an attacker compromises a server with unconstrained delegation, they can extract cached TGTs and impersonate any user — including Domain Admins. Domain controllers are excluded from this report as they inherently require delegation. All other flagged objects should be reviewed and migrated to constrained or resource-based delegation.'
        },
        {
          id: 'asrep',
          q: 'What is AS-REP Roasting?',
          a: 'AS-REP Roasting targets accounts that have Kerberos pre-authentication disabled. An attacker can request an AS-REP for these accounts without knowing their password, then crack the encrypted portion offline. The Kerberos Pre-Auth Disabled report identifies all such accounts so they can be remediated by enabling pre-authentication.'
        },
        {
          id: 'nested-priv',
          q: 'How does nested privileged group analysis work?',
          a: 'The report recursively enumerates membership of privileged groups (Domain Admins, Enterprise Admins, Schema Admins, Administrators, Account Operators, Backup Operators, Server Operators). It traces nested group paths to reveal users who have indirect privileged access through group nesting — a common blind spot in access reviews. Each result shows the full membership path.'
        },
      ]
    },
    {
      category: 'Troubleshooting',
      items: [
        {
          id: 'slow',
          q: 'Reports are running slowly. What can I do?',
          a: 'Forest-wide reports query every domain controller and aggregate results, which can be slow in large environments. Try using Domain-Specific reports to target a single domain. Ensure the machine running AD Report Hub has good network connectivity to domain controllers. If a specific report consistently times out, it may indicate an unreachable DC or DNS resolution issue.'
        },
        {
          id: 'no-data',
          q: 'A report returns no data. What should I check?',
          a: 'Verify that (1) the domain name is entered correctly for domain-specific reports, (2) your account has read permissions to the target domain, (3) the Active Directory PowerShell module is installed on the server, and (4) the target domain controller is reachable. Check the Activity Logs page for error details — failed queries are logged with the error message.'
        },
        {
          id: 'lockout',
          q: 'The lockout diagnostics report shows no events. Why?',
          a: 'Lockout event correlation requires Security event log access on domain controllers (Event ID 4740). If the running account lacks permission to read DC security logs, the report will return the AD lockout state but no event details. Run the application with START_AS_ADMIN.bat or ensure the account has "Read" access to the Security event log on DCs.'
        },
        {
          id: 'keyboard',
          q: 'What keyboard shortcuts are available?',
          a: 'Press ? (question mark) to open the keyboard shortcuts overlay. Key shortcuts: Ctrl+K or Cmd+K opens global search, Escape closes any open modal or the help overlay, Alt+H navigates to the dashboard. Shortcuts are disabled when focus is inside a text input field.'
        },
      ]
    },
  ];

  const renderReleaseNotes = () => (
    <div className="release-notes">
      {releases.map((release) => (
        <div key={release.version} className="release-block">
          <div className="release-block-header">
            <div className="release-meta">
              <span className="release-badge">{release.version}</span>
              <span className="release-codename">{release.codename}</span>
              <span className="release-block-date">{release.date}</span>
            </div>
            <p className="release-summary">{release.summary}</p>
          </div>
          <div className="release-sections">
            {release.sections.map((section, si) => (
              <div key={si} className="release-section">
                <div className="release-section-header">
                  <span className={`release-section-tag tag-${section.type}`}>
                    {section.type === 'new' ? 'NEW' : section.type === 'improved' ? 'IMPROVED' : 'FIXED'}
                  </span>
                  <h4>{section.title}</h4>
                </div>
                <ul className="release-section-items">
                  {section.items.map((item, ii) => (
                    <li key={ii}>{item}</li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );

  const renderFaq = () => (
    <div className="faq-container">
      {faqs.map((category) => (
        <div key={category.category} className="faq-category">
          <h3 className="faq-category-title">{category.category}</h3>
          <div className="faq-items">
            {category.items.map((item) => (
              <div
                key={item.id}
                className={`faq-item ${expandedFaq === item.id ? 'expanded' : ''}`}
              >
                <button
                  className="faq-question"
                  onClick={() => toggleFaq(item.id)}
                  aria-expanded={expandedFaq === item.id}
                >
                  <span className="faq-q-text">{item.q}</span>
                  <span className="faq-chevron">{expandedFaq === item.id ? '−' : '+'}</span>
                </button>
                {expandedFaq === item.id && (
                  <div className="faq-answer">
                    <p>{item.a}</p>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );

  const renderGettingStarted = () => (
    <div className="getting-started">
      <div className="gs-section">
        <div className="gs-step-number">01</div>
        <div className="gs-step-content">
          <h3>Launch the Application</h3>
          <p>
            Run <code>START.bat</code> for standard user mode or <code>START_AS_ADMIN.bat</code> for
            elevated operations (DC health checks, event log access). The application opens automatically
            at <code>http://localhost:3000</code>.
          </p>
        </div>
      </div>

      <div className="gs-section">
        <div className="gs-step-number">02</div>
        <div className="gs-step-content">
          <h3>Set Your Domain Context</h3>
          <p>
            Use the forest/domain selector in the navigation bar to set your target domain. This
            pre-fills domain fields across all report pages. For forest-wide reports, no domain
            selection is needed.
          </p>
        </div>
      </div>

      <div className="gs-section">
        <div className="gs-step-number">03</div>
        <div className="gs-step-content">
          <h3>Run Reports</h3>
          <p>
            Navigate to any report page using the tabs. Forest-Wide reports run with a single click.
            Domain-Specific reports require a domain FQDN. Results appear in sortable, paginated tables
            with full export capability.
          </p>
        </div>
      </div>

      <div className="gs-section">
        <div className="gs-step-number">04</div>
        <div className="gs-step-content">
          <h3>Export & Analyze</h3>
          <p>
            Click <strong>Export All Results</strong> on any report to download as Excel, CSV, JSON,
            or PDF. Use favorites (star icon) to bookmark reports you run frequently — they appear on
            the dashboard for quick access.
          </p>
        </div>
      </div>

      <div className="gs-section">
        <div className="gs-step-number">05</div>
        <div className="gs-step-content">
          <h3>Security Assessment</h3>
          <p>
            Visit the Security page for a posture overview. The summary dashboard shows critical,
            high, and medium risk counts. Run individual reports to drill into AdminSDHolder, delegation,
            Kerberos, and privileged access findings.
          </p>
        </div>
      </div>

      <div className="gs-shortcuts card">
        <h3>Keyboard Shortcuts</h3>
        <div className="gs-shortcut-grid">
          <div className="gs-shortcut"><kbd>Ctrl</kbd> + <kbd>K</kbd><span>Open Search</span></div>
          <div className="gs-shortcut"><kbd>?</kbd><span>Shortcuts Help</span></div>
          <div className="gs-shortcut"><kbd>Esc</kbd><span>Close Modal</span></div>
          <div className="gs-shortcut"><kbd>Alt</kbd> + <kbd>H</kbd><span>Go to Dashboard</span></div>
        </div>
      </div>
    </div>
  );

  return (
    <div className="help-page">
      <h1>Help & Release Notes</h1>

      <div className="help-tabs">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            className={`help-tab ${activeTab === tab.id ? 'active' : ''}`}
            onClick={() => setActiveTab(tab.id)}
          >
            {tab.label}
          </button>
        ))}
      </div>

      <div className="help-content">
        {activeTab === 'release-notes' && renderReleaseNotes()}
        {activeTab === 'faq' && renderFaq()}
        {activeTab === 'getting-started' && renderGettingStarted()}
      </div>
    </div>
  );
};

export default Help;

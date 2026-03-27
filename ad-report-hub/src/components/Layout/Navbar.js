import React, { useContext, useState, useEffect, useRef } from 'react';
import { NavLink, useNavigate, useLocation } from 'react-router-dom';
import { ThemeContext } from '../../App';
import { useApp } from '../../context/AppContext';
import axios from 'axios';
import ForestDomainSelector from '../common/ForestDomainSelector';
import './Navbar.css';
// AD Report Hub logos - using direct paths for compatibility
const adReportHubLogo = '/ad-report-hub-logo.svg';
const adReportHubLogoDark = '/ad-report-hub-logo-dark.svg';
const adReportHubIcon = '/ad-report-hub-icon.svg';

// Custom colored SVG icons
const UserIcon = ({ color = '#F59E0B' }) => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill={color} xmlns="http://www.w3.org/2000/svg">
    <circle cx="12" cy="8" r="4"/>
    <path d="M12 14c-4.42 0-8 1.79-8 4v2h16v-2c0-2.21-3.58-4-8-4z"/>
  </svg>
);

const GroupIcon = ({ color = '#EA580C' }) => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill={color} xmlns="http://www.w3.org/2000/svg">
    <circle cx="9" cy="8" r="3.5"/>
    <circle cx="16" cy="9" r="2.5"/>
    <path d="M9 13c-3.31 0-6 1.34-6 3v2h12v-2c0-1.66-2.69-3-6-3z"/>
    <path d="M16 13.5c-1.18 0-2.24.32-3 .82.93.6 1.5 1.38 1.5 2.18v1.5h6v-1.5c0-1.38-2.01-3-4.5-3z"/>
  </svg>
);

const Navbar = () => {
  const { theme, toggleTheme } = useContext(ThemeContext);
  const { isLoading, recentActivity, addRecentActivity, requestNavigation, confirmAbort } = useApp();
  const navigate = useNavigate();
  const location = useLocation();
  
  const [userInfo, setUserInfo] = useState(null);
  const [showUserMenu, setShowUserMenu] = useState(false);
  const [showUserDetails, setShowUserDetails] = useState(false);
  const [showRecentActivity, setShowRecentActivity] = useState(false);
  const [tabAlerts, setTabAlerts] = useState({});
  const userMenuRef = useRef(null);
  const recentRef = useRef(null);

  useEffect(() => {
    const checkSecurityAlerts = async () => {
      try {
        const res = await axios.get('/api/security/summary');
        const data = res.data?.data;
        if (data) {
          const total = (data.criticalIssues || 0) + (data.highIssues || 0);
          if (total > 0) {
            setTabAlerts(prev => ({
              ...prev,
              '/security': { count: total, level: data.criticalIssues > 0 ? 'critical' : 'high' },
              '/compliance': data.disabledInGroupsCount > 0 ? { count: data.disabledInGroupsCount, level: 'medium' } : null
            }));
          }
        }
      } catch (e) { /* silent */ }
    };
    checkSecurityAlerts();
  }, []);

  const s = 16;
  const sw = "2";
  const lc = "round";
  const lj = "round";

  const tabIcons = {
    // Identity (filled style matching Users/Groups)
    '/users': <UserIcon />,
    '/groups': <GroupIcon />,
    '/service-accounts': <svg width={s} height={s} viewBox="0 0 24 24" fill="#6366f1" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="7" r="4"/><path d="M12 13c-4.42 0-8 1.79-8 4v2h16v-2c0-2.21-3.58-4-8-4z"/><path d="M19 3l-1.5 1.5M19 3l1.5 1.5M19 3v2.5" stroke="#6366f1" strokeWidth="2" strokeLinecap="round" fill="none"/></svg>,
    '/contacts': <svg width={s} height={s} viewBox="0 0 24 24" fill="#0ea5e9" xmlns="http://www.w3.org/2000/svg"><rect x="3" y="2" width="18" height="20" rx="2" fill="none" stroke="#0ea5e9" strokeWidth="2"/><circle cx="12" cy="10" r="3"/><path d="M7 18c0-2.21 2.24-4 5-4s5 1.79 5 4"/></svg>,

    // Infrastructure (greens / teals / earth tones)
    '/domain-controllers': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#dc2626" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><rect x="2" y="2" width="20" height="8" rx="2"/><rect x="2" y="14" width="20" height="8" rx="2"/><circle cx="6" cy="6" r="1" fill="#dc2626"/><circle cx="6" cy="18" r="1" fill="#dc2626"/></svg>,
    '/computers': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#0891b2" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>,
    '/sites-subnets': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#059669" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 014 10 15.3 15.3 0 01-4 10 15.3 15.3 0 01-4-10 15.3 15.3 0 014-10z"/></svg>,
    '/topology': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#2563eb" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><circle cx="12" cy="5" r="3"/><circle cx="5" cy="19" r="3"/><circle cx="19" cy="19" r="3"/><line x1="12" y1="8" x2="5" y2="16"/><line x1="12" y1="8" x2="19" y2="16"/></svg>,
    '/gpos': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#d97706" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>,
    '/printers': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="1.8" strokeLinecap={lc} strokeLinejoin={lj}><path d="M6 9V2h12v7"/><rect x="2" y="9" width="20" height="8" rx="2"/><rect x="6" y="14" width="12" height="8" rx="1"/><circle cx="18" cy="13" r="1" fill="#7c3aed"/></svg>,
    '/containers': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#ca8a04" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/></svg>,
    '/security': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#e11d48" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>,
    '/compliance': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>,

    // Quick Links (purple / neutral tones)
    '/search': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#8b5cf6" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>,
    '/activity-logs': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#f59e0b" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>,
    '/help': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#6366f1" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 015.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>,
    '/license': <svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth={sw} strokeLinecap={lc} strokeLinejoin={lj}><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>,
  };

  // Organized tabs - Identity reports
  const mainTabs = [
    { path: '/users', label: 'USERS' },
    { path: '/groups', label: 'GROUPS' },
    { path: '/service-accounts', label: 'SERVICE ACCOUNTS' },
    { path: '/contacts', label: 'CONTACTS' },
  ];

  // Infrastructure tabs - AD infrastructure reports
  const infraTabs = [
    { path: '/domain-controllers', label: 'DOMAIN CONTROLLERS' },
    { path: '/computers', label: 'COMPUTERS' },
    { path: '/sites-subnets', label: 'SITES & SUBNETS' },
    { path: '/topology', label: 'AD TOPOLOGY' },
    { path: '/gpos', label: 'GROUP POLICY' },
    { path: '/containers', label: 'CONTAINERS' },
  ];

  // Security & Governance tabs
  const securityTabs = [
    { path: '/security', label: 'SECURITY' },
    { path: '/compliance', label: 'COMPLIANCE' },
  ];

  // Quick links - utility/tool tabs
  const utilityTabs = [
    { path: '/search', label: 'SEARCH' },
    { path: '/activity-logs', label: 'LOGS' },
    { path: '/help', label: 'HELP' },
    { path: '/license', label: 'LICENSE' },
  ];

  useEffect(() => {
    const fetchUserInfo = async () => {
      try {
        const response = await axios.get('/api/system/current-user');
        const raw = response?.data?.data;
        const normalized = Array.isArray(raw) ? raw[0] : raw;
        // Ensure we store an object (some endpoints may still return arrays)
        if (normalized && typeof normalized === 'object') {
          setUserInfo(normalized);
        } else {
          setUserInfo({ 
            username: 'Unknown User',
            domain: 'Unknown Domain',
            displayName: 'User'
          });
        }
      } catch (err) {
        setUserInfo({ 
          username: 'Unknown User',
          domain: 'Unknown Domain',
          displayName: 'User'
        });
      }
    };
    fetchUserInfo();
  }, []);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (userMenuRef.current && !userMenuRef.current.contains(event.target)) {
        setShowUserMenu(false);
      }
      if (recentRef.current && !recentRef.current.contains(event.target)) {
        setShowRecentActivity(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Track navigation for recent activity
  useEffect(() => {
    const allTabs = [...mainTabs, ...infraTabs, ...securityTabs, ...utilityTabs];
    const currentTab = allTabs.find(t => t.path === location.pathname);
    if (currentTab) {
      addRecentActivity({
        name: currentTab.label,
        path: currentTab.path,
        icon: tabIcons[currentTab.path] || '📄'
      });
    }
  }, [location.pathname]);

  const handleNavClick = (e, path) => {
    if (isLoading) {
      e.preventDefault();
      const canNavigate = requestNavigation(path, () => {
        navigate(path);
      });
      if (!canNavigate) {
        return;
      }
    }
  };

  const handleRecentClick = (item) => {
    setShowRecentActivity(false);
    if (isLoading) {
      requestNavigation(item.path, () => {
        navigate(item.path);
      });
    } else {
      navigate(item.path);
    }
  };

  const formatTimeAgo = (timestamp) => {
    const now = new Date();
    const then = new Date(timestamp);
    const diffMs = now - then;
    const diffMins = Math.floor(diffMs / 60000);
    
    if (diffMins < 1) return 'Just now';
    if (diffMins < 60) return `${diffMins}m ago`;
    const diffHours = Math.floor(diffMins / 60);
    if (diffHours < 24) return `${diffHours}h ago`;
    return then.toLocaleDateString();
  };

  const handleCloseApp = async () => {
    if (window.confirm('Are you sure you want to close the application?')) {
      try {
        await axios.post('/api/system/shutdown');
        window.close();
        setTimeout(() => {
          document.body.innerHTML = '<div style="display:flex;justify-content:center;align-items:center;height:100vh;background:#0a1929;color:white;font-family:Arial,sans-serif;flex-direction:column;"><h1>Application Closed</h1><p>You can safely close this browser tab.</p></div>';
        }, 500);
      } catch (err) {
        window.close();
      }
    }
  };

  const handleSignOut = async () => {
    if (window.confirm('Are you sure you want to sign out?')) {
      try {
        await axios.post('/api/system/shutdown');
        window.close();
        setTimeout(() => {
          document.body.innerHTML = '<div style="display:flex;justify-content:center;align-items:center;height:100vh;background:#0a1929;color:white;font-family:Arial,sans-serif;flex-direction:column;"><h1>Signed Out</h1><p>You can safely close this browser tab.</p></div>';
        }, 500);
      } catch (err) {
        window.close();
      }
    }
  };

  const getInitials = (name) => {
    if (!name) return 'U';
    return name.split(' ').map(n => n[0]).join('').toUpperCase().substring(0, 2);
  };

  return (
    <>
      {/* AD Report Hub Brand Header */}
      <div className="brand-header">
        <img 
          src={theme === 'dark' ? adReportHubLogoDark : adReportHubLogo} 
          alt="AD Report Hub - Active Directory Reporting" 
          className="brand-header-logo" 
        />
      </div>
      
      <nav className="navbar">
        <div className="navbar-content">
          <div className="navbar-brand">
            <div className="brand-left">
              <img src={adReportHubIcon} alt="AD Report Hub" className="brand-icon-img" />
              <span className="brand-text">AD Report Hub</span>
              {isLoading && (
                <span className="loading-indicator">
                  <span className="loading-dot"></span>
                  <span className="loading-text">Loading...</span>
                </span>
              )}
            </div>
            <div className="navbar-actions">
              {/* Home Button */}
              <button 
                className="home-btn" 
                onClick={() => navigate('/dashboard')}
                title="Go to Home"
                aria-label="Go to Home Dashboard"
              >
                🏠 Home
              </button>

              {/* Recent Activity Button */}
              <div className="recent-activity-wrapper" ref={recentRef}>
                <button 
                  className="recent-btn" 
                  onClick={() => setShowRecentActivity(!showRecentActivity)}
                  title="Recent Activity"
                >
                  🕐 Recent
                </button>
                
                {showRecentActivity && (
                  <div className="recent-activity-panel">
                    <div className="recent-activity-header">
                      <h4>Recent Activity</h4>
                    </div>
                    <div className="recent-activity-list">
                      {recentActivity.length === 0 ? (
                        <div className="recent-activity-empty">
                          No recent activity
                        </div>
                      ) : (
                        recentActivity.map((item) => (
                          <div 
                            key={item.id} 
                            className="recent-activity-item"
                            onClick={() => handleRecentClick(item)}
                          >
                            <span className="recent-activity-icon">{item.icon}</span>
                            <div className="recent-activity-info">
                              <div className="recent-activity-name">{item.name}</div>
                              <div className="recent-activity-path">{item.path}</div>
                            </div>
                            <span className="recent-activity-time">
                              {formatTimeAgo(item.timestamp)}
                            </span>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                )}
              </div>

              <ForestDomainSelector />

              <button 
                className="theme-toggle" 
                onClick={toggleTheme} 
                title={`Switch to ${theme === 'light' ? 'dark' : 'light'} mode`}
                aria-label={`Switch to ${theme === 'light' ? 'dark' : 'light'} mode`}
              >
                {theme === 'light' ? '🌙' : '☀️'}
              </button>
              
              {/* User Profile */}
              <div className="user-profile" ref={userMenuRef}>
                <button 
                  className="user-profile-btn" 
                  onClick={() => setShowUserMenu(!showUserMenu)}
                  title={userInfo?.username || 'User'}
                >
                  <span className="user-avatar-small">{getInitials(userInfo?.displayName || userInfo?.username)}</span>
                  <span className="user-name">{userInfo?.displayName || userInfo?.username || 'User'}</span>
                  <span className="dropdown-arrow">▼</span>
                </button>
                
                {showUserMenu && (
                  <div className="user-dropdown">
                    <div className="user-dropdown-header">
                      <span className="user-avatar-large">{getInitials(userInfo?.displayName || userInfo?.username)}</span>
                      <div className="user-dropdown-info">
                        <span className="user-dropdown-name">{userInfo?.displayName || userInfo?.username}</span>
                        <span className="user-dropdown-domain">{userInfo?.domain}</span>
                      </div>
                    </div>
                    <div className="user-dropdown-divider"></div>
                    <button className="user-dropdown-item" onClick={() => { setShowUserDetails(true); setShowUserMenu(false); }}>
                      <span>👤</span> User Info
                    </button>
                    <button className="user-dropdown-item" onClick={handleSignOut}>
                      <span>🚪</span> Sign Out
                    </button>
                  </div>
                )}
              </div>

              <button className="close-app-btn" onClick={handleCloseApp} title="Close Application">
                ✕
              </button>
            </div>
          </div>

          {/* User Details Modal */}
          {showUserDetails && userInfo && (
            <div className="modal-overlay" onClick={() => setShowUserDetails(false)}>
              <div className="modal user-details-modal user-profile-modal" onClick={(e) => e.stopPropagation()}>
                <div className="modal-header">
                  <h3>User Profile</h3>
                  <button className="modal-close" onClick={() => setShowUserDetails(false)} aria-label="Close user details">×</button>
                </div>
                <div className="modal-body user-profile-body">
                  {/* Profile Header */}
                  <div className="profile-header">
                    <div className="profile-avatar">{getInitials(userInfo.displayName || userInfo.username)}</div>
                    <div className="profile-header-info">
                      <h2 className="profile-name">{userInfo.displayName || userInfo.username}</h2>
                      <p className="profile-title">{userInfo.title || 'No title'}</p>
                      <p className="profile-email">{userInfo.email || userInfo.upn || 'No email'}</p>
                    </div>
                    <div className="profile-status">
                      <span className={`status-badge ${userInfo.enabled ? 'enabled' : 'disabled'}`}>
                        {userInfo.enabled ? '● Active' : '○ Disabled'}
                      </span>
                      {userInfo.lockedOut && <span className="status-badge locked">🔒 Locked</span>}
                    </div>
                  </div>

                  {/* Profile Sections */}
                  <div className="profile-sections">
                    {/* Basic Info */}
                    <div className="profile-section">
                      <h4 className="section-title">👤 Basic Information</h4>
                      <div className="profile-grid">
                        <div className="profile-field">
                          <span className="field-label">Username</span>
                          <span className="field-value">{userInfo.username}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Domain</span>
                          <span className="field-value">{userInfo.domain}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">First Name</span>
                          <span className="field-value">{userInfo.firstName || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Last Name</span>
                          <span className="field-value">{userInfo.lastName || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Employee ID</span>
                          <span className="field-value">{userInfo.employeeId || 'N/A'}</span>
                        </div>
                        <div className="profile-field full-width">
                          <span className="field-label">Description</span>
                          <span className="field-value">{userInfo.description || 'N/A'}</span>
                        </div>
                      </div>
                    </div>

                    {/* Job Info */}
                    <div className="profile-section">
                      <h4 className="section-title">💼 Organization</h4>
                      <div className="profile-grid">
                        <div className="profile-field">
                          <span className="field-label">Title</span>
                          <span className="field-value">{userInfo.title || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Department</span>
                          <span className="field-value">{userInfo.department || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Company</span>
                          <span className="field-value">{userInfo.company || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Manager</span>
                          <span className="field-value">{userInfo.manager || 'N/A'}</span>
                        </div>
                      </div>
                    </div>

                    {/* Contact Info */}
                    <div className="profile-section">
                      <h4 className="section-title">📞 Contact</h4>
                      <div className="profile-grid">
                        <div className="profile-field">
                          <span className="field-label">Email</span>
                          <span className="field-value">{userInfo.email || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Phone</span>
                          <span className="field-value">{userInfo.phone || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Mobile</span>
                          <span className="field-value">{userInfo.mobile || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Office</span>
                          <span className="field-value">{userInfo.office || 'N/A'}</span>
                        </div>
                      </div>
                    </div>

                    {/* Location */}
                    {(userInfo.city || userInfo.country || userInfo.address) && (
                      <div className="profile-section">
                        <h4 className="section-title">📍 Location</h4>
                        <div className="profile-grid">
                          <div className="profile-field full-width">
                            <span className="field-label">Address</span>
                            <span className="field-value">{userInfo.address || 'N/A'}</span>
                          </div>
                          <div className="profile-field">
                            <span className="field-label">City</span>
                            <span className="field-value">{userInfo.city || 'N/A'}</span>
                          </div>
                          <div className="profile-field">
                            <span className="field-label">State</span>
                            <span className="field-value">{userInfo.state || 'N/A'}</span>
                          </div>
                          <div className="profile-field">
                            <span className="field-label">Postal Code</span>
                            <span className="field-value">{userInfo.postalCode || 'N/A'}</span>
                          </div>
                          <div className="profile-field">
                            <span className="field-label">Country</span>
                            <span className="field-value">{userInfo.country || 'N/A'}</span>
                          </div>
                        </div>
                      </div>
                    )}

                    {/* Account Info */}
                    <div className="profile-section">
                      <h4 className="section-title">🔐 Account Details</h4>
                      <div className="profile-grid">
                        <div className="profile-field">
                          <span className="field-label">Domain</span>
                          <span className="field-value">{userInfo.domainDns || userInfo.domain || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Logon Server (DC)</span>
                          <span className="field-value">{userInfo.logonServer || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Group Memberships</span>
                          <span className="field-value">{userInfo.groupCount || 0} groups</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Account Created</span>
                          <span className="field-value">{userInfo.created || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Last Logon</span>
                          <span className="field-value">{userInfo.lastLogon || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Password Last Set</span>
                          <span className="field-value">{userInfo.passwordLastSet || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Account Expires</span>
                          <span className="field-value">{userInfo.accountExpires || 'Never'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">UPN</span>
                          <span className="field-value">{userInfo.upn || 'N/A'}</span>
                        </div>
                        <div className="profile-field full-width">
                          <span className="field-label">Distinguished Name</span>
                          <span className="field-value dn-value">{userInfo.distinguishedName || 'N/A'}</span>
                        </div>
                      </div>
                    </div>

                    {/* Computer Info */}
                    <div className="profile-section">
                      <h4 className="section-title">💻 Computer Details</h4>
                      <div className="profile-grid">
                        <div className="profile-field">
                          <span className="field-label">Computer Name</span>
                          <span className="field-value">{userInfo.computer || 'N/A'}</span>
                        </div>
                        <div className="profile-field">
                          <span className="field-label">Operating System</span>
                          <span className="field-value">{userInfo.computerOS || 'N/A'}</span>
                        </div>
                        <div className="profile-field full-width">
                          <span className="field-label">Computer DN</span>
                          <span className="field-value dn-value">{userInfo.computerDN || 'N/A'}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Navigation Tabs */}
          <div className="navbar-tabs">
            <div className="tab-group group-identity">
              <span className="tab-group-label">IDENTITY</span>
              <div className="tab-group-items">
                {mainTabs.map((tab) => (
                  <NavLink
                    key={tab.path}
                    to={tab.path}
                    className={({ isActive }) => `nav-tab ${isActive ? 'active' : ''}`}
                    onClick={(e) => handleNavClick(e, tab.path)}
                  >
                    <span className="tab-icon">{tabIcons[tab.path]}</span>
                    {tab.label}
                  </NavLink>
                ))}
              </div>
            </div>
            
            <div className="tab-divider"></div>
            
            <div className="tab-group group-infra">
              <span className="tab-group-label">INFRASTRUCTURE</span>
              <div className="tab-group-items">
                {infraTabs.map((tab) => (
                  <NavLink
                    key={tab.path}
                    to={tab.path}
                    className={({ isActive }) => `nav-tab ${isActive ? 'active' : ''}`}
                    onClick={(e) => handleNavClick(e, tab.path)}
                  >
                    <span className="tab-icon">{tabIcons[tab.path]}</span>
                    {tab.label}
                  </NavLink>
                ))}
              </div>
            </div>
            
            <div className="tab-divider"></div>
            
            <div className="tab-group group-security">
              <span className="tab-group-label">SECURITY & GOVERNANCE</span>
              <div className="tab-group-items">
                {securityTabs.map((tab) => (
                  <NavLink
                    key={tab.path}
                    to={tab.path}
                    className={({ isActive }) => `nav-tab ${isActive ? 'active' : ''}`}
                    onClick={(e) => handleNavClick(e, tab.path)}
                  >
                    <span className="tab-icon">{tabIcons[tab.path]}</span>
                    {tab.label}
                    {tabAlerts[tab.path] && (
                      <span className={`tab-alert-dot alert-${tabAlerts[tab.path].level}`} title={`${tabAlerts[tab.path].count} issue(s) detected`}></span>
                    )}
                  </NavLink>
                ))}
              </div>
            </div>
            
            <div className="tab-divider"></div>
            
            <div className="tab-group group-utility">
              <span className="tab-group-label">QUICK LINKS</span>
              <div className="tab-group-items">
                {utilityTabs.map((tab) => (
                  <NavLink
                    key={tab.path}
                    to={tab.path}
                    className={({ isActive }) => `nav-tab secondary ${isActive ? 'active' : ''}`}
                    onClick={(e) => handleNavClick(e, tab.path)}
                  >
                    <span className="tab-icon">{tabIcons[tab.path]}</span>
                    {tab.label}
                  </NavLink>
                ))}
              </div>
            </div>
          </div>
        </div>
      </nav>
    </>
  );
};

export default Navbar;

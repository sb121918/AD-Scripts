import React, { useState, useEffect, createContext } from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import './App.css';

// Context
import { AppProvider, useApp } from './context/AppContext';

// Layout Components
import Navbar from './components/Layout/Navbar';
import Footer from './components/Layout/Footer';

// Page Components
import Dashboard from './components/Dashboard/Dashboard';
import Computers from './components/Computers/Computers';
import Contacts from './components/Contacts/Contacts';
import Containers from './components/Containers/Containers';
import Groups from './components/Groups/Groups';
import GPOs from './components/GPOs/GPOs';
import Printers from './components/Printers/Printers';
import Users from './components/Users/Users';
import ActivityLogs from './components/ActivityLogs/ActivityLogs';
import License from './components/License/License';

// New Tabs
import DomainControllers from './components/DomainControllers/DomainControllers';
import ADSitesSubnets from './components/ADSitesSubnets/ADSitesSubnets';
import ServiceAccounts from './components/ServiceAccounts/ServiceAccounts';
import Topology from './components/Topology/Topology';
import Search from './components/Search/Search';
import Compliance from './components/Compliance/Compliance';
import Security from './components/Security/Security';
import Help from './components/Help/Help';

// Hooks & Common Components
import useKeyboardShortcuts from './hooks/useKeyboardShortcuts';
import ShortcutsHelp from './components/common/ShortcutsHelp';
import Breadcrumb from './components/common/Breadcrumb';

// Theme Context
export const ThemeContext = createContext();

// Abort Confirmation Modal Component
const AbortConfirmModal = () => {
  const { showAbortConfirm, currentReport, confirmAbort, cancelAbort } = useApp();

  if (!showAbortConfirm) return null;

  return (
    <div className="modal-overlay abort-modal-overlay">
      <div className="modal abort-modal">
        <div className="modal-header">
          <h3>Report In Progress</h3>
        </div>
        <div className="modal-body">
          <div className="abort-icon">⚠️</div>
          <p className="abort-message">
            A report is currently being generated:
          </p>
          <p className="abort-report-name">
            <strong>{currentReport?.name || 'Report'}</strong>
          </p>
          <p className="abort-question">
            Do you want to abort the current report and navigate away?
          </p>
        </div>
        <div className="modal-footer abort-footer">
          <button className="btn btn-secondary" onClick={cancelAbort}>
            Continue Report
          </button>
          <button className="btn btn-danger" onClick={confirmAbort}>
            Abort & Navigate
          </button>
        </div>
      </div>
    </div>
  );
};

function AppContent() {
  const [theme, setTheme] = useState(() => {
    const savedTheme = localStorage.getItem('ad-report-hub-theme');
    return savedTheme || 'dark';
  });
  const [showShortcutsHelp, setShowShortcutsHelp] = useState(false);

  useKeyboardShortcuts({ onShowHelp: setShowShortcutsHelp });

  useEffect(() => {
    localStorage.setItem('ad-report-hub-theme', theme);
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prevTheme => prevTheme === 'light' ? 'dark' : 'light');
  };

  return (
    <ThemeContext.Provider value={{ theme, toggleTheme }}>
      <div className={`app ${theme}`}>
        <Navbar />
        <main className="main-content">
          <Breadcrumb />
          <Routes>
            <Route path="/" element={<Navigate to="/dashboard" replace />} />
            <Route path="/dashboard" element={<Dashboard />} />
            <Route path="/computers" element={<Computers />} />
            <Route path="/contacts" element={<Contacts />} />
            <Route path="/containers" element={<Containers />} />
            <Route path="/groups" element={<Groups />} />
            <Route path="/gpos" element={<GPOs />} />
            <Route path="/printers" element={<Printers />} />
            <Route path="/users" element={<Users />} />
            <Route path="/domain-controllers" element={<DomainControllers />} />
            <Route path="/sites-subnets" element={<ADSitesSubnets />} />
            <Route path="/topology" element={<Topology />} />
            <Route path="/service-accounts" element={<ServiceAccounts />} />
            <Route path="/search" element={<Search />} />
            <Route path="/compliance" element={<Compliance />} />
            <Route path="/security" element={<Security />} />
            <Route path="/activity-logs" element={<ActivityLogs />} />
            <Route path="/license" element={<License />} />
            <Route path="/help" element={<Help />} />
          </Routes>
        </main>
        <Footer />
        <AbortConfirmModal />
        <ShortcutsHelp isOpen={showShortcutsHelp} onClose={() => setShowShortcutsHelp(false)} />
      </div>
    </ThemeContext.Provider>
  );
}

function App() {
  return (
    <AppProvider>
      <AppContent />
    </AppProvider>
  );
}

export default App;

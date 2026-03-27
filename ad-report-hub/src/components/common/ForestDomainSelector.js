import React, { useState, useRef, useEffect, useMemo } from 'react';
import { useApp } from '../../context/AppContext';

const ForestDomainSelector = () => {
  const { forestData, forestLoading, selectedDomain, changeSelectedDomain } = useApp();
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Group domains by forest
  const domainsByForest = useMemo(() => {
    if (!forestData?.domains) return {};
    const groups = {};
    for (const d of forestData.domains) {
      if (!groups[d.forest]) groups[d.forest] = [];
      groups[d.forest].push(d);
    }
    return groups;
  }, [forestData]);

  const handleSelect = (domainName) => {
    changeSelectedDomain(domainName);
    setIsOpen(false);
  };

  const getDisplayLabel = () => {
    if (!selectedDomain) return 'Select Domain';
    const found = forestData?.domains?.find(d => d.name === selectedDomain);
    return found ? found.label : selectedDomain;
  };

  // Check if selected domain requires elevated access
  const selectedRequiresElevated = useMemo(() => {
    if (!selectedDomain || !forestData?.domains) return false;
    const found = forestData.domains.find(d => d.name === selectedDomain);
    return found?.requiresElevated || false;
  }, [selectedDomain, forestData]);

  if (forestLoading) {
    return (
      <div style={selectorStyle}>
        <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>Loading domains...</span>
      </div>
    );
  }

  return (
    <div ref={dropdownRef} style={selectorStyle}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        style={{
          ...buttonStyle,
          borderColor: selectedRequiresElevated ? '#f59e0b' : 'var(--border-color)',
        }}
        title="Select target domain for reports"
      >
        <span style={{ fontSize: '14px' }}>{'\uD83C\uDF10'}</span>
        <span style={{ fontSize: '13px', fontWeight: 500 }}>{getDisplayLabel()}</span>
        {selectedRequiresElevated && <span title="Requires A/DA account" style={{ fontSize: '13px' }}>{'\u26A0\uFE0F'}</span>}
        <span style={{ fontSize: '10px', marginLeft: '2px' }}>{isOpen ? '\u25B2' : '\u25BC'}</span>
      </button>

      {isOpen && (
        <div style={dropdownStyle}>
          {/* Domains grouped by forest */}
          {Object.entries(domainsByForest).map(([forestName, domains]) => (
            <div key={forestName}>
              <div style={forestHeaderStyle}>
                <span>{'\uD83C\uDF32'}</span>
                <span>{forestName} Forest</span>
              </div>
              {domains.map((domain) => (
                <div
                  key={domain.name}
                  onClick={() => handleSelect(domain.name)}
                  style={{
                    ...itemStyle,
                    paddingLeft: '28px',
                    fontWeight: selectedDomain === domain.name ? 600 : 400,
                    backgroundColor: selectedDomain === domain.name ? 'var(--accent-light)' : 'transparent',
                  }}
                >
                  <span style={{ fontSize: '13px', flexShrink: 0 }}>
                    {domain.requiresElevated ? '\uD83D\uDD12' : '\uD83D\uDCE6'}
                  </span>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {domain.label}
                      </span>
                    </div>
                    <div style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                      {domain.description}
                    </div>
                  </div>
                  {domain.requiresElevated && (
                    <span
                      title="Requires A/DA account"
                      style={{
                        fontSize: '10px',
                        padding: '2px 6px',
                        borderRadius: '3px',
                        backgroundColor: '#fef3c7',
                        color: '#92400e',
                        fontWeight: 600,
                        flexShrink: 0,
                        whiteSpace: 'nowrap',
                      }}
                    >
                      {'\u26A0'} Elevated
                    </span>
                  )}
                  {selectedDomain === domain.name && <span style={checkStyle}>{'\u2713'}</span>}
                </div>
              ))}
            </div>
          ))}

          {!forestData?.domains?.length && (
            <div style={{ padding: '12px 16px', color: 'var(--text-muted)', fontSize: '13px', textAlign: 'center' }}>
              No domains configured. Edit server/config/domains.json.
            </div>
          )}
        </div>
      )}
    </div>
  );
};

const selectorStyle = {
  position: 'relative',
  display: 'inline-flex',
};

const buttonStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: '6px',
  padding: '6px 12px',
  border: '1px solid var(--border-color)',
  borderRadius: '6px',
  backgroundColor: 'var(--bg-secondary)',
  color: 'var(--text-primary)',
  cursor: 'pointer',
  transition: 'all 0.15s ease',
  whiteSpace: 'nowrap',
};

const dropdownStyle = {
  position: 'absolute',
  top: '100%',
  right: 0,
  marginTop: '4px',
  width: '360px',
  maxHeight: '420px',
  overflowY: 'auto',
  backgroundColor: 'var(--bg-primary)',
  border: '1px solid var(--border-color)',
  borderRadius: '8px',
  boxShadow: '0 8px 24px rgba(0,0,0,0.15)',
  zIndex: 1000,
};

const forestHeaderStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: '6px',
  padding: '8px 16px 4px 16px',
  fontSize: '11px',
  fontWeight: 700,
  color: 'var(--text-secondary)',
  textTransform: 'uppercase',
  letterSpacing: '0.5px',
};

const itemStyle = {
  display: 'flex',
  alignItems: 'center',
  gap: '8px',
  padding: '8px 16px',
  cursor: 'pointer',
  transition: 'background-color 0.1s ease',
  color: 'var(--text-primary)',
};

const dividerStyle = {
  height: '1px',
  backgroundColor: 'var(--border-color)',
  margin: '4px 0',
};

const checkStyle = {
  color: 'var(--accent-primary)',
  fontWeight: 700,
  fontSize: '14px',
  flexShrink: 0,
};

export default ForestDomainSelector;

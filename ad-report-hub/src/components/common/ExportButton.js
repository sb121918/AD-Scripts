import React, { useState, useRef, useEffect } from 'react';
import { handleExportWithFormat } from '../../utils/exportUtils';

const ExportButton = ({ data, filename, title }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const exportFormats = [
    { id: 'excel', label: 'Excel (.xlsx)', icon: '📊' },
    { id: 'csv', label: 'CSV (.csv)', icon: '📄' },
    { id: 'json', label: 'JSON (.json)', icon: '{ }' },
    { id: 'pdf', label: 'PDF (Print)', icon: '📑' },
  ];

  const handleExport = (format) => {
    handleExportWithFormat(data, filename, format, title);
    setIsOpen(false);
  };

  if (!data || (Array.isArray(data) && data.length === 0)) {
    return null;
  }

  return (
    <div ref={dropdownRef} style={{ position: 'relative', display: 'inline-block' }}>
      <button
        className="btn btn-secondary"
        onClick={() => setIsOpen(!isOpen)}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: '8px',
          padding: '8px 16px',
        }}
      >
        <span>Export All Results</span>
        <span style={{ fontSize: '10px' }}>{isOpen ? '▲' : '▼'}</span>
      </button>

      {isOpen && (
        <div
          style={{
            position: 'absolute',
            top: '100%',
            right: 0,
            marginTop: '4px',
            backgroundColor: 'var(--bg-secondary, #1e293b)',
            border: '1px solid var(--border-color, #334155)',
            borderRadius: '8px',
            boxShadow: '0 4px 12px rgba(0, 0, 0, 0.3)',
            zIndex: 1000,
            minWidth: '180px',
            overflow: 'hidden',
          }}
        >
          <div style={{ padding: '8px 12px', borderBottom: '1px solid var(--border-color, #334155)', fontSize: '12px', color: 'var(--text-secondary, #94a3b8)' }}>
            Export Format
          </div>
          {exportFormats.map((format) => (
            <button
              key={format.id}
              onClick={() => handleExport(format.id)}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '10px',
                width: '100%',
                padding: '10px 12px',
                border: 'none',
                backgroundColor: 'transparent',
                color: 'var(--text-primary, #e2e8f0)',
                cursor: 'pointer',
                textAlign: 'left',
                fontSize: '14px',
                transition: 'background-color 0.2s',
              }}
              onMouseEnter={(e) => e.target.style.backgroundColor = 'var(--accent-light, #1e3a5f)'}
              onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
            >
              <span style={{ width: '24px', textAlign: 'center' }}>{format.icon}</span>
              <span>{format.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

export default ExportButton;

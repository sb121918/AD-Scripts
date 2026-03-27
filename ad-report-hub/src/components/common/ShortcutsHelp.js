import React from 'react';

const shortcuts = [
  { keys: ['Ctrl', 'K'], description: 'Open Global Search' },
  { keys: ['Escape'], description: 'Close any open modal or dialog' },
  { keys: ['Alt', 'H'], description: 'Navigate to Home / Dashboard' },
  { keys: ['?'], description: 'Show this shortcuts help' },
];

const ShortcutsHelp = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()} style={{ maxWidth: '480px' }}>
        <div className="modal-header">
          <h3>Keyboard Shortcuts</h3>
          <button className="modal-close" onClick={onClose} aria-label="Close shortcuts help">&times;</button>
        </div>
        <div className="modal-body" style={{ padding: '20px' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              {shortcuts.map((shortcut, i) => (
                <tr key={i} style={{ borderBottom: '1px solid var(--border-color)' }}>
                  <td style={{ padding: '12px 8px', width: '160px' }}>
                    <div style={{ display: 'flex', gap: '4px' }}>
                      {shortcut.keys.map((key, j) => (
                        <React.Fragment key={j}>
                          {j > 0 && <span style={{ color: 'var(--text-muted)', alignSelf: 'center' }}>+</span>}
                          <kbd style={{
                            display: 'inline-block',
                            padding: '4px 8px',
                            fontSize: '12px',
                            fontFamily: 'monospace',
                            fontWeight: 600,
                            color: 'var(--text-primary)',
                            backgroundColor: 'var(--bg-secondary)',
                            border: '1px solid var(--border-color)',
                            borderRadius: '4px',
                            boxShadow: '0 1px 2px rgba(0,0,0,0.1)',
                            minWidth: '24px',
                            textAlign: 'center',
                          }}>
                            {key}
                          </kbd>
                        </React.Fragment>
                      ))}
                    </div>
                  </td>
                  <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontSize: '14px' }}>
                    {shortcut.description}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p style={{ marginTop: '16px', color: 'var(--text-muted)', fontSize: '12px', textAlign: 'center' }}>
            Press <kbd style={{ padding: '2px 6px', fontSize: '11px', border: '1px solid var(--border-color)', borderRadius: '3px', backgroundColor: 'var(--bg-secondary)' }}>Escape</kbd> to close
          </p>
        </div>
      </div>
    </div>
  );
};

export default ShortcutsHelp;

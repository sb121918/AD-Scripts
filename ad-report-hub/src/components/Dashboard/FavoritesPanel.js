import React from 'react';
import { useNavigate } from 'react-router-dom';
import { useApp } from '../../context/AppContext';

const pageIcons = {
  '/users': '\uD83D\uDC64',
  '/groups': '\uD83D\uDC65',
  '/computers': '\uD83D\uDCBB',
  '/domain-controllers': '\uD83D\uDDA5\uFE0F',
  '/sites-subnets': '\uD83C\uDF10',
  '/topology': '\uD83D\uDD17',
  '/service-accounts': '\u2699\uFE0F',
  '/containers': '\uD83D\uDCC1',
  '/compliance': '\uD83D\uDEE1\uFE0F',
  '/gpos': '\uD83D\uDCCB',
  '/printers': '\uD83D\uDDA8\uFE0F',
  '/contacts': '\uD83D\uDCC7',
};

const FavoritesPanel = () => {
  const { favorites, removeFavorite } = useApp();
  const navigate = useNavigate();

  if (favorites.length === 0) return null;

  const handleClick = (fav) => {
    navigate(fav.page);
  };

  return (
    <div className="favorites-panel" style={{ marginBottom: '32px' }}>
      <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <span>\u2B50</span> Favorites
      </h2>
      <div className="actions-grid">
        {favorites.map((fav) => (
          <div
            key={fav.id}
            className="action-card"
            style={{
              cursor: 'pointer',
              position: 'relative',
              borderLeftColor: 'var(--accent-primary)',
            }}
            onClick={() => handleClick(fav)}
          >
            <span className="action-icon">{pageIcons[fav.page] || '\uD83D\uDCC4'}</span>
            <span className="action-label">{fav.label}</span>
            <button
              onClick={(e) => {
                e.stopPropagation();
                removeFavorite(fav.page, fav.queryId);
              }}
              title="Remove from favorites"
              style={{
                position: 'absolute',
                top: '8px',
                right: '8px',
                background: 'none',
                border: 'none',
                cursor: 'pointer',
                fontSize: '14px',
                opacity: 0.5,
                color: 'var(--text-secondary)',
                padding: '2px 4px',
                lineHeight: 1,
              }}
              onMouseEnter={(e) => { e.target.style.opacity = '1'; }}
              onMouseLeave={(e) => { e.target.style.opacity = '0.5'; }}
            >
              &times;
            </button>
          </div>
        ))}
      </div>
    </div>
  );
};

export default FavoritesPanel;

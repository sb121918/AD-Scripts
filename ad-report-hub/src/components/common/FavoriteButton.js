import React from 'react';
import { useApp } from '../../context/AppContext';

/**
 * Star icon toggle button for bookmarking queries
 * @param {string} page - The page path (e.g., '/users', '/topology')
 * @param {string} queryId - The query identifier (e.g., 'all-users', 'forest-info')
 * @param {string} label - Human-readable label for the favorite
 * @param {Object} params - Optional additional params to store with the favorite
 */
const FavoriteButton = ({ page, queryId, label, params = {} }) => {
  const { isFavorite, addFavorite, removeFavorite } = useApp();
  const favorited = isFavorite(page, queryId);

  const handleToggle = (e) => {
    e.stopPropagation(); // Prevent triggering parent click handlers (query cards)
    if (favorited) {
      removeFavorite(page, queryId);
    } else {
      addFavorite({ page, queryId, label, params });
    }
  };

  return (
    <button
      className={`favorite-btn ${favorited ? 'active' : ''}`}
      onClick={handleToggle}
      title={favorited ? 'Remove from favorites' : 'Add to favorites'}
      style={{
        background: 'none',
        border: 'none',
        cursor: 'pointer',
        fontSize: '16px',
        padding: '4px',
        lineHeight: 1,
        color: favorited ? '#facc15' : 'var(--text-secondary)',
        opacity: favorited ? 1 : 0.4,
        transition: 'opacity 0.2s ease, transform 0.2s ease, color 0.2s ease',
        transform: favorited ? 'scale(1.1)' : 'scale(1)',
        filter: favorited ? 'drop-shadow(0 0 2px #facc15)' : 'none',
      }}
      onMouseEnter={(e) => { e.target.style.opacity = '1'; e.target.style.transform = 'scale(1.2)'; e.target.style.color = '#facc15'; }}
      onMouseLeave={(e) => { e.target.style.opacity = favorited ? '1' : '0.4'; e.target.style.transform = favorited ? 'scale(1.1)' : 'scale(1)'; e.target.style.color = favorited ? '#facc15' : 'var(--text-secondary)'; }}
    >
      {favorited ? '\u2605' : '\u2606'}
    </button>
  );
};

export default FavoriteButton;

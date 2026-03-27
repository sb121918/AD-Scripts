import React from 'react';
import './SearchHistory.css';

const MAX_HISTORY = 10;

const SearchHistory = ({ history, onSelect }) => {
  if (!history || history.length === 0) return null;

  return (
    <div className="search-history">
      <div className="search-history-header">
        <h4>Recent Searches</h4>
      </div>
      <ul className="search-history-list">
        {history.slice(0, MAX_HISTORY).map((item, index) => (
          <li key={index}>
            <button
              type="button"
              className="search-history-item"
              onClick={() => onSelect(item)}
            >
              {item.mode === 'ldap' ? (
                <span className="history-filter">{item.ldapFilter || 'LDAP query'}</span>
              ) : (
                <span className="history-query">{item.query}</span>
              )}
              {item.types && (
                <span className="history-types">[{item.types}]</span>
              )}
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
};

export default SearchHistory;

export const addToHistory = (history, item) => {
  const next = [item, ...history.filter((h) => h.ldapFilter !== item.ldapFilter || h.query !== item.query)];
  return next.slice(0, MAX_HISTORY);
};

import React, { useState, useEffect } from 'react';
import axios from 'axios';
import './SavedSearches.css';

const SavedSearches = ({ onLoadSearch, saveModal = { open: false, data: null }, onSaveComplete }) => {
  const [savedSearches, setSavedSearches] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saveName, setSaveName] = useState('');
  const showSaveModal = saveModal?.open || false;
  const currentSearchToSave = saveModal?.data || null;

  useEffect(() => {
    fetchSavedSearches();
  }, []);

  const fetchSavedSearches = async () => {
    try {
      const res = await axios.get('/api/search/saved');
      setSavedSearches(res.data.data || []);
    } catch (err) {
      setSavedSearches([]);
    } finally {
      setLoading(false);
    }
  };

  const confirmSave = async () => {
    if (!saveName.trim() || !currentSearchToSave) return;
    try {
      await axios.post('/api/search/saved', {
        name: saveName.trim(),
        ...currentSearchToSave,
      });
      setSaveName('');
      onSaveComplete?.();
      fetchSavedSearches();
    } catch (err) {
      console.error('Save failed:', err);
    }
  };

  const handleLoad = (search) => {
    onLoadSearch(search);
  };

  const handleDelete = async (id) => {
    if (!window.confirm('Delete this saved search?')) return;
    try {
      await axios.delete(`/api/search/saved/${id}`);
      fetchSavedSearches();
    } catch (err) {
      console.error('Delete failed:', err);
    }
  };

  if (loading) return null;

  return (
    <div className="saved-searches">
      <div className="saved-searches-header">
        <h4>Saved Searches</h4>
      </div>
      {savedSearches.length === 0 ? (
        <div className="saved-searches-empty">No saved searches yet</div>
      ) : (
        <ul className="saved-searches-list">
          {savedSearches.map((s) => (
            <li key={s.id} className="saved-search-item">
              <button
                type="button"
                className="saved-search-load"
                onClick={() => handleLoad(s)}
              >
                {s.name}
              </button>
              <button
                type="button"
                className="saved-search-delete"
                onClick={() => handleDelete(s.id)}
                title="Delete"
              >
                ×
              </button>
            </li>
          ))}
        </ul>
      )}

      {showSaveModal && (
        <div className="modal-overlay" onClick={() => onSaveComplete?.()}>
          <div className="modal saved-search-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Save Search</h3>
              <button className="modal-close" onClick={() => onSaveComplete?.()} aria-label="Close save search modal">×</button>
            </div>
            <div className="modal-body">
              <label>
                Save as:
                <input
                  type="text"
                  value={saveName}
                  onChange={(e) => setSaveName(e.target.value)}
                  placeholder="Search name..."
                  autoFocus
                  onKeyPress={(e) => e.key === 'Enter' && confirmSave()}
                />
              </label>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => onSaveComplete?.()}>
                Cancel
              </button>
              <button
                className="btn btn-primary"
                onClick={confirmSave}
                disabled={!saveName.trim()}
              >
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default SavedSearches;

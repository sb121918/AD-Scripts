import { useEffect, useCallback } from 'react';
import { useNavigate } from 'react-router-dom';

/**
 * Global keyboard shortcuts hook
 * @param {Object} options
 * @param {Function} options.onShowHelp - Callback to show shortcuts help modal
 */
const useKeyboardShortcuts = ({ onShowHelp } = {}) => {
  const navigate = useNavigate();

  const handleKeyDown = useCallback((e) => {
    const target = e.target;
    const isInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT' || target.isContentEditable;

    // Ctrl+K / Cmd+K - Navigate to search
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      navigate('/search');
      // Focus the search input after navigation
      setTimeout(() => {
        const searchInput = document.querySelector('.search-page input[type="text"]');
        if (searchInput) searchInput.focus();
      }, 100);
      return;
    }

    // Escape - Close any open modal
    if (e.key === 'Escape') {
      const modalOverlay = document.querySelector('.modal-overlay');
      if (modalOverlay) {
        const closeBtn = modalOverlay.querySelector('.modal-close');
        if (closeBtn) {
          closeBtn.click();
          return;
        }
        // Fallback: click the overlay itself
        modalOverlay.click();
        return;
      }
      // Close shortcuts help if open
      if (onShowHelp) {
        onShowHelp(false);
      }
      return;
    }

    // Skip remaining shortcuts if in an input field
    if (isInput) return;

    // Alt+H - Navigate to home/dashboard
    if (e.altKey && e.key === 'h') {
      e.preventDefault();
      navigate('/dashboard');
      return;
    }

    // ? - Show shortcuts help
    if (e.key === '?' && onShowHelp) {
      e.preventDefault();
      onShowHelp(true);
      return;
    }
  }, [navigate, onShowHelp]);

  useEffect(() => {
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [handleKeyDown]);
};

export default useKeyboardShortcuts;

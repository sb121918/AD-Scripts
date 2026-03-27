/**
 * Safe date formatting utilities for AD Identity Tool
 */

/**
 * Safely format a date value to a localized date string
 * Handles various date formats including ISO 8601 strings from PowerShell
 * @param {string|Date|null} dateValue - The date value to format
 * @param {string} fallback - The fallback string if date is invalid (default: '-')
 * @returns {string} - Formatted date string or fallback
 */
export const formatDate = (dateValue, fallback = '-') => {
  if (!dateValue) return fallback;
  
  try {
    // Handle PowerShell /Date()/ format
    if (typeof dateValue === 'string' && dateValue.includes('/Date(')) {
      const match = dateValue.match(/\/Date\((\d+)([+-]\d{4})?\)\//);
      if (match) {
        const timestamp = parseInt(match[1], 10);
        const date = new Date(timestamp);
        if (!isNaN(date.getTime())) {
          return date.toLocaleDateString();
        }
      }
    }
    
    // Handle standard date formats
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return fallback;
    return date.toLocaleDateString();
  } catch {
    return fallback;
  }
};

/**
 * Format a date value to include time
 * @param {string|Date|null} dateValue - The date value to format
 * @param {string} fallback - The fallback string if date is invalid (default: '-')
 * @returns {string} - Formatted date/time string or fallback
 */
export const formatDateTime = (dateValue, fallback = '-') => {
  if (!dateValue) return fallback;
  
  try {
    // Handle PowerShell /Date()/ format
    if (typeof dateValue === 'string' && dateValue.includes('/Date(')) {
      const match = dateValue.match(/\/Date\((\d+)([+-]\d{4})?\)\//);
      if (match) {
        const timestamp = parseInt(match[1], 10);
        const date = new Date(timestamp);
        if (!isNaN(date.getTime())) {
          return date.toLocaleString();
        }
      }
    }
    
    // Handle standard date formats
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) return fallback;
    return date.toLocaleString();
  } catch {
    return fallback;
  }
};

/**
 * Check if a date value represents "Never" (for LastLogon scenarios)
 * @param {string|Date|null} dateValue - The date value to check
 * @returns {boolean} - True if the date represents "never logged on"
 */
export const isNeverLoggedOn = (dateValue) => {
  if (!dateValue) return true;
  
  try {
    const date = new Date(dateValue);
    // Check for year 1601 (Windows FILETIME epoch) or invalid date
    return isNaN(date.getTime()) || date.getFullYear() <= 1601;
  } catch {
    return true;
  }
};

/**
 * Format LastLogon date with special handling for "Never" values
 * @param {string|Date|null} dateValue - The date value to format
 * @returns {string} - Formatted date string or "Never"
 */
export const formatLastLogon = (dateValue) => {
  if (isNeverLoggedOn(dateValue)) return 'Never';
  return formatDate(dateValue, 'Never');
};

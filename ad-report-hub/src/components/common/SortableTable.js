import React, { useState, useMemo } from 'react';
import './SortableTable.css';

/**
 * SortableTable component - Provides sortable headers for any table
 * 
 * Usage:
 * <SortableTable 
 *   data={arrayOfObjects}
 *   columns={[
 *     { key: 'name', label: 'Name' },
 *     { key: 'email', label: 'Email', render: (value, row) => <a href={`mailto:${value}`}>{value}</a> }
 *   ]}
 *   pageSize={25}
 *   currentPage={1}
 * />
 */

const SortableTable = ({ 
  data = [], 
  columns = [], 
  pageSize = 25, 
  currentPage = 1,
  onPageChange,
  onPageSizeChange,
  showPagination = true
}) => {
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  // Sort the data
  const sortedData = useMemo(() => {
    if (!data || !Array.isArray(data)) return [];
    
    let sortableItems = [...data];
    if (sortConfig.key !== null) {
      sortableItems.sort((a, b) => {
        let aValue = a[sortConfig.key];
        let bValue = b[sortConfig.key];
        
        // Handle null/undefined
        if (aValue === null || aValue === undefined) aValue = '';
        if (bValue === null || bValue === undefined) bValue = '';
        
        // Handle boolean
        if (typeof aValue === 'boolean') aValue = aValue ? 1 : 0;
        if (typeof bValue === 'boolean') bValue = bValue ? 1 : 0;
        
        // Handle numbers
        if (typeof aValue === 'number' && typeof bValue === 'number') {
          return sortConfig.direction === 'asc' ? aValue - bValue : bValue - aValue;
        }
        
        // Convert to string for comparison
        const aStr = String(aValue).toLowerCase();
        const bStr = String(bValue).toLowerCase();
        
        if (aStr < bStr) {
          return sortConfig.direction === 'asc' ? -1 : 1;
        }
        if (aStr > bStr) {
          return sortConfig.direction === 'asc' ? 1 : -1;
        }
        return 0;
      });
    }
    return sortableItems;
  }, [data, sortConfig]);

  // Pagination
  const totalPages = Math.ceil(sortedData.length / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedData = sortedData.slice(startIndex, endIndex);

  const handleSort = (key) => {
    let direction = 'asc';
    if (sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  const getSortIndicator = (key) => {
    if (sortConfig.key !== key) {
      return <span className="sort-indicator unsorted">↕</span>;
    }
    return (
      <span className="sort-indicator sorted">
        {sortConfig.direction === 'asc' ? '↑' : '↓'}
      </span>
    );
  };

  const renderCellValue = (column, row) => {
    const value = row[column.key];
    
    if (column.render) {
      return column.render(value, row);
    }
    
    if (value === null || value === undefined) {
      return '-';
    }
    
    if (typeof value === 'boolean') {
      return (
        <span className={`status-badge ${value ? 'enabled' : 'disabled'}`}>
          {value ? 'Yes' : 'No'}
        </span>
      );
    }
    
    return value;
  };

  return (
    <div className="sortable-table-wrapper">
      <div className="table-scroll">
        <table className="data-table sortable-table">
          <thead>
            <tr>
              {columns.map((column) => (
                <th
                  key={column.key}
                  className={`sortable-header ${sortConfig.key === column.key ? 'sorted' : ''}`}
                  onClick={() => handleSort(column.key)}
                  style={column.width ? { width: column.width } : {}}
                >
                  <div className="header-content">
                    <span className="header-label">{column.label}</span>
                    {getSortIndicator(column.key)}
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {paginatedData.length === 0 ? (
              <tr>
                <td colSpan={columns.length} className="no-data">
                  No data available
                </td>
              </tr>
            ) : (
              paginatedData.map((row, rowIndex) => (
                <tr key={row.id || rowIndex}>
                  {columns.map((column) => (
                    <td key={column.key}>
                      {renderCellValue(column, row)}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
      
      {showPagination && sortedData.length > 0 && (
        <div className="table-pagination">
          <div className="pagination-info">
            Showing {startIndex + 1}-{Math.min(endIndex, sortedData.length)} of {sortedData.length} results
          </div>
          <div className="pagination-controls">
            <div className="page-size-selector">
              <label>Per page:</label>
              <select 
                value={pageSize} 
                onChange={(e) => onPageSizeChange && onPageSizeChange(parseInt(e.target.value))}
              >
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>
            <div className="page-nav">
              <button 
                onClick={() => onPageChange && onPageChange(currentPage - 1)} 
                disabled={currentPage <= 1}
                className="btn btn-secondary"
              >
                Back
              </button>
              <span className="page-indicator">
                Page {currentPage} of {totalPages}
              </span>
              <button 
                onClick={() => onPageChange && onPageChange(currentPage + 1)} 
                disabled={currentPage >= totalPages}
                className="btn btn-secondary"
              >
                Next
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default SortableTable;

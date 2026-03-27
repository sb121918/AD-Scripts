import React, { useState, useCallback } from 'react';
import axios from 'axios';
import './OUTree.css';

/**
 * Recursive tree node component
 */
const TreeNode = ({ node, level = 0, selectedDN, onSelect }) => {
  const [expanded, setExpanded] = useState(level < 1); // Auto-expand first level
  const hasChildren = node.Children && node.Children.length > 0;
  const isSelected = selectedDN === node.DistinguishedName;

  const handleToggle = (e) => {
    e.stopPropagation();
    setExpanded(!expanded);
  };

  const handleSelect = () => {
    onSelect(node);
  };

  return (
    <div className="ou-tree-node">
      <div
        className={`ou-tree-node-header ${isSelected ? 'selected' : ''}`}
        onClick={handleSelect}
      >
        <span
          className={`ou-tree-toggle ${expanded ? 'expanded' : ''} ${!hasChildren ? 'empty' : ''}`}
          onClick={handleToggle}
        >
          &#9654;
        </span>
        <span className="ou-tree-icon">
          {level === 0 ? '\uD83C\uDF10' : '\uD83D\uDCC1'}
        </span>
        <span className="ou-tree-name" title={node.Name}>
          {node.Name}
        </span>
        {node.Protected && <span className="ou-tree-protected" title="Protected from accidental deletion">&#x1F512;</span>}
        <div className="ou-tree-badges">
          {node.UserCount > 0 && (
            <span className="ou-tree-badge users" title={`${node.UserCount} users`}>
              &#x1F464; {node.UserCount}
            </span>
          )}
          {node.GroupCount > 0 && (
            <span className="ou-tree-badge groups" title={`${node.GroupCount} groups`}>
              &#x1F465; {node.GroupCount}
            </span>
          )}
          {node.ComputerCount > 0 && (
            <span className="ou-tree-badge computers" title={`${node.ComputerCount} computers`}>
              &#x1F4BB; {node.ComputerCount}
            </span>
          )}
        </div>
      </div>

      {/* GPO links */}
      {node.GPOLinks && node.GPOLinks.length > 0 && expanded && (
        <div className="ou-tree-gpo-tags">
          {node.GPOLinks.map((gpo, i) => (
            <span key={i} className="ou-tree-gpo-tag">
              &#x1F4CB; {gpo}
            </span>
          ))}
        </div>
      )}

      {/* Children */}
      {hasChildren && expanded && (
        <div className="ou-tree-children">
          {node.Children.map((child, i) => (
            <TreeNode
              key={child.DistinguishedName || i}
              node={child}
              level={level + 1}
              selectedDN={selectedDN}
              onSelect={onSelect}
            />
          ))}
        </div>
      )}
    </div>
  );
};

/**
 * OU Tree component with details panel
 */
const OUTree = ({ treeData, domain }) => {
  const [selectedNode, setSelectedNode] = useState(null);
  const [childDetails, setChildDetails] = useState(null);
  const [detailsLoading, setDetailsLoading] = useState(false);

  const handleSelect = useCallback(async (node) => {
    setSelectedNode(node);
    setChildDetails(null);

    // Fetch child objects if this is an actual OU (not the domain root with no DN match)
    if (node.DistinguishedName && node.DistinguishedName.startsWith('OU=')) {
      setDetailsLoading(true);
      try {
        const response = await axios.get('/api/containers/children', {
          params: { dn: node.DistinguishedName, domain }
        });
        setChildDetails(response.data.data);
      } catch (err) {
        console.error('Failed to load OU children:', err);
      } finally {
        setDetailsLoading(false);
      }
    }
  }, []);

  if (!treeData) return null;

  return (
    <div className="ou-tree">
      <TreeNode
        node={treeData}
        level={0}
        selectedDN={selectedNode?.DistinguishedName}
        onSelect={handleSelect}
      />

      {/* Details panel for selected OU */}
      {selectedNode && selectedNode.DistinguishedName && (
        <div className="ou-tree-details">
          <h4>
            {'\uD83D\uDCC1'} {selectedNode.Name}
          </h4>
          <div className="ou-tree-details-dn">{selectedNode.DistinguishedName}</div>
          {selectedNode.Description && (
            <p style={{ color: 'var(--text-secondary)', fontSize: '13px', margin: '0 0 12px 0' }}>
              {selectedNode.Description}
            </p>
          )}

          {/* Object counts summary */}
          <div style={{ display: 'flex', gap: '16px', marginBottom: '12px', flexWrap: 'wrap' }}>
            <span className="ou-tree-badge users">&#x1F464; {selectedNode.UserCount || 0} Users</span>
            <span className="ou-tree-badge groups">&#x1F465; {selectedNode.GroupCount || 0} Groups</span>
            <span className="ou-tree-badge computers">&#x1F4BB; {selectedNode.ComputerCount || 0} Computers</span>
          </div>

          {/* GPO Links */}
          {selectedNode.GPOLinks && selectedNode.GPOLinks.length > 0 && (
            <div className="ou-tree-details-section">
              <h5>Linked GPOs ({selectedNode.GPOLinks.length})</h5>
              <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                {selectedNode.GPOLinks.map((gpo, i) => (
                  <span key={i} className="ou-tree-gpo-tag">&#x1F4CB; {gpo}</span>
                ))}
              </div>
            </div>
          )}

          {/* Child object details */}
          {detailsLoading && (
            <div className="ou-tree-loading">
              <div className="spinner" style={{ width: '16px', height: '16px' }}></div>
              Loading child objects...
            </div>
          )}

          {childDetails && (
            <>
              {childDetails.Users && childDetails.Users.length > 0 && (
                <div className="ou-tree-details-section">
                  <h5>Users ({childDetails.Users.length})</h5>
                  <div style={{ overflowX: 'auto' }}>
                    <table className="data-table" style={{ fontSize: '13px' }}>
                      <thead>
                        <tr>
                          <th>Name</th>
                          <th>SAM Account</th>
                          <th>Email</th>
                          <th>Enabled</th>
                        </tr>
                      </thead>
                      <tbody>
                        {childDetails.Users.map((u, i) => (
                          <tr key={i}>
                            <td>{u.Name}</td>
                            <td>{u.SamAccountName}</td>
                            <td>{u.Email || '-'}</td>
                            <td>{u.Enabled ? 'Yes' : 'No'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {childDetails.Groups && childDetails.Groups.length > 0 && (
                <div className="ou-tree-details-section">
                  <h5>Groups ({childDetails.Groups.length})</h5>
                  <div style={{ overflowX: 'auto' }}>
                    <table className="data-table" style={{ fontSize: '13px' }}>
                      <thead>
                        <tr>
                          <th>Name</th>
                          <th>Category</th>
                          <th>Scope</th>
                          <th>Description</th>
                        </tr>
                      </thead>
                      <tbody>
                        {childDetails.Groups.map((g, i) => (
                          <tr key={i}>
                            <td>{g.Name}</td>
                            <td>{g.Category}</td>
                            <td>{g.Scope}</td>
                            <td>{g.Description || '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {childDetails.Computers && childDetails.Computers.length > 0 && (
                <div className="ou-tree-details-section">
                  <h5>Computers ({childDetails.Computers.length})</h5>
                  <div style={{ overflowX: 'auto' }}>
                    <table className="data-table" style={{ fontSize: '13px' }}>
                      <thead>
                        <tr>
                          <th>Name</th>
                          <th>Operating System</th>
                          <th>Enabled</th>
                        </tr>
                      </thead>
                      <tbody>
                        {childDetails.Computers.map((c, i) => (
                          <tr key={i}>
                            <td>{c.Name}</td>
                            <td>{c.OperatingSystem || '-'}</td>
                            <td>{c.Enabled ? 'Yes' : 'No'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {childDetails.Users?.length === 0 && childDetails.Groups?.length === 0 && childDetails.Computers?.length === 0 && (
                <p style={{ color: 'var(--text-muted)', fontSize: '13px', marginTop: '12px' }}>
                  No direct child objects in this OU.
                </p>
              )}
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default OUTree;

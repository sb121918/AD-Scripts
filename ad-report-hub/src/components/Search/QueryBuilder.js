import React, { useState, useCallback } from 'react';
import './QueryBuilder.css';

// LDAP attributes by object type
const ATTRIBUTES = {
  user: [
    { value: 'samAccountName', label: 'Username (SAM)', type: 'string' },
    { value: 'displayName', label: 'Display Name', type: 'string' },
    { value: 'mail', label: 'Email', type: 'string' },
    { value: 'userPrincipalName', label: 'UPN', type: 'string' },
    { value: 'name', label: 'Name', type: 'string' },
    { value: 'department', label: 'Department', type: 'string' },
    { value: 'title', label: 'Title', type: 'string' },
    { value: 'description', label: 'Description', type: 'string' },
    { value: 'employeeID', label: 'Employee ID', type: 'string' },
    { value: 'telephoneNumber', label: 'Phone', type: 'string' },
    { value: 'mobile', label: 'Mobile', type: 'string' },
    { value: 'manager', label: 'Manager', type: 'string' },
    { value: 'memberOf', label: 'Member Of', type: 'string' },
    { value: 'userAccountControl', label: 'Enabled', type: 'boolean' },
    { value: 'lockoutTime', label: 'Locked Out', type: 'string' },
  ],
  computer: [
    { value: 'name', label: 'Name', type: 'string' },
    { value: 'operatingSystem', label: 'Operating System', type: 'string' },
    { value: 'operatingSystemVersion', label: 'OS Version', type: 'string' },
    { value: 'description', label: 'Description', type: 'string' },
    { value: 'dNSHostName', label: 'DNS Host Name', type: 'string' },
    { value: 'userAccountControl', label: 'Enabled', type: 'boolean' },
  ],
  group: [
    { value: 'name', label: 'Name', type: 'string' },
    { value: 'displayName', label: 'Display Name', type: 'string' },
    { value: 'mail', label: 'Email', type: 'string' },
    { value: 'description', label: 'Description', type: 'string' },
    { value: 'groupType', label: 'Group Type', type: 'string' },
    { value: 'member', label: 'Member', type: 'string' },
  ],
  contact: [
    { value: 'name', label: 'Name', type: 'string' },
    { value: 'displayName', label: 'Display Name', type: 'string' },
    { value: 'mail', label: 'Email', type: 'string' },
    { value: 'company', label: 'Company', type: 'string' },
    { value: 'department', label: 'Department', type: 'string' },
    { value: 'telephoneNumber', label: 'Phone', type: 'string' },
    { value: 'description', label: 'Description', type: 'string' },
  ],
  ou: [
    { value: 'name', label: 'Name', type: 'string' },
    { value: 'description', label: 'Description', type: 'string' },
    { value: 'managedBy', label: 'Managed By', type: 'string' },
    { value: 'gPLink', label: 'GPO Link', type: 'string' },
  ],
  gpo: [
    { value: 'displayName', label: 'Display Name', type: 'string' },
    { value: 'name', label: 'Name (GUID)', type: 'string' },
    { value: 'description', label: 'Description', type: 'string' },
    { value: 'gPCFileSysPath', label: 'File System Path', type: 'string' },
  ],
};

// Operators by attribute type
const OPERATORS = {
  string: [
    { value: 'equals', label: 'Equals', ldap: '=' },
    { value: 'contains', label: 'Contains', ldap: 'substring' },
    { value: 'startsWith', label: 'Starts With', ldap: 'startsWith' },
    { value: 'endsWith', label: 'Ends With', ldap: 'endsWith' },
    { value: 'notEquals', label: 'Not Equals', ldap: '!' },
    { value: 'exists', label: 'Exists (has value)', ldap: 'exists' },
    { value: 'notExists', label: 'Does Not Exist', ldap: 'notExists' },
  ],
  boolean: [
    { value: 'equals', label: 'Is', ldap: '=' },
    { value: 'notEquals', label: 'Is Not', ldap: '!' },
  ],
};

// Query templates
const QUERY_TEMPLATES = [
  {
    id: 'disabled-users',
    label: 'Disabled Users',
    objectType: 'user',
    conditions: [{ attribute: 'userAccountControl', operator: 'equals', value: 'false', logic: 'AND' }],
  },
  {
    id: 'locked-accounts',
    label: 'Locked Accounts',
    objectType: 'user',
    conditions: [{ attribute: 'lockoutTime', operator: 'exists', value: '', logic: 'AND' }],
  },
  {
    id: 'department-it',
    label: 'IT Department',
    objectType: 'user',
    conditions: [{ attribute: 'department', operator: 'equals', value: 'IT', logic: 'AND' }],
  },
  {
    id: 'mail-contains',
    label: 'Email Contains',
    objectType: 'user',
    conditions: [{ attribute: 'mail', operator: 'contains', value: '', logic: 'AND' }],
  },
  {
    id: 'disabled-computers',
    label: 'Disabled Computers',
    objectType: 'computer',
    conditions: [{ attribute: 'userAccountControl', operator: 'equals', value: 'false', logic: 'AND' }],
  },
  {
    id: 'empty-groups',
    label: 'Empty Groups',
    objectType: 'group',
    conditions: [{ attribute: 'member', operator: 'notExists', value: '', logic: 'AND' }],
  },
];

const DEFAULT_CONDITION = {
  attribute: 'samAccountName',
  operator: 'contains',
  value: '',
  logic: 'AND',
};

/**
 * Convert conditions to LDAP filter string
 */
export function buildLdapFilter(conditions, objectType, logicBetween = 'AND') {
  if (!conditions || conditions.length === 0) return '';

  const objectClassFilter = {
    user: '(objectClass=user)',
    computer: '(objectClass=computer)',
    group: '(objectClass=group)',
    contact: '(objectClass=contact)',
    ou: '(objectClass=organizationalUnit)',
    gpo: '(objectClass=groupPolicyContainer)',
  };

  const escapeLdapValue = (val) => {
    if (!val) return '';
    return String(val)
      .replace(/\\/g, '\\5c')
      .replace(/\*/g, '\\2a')
      .replace(/\(/g, '\\28')
      .replace(/\)/g, '\\29')
      .replace(/\0/g, '\\00');
  };

  const buildCondition = (cond) => {
    const attr = cond.attribute;
    const op = cond.operator;
    const val = escapeLdapValue(cond.value);
    const attrInfo = ATTRIBUTES[objectType]?.find((a) => a.value === attr);
    const attrType = attrInfo?.type || 'string';

    if (op === 'exists' || op === 'notExists') {
      const filter = attr === 'lockoutTime'
        ? '(lockoutTime>=1)'
        : `(${attr}=*)`;
      return op === 'notExists' ? `(!${filter})` : filter;
    }

    if (attrType === 'boolean') {
      const boolVal = val === 'true' || val === '1' || val === 'yes' ? '0' : '544'; // 544 = disabled
      if (attr === 'userAccountControl') {
        return val === 'true' || val === '1' || val === 'yes'
          ? '(!(userAccountControl:1.2.840.113556.1.4.803:=2))' // Not disabled
          : '(userAccountControl:1.2.840.113556.1.4.803:=2)'; // Disabled bit
      }
    }

    switch (op) {
      case 'equals':
        return `(${attr}=${val})`;
      case 'notEquals':
        return `(!(${attr}=${val}))`;
      case 'contains':
        return `(${attr}=*${val}*)`;
      case 'startsWith':
        return `(${attr}=${val}*)`;
      case 'endsWith':
        return `(${attr}=*${val})`;
      default:
        return `(${attr}=*${val}*)`;
    }
  };

  const conditionFilters = conditions
    .filter((c) => c.attribute && (c.operator === 'exists' || c.operator === 'notExists' || c.value))
    .map(buildCondition);

  if (conditionFilters.length === 0) return '';

  let combinedFilter;
  if (conditionFilters.length === 1) {
    combinedFilter = conditionFilters[0];
  } else {
    const op = logicBetween === 'OR' ? '|' : '&';
    combinedFilter = `(${op}${conditionFilters.join('')})`;
  }

  const objectFilter = objectClassFilter[objectType] || '(objectClass=*)';
  return `(&${objectFilter}${combinedFilter})`;
}

const QueryBuilder = ({ objectType, conditions, onConditionsChange, onLdapFilterChange, logicBetween, onLogicBetweenChange }) => {
  const attributes = ATTRIBUTES[objectType] || ATTRIBUTES.user;

  const getOperatorsForAttribute = (attr) => {
    const attrInfo = attributes.find((a) => a.value === attr);
    return OPERATORS[attrInfo?.type || 'string'] || OPERATORS.string;
  };

  const addCondition = useCallback(() => {
    const lastLogic = conditions.length > 0 ? conditions[conditions.length - 1].logic : 'AND';
    const newCond = {
      ...DEFAULT_CONDITION,
      attribute: attributes[0]?.value || 'samAccountName',
      operator: 'contains',
      value: '',
      logic: lastLogic,
    };
    onConditionsChange([...conditions, newCond]);
  }, [conditions, attributes, onConditionsChange]);

  const updateCondition = (index, field, value) => {
    const updated = [...conditions];
    updated[index] = { ...updated[index], [field]: value };
    if (field === 'attribute') {
      const attrInfo = attributes.find((a) => a.value === value);
      const ops = getOperatorsForAttribute(value);
      updated[index].operator = ops[0]?.value || 'contains';
      if (attrInfo?.type === 'boolean') {
        updated[index].value = 'true';
      }
    }
    onConditionsChange(updated);
  };

  const removeCondition = (index) => {
    const updated = conditions.filter((_, i) => i !== index);
    if (updated.length > 0 && index > 0) {
      updated[index - 1].logic = updated[index - 1].logic || 'AND';
    }
    onConditionsChange(updated);
  };

  const applyTemplate = (template) => {
    onConditionsChange(template.conditions.map((c) => ({ ...c })));
    if (template.objectType && onLogicBetweenChange) {
      onLogicBetweenChange('AND');
    }
  };

  const ldapPreview = buildLdapFilter(conditions, objectType, logicBetween);
  if (onLdapFilterChange) {
    onLdapFilterChange(ldapPreview);
  }

  return (
    <div className="query-builder">
      <div className="query-builder-header">
        <h3>Visual Query Builder</h3>
        <div className="logic-toggle">
          <span className="logic-label">Combine conditions with:</span>
          <button
            type="button"
            className={`logic-btn ${logicBetween === 'AND' ? 'active' : ''}`}
            onClick={() => onLogicBetweenChange?.('AND')}
          >
            AND
          </button>
          <button
            type="button"
            className={`logic-btn ${logicBetween === 'OR' ? 'active' : ''}`}
            onClick={() => onLogicBetweenChange?.('OR')}
          >
            OR
          </button>
        </div>
      </div>

      <div className="query-templates">
        <span className="templates-label">Quick templates:</span>
        {QUERY_TEMPLATES.filter((t) => t.objectType === objectType).map((t) => (
          <button
            key={t.id}
            type="button"
            className="template-btn"
            onClick={() => applyTemplate(t)}
            title={t.label}
          >
            {t.label}
          </button>
        ))}
      </div>

      <div className="conditions-list">
        {conditions.map((cond, index) => (
          <div key={index} className="condition-row">
            {index > 0 && (
              <span className="condition-logic-badge">{cond.logic || logicBetween}</span>
            )}
            <select
              value={cond.attribute}
              onChange={(e) => updateCondition(index, 'attribute', e.target.value)}
              className="condition-select attr-select"
            >
              {attributes.map((a) => (
                <option key={a.value} value={a.value}>
                  {a.label}
                </option>
              ))}
            </select>
            <select
              value={cond.operator}
              onChange={(e) => updateCondition(index, 'operator', e.target.value)}
              className="condition-select op-select"
            >
              {getOperatorsForAttribute(cond.attribute).map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
            {cond.operator !== 'exists' && cond.operator !== 'notExists' && (
              <input
                type={attributes.find((a) => a.value === cond.attribute)?.type === 'boolean' ? 'text' : 'text'}
                value={cond.value}
                onChange={(e) => updateCondition(index, 'value', e.target.value)}
                placeholder="Value..."
                className="condition-input"
              />
            )}
            {index > 0 && (
              <select
                value={cond.logic || 'AND'}
                onChange={(e) => updateCondition(index, 'logic', e.target.value)}
                className="condition-select logic-select"
              >
                <option value="AND">AND</option>
                <option value="OR">OR</option>
              </select>
            )}
            <button
              type="button"
              className="condition-remove"
              onClick={() => removeCondition(index)}
              title="Remove condition"
            >
              ×
            </button>
          </div>
        ))}
      </div>

      <button type="button" className="btn btn-secondary add-condition-btn" onClick={addCondition}>
        + Add Condition
      </button>

      {ldapPreview && (
        <div className="ldap-preview">
          <span className="preview-label">Generated LDAP filter:</span>
          <code className="ldap-filter">{ldapPreview}</code>
        </div>
      )}
    </div>
  );
};

export default QueryBuilder;
export { ATTRIBUTES, OPERATORS, QUERY_TEMPLATES };

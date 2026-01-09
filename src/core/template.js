import { getByPath } from './utils.js';
import { isPlainObject } from './styles.js';

function stringifyMaybe(val) {
  if (val == null) return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'number' || typeof val === 'boolean') return String(val);
  return JSON.stringify(val);
}

function applyTemplateToString(str, variables, maxPasses = 5) {
  let cur = String(str);
  for (let i = 0; i < maxPasses; i += 1) {
    const next = cur.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_m, expr) => {
      const v = getByPath(variables, expr);
      return stringifyMaybe(v);
    });
    if (next === cur) return cur;
    cur = next;
  }
  return cur;
}

export function applyTemplate(input, variables) {
  if (typeof input === 'string') {
    return applyTemplateToString(input, variables);
  }

  if (Array.isArray(input)) {
    return input.map((v) => applyTemplate(v, variables));
  }

  if (isPlainObject(input)) {
    if (typeof input.$var === 'string' && Object.keys(input).length === 1) {
      const v = getByPath(variables, input.$var);
      return applyTemplate(v, variables);
    }

    const out = {};
    for (const [k, v] of Object.entries(input)) {
      out[k] = applyTemplate(v, variables);
    }
    return out;
  }

  return input;
}

const { spawn } = require('child_process');
const path = require('path');

/**
 * Clean PowerShell stdout to extract valid JSON.
 * Handles BOM, progress messages, ANSI codes, and mixed output.
 */
function cleanPowerShellOutput(raw) {
  let cleaned = raw;
  // Remove UTF-8 BOM
  cleaned = cleaned.replace(/^\uFEFF/, '');
  // Remove ANSI escape codes
  cleaned = cleaned.replace(/\x1B\[[0-9;]*[a-zA-Z]/g, '');
  // Remove carriage returns
  cleaned = cleaned.replace(/\r/g, '');
  // Drop common non-JSON noise lines (warnings/verbose/progress)
  cleaned = cleaned
    .split('\n')
    .filter((line) => {
      const t = line.trim();
      if (!t) return true;
      return !(
        t.startsWith('WARNING:') ||
        t.startsWith('VERBOSE:') ||
        t.startsWith('DEBUG:') ||
        t.startsWith('Progress') ||
        t.startsWith('Write-Progress')
      );
    })
    .join('\n');
  // Trim whitespace
  cleaned = cleaned.trim();
  // If output contains non-JSON lines before the JSON array/object, extract just the JSON
  // PowerShell sometimes outputs progress or warning text before the JSON
  const jsonStart = cleaned.search(/[\[{]/);
  if (jsonStart === -1) {
    return '';
  }
  if (jsonStart > 0) {
    cleaned = cleaned.substring(jsonStart);
  }
  // Find the last ] or } to trim any trailing non-JSON output
  const lastBracket = cleaned.lastIndexOf(']');
  const lastBrace = cleaned.lastIndexOf('}');
  const jsonEnd = Math.max(lastBracket, lastBrace);
  if (jsonEnd > 0 && jsonEnd < cleaned.length - 1) {
    cleaned = cleaned.substring(0, jsonEnd + 1);
  }
  return cleaned;
}

/**
 * Fix unescaped double quotes inside JSON string values.
 * PowerShell's ConvertTo-Json sometimes fails to escape quotes in AD field values
 * e.g. "Description":"RR Level described as "vpnlab.bc""  (invalid)
 * becomes "Description":"RR Level described as vpnlab.bc"  (fixed)
 *
 * Strategy: walk through the string character by character tracking JSON state.
 */
function fixUnescapedQuotes(jsonStr) {
  const chars = [];
  let i = 0;
  const len = jsonStr.length;

  const escapeControlChar = (c) => {
    // Escape any control character that would break JSON parsing inside a string
    if (c === '\n') return '\\n';
    if (c === '\t') return '\\t';
    if (c === '\r') return '\\r';
    const code = c.charCodeAt(0);
    if (code < 0x20) {
      return `\\u${code.toString(16).padStart(4, '0')}`;
    }
    return c;
  };

  while (i < len) {
    const ch = jsonStr[i];

    // Outside a string value - pass through
    if (ch !== '"') {
      chars.push(ch);
      i++;
      continue;
    }

    // We hit a quote - this should be the opening quote of a JSON key or value string.
    // Push the opening quote.
    chars.push('"');
    i++;

    // Now collect everything until the *real* closing quote.
    // The real closing quote is followed by one of: , } ] :
    // (with optional whitespace in between)
    while (i < len) {
      const c = jsonStr[i];

      if (c === '\\') {
        // Already-escaped character - pass through both chars
        chars.push(c);
        i++;
        if (i < len) {
          chars.push(jsonStr[i]);
          i++;
        }
        continue;
      }

      if (c === '"') {
        // Is this the real closing quote?
        // Look ahead past any whitespace for , } ] :
        let lookAhead = i + 1;
        while (lookAhead < len && (jsonStr[lookAhead] === ' ' || jsonStr[lookAhead] === '\n' || jsonStr[lookAhead] === '\t')) {
          lookAhead++;
        }
        const nextSignificant = jsonStr[lookAhead];
        if (nextSignificant === ',' || nextSignificant === '}' || nextSignificant === ']' || nextSignificant === ':') {
          // This is the real closing quote
          chars.push('"');
          i++;
          break;
        } else {
          // This is an unescaped quote inside the string value - remove it
          i++;
          continue;
        }
      }

      // Regular character inside string
      // Escape any control characters that would break JSON
      chars.push(escapeControlChar(c));
      i++;
    }
  }

  return chars.join('');
}

/**
 * Execute a PowerShell script and return JSON results
 * @param {string} scriptName - Name of the script file (in scripts folder)
 * @param {object} args - Arguments to pass to the script
 * @param {string} subfolder - Optional subfolder within scripts (e.g., 'users', 'groups')
 * @returns {Promise<object>} - Parsed JSON result
 */
async function runPowerShell(scriptName, args = {}, subfolder = '') {
  const scriptsDir = path.join(__dirname, '../scripts', subfolder);
  const scriptPath = path.join(scriptsDir, scriptName);

  const formatRemediation = (remediation) => {
    if (!remediation || typeof remediation !== 'object') return '';
    const lines = [];
    for (const [k, v] of Object.entries(remediation)) {
      if (v === undefined || v === null || v === '') continue;
      lines.push(`- ${k}: ${String(v)}`);
    }
    return lines.length ? `\n\nRemediation:\n${lines.join('\n')}` : '';
  };

  // Build argument list
  const psArgs = [
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', scriptPath
  ];

  // Add script arguments (boolean true => PowerShell switch with no value)
  Object.entries(args).forEach(([key, value]) => {
    if (value === undefined || value === null || value === '') return;
    if (value === false) return;
    if (value === true) {
      psArgs.push(`-${key}`);
      return;
    }
    psArgs.push(`-${key}`, String(value));
  });

  return new Promise((resolve, reject) => {
    const startTime = Date.now();
    const stdoutChunks = [];
    let stderr = '';

    const ps = spawn('powershell.exe', psArgs, {
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
    });

    ps.stdout.on('data', (data) => {
      stdoutChunks.push(data);
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('error', (error) => {
      reject(new Error(`Failed to start PowerShell: ${error.message}`));
    });

    ps.on('close', (code) => {
      const duration = Date.now() - startTime;
      const stdout = Buffer.concat(stdoutChunks).toString('utf8');

      if (code !== 0) {
        // Try to get error from stdout (scripts output JSON errors)
        let errorMsg = stderr || `PowerShell exited with code ${code}`;
        try {
          const cleaned = cleanPowerShellOutput(stdout);
          const parsed = JSON.parse(cleaned);
          if (parsed && parsed.Error) {
            errorMsg = String(parsed.Error) + formatRemediation(parsed.Remediation);
          }
        } catch (e) {
          // If stdout has content but isn't JSON, include it
          if (stdout.trim()) {
            errorMsg = stdout.trim().substring(0, 500);
          }
        }
        reject(new Error(errorMsg));
        return;
      }

      // Clean and parse the output
      const cleaned = cleanPowerShellOutput(stdout);

      if (!cleaned) {
        resolve({ data: [], duration });
        return;
      }

      let result;
      try {
        result = JSON.parse(cleaned);
      } catch (firstParseError) {
        // JSON parse failed - likely unescaped quotes in AD field values.
        // Try fixing and re-parsing.
        console.warn(`Initial JSON parse failed for ${scriptName} (${cleaned.length} chars): ${firstParseError.message}`);
        console.warn('Attempting to fix unescaped quotes...');
        try {
          const fixed = fixUnescapedQuotes(cleaned);
          result = JSON.parse(fixed);
          console.warn('Fixed successfully - parsed after quote repair.');
        } catch (secondParseError) {
          // Still failing - log details for debugging
          const outputLen = cleaned.length;
          const firstChars = cleaned.substring(0, 200);
          const lastChars = cleaned.substring(Math.max(0, outputLen - 200));
          console.error(`JSON parse failed for ${scriptName} even after fix (${outputLen} chars)`);
          console.error(`First 200: ${firstChars}`);
          console.error(`Last 200: ${lastChars}`);
          console.error(`Parse error: ${secondParseError.message}`);
          reject(new Error(`Failed to parse output from ${scriptName}: ${secondParseError.message}`));
          return;
        }
      }

      // Check if the result itself contains an error
      if (result && result.Error) {
        const msg = String(result.Error) + formatRemediation(result.Remediation);
        reject(new Error(msg));
        return;
      }
      
      // Ensure data is always an array for consistency
      const dataArray = Array.isArray(result) ? result : [result];
      
      resolve({
        data: dataArray,
        duration
      });
    });
  });
}

/**
 * Execute inline PowerShell command
 * @param {string} command - PowerShell command to execute
 * @returns {Promise<object>} - Result
 */
async function runPowerShellCommand(command) {
  return new Promise((resolve, reject) => {
    const startTime = Date.now();
    let stdout = '';
    let stderr = '';

    const ps = spawn('powershell.exe', [
      '-NoProfile',
      '-ExecutionPolicy', 'Bypass',
      '-Command', command
    ], {
      windowsHide: true
    });

    ps.stdout.on('data', (data) => {
      stdout += data.toString();
    });

    ps.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ps.on('error', (error) => {
      reject(new Error(`Failed to start PowerShell: ${error.message}`));
    });

    ps.on('close', (code) => {
      const duration = Date.now() - startTime;

      if (code !== 0) {
        reject(new Error(stderr || `PowerShell exited with code ${code}`));
        return;
      }

      try {
        const result = JSON.parse(stdout.trim());
        resolve({ data: result, duration });
      } catch (parseError) {
        resolve({ data: stdout.trim(), duration, raw: true });
      }
    });
  });
}

module.exports = {
  runPowerShell,
  runPowerShellCommand
};

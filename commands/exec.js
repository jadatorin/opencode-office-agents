#!/usr/bin/env node
/**
 * office-agents exec command
 * Ejecuta código JavaScript en el documento de Office
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');
const BRIDGE_URL = 'https://localhost:4017';

function log(msg) {
  console.log(msg);
}

function getConfig() {
  if (fs.existsSync(CONFIG_PATH)) {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  }
  return null;
}

function parseArgs(args) {
  const result = { app: null, code: null, sandbox: false };
  
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--code' || args[i] === '-c') {
      result.code = args[i + 1];
      i++;
    } else if (args[i] === '--sandbox' || args[i] === '-s') {
      result.sandbox = true;
    } else if (!result.app && ['excel', 'word', 'ppt', 'powerpoint'].includes(args[i].toLowerCase())) {
      result.app = args[i].toLowerCase();
    }
  }
  
  return result;
}

async function exec(app, code, options = {}) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  const appName = app === 'ppt' || app === 'powerpoint' ? 'powerpoint' : app;
  
  // Usar el Bridge CLI
  const cmd = [
    'pnpm',
    ['exec', 'office-bridge', 'exec', appName, '--code', code, ...(options.sandbox ? ['--sandbox'] : [])],
    { cwd: installDir, stdio: 'inherit' }
  ];
  
  try {
    execSync(cmd[0], cmd[1], cmd[2]);
  } catch (error) {
    log(`Error: ${error.message}`);
    process.exit(1);
  }
}

const args = process.argv.slice(2);
const parsed = parseArgs(args);

if (!parsed.app) {
  log('Uso: office-agents exec <excel|word|ppt> --code "..."');
  log('       office-agents exec excel -c "return workbook.worksheets.items[0].name;"');
  process.exit(1);
}

if (!parsed.code) {
  log('Error: Falta --code');
  process.exit(1);
}

exec(parsed.app, parsed.code, { sandbox: parsed.sandbox }).catch(err => {
  log(`Error: ${err.message}`);
  process.exit(1);
});
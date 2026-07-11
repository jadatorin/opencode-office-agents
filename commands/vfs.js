#!/usr/bin/env node
/**
 * office-agents vfs command
 * Administra el sistema de archivos virtual del add-in
 */

const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');

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
  const result = { 
    action: null, 
    app: null, 
    source: null, 
    dest: null 
  };
  
  const validActions = ['ls', 'pull', 'push', 'rm'];
  
  for (let i = 0; i < args.length; i++) {
    const arg = args[i].toLowerCase();
    
    if (validActions.includes(arg)) {
      result.action = arg;
    } else if (!result.app && ['excel', 'word', 'ppt', 'powerpoint'].includes(arg)) {
      result.app = arg;
    } else if (result.action && !result.source) {
      result.source = args[i];
    } else if (result.action && result.source && !result.dest) {
      result.dest = args[i];
    }
  }
  
  return result;
}

async function vfs(action, app, options = {}) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  const appName = app === 'ppt' || app === 'powerpoint' ? 'powerpoint' : app;
  
  // Construir comando
  const cmdArgs = ['exec', 'office-bridge', 'vfs', action, appName];
  
  if (options.source) cmdArgs.push(options.source);
  if (options.dest) cmdArgs.push(options.dest);
  
  try {
    const result = spawnSync('pnpm', cmdArgs, { cwd: installDir, stdio: 'inherit', shell: true });
    if (result.status !== 0) {
      process.exit(result.status || 1);
    }
  } catch (error) {
    log(`Error: ${error.message}`);
    process.exit(1);
  }
}

const args = process.argv.slice(2);
const parsed = parseArgs(args);

if (!parsed.action || !parsed.app) {
  log('Uso: office-agents vfs <ls|pull|push|rm> <excel|word|ppt> [origen] [destino]');
  log('');
  log('Comandos:');
  log('  ls <app> [ruta]      - Lista archivos en el VFS');
  log('  pull <app> <VFS> <local> - Descarga archivo del VFS');
  log('  push <app> <local> <VFS>  - Sube archivo al VFS');
  log('  rm <app> <VFS>        - Elimina archivo del VFS');
  log('');
  log('Ejemplos:');
  log('  office-agents vfs ls word /home/user');
  log('  office-agents vfs pull word /home/user/uploads/reporte.docx ./reporte.docx');
  log('  office-agents vfs push word ./datos.xlsx /home/user/uploads/datos.xlsx');
  log('  office-agents vfs rm word /home/user/uploads/temp.txt');
  process.exit(1);
}

vfs(parsed.action, parsed.app, {
  source: parsed.source,
  dest: parsed.dest
}).catch(err => {
  log(`Error: ${err.message}`);
  process.exit(1);
});
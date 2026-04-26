#!/usr/bin/env node
/**
 * office-agents inspect command
 * Muestra herramientas disponibles en la sesión
 */

const { execSync } = require('child_process');
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
  if (args.length === 0) return null;
  const app = args[0].toLowerCase();
  if (['excel', 'word', 'ppt', 'powerpoint'].includes(app)) {
    return app;
  }
  return null;
}

async function inspect(app) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  const appName = app === 'ppt' || app === 'powerpoint' ? 'powerpoint' : app;
  
  try {
    execSync('pnpm exec office-bridge inspect ' + appName, { 
      cwd: installDir, 
      stdio: 'inherit' 
    });
  } catch (error) {
    log('Error: No se pudo inspeccionar la sesión.');
    log('Asegúrate de que el add-in está corriendo: office-agents start excel');
    process.exit(1);
  }
}

const app = parseArgs(process.argv.slice(2));

if (!app) {
  log('Uso: office-agents inspect <excel|word|ppt>');
  log('Ejemplo: office-agents inspect excel');
  process.exit(1);
}

inspect(app);
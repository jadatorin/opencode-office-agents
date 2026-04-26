#!/usr/bin/env node
/**
 * office-agents list command
 * Lista sesiones de Office conectadas al bridge
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

async function list() {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  
  try {
    execSync('pnpm exec office-bridge list', { 
      cwd: installDir, 
      stdio: 'inherit' 
    });
  } catch (error) {
    log('Error: No se pudo conectar al bridge.');
    log('Asegúrate de que el add-in está corriendo: office-agents start excel');
    process.exit(1);
  }
}

list();
#!/usr/bin/env node
/**
 * office-agents screenshot command
 * Captura documentos de Office como imágenes
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
  const result = { app: null, pages: null, sheetId: null, range: null, out: null, slideIndex: null };
  
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--pages') {
      result.pages = args[i + 1];
      i++;
    } else if (args[i] === '--sheet-id') {
      result.sheetId = args[i + 1];
      i++;
    } else if (args[i] === '--range') {
      result.range = args[i + 1];
      i++;
    } else if (args[i] === '--slide-index') {
      result.slideIndex = args[i + 1];
      i++;
    } else if (args[i] === '--out' || args[i] === '-o') {
      result.out = args[i + 1];
      i++;
    } else if (!result.app && ['excel', 'word', 'ppt', 'powerpoint'].includes(args[i].toLowerCase())) {
      result.app = args[i].toLowerCase();
    }
  }
  
  return result;
}

async function screenshot(app, options = {}) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  const appName = app === 'ppt' || app === 'powerpoint' ? 'powerpoint' : app;
  
  // Construir comando
  const cmdArgs = ['exec', 'office-bridge', 'screenshot', appName];
  
  if (options.pages) cmdArgs.push('--pages', options.pages);
  if (options.sheetId) cmdArgs.push('--sheet-id', options.sheetId);
  if (options.range) cmdArgs.push('--range', options.range);
  if (options.slideIndex) cmdArgs.push('--slide-index', options.slideIndex);
  if (options.out) cmdArgs.push('--out', options.out);
  
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

if (!parsed.app) {
  log('Uso: office-agents screenshot <excel|word|ppt> [opciones]');
  log('');
  log('Opciones Word:');
  log('  --pages N        - Número de páginas a capturar');
  log('  --out ARCHIVO    - Archivo de salida');
  log('');
  log('Opciones Excel:');
  log('  --sheet-id N    - ID de la hoja');
  log('  --range RANGO   - Rango a capturar (ej: A1:F20)');
  log('  --out ARCHIVO   - Archivo de salida');
  log('');
  log('Opciones PowerPoint:');
  log('  --slide-index N  - Índice del slide');
  log('  --out ARCHIVO   - Archivo de salida');
  log('');
  log('Ejemplos:');
  log('  office-agents screenshot word --pages 1 --out pagina.png');
  log('  office-agents screenshot excel --range A1:F20 --out rango.png');
  log('  office-agents screenshot ppt --slide-index 0 --out slide.png');
  process.exit(1);
}

if (!parsed.out) {
  log('Error: Falta --out ARCHIVO');
  process.exit(1);
}

screenshot(parsed.app, {
  pages: parsed.pages,
  sheetId: parsed.sheetId,
  range: parsed.range,
  slideIndex: parsed.slideIndex,
  out: parsed.out
}).catch(err => {
  log(`Error: ${err.message}`);
  process.exit(1);
});
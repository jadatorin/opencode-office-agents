#!/usr/bin/env node
/**
 * office-agents tool command
 * Invoca herramientas específicas del add-in
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
  const result = { app: null, tool: null, input: null, out: null };
  
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--input' || args[i] === '-i') {
      result.input = args[i + 1];
      i++;
    } else if (args[i] === '--out' || args[i] === '-o') {
      result.out = args[i + 1];
      i++;
    } else if (!result.app && ['excel', 'word', 'ppt', 'powerpoint'].includes(args[i].toLowerCase())) {
      result.app = args[i].toLowerCase();
    } else if (!result.tool && args[i].toLowerCase() !== '--input' && args[i].toLowerCase() !== '--out') {
      result.tool = args[i];
    }
  }
  
  return result;
}

async function tool(app, toolName, options = {}) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  const appName = app === 'ppt' || app === 'powerpoint' ? 'powerpoint' : app;
  
  // Construir comando
  const cmdArgs = ['exec', 'office-bridge', 'tool', appName, toolName];
  
  if (options.input) {
    cmdArgs.push('--input', options.input);
  }
  
  if (options.out) {
    cmdArgs.push('--out', options.out);
  }
  
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

if (!parsed.app || !parsed.tool) {
  log('Uso: office-agents tool <excel|word|ppt> <tool_name> --input "..." --out archivo');
  log('');
  log('Herramientas comunes Excel:');
  log('  get_range          - Leer un rango de celdas');
  log('  set_values        - Escribir valores');
  log('  get_sheet_names   - Listar hojas');
  log('  screenshot_range  - Capturar rango como imagen');
  log('');
  log('Herramientas comunes Word:');
  log('  get_document_text  - Leer texto del documento');
  log('  get_document_properties - Propiedades del doc');
  log('');
  log('Ejemplos:');
  log("  office-agents tool excel get_range --input '{\"range\":\"A1:D10\"}'");
  log('  office-agents tool word get_document_text');
  process.exit(1);
}

tool(parsed.app, parsed.tool, { 
  input: parsed.input,
  out: parsed.out 
}).catch(err => {
  log(`Error: ${err.message}`);
  process.exit(1);
});
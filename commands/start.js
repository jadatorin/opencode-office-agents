#!/usr/bin/env node
/**
 * office-agents start command
 * Inicia el servidor bridge y la app Office especificada
 */

const { execSync, spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const net = require('net');

const CONFIG_PATH = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');
const DEFAULT_PORT = 4017;

function log(msg, type = 'info') {
  const symbols = { info: 'ℹ', success: '✓', error: '✗', warn: '⚠' };
  console.log(`${symbols[type] || '•'} ${msg}`);
}

function getConfig() {
  if (fs.existsSync(CONFIG_PATH)) {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  }
  return null;
}

function checkPort(port) {
  return new Promise((resolve) => {
    const client = new net.Socket();
    client.once('error', () => {
      client.destroy();
      resolve(false);
    });
    client.once('connect', () => {
      client.destroy();
      resolve(true);
    });
    client.connect(port, '127.0.0.1');
  });
}

async function start(app) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Office-agents no está instalado. Ejecuta: office-agents install', 'error');
    process.exit(1);
  }
  
  const installDir = config.installDir;
  const appMap = {
    excel: { port: 3000, name: 'Excel', script: 'dev-server:excel' },
    word: { port: 3002, name: 'Word', script: 'dev-server:word' },
    ppt: { port: 3001, name: 'PowerPoint', script: 'dev-server:ppt' },
    powerpoint: { port: 3001, name: 'PowerPoint', script: 'dev-server:ppt' }
  };
  
  const appConfig = appMap[app.toLowerCase()];
  
  if (!appConfig) {
    log(`App desconocida: ${app}. Usa: excel, word, ppt`, 'error');
    process.exit(1);
  }
  
  console.log(`\n🚀 Iniciando ${appConfig.name}...\n`);
  
  // Verificar si el bridge ya está corriendo
  const bridgeRunning = await checkPort(DEFAULT_PORT);
  
  if (!bridgeRunning) {
    log('Iniciando bridge server...', 'info');
    const bridge = spawn('pnpm', ['bridge:serve'], {
      cwd: installDir,
      stdio: 'inherit',
      shell: true,
      detached: false
    });
    
    // Esperar a que inicie
    await new Promise(r => setTimeout(r, 3000));
  } else {
    log('Bridge server ya está corriendo', 'warn');
  }
  
  // Iniciar el dev server de la app
  log(`Iniciando ${appConfig.name} dev server...`, 'info');
  const devServer = spawn('pnpm', [appConfig.script], {
    cwd: installDir,
    stdio: 'inherit',
    shell: true,
    detached: false
  });
  
  // Esperar a que inicie
  await new Promise(r => setTimeout(r, 5000));
  
  // Abrir Office con el add-in sideloaded
  log(`Abriendo ${appConfig.name}...`, 'info');
  
  const officeExe = path.join(
    'C:\\Program Files\\Microsoft Office\\root\\Office16',
    app === 'ppt' || app === 'powerpoint' ? 'POWERPNT.EXE' :
    app === 'word' ? 'WINWORD.EXE' : 'EXCEL.EXE'
  );
  
  spawn(officeExe, [], {
    detached: true,
    shell: true,
    stdio: 'ignore'
  });
  
  log(`\n✓ ${appConfig.name} iniciado!`, 'success');
  log(`Bridge: https://localhost:${DEFAULT_PORT}`, 'info');
  log(`Dev server: https://localhost:${appConfig.port}`, 'info');
  log(`\nPuedes usar los comandos:`, 'info');
  log(`  office-agents list`, 'info');
  log(`  office-agents inspect ${app}`, 'info');
  log(`  office-agents exec ${app} --code "..."`, 'info');
}

// Obtener app del argumento
const app = process.argv[2];

if (!app) {
  console.log('Uso: office-agents start <excel|word|ppt>');
  console.log('Ejemplo: office-agents start excel');
  process.exit(1);
}

start(app).catch(err => {
  log(err.message, 'error');
  process.exit(1);
});
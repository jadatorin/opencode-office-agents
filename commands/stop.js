#!/usr/bin/env node
/**
 * office-agents stop command
 * Detiene todos los servicios (bridge server y dev servers)
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');

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

async function stop() {
  const config = getConfig();

  if (!config || !config.installDir) {
    log('Office-agents no está instalado. Nada que detener.', 'warn');
    return;
  }

  console.log('\n🛑 Deteniendo servicios...\n');

  // Matar bridge server por PID si existe, o por coincidencia de línea de comandos
  log('Deteniendo bridge server...', 'info');
  try {
    const pidFile = path.join(process.env.APPDATA, 'opencode', 'office-bridge.pid');
    if (fs.existsSync(pidFile)) {
      const pid = fs.readFileSync(pidFile, 'utf8').trim();
      execSync(`taskkill /f /pid ${pid} 2>nul || echo OK`, { stdio: 'pipe', shell: true });
      fs.unlinkSync(pidFile);
      log('Bridge server detenido', 'success');
    } else {
      // Fallback: buscar proceso node con 'bridge' en cmdline usando wmic
      try {
        const result = execSync(
          'wmic process where "name=\'node.exe\' and commandline like \'%%bridge%%\'" get processid 2>nul | findstr /r /c:"^[0-9]"',
          { encoding: 'utf8', shell: true }
        ).trim();
        if (result) {
          const pids = result.split(/\s+/).filter(p => /^\d+$/.test(p));
          pids.forEach(pid => {
            try {
              execSync(`taskkill /f /pid ${pid} 2>nul || echo OK`, { stdio: 'pipe', shell: true });
            } catch { /* already gone */ }
          });
          log('Bridge server detenido', 'success');
        } else {
          log('No se encontró bridge server corriendo', 'warn');
        }
      } catch {
        log('No se encontró bridge server corriendo', 'warn');
      }
    }
  } catch {
    log('No se encontró bridge server corriendo', 'warn');
  }

  // Matar dev servers (procesos node con 'dev-server' en cmdline)
  log('Deteniendo dev servers...', 'info');
  try {
    const devResult = execSync(
      'wmic process where "name=\'node.exe\' and commandline like \'%%dev-server%%\'" get processid 2>nul | findstr /r /c:"^[0-9]"',
      { encoding: 'utf8', shell: true }
    ).trim();
    if (devResult) {
      const pids = devResult.split(/\s+/).filter(p => /^\d+$/.test(p));
      pids.forEach(pid => {
        try {
          execSync(`taskkill /f /pid ${pid} 2>nul || echo OK`, { stdio: 'pipe', shell: true });
        } catch { /* already gone */ }
      });
      log('Dev servers detenidos', 'success');
    } else {
      log('No se encontraron dev servers corriendo', 'warn');
    }
  } catch {
    log('No se encontraron dev servers corriendo', 'warn');
  }

  log('\n✓ Todos los servicios detenidos', 'success');
}

stop().catch(err => {
  log(err.message, 'error');
  process.exit(1);
});

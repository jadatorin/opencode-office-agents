#!/usr/bin/env node
/**
 * office-agents setup command
 * Instala y configura el repo office-agents
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const INSTALL_DIR = path.join(process.env.USERPROFILE || process.env.HOME, 'office-agents');
// (ALT_INSTALL_DIR removed — was developer-specific artifact)

function log(msg, type = 'info') {
  const symbols = { info: 'ℹ', success: '✓', error: '✗', warn: '⚠' };
  console.log(`${symbols[type] || '•'} ${msg}`);
}

function run(cmd, options = {}) {
  try {
    return execSync(cmd, { 
      stdio: 'inherit',
      shell: true,
      ...options 
    });
  } catch (error) {
    if (options.shell !== false) {
      throw error;
    }
    return null;
  }
}

async function install() {
  console.log('\n🔧 Configurando office-agents...\n');
  
  // 1. Verificar prerequisites
  log('Verificando prerequisitos...', 'info');
  
  // Node.js
  try {
    const nodeVersion = execSync('node --version', { encoding: 'utf8' }).trim();
    log(`Node.js: ${nodeVersion}`, 'success');
  } catch {
    log('Node.js no está instalado. Instala Node.js 18+', 'error');
    process.exit(1);
  }
  
  // pnpm - Try global first, then use npx
  let pnpmAvailable = false;
  try {
    execSync('pnpm --version', { encoding: 'utf8' });
    pnpmAvailable = true;
  } catch {
    log('Instalando pnpm...', 'warn');
    try {
      execSync('npm install -g pnpm', { stdio: 'inherit', shell: true });
      pnpmAvailable = true;
    } catch {
      // try npx as fallback
      try {
        execSync('npx pnpm --version', { encoding: 'utf8' });
        pnpmAvailable = true;
      } catch {
        // use npm instead
        log('Usando npm como fallback...', 'warn');
      }
    }
  }
  
  if (pnpmAvailable) {
    log('pnpm: instalado', 'success');
  }
  
  // Microsoft 365 Apps (verificación básica)
  const officePaths = [
    'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE',
    'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE',
    'C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE',
    'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE',
    'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE',
    'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE'
  ];
  
  const officeApps = ['Excel', 'Word', 'PowerPoint'];
  const officeFiles = ['EXCEL.EXE', 'WINWORD.EXE', 'POWERPNT.EXE'];
  let hasOffice = false;
  
  for (let i = 0; i < officeApps.length; i++) {
    const exists = officePaths.some(p => p.endsWith(officeFiles[i]) && fs.existsSync(p));
    if (exists) {
      log(`${officeApps[i]}: ✓`, 'success');
      hasOffice = true;
    }
  }
  
  if (!hasOffice) {
    log('Office 365 no está instalado', 'error');
    process.exit(1);
  }
  
// 2. Clonar repo (check both possible locations)
  let targetDir = INSTALL_DIR;
  if (fs.existsSync(INSTALL_DIR)) {
    log('Repo ya existe, actualizando...', 'warn');
    execSync('git pull', { cwd: targetDir, stdio: 'inherit', shell: true });
  } else {
    log('Clonando repo office-agents...', 'info');
    execSync('git clone https://github.com/hewliyang/office-agents.git', {
      cwd: path.dirname(INSTALL_DIR),
      stdio: 'inherit'
    });
  }
  
  // Use the found dir for all subsequent operations
  const finalDir = INSTALL_DIR;

  // 3. Instalar dependencias (use npx pnpm as fallback)
  log('Instalando dependencias...', 'info');
  const pnpmCmd = 'npx pnpm';
  try {
    execSync(pnpmCmd + ' install', { cwd: finalDir, stdio: 'inherit', shell: true });
  } catch {
    log('pnpm install falló. Intentá correrlo manualmente en ' + finalDir, 'error');
    process.exit(1);
  }

  // 4. Configurar certificados de desarrollo
  log('Configurando certificados de desarrollo...', 'info');
  try {
    execSync('npx office-addin-dev-certs install', { 
      cwd: finalDir, 
      stdio: 'inherit',
      shell: true
    });
  } catch {
    log('Certificados ya configurados o requiere atención manual', 'warn');
  }

  // 5. Guardar ubicación en config
  const configPath = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');
  const config = {
    installDir: finalDir,
    installedAt: new Date().toISOString(),
    version: 'latest'
  };

  fs.mkdirSync(path.dirname(configPath), { recursive: true });
  fs.writeFileSync(configPath, JSON.stringify(config, null, 2));
  
  log(`\n✓ Setup completo!`, 'success');
  log(`Ubicación: ${finalDir}`, 'info');
  }

install();
#!/usr/bin/env node
/**
 * office-agents - Main command dispatcher
 * Microsoft Office AI integration via office-agents bridge
 */

const path = require('path');
const fs = require('fs');

const COMMANDS_DIR = __dirname;

function log(msg) {
  console.log(msg);
}

function getConfig() {
  const configPath = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');
  if (fs.existsSync(configPath)) {
    return JSON.parse(fs.readFileSync(configPath, 'utf8'));
  }
  return null;
}

function showHelp() {
  log(`
🎯 Office Agents - Microsoft Office AI Integration

📦 Instalación
   office-agents install          - Instalar y configurar office-agents

🚀 Inicio
   office-agents start excel    - Iniciar Excel con add-in
   office-agents start word    - Iniciar Word con add-in
   office-agents start ppt     - Iniciar PowerPoint con add-in
   office-agents stop         - Detener servicios

📋 Documentos
   office-agents list           - Listar sesiones conectadas
   office-agents inspect excel - Ver herramientas de Excel

⚡ Ejecución
   office-agents exec excel --code "..."
   office-agents tool excel get_range --input "{...}"

🖼️ Captura
   office-agents screenshot word --pages 1 --out page.png
   office-agents screenshot excel --range A1:F20 --out range.png

📁 VFS
   office-agents vfs ls word /home/user
   office-agents vfs push word ./file.xlsx /home/user/uploads/
   office-agents vfs pull word /home/user/uploads/file.xlsx ./file.xlsx

📝 VBA
   office-agents vba run excel MiMacro
   office-agents vba list excel
   office-agents vba export excel --out macros.bas

🔧 Utilidades
   office-agents status        - Ver estado de instalación
   office-agents update       - Actualizar office-agents
   office-agents uninstall   - Desinstalar

ℹ️ Ejemplos
   # Leer datos de Excel
   office-agents exec excel --code "return workbook.worksheets.items[0].name;"

   # Escribir en Word
   office-agents exec word --code "context.document.body.insertText('Hola!', 'end');"

   # Capturar documento
   office-agents screenshot word --out doc.png

   # Listar herramientas disponibles
   office-agents inspect excel

   # Ver estado
   office-agents status
`);
}

function showStatus() {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('❌ Office-agents no está instalado');
    log('   Ejecuta: office-agents install');
    return;
  }
  
  log('✅ Office-agents instalado');
  log(`   Ubicación: ${config.installDir}`);
  log(`   Instalado: ${config.installedAt}`);
  
  // Verificar si el directorio existe
  if (fs.existsSync(config.installDir)) {
    log('   Estado: ✓ Directorio existe');
  } else {
    log('   Estado: ⚠ Directorio no encontrado');
  }
}

async function update() {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('❌ Office-agents no está instalado');
    log('   Ejecuta: office-agents install');
    return;
  }
  
  const { execSync } = require('child_process');
  log('🔄 Actualizando office-agents...');
  
  try {
    execSync('git pull', { cwd: config.installDir, stdio: 'inherit' });
    execSync('pnpm install', { cwd: config.installDir, stdio: 'inherit' });
    log('✅ Actualización completa');
  } catch (error) {
    log('❌ Error en actualización: ' + error.message);
  }
}

function uninstall() {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('❌ Office-agents no está instalado');
    return;
  }
  
  const { execSync } = require('child_process');
  
  log('⚠️ Desinstalando office-agents...');
  log('   Directorio: ' + config.installDir);
  log('   Para eliminar manualmente: rmdir /s /q "' + config.installDir + '"');
  
  // Solo eliminar config
  const configPath = path.join(process.env.APPDATA, 'opencode', 'office-agents.json');
  if (fs.existsSync(configPath)) {
    fs.unlinkSync(configPath);
  }
  
  log('✅ Configuración eliminada');
}

// Dispatcher principal
async function main() {
  const command = process.argv[2];
  const subcommand = process.argv[3];
  
  const validCommands = [
    'install', 'start', 'stop', 'list', 'inspect', 
    'exec', 'tool', 'screenshot', 'vfs', 'vba',
    'status', 'update', 'uninstall', 'help'
  ];
  
  // Sin comando o help
  if (!command || command === 'help' || command === '--help' || command === '-h') {
    showHelp();
    return;
  }
  
  // Status
  if (command === 'status') {
    showStatus();
    return;
  }
  
  // Update
  if (command === 'update') {
    await update();
    return;
  }
  
  // Uninstall
  if (command === 'uninstall') {
    uninstall();
    return;
  }
  
  // Verificar comando válido
  if (!validCommands.includes(command)) {
    log(`Comando desconocido: ${command}`);
    log('Usa: office-agents help para ver comandos disponibles');
    process.exit(1);
  }
  
  // Cargar subcomando
  const commandFileMap = {
    'start': 'start.js',
    'install': 'setup.js',
    'stop': 'stop.js'
  };
  
  const commandFile = commandFileMap[command] || (command + '.js');
  const commandPath = path.join(COMMANDS_DIR, commandFile);
  
  if (!fs.existsSync(commandPath)) {
    log(`Comando no implementado: ${command}`);
    process.exit(1);
  }
  
  // Ejecutar subcomando
  try {
    require(commandPath);
  } catch (error) {
    log(`Error: ${error.message}`);
    process.exit(1);
  }
}

main();
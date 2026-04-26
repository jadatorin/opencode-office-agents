#!/usr/bin/env node
/**
 * office-agents vba commands
 * Manejo de macros VBA en documentos Office
 */

const { execSync, spawn } = require('child_process');
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
    macro: null, 
    params: null,
    code: null,
    out: null 
  };
  
  const actions = ['run', 'new', 'list', 'export'];
  
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    
    if (actions.includes(arg.toLowerCase())) {
      result.action = arg.toLowerCase();
    } else if (!result.app && ['excel', 'word', 'ppt'].includes(arg.toLowerCase())) {
      result.app = arg.toLowerCase();
    } else if (arg === '--params' || arg === '-p') {
      result.params = args[i + 1];
      i++;
    } else if (arg === '--code' || arg === '-c') {
      result.code = args[i + 1];
      i++;
    } else if (arg === '--out' || arg === '-o') {
      result.out = args[i + 1];
      i++;
    } else if (!result.action && !result.app && arg) {
      result.macro = arg;
    }
  }
  
  return result;
}

// Obtener la aplicación Office como objeto COM
function getOfficeApp(appName) {
  const appMap = {
    excel: 'Excel.Application',
    word: 'Word.Application',
    ppt: 'PowerPoint.Application',
    powerpoint: 'PowerPoint.Application'
  };
  
  return appMap[appName] || appMap.excel;
}

// PowerShell para ejecutar VBA
function runVbaMacro(appName, macroName, params = '') {
  const app = getOfficeApp(appName);
  const psScript = `
    $excel = New-Object -ComObject ${app}
    $wb = $excel.Workbooks.Open((Get-Item "${process.cwd()}").FullName)
    $excel.Run("${macroName}"${params ? ', ' + params : ''})
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
  `;
  
  return execSync(`powershell -Command "${psScript}"`, { 
    stdio: 'inherit',
    shell: true 
  });
}

// Listar macros VBA disponibles
function listVbaMacros(appName, docPath) {
  const app = getOfficeApp(appName);
  const psScript = `
    $app = New-Object -ComObject ${app}
    $wb = $app.Workbooks.Open("${docPath}")
    $vbPro = $wb.VBProject
    $vbPro.VBComponents | ForEach-Object {
      $_.CodeModule.Lines(1, $_.CodeModule.CountOfLines) -split "$(10)" | Select-String "Sub " -SimpleMatch
    } | Select-Object -Unique
    $wb.Close($false)
    $app.Quit()
  `;
  
  return execSync(`powershell -Command "${psScript}"`, { 
    stdio: 'inherit',
    shell: true 
  });
}

async function vba(action, app, options = {}) {
  const config = getConfig();
  
  if (!config || !config.installDir) {
    log('Error: Office-agents no instalado. Ejecuta: office-agents install');
    process.exit(1);
  }
  
  switch(action) {
    case 'run':
      if (!options.macro) {
        log('Error: Falta nombre de macro');
        process.exit(1);
      }
      log(`Ejecutando macro: ${options.macro}`);
      runVbaMacro(app, options.macro, options.params);
      break;
      
    case 'list':
      log('Listando macros VBA...');
      // Por ahora mostrar mensaje informativo
      log('Para listar macros, Abre el documento en Office y presiona Alt+F11');
      break;
      
    case 'new':
      if (!options.code) {
        log('Error: Falta código de macro (--code)');
        process.exit(1);
      }
      log('Creando macro VBA...');
      log('Nota: La creación de macros VBA requiere intervención manual.');
      log('Abre Office y usa el editor VBA (Alt+F11) para crear la macro.');
      log(`\nCódigo sugerido:\n${options.code}`);
      break;
      
    case 'export':
      log('Exportando código VBA...');
      log('Para exportar VBA:');
      log('1. Abre el documento en Office');
      log('2. Presiona Alt+F11 para abrir el Editor VBA');
      log('3. Archivo -> Exportar -> guardar como .bas');
      break;
      
    default:
      log('Acción desconocida');
      process.exit(1);
  }
}

const args = process.argv.slice(2);
const parsed = parseArgs(args);

if (!parsed.action) {
  log('Uso: office-agents vba <run|new|list|export> <excel|word|ppt> [opciones]');
  log('');
  log('Comandos:');
  log('  run <app> <nombre_macro> --params "p1,p2"  - Ejecuta una macro VBA');
  log('  new <app> --code "Sub MiMacro()..."    - Crea nueva macro');
  log('  list <app>                        - Lista macros del documento');
  log('  export <app> --out archivo.bas       - Exporta código VBA');
  log('');
  log('Ejemplos:');
  log('  office-agents vba run excel MiMacro');
  log('  office-agents vba run excel MiMacro -p "param1,param2"');
  log('  office-agents vba new excel -c "Sub Hola()${String.fromCharCode(10)}    MsgBox \\"Hola\\"${String.fromCharCode(10)}End Sub"');
  process.exit(1);
}

if (!parsed.app) {
  log('Error: Falta app (excel, word, ppt)');
  process.exit(1);
}

vba(parsed.action, parsed.app, {
  macro: parsed.macro,
  params: parsed.params,
  code: parsed.code,
  out: parsed.out
});
<img width="896" height="500" alt="image" src="https://github.com/user-attachments/assets/6b4d9126-ddbf-4960-a9ce-059ab43cc98e" />

---------------------------------------------------------------------------------------------------------------------------------------------------------------------
# 🎯 Office Agents Skill

> **Skill de opencode** para integrar Microsoft Office (Excel, Word, PowerPoint) con AI mediante el puente de [office-agents](https://github.com/hewliyang/office-agents).

## ⚠️ Setup Requerido

Para desarrollar con office-agents necesitás **3 terminals simultáneas**. La forma más fácil es usar el script automático:

```powershell
# Ejecutar el script (abre las 3 terminals automáticamente)
powershell -ExecutionPolicy Bypass -File "scripts\start-excel.ps1"
```

O manualmente:
1. Bridge: `node packages\bridge\dist\cli.js serve`
2. Dev Server: `cd packages/excel && npx vite --port 3000`
3. Sideload: `npx office-addin-dev-settings sideload packages\excel\manifest.xml --app Excel`

---

## Requisitos

| Requisito | Versión mínima |
|-----------|----------------|
| Node.js | 18+ |
| pnpm | 8+ |
| Microsoft 365 Apps | Office 365 / M365 |

## Instalación

### Opción 1: Automática (recomendada)

```bash
# Desde opencode, carga la skill:
skill: load office-agents

# O simplemente usa los comandos:
office-agents install
```

### Opción 2: Manual

> **Nota**: La ruta `~/.config/opencode/skills/office-agents` es solo un **ejemplo**. Podés instalar en cualquier lugar, solo asegurate de que la skill sea accesible desde opencode.

```bash
# Clonar este repo a cualquier directorio
git clone https://github.com/jadatorin/opencode-office-agents.git ~/opencode-office-agents

# Luego configurar manually en %APPDATA%\opencode\office-agents.json:
# {"installDir": "C:\\Users\\tu usuario\\opencode-office-agents"}
```

## Uso Rápido

```bash
# 1. Instalar office-agents
office-agents install

# 2. Iniciar Excel con add-in
office-agents start excel

# 3. Manipular documentos
office-agents exec excel --code "return workbook.worksheets.items[0].name;"
office-agents tool excel get_range --input '{"range":"A1:D10"}'

# 4. Capturar screenshot
office-agents screenshot word --pages 1 --out page.png
```

## Comandos

| Comando | Descripción |
|---------|-------------|
| `install` | Instalar y configurar office-agents |
| `start <app>` | Iniciar app Office con add-in |
| `stop` | Detener servicios |
| `list` | Listar sesiones conectadas |
| `inspect` | Ver herramientas disponibles |
| `exec` | Ejecutar JavaScript en documento |
| `tool` | Invocar herramientas del add-in |
| `screenshot` | Capturar documento como imagen |
| `vfs` | Administrar archivos virtuales |
| `vba` | Manejar macros VBA |

## Instalación de office-agents

El comando `install` automáticamente:

1. ✅ Verifica Node.js y pnpm
2. ✅ Verifica Office 365 Apps instalado
3. ✅ Clona el repo [hewliyang/office-agents](https://github.com/hewliyang/office-agents)
4. ✅ Instala dependencias con pnpm
5. ✅ Configura certificados de desarrollo
6. ✅ Guarda configuración en `%APPDATA%\opencode\office-agents.json`

## Solución de Problemas

### Error: bridge no conecta

```bash
# Reiniciar servicios
office-agents stop
office-agents start excel
```

### Error: certificados

```bash
# Regenerar certificados
cd %USERPROFILE%\office-agents
npx office-addin-dev-certs install
```

### Error: Office no encontrado

Verificar que Microsoft 365 Apps esté instalado.
Verificar en: `C:\Program Files\Microsoft Office\root\Office16\`

## Ejemplos

### 📊 Excel

```bash
# Leer rango de celdas
office-agents exec excel --code "
  const ws = workbook.worksheets.items[0];
  const range = ws.getRange('A1:C10');
  range.load('values');
  await context.sync();
  return range.values;
"

# Escribir datos
office-agents tool excel set_values --input '{"range":"A1","values":[["Nombre","Edad"],["Juan",30]]}'

# Get sheet names
office-agents tool excel get_sheet_names
```

### 📝 Word

```bash
# Leer texto
office-agents exec word --code "
  const body = context.document.body;
  body.load('text');
  await context.sync();
  return body.text;
"

# Insertar texto
office-agents exec word --code "
  context.document.body.insertText('Hola desde opencode!', 'end');
"
```

### 🎬 PowerPoint

```bash
# Obtener número de slides
office-agents exec ppt --code "
  return Presentation.slides.items.length;
"

# Exportar slide
office-agents screenshot ppt --slide-index 0 --out slide.png
```

## 📝 VBA

```bash
# Ejecutar macro existente
office-agents vba run excel MiMacro

# Con parámetros
office-agents vba run excel MiMacro -p "param1,param2"

# Exportar código VBA
office-agents vba export excel --out misMacros.bas
```

## Estructura

```
office-agents/
├── SKILL.md                 # Este archivo
├── README.md               # Documentación
├── package.json             # Config npm
├── commands/
│   ├── office-agents.js    # Dispatcher principal
│   ├── setup.js           # Instalación
│   ├── start.js          # Iniciar app
│   ├── list.js           # Listar sesiones
│   ├── inspect.js        # Inspeccionar
│   ├── exec.js          # Ejecutar JS
│   ├── tool.js          # Herramientas
│   ├── screenshot.js   # Captura
│   ├── vfs.js          # VFS
│   └── vba.js           # VBA
└── config/
    └── default.json     # Configuración por defecto
```

##️ Estado del Skill

- **Estado**: ✅ Completado
- **Versión**: 1.0.0
- **Autor**: BITORIN
- **Repo**: https://github.com/jadatorin/opencode-office-agents

## Licencia

MIT - Ver LICENSE para más detalles.

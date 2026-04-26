# Office Agents Commands

Directorio de comandos disponibles para la skill office-agents.

## Índice

| Comando | Descripción |
|---------|-------------|
| `office-agents.js` | Dispatcher principal |
| `setup.js` | Instalación y configuración |
| `start.js` | Iniciar app Office |
| `list.js` | Listar sesiones |
| `inspect.js` | Inspeccionar herramientas |
| `exec.js` | Ejecutar JavaScript |
| `tool.js` | Invocar herramientas |
| `screenshot.js` | Capturar documentos |
| `vfs.js` | Sistema de archivos virtual |
| `vba.js` | Macros VBA |

## Uso Rápido

```bash
# Instalación inicial
node commands/office-agents.js install

# IniciarExcel
node commands/office-agents.js start excel

# Ver ayuda
node commands/office-agents.js help
```

## Alias (agregar al PATH)

Para usar `office-agents` directamente desde cualquier directorio:

```bash
# Agregar al PATH temporal (PowerShell)
$env:PATH += ";C:\Users\5 de julio\.config\opencode\skills\office-agents\commands"

# O agregar permanentemente
[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";C:\Users\5 de julio\.config\opencode\skills\office-agents\commands", "User")
```

Luego usar:
```bash
office-agents install
office-agents start excel
office-agents exec excel --code "..."
```
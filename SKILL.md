---
name: office-agents
description: Microsoft Office AI integration via office-agents bridge. Manipulate Excel, Word, and PowerPoint documents with AI-powered tools.
trigger: When working with Office documents (xlsx, docx, pptx), Office add-ins, or Microsoft automation.
---

# Office Agents Skill

Skill para integrar opencode con Microsoft Office a través del puente de [office-agents](https://github.com/hewliyang/office-agents).

## Requisitos Previos

### Software necesario
- [ ] Microsoft 365 Apps (Excel, Word, PowerPoint)
- [ ] Node.js 18+
- [ ] pnpm (`npm install -g pnpm`)

### Instalación automática (recomendada)
```bash
# La skill maneja todo automáticamente
```

## Comandos

### Setup

```bash
# install - Instala el repo office-agents y configura el entorno
office-agents install

# start - Inicia el servidor bridge y la app Office especificada
office-agents start excel    # Excel
office-agents start word    # Word
office-agents start ppt     # PowerPoint

# stop - Detiene todos los servicios
office-agents stop
```

### Manipulación de Documentos

```bash
# list - Lista sesiones de Office conectadas
office-agents list

# inspect - Muestra herramientas disponibles en la sesión
office-agents inspect excel
office-agents inspect word

# exec - Ejecuta código JavaScript en el documento
office-agents exec excel --code "..."

# tool - Invoca herramientas específicas
office-agents tool excel get_range --input '{"sheetId":1,"range":"A1:D10"}'
office-agents tool word get_document_text

# screenshot - Captura el documento
office-agents screenshot word --pages 1 --out page.png
office-agents screenshot excel --sheet-id 1 --range A1:F20 --out range.png

# vfs - Administra el sistema de archivos virtual
office-agents vfs ls word /home/user
office-agents vfs push word ./file.xlsx /home/user/uploads/
office-agents vfs pull word /home/user/uploads/file.xlsx ./local.xlsx
```

### VBA (Macros)

```bash
# run - Ejecuta una macro VBA existente
office-agents vba run excel MiMacro --params "p1,p2"

# new - Crea una nueva macro VBA
office-agents vba new excel --code "Sub MiMacro()\n    MsgBox \"Hola\"\nEnd Sub"

# list - Lista macros disponibles en el documento
office-agents vba list excel

# export - Exporta el código VBA del documento
office-agents vba export excel --out macros.bas
```

## Ejemplos de Uso

### Excel

```bash
# Leer datos de un rango
office-agents tool excel get_range --input '{"sheetId":1,"range":"A1:C10"}'

# Escribir datos
office-agents tool excel set_values --input '{"sheetId":1,"range":"A1","values":[["Nombre","Edad"],["Juan",30]]}'

# Obtener nombres de hojas
office-agents exec excel --code "return workbook.worksheets.items.map(w => w.name);"

# Ejecutar macro VBA
office-agents vba run excel CalculateAll
```

### Word

```bash
# Leer texto completo
office-agents exec word --code "
  const body = context.document.body;
  body.load('text');
  await context.sync();
  return body.text;
"

# Insertar texto
office-agents exec word --code "
  const body = context.document.body;
  body.insertText('Hola desde opencode!', 'end');
  await context.sync();
"

# Leer propiedades del documento
office-agents tool word get_document_properties
```

### PowerPoint

```bash
# Obtener slides
office-agents exec ppt --code "
  const slides = Presentation.slides.items;
  return slides.length;
"

# Exportar slide como imagen
office-agents screenshot ppt --slide-index 0 --out slide1.png
```

## Solución de Problemas

### Error: No se puede conectar al bridge
```bash
# Reiniciar el bridge
office-agents stop
office-agents start excel
```

### Error: Certificados de desarrollo
```bash
# Regenerar certificados
cd office-agents
npx office-addin-dev-certs install
```

### Error: Office no está instalado
Verificar que Microsoft 365 Apps esté instalado y activo.

## Archivos

- `commands/` - Subcomandos de la skill
- `SKILL.md` - Documentación de la skill
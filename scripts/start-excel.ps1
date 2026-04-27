# Office Agents - Start Excel (3 Terminals)
# IMPORTANTE: Abrir una a la vez para evitar errores
# 
# USO: Edita $officeAgentsPath abajo con tu ruta al repo de office-agents
# Por defecto apunta a ~/office-agents

$officeAgentsPath = "$env:USERPROFILE\office-agents"

Write-Host "Office Agents - Starting Excel..." -ForegroundColor Cyan
Write-Host ""

# Terminal 1: Bridge Server
Write-Host "[Terminal 1] Iniciando Bridge Server..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$officeAgentsPath'; node packages\bridge\dist\cli.js serve"

# Esperar a que el bridge esté listo
Start-Sleep -Seconds 5

# Terminal 2: Excel Dev Server
Write-Host "[Terminal 2] Iniciando Excel Dev Server..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$officeAgentsPath\packages\excel'; npx vite --port 3000"

# Esperar a que el dev server esté listo
Start-Sleep -Seconds 5

# Terminal 3: Excel con Add-in (sideload)
Write-Host "[Terminal 3] Lanzando Excel con Add-in..." -ForegroundColor Yellow
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$officeAgentsPath'; & 'node_modules\.pnpm\node_modules\.bin\office-addin-dev-settings.CMD' sideload 'packages\excel\manifest.xml' --app Excel"

Write-Host ""
Write-Host "Listo!" -ForegroundColor Green
Write-Host "   1. Bridge Server -> http://localhost:4017"
Write-Host "   2. Excel Dev Server -> http://localhost:3000"
Write-Host "   3. Excel con Add-in"
Script para obtener los totales a partir de los PDF del S-21 

## Requisitos

Powershell 7 o superior

Bajar Powershell https://github.com/PowerShell/PowerShell/releases/tag/v7.1.3

## Modo de uso 

Crear archivo `config.json` en base a `config.sample.json`

```powershell 
.\Get-Totals.ps1 2021 09 
```

Primero parámetro: año de servicio
Segundo parámetro: número de mes

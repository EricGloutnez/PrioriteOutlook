<#
.SYNOPSIS
    Génère manifest.xml à partir de manifest.template.xml et config.json.

.DESCRIPTION
    Remplace le jeton {{BASE_URL}} du gabarit par l'adresse d'hébergement.
    Pour changer d'hébergement (Azure, serveur interne), modifier
    "baseUrl" dans config.json, puis relancer ce script.

.PARAMETER Dev
    Utilise "devBaseUrl" (https://localhost:3000) au lieu de "baseUrl"
    et produit manifest.dev.xml.

.PARAMETER BaseUrl
    Remplace l'adresse de config.json pour cette exécution seulement.

.EXAMPLE
    .\scripts\Generer-Manifeste.ps1
.EXAMPLE
    .\scripts\Generer-Manifeste.ps1 -Dev
.EXAMPLE
    .\scripts\Generer-Manifeste.ps1 -BaseUrl "https://complements.bonneville.ca/priorite"
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Dev,
    [string]$BaseUrl
)

$ErrorActionPreference = 'Stop'
$racine  = Split-Path -Parent $PSScriptRoot
$gabarit = Join-Path $racine 'manifest.template.xml'
$config  = Get-Content (Join-Path $racine 'config.json') -Raw -Encoding UTF8 | ConvertFrom-Json

if (-not $BaseUrl) {
    $BaseUrl = if ($Dev) { $config.devBaseUrl } else { $config.baseUrl }
}
$BaseUrl = $BaseUrl.TrimEnd('/')

if ($BaseUrl -notmatch '^https://') {
    throw "L'adresse doit commencer par https:// (reçu : $BaseUrl)."
}

$nomSortie = if ($Dev) { 'manifest.dev.xml' } else { 'manifest.xml' }
$sortie    = Join-Path $racine $nomSortie
$contenu   = (Get-Content $gabarit -Raw -Encoding UTF8).Replace('{{BASE_URL}}', $BaseUrl)

# Validation XML avant l'écriture
try {
    $null = [xml]$contenu
} catch {
    throw "Le manifeste généré n'est pas un XML valide : $($_.Exception.Message)"
}

if ($PSCmdlet.ShouldProcess($sortie, "Écrire le manifeste pour $BaseUrl")) {
    $utf8SansBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($sortie, $contenu, $utf8SansBom)
    Write-Host "Manifeste généré : $sortie"
    Write-Host "Adresse de base  : $BaseUrl"
}

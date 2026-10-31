param(
  # Modo GitHub Pages: informe o usuario.
  [string]$Usuario,

  [string]$Repo = "protocolo-scpm",

  # Modo generico: informe a URL base do painel (ex.: https://site.pmmg.mg.gov.br/protocolo).
  [string]$Endereco
)

# Troca o endereco de exemplo do manifesto pelo endereco real do painel.
#   .\ajustar-endereco.ps1 -Usuario meusuario
#   .\ajustar-endereco.ps1 -Endereco https://servidor.exemplo/protocolo
$ErrorActionPreference = "Stop"

if (-not $Endereco) {
  if (-not $Usuario) {
    throw "Informe -Usuario (GitHub Pages) ou -Endereco (URL base do painel)."
  }
  $Endereco = "https://$Usuario.github.io/$Repo"
}

$Endereco = $Endereco.Trim().TrimEnd('/')

if ($Endereco -notmatch '^https://') {
  Write-Warning "O Excel exige HTTPS. O endereco informado comeca com '$($Endereco.Substring(0, [Math]::Min(8, $Endereco.Length)))'."
}

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifest = Join-Path $root "manifest.xml"

if (-not (Test-Path $manifest)) {
  throw "manifest.xml nao encontrado em $root."
}

# Cobre o endereco de exemplo e tambem uma publicacao anterior no GitHub Pages.
$padrao = 'https://scpm5rpm\.github\.io/protocolo-scpm|https://[A-Za-z0-9._-]+\.github\.io/[A-Za-z0-9._-]+'

$utf8 = New-Object System.Text.UTF8Encoding($false)
$texto = [System.IO.File]::ReadAllText($manifest, [System.Text.Encoding]::UTF8)
$novoTexto = [regex]::Replace($texto, $padrao, $Endereco)

if ($novoTexto -eq $texto) {
  Write-Output "Nada para alterar: o manifesto ja aponta para $Endereco (ou nao tem endereco de exemplo)."
} else {
  [System.IO.File]::WriteAllText($manifest, $novoTexto, $utf8)
  $ocorrencias = ([regex]::Matches($texto, $padrao)).Count
  Write-Output "manifest.xml atualizado ($ocorrencias ocorrencia(s))."
  Write-Output "Endereco do painel agora e: $Endereco"
}

Write-Output ""
Write-Output "Proximos passos:"
Write-Output "  1. Envie o manifest.xml atualizado para a pasta do OneDrive (e para o repositorio, se usar GitHub Pages)."
Write-Output "  2. Cada usuario carrega o manifest.xml uma vez no Excel Online (Pagina Inicial > Suplementos)."

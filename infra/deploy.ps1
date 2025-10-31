# infra/deploy.ps1
# Build → Deploy no Cloud Run → dispara 1x → agenda diário 00:00 (America/Sao_Paulo)
# Requer Google Cloud SDK (gcloud) autenticado: gcloud auth login

param(
  [string]$ImageParam  # opcional: .\infra\deploy.ps1 -ImageParam pautasonp
)

$ErrorActionPreference = 'Stop'

function Assert-NonEmpty {
  param([string]$Value, [string]$Name)
  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw "Config inválida: '$Name' está vazio."
  }
}

# =========================
# CONFIG
# =========================
$Project          = "1025753172637"          # número ou projectId
$Region           = "southamerica-east1"
$Repo             = "pauta-repo"
$Tag              = "latest"

# Nome da imagem: parâmetro → env → default
$Image = if ($ImageParam) { $ImageParam } elseif ($env:IMAGE) { $env:IMAGE } else { "pautasonp" }

$Service          = "pauta-automation"
$JobName          = "pauta-automation-nightly"
$SchedulerSAName  = "pauta-scheduler-sa"
$TimeZone         = "America/Sao_Paulo"

# Parâmetros do app
$Sessao           = "71"
$DataDe           = "29/09/2025"
$DataAte          = "29/10/2025"
$Headless         = "true"
$HeaderTemplate   = "papel_timbrado_tcm.docx"
$DownloadDir      = "/workspace/downloads"
$OutputDir        = "/workspace/output"

# SMTP (Office 365) – senha via Secret Manager
$SmtpHost         = "smtp.office365.com"
$SmtpPort         = "587"
$SmtpUser         = "raphael.goncalves@tcmsp.tc.br"
$SmtpPassValue    = 'rhg#1004'

$EmailSender      = "raphael.goncalves@tcmsp.tc.br"
$EmailTo          = "raphael.goncalves@tcmsp.tc.br"
$EmailSubject     = "TESTE – Pauta SONP 71 gerada automaticamente (GCP)"
$EmailBody        = "<p>Envio automático de <b>teste</b> pós-deploy pelo GCP.</p>"

# e-TCM (via Secret)
$EtcmUserValue    = "20386"
$EtcmPassValue    = 'rhg#1004'

# =========================
# Resolver projeto
# =========================
Write-Host ">> Resolvendo projeto..." -ForegroundColor Cyan
$ProjectId  = (gcloud projects describe $Project --format 'value(projectId)')
$ProjectNum = (gcloud projects describe $Project --format 'value(projectNumber)')
Assert-NonEmpty $ProjectId  'ProjectId'
Assert-NonEmpty $ProjectNum 'ProjectNumber'
Write-Host ">> Projeto: id=$ProjectId  num=$ProjectNum" -ForegroundColor Green
gcloud config set project $ProjectId | Out-Null

# SA padrão do Cloud Run (Compute Default)
$RunSA = "$ProjectNum-compute@developer.gserviceaccount.com"

# Confere repo/imagem
if ([string]::IsNullOrWhiteSpace($Repo))  { $Repo  = "pauta-repo" }
if ([string]::IsNullOrWhiteSpace($Image)) { $Image = "pautasonp" }
Assert-NonEmpty $Repo  'Repo'
Assert-NonEmpty $Image 'Image'
Write-Host ">> Repo:   $Repo"  -ForegroundColor Yellow
Write-Host ">> Imagem: $Image" -ForegroundColor Yellow

# Monta o tag usando formatação (evita erro de ':')
$ImageUri = "{0}-docker.pkg.dev/{1}/{2}/{3}:{4}" -f $Region, $ProjectId, $Repo, $Image, $Tag
Write-Host ">> Tag da imagem: $ImageUri" -ForegroundColor Green
if ($ImageUri -notmatch '^[a-z0-9-]+-docker\.pkg\.dev\/[^\/]+\/[^\/]+\/[^:]+:[^:]+$') {
  throw "Tag de imagem inválido: $ImageUri"
}

# =========================
# Habilitar APIs
# =========================
Write-Host ">> Habilitando APIs..." -ForegroundColor Cyan
gcloud services enable `
  run.googleapis.com `
  cloudbuild.googleapis.com `
  cloudscheduler.googleapis.com `
  artifactregistry.googleapis.com `
  secretmanager.googleapis.com | Out-Null

# =========================
# Artifact Registry
# =========================
Write-Host ">> Verificando repositório Artifact Registry..." -ForegroundColor Cyan
gcloud artifacts repositories describe $Repo --location=$Region *> $null
if ($LASTEXITCODE -ne 0) {
  Write-Host ">> Criando repositório: $Repo" -ForegroundColor Yellow
  gcloud artifacts repositories create $Repo `
    --repository-format=docker `
    --location=$Region `
    --description="Imagens da pauta SONP" | Out-Null
} else {
  Write-Host ">> Repositório $Repo já existe" -ForegroundColor Green
}

# =========================
# Build & Push
# =========================
Write-Host ">> Build & Push: $ImageUri" -ForegroundColor Cyan
gcloud builds submit --tag $ImageUri
if ($LASTEXITCODE -ne 0) { throw "Build falhou. Veja os logs do Cloud Build e corrija antes de continuar." }

# =========================
# Segredos
# =========================
function Ensure-Secret {
  param([string]$Name, [string]$Value)
  gcloud secrets describe $Name *> $null
  if ($LASTEXITCODE -eq 0) {
    $Value | gcloud secrets versions add $Name --data-file=- | Out-Null
  } else {
    $Value | gcloud secrets create $Name --data-file=- | Out-Null
  }
  gcloud secrets add-iam-policy-binding $Name `
    --member=("serviceAccount:$RunSA") `
    --role="roles/secretmanager.secretAccessor" *> $null
}
Write-Host ">> Gravando segredos..." -ForegroundColor Cyan
Ensure-Secret -Name "ETCM_USER" -Value $EtcmUserValue
Ensure-Secret -Name "ETCM_PASS" -Value $EtcmPassValue
Ensure-Secret -Name "SMTP_PASS" -Value $SmtpPassValue

# =========================
# Deploy Cloud Run
# =========================
Write-Host ">> Deploy Cloud Run: $Service" -ForegroundColor Cyan

# monta envs como array (evita problema com vírgulas no PowerShell)
$envVars = @(
  "TZ=$TimeZone",
  "BASE_URL=https://etcm.tcm.sp.gov.br",
  "SESSAO=$Sessao",
  "DATA_DE=$DataDe",
  "DATA_ATE=$DataAte",
  "HEADER_TEMPLATE=$HeaderTemplate",
  "DOWNLOAD_DIR=$DownloadDir",
  "OUTPUT_DIR=$OutputDir",
  "HEADLESS=$Headless",
  "SMTP_HOST=$SmtpHost",
  "SMTP_PORT=$SmtpPort",
  "SMTP_USER=$SmtpUser",
  "EMAIL_SENDER=$EmailSender",
  "EMAIL_TO=$EmailTo",
  "EMAIL_SUBJECT=$EmailSubject",
  "EMAIL_BODY=$EmailBody"
)

gcloud run deploy $Service `
  --image $ImageUri `
  --region $Region `
  --allow-unauthenticated `
  --service-account $RunSA `
  --max-instances 1 `
  --cpu 1 --memory 1Gi `
  --set-env-vars $envVars `
  --set-secrets "ETCM_USER=ETCM_USER:latest,ETCM_PASS=ETCM_PASS:latest,SMTP_PASS=SMTP_PASS:latest"
if ($LASTEXITCODE -ne 0) { throw "Deploy falhou. Corrija o erro do Cloud Run e rode o script de novo." }

$ServiceUrl = (gcloud run services describe $Service --region $Region --format 'value(status.url)')
if ([string]::IsNullOrWhiteSpace($ServiceUrl)) { throw "Sem URL do serviço; não é possível criar o Scheduler. Verifique o deploy e rode novamente." }
Write-Host ">> Serviço disponível em: $ServiceUrl" -ForegroundColor Green

# =========================
# Cloud Scheduler (OIDC)
# =========================
Write-Host ">> Configurando Cloud Scheduler..." -ForegroundColor Cyan
$SchedulerSA = "$SchedulerSAName@$ProjectId.iam.gserviceaccount.com"
gcloud iam service-accounts describe $SchedulerSA *> $null
if ($LASTEXITCODE -ne 0) {
  Write-Host ">> Criando SA do Scheduler: $SchedulerSAName" -ForegroundColor Yellow
  gcloud iam service-accounts create $SchedulerSAName --display-name="Scheduler SA" | Out-Null
}
gcloud run services add-iam-policy-binding $Service `
  --region $Region `
  --member ("serviceAccount:$SchedulerSA") `
  --role roles/run.invoker *> $null

$JobUri   = "$ServiceUrl/run"
$Schedule = "0 0 * * *"

gcloud scheduler jobs describe $JobName --location $Region *> $null
if ($LASTEXITCODE -eq 0) {
  Write-Host ">> Atualizando job do Scheduler: $JobName" -ForegroundColor Yellow
  gcloud scheduler jobs update http $JobName `
    --location $Region `
    --schedule $Schedule `
    --time-zone $TimeZone `
    --http-method POST `
    --uri $JobUri `
    --oidc-service-account-email $SchedulerSA `
    --oidc-token-audience $JobUri | Out-Null
} else {
  Write-Host ">> Criando job do Scheduler: $JobName" -ForegroundColor Yellow
  gcloud scheduler jobs create http $JobName `
    --location $Region `
    --schedule $Schedule `
    --time-zone $TimeZone `
    --http-method POST `
    --uri $JobUri `
    --oidc-service-account-email $SchedulerSA `
    --oidc-token-audience $JobUri | Out-Null
}

# =========================
# Disparo imediato
# =========================
Write-Host ">> Disparando execução imediata via Scheduler..." -ForegroundColor Cyan
gcloud scheduler jobs run $JobName --location $Region | Out-Null

Write-Host "✔ Finalizado." -ForegroundColor Green

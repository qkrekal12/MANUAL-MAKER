param(
  [Parameter(Mandatory = $true)]
  [ValidateSet("accounts", "send", "pdf")]
  [string]$Action
)

$ErrorActionPreference = "Stop"
$tempAttachment = $null
$outlook = $null
$word = $null
$document = $null
$tempDocx = $null
$tempPdf = $null

try {
  if ($Action -eq "accounts") {
    $outlook = New-Object -ComObject Outlook.Application
    $session = $outlook.Session
    $accounts = @()
    foreach ($account in $session.Accounts) {
      $address = [string]$account.SmtpAddress
      if ($address) {
        $accounts += [pscustomobject]@{
          address = $address
          name = [string]$account.DisplayName
        }
      }
    }
    @{ accounts = $accounts } | ConvertTo-Json -Depth 4 -Compress
    exit 0
  }

  $reader = [System.IO.StreamReader]::new([Console]::OpenStandardInput())
  try { $payload = $reader.ReadToEnd() | ConvertFrom-Json }
  finally { $reader.Dispose() }

  if ($Action -eq "pdf") {
    $docxBase64 = [string]$payload.docx
    $filename = [System.IO.Path]::GetFileName([string]$payload.filename)
    if (-not $docxBase64) { throw "PDF로 변환할 Word 문서가 없습니다." }
    if ($docxBase64.Length -gt 45000000) { throw "Word 문서가 너무 큽니다." }
    if (-not $filename.EndsWith(".docx", [System.StringComparison]::OrdinalIgnoreCase)) { $filename = "사용자_매뉴얼.docx" }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filename)
    $tempDocx = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + "_" + $filename)
    $tempPdf = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + "_" + $baseName + ".pdf")
    [System.IO.File]::WriteAllBytes($tempDocx, [Convert]::FromBase64String($docxBase64))
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($tempDocx, $false, $true)
    $document.ExportAsFixedFormat($tempPdf, 17)
    $document.Close(0)
    $document = $null
    $pdf = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($tempPdf))
    @{ pdf = $pdf } | ConvertTo-Json -Depth 3 -Compress
    exit 0
  }

  $outlook = New-Object -ComObject Outlook.Application
  $session = $outlook.Session

  $to = [string]$payload.to
  $from = [string]$payload.from
  $subject = [string]$payload.subject
  $body = [string]$payload.body
  $filename = [string]$payload.filename
  $attachmentBase64 = [string]$payload.attachment

  if (-not $to -or -not $from -or -not $subject -or -not $body -or -not $attachmentBase64) {
    throw "받는 사람, 보내는 계정, 제목, 본문과 매뉴얼 파일이 모두 필요합니다."
  }
  if ($to.Length -gt 2000 -or $subject.Length -gt 255 -or $body.Length -gt 100000) {
    throw "이메일 입력 내용이 허용된 길이를 초과했습니다."
  }
  $recipients = $to -split "[,;]" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
  if (-not $recipients.Count -or @($recipients | Where-Object { $_ -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' }).Count -gt 0) {
    throw "받는 사람 이메일 주소를 확인해 주세요."
  }
  if ($attachmentBase64.Length -gt 45000000) {
    throw "매뉴얼 첨부 파일이 너무 큽니다."
  }

  $selectedAccount = $null
  foreach ($account in $session.Accounts) {
    if ([string]::Equals([string]$account.SmtpAddress, $from, [System.StringComparison]::OrdinalIgnoreCase)) {
      $selectedAccount = $account
      break
    }
  }
  if (-not $selectedAccount) {
    throw "선택한 발신 계정이 Outlook에 설정되어 있지 않습니다. 계정 목록을 새로고침해 주세요."
  }

  $safeName = [System.IO.Path]::GetFileName($filename)
  if (-not $safeName.EndsWith(".docx", [System.StringComparison]::OrdinalIgnoreCase)) {
    $safeName = "사용자_매뉴얼.docx"
  }
  $tempAttachment = Join-Path ([System.IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString() + "_" + $safeName)
  [System.IO.File]::WriteAllBytes($tempAttachment, [Convert]::FromBase64String($attachmentBase64))

  $mail = $outlook.CreateItem(0)
  $mail.SendUsingAccount = $selectedAccount
  $mail.To = ($recipients -join ";")
  $mail.Subject = $subject
  $mail.Body = $body
  [void]$mail.Attachments.Add($tempAttachment)
  $mail.Send()
  @{ sent = $true; from = [string]$selectedAccount.SmtpAddress; to = ($recipients -join ";") } | ConvertTo-Json -Depth 4 -Compress
}
catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
}
finally {
  if ($tempAttachment -and (Test-Path -LiteralPath $tempAttachment)) {
    Remove-Item -LiteralPath $tempAttachment -Force -ErrorAction SilentlyContinue
  }
  if ($tempDocx -and (Test-Path -LiteralPath $tempDocx)) {
    Remove-Item -LiteralPath $tempDocx -Force -ErrorAction SilentlyContinue
  }
  if ($tempPdf -and (Test-Path -LiteralPath $tempPdf)) {
    Remove-Item -LiteralPath $tempPdf -Force -ErrorAction SilentlyContinue
  }
  if ($word) {
    try { $word.Quit(0) } catch {}
  }
  if ($outlook -and [Runtime.InteropServices.Marshal]::IsComObject($outlook)) {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
  }
  if ($document -and [Runtime.InteropServices.Marshal]::IsComObject($document)) {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
  }
  if ($word -and [Runtime.InteropServices.Marshal]::IsComObject($word)) {
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
}

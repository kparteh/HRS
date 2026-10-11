$ErrorActionPreference = "Stop"
$root = [System.IO.Path]::GetFullPath($PSScriptRoot)
$rootPrefix = $root.TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:8787/")
$listener.Start()
Write-Output "WorkSmart HR is serving at http://127.0.0.1:8787/"

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $response = $context.Response
    try {
      if ($context.Request.HttpMethod -notin @("GET", "HEAD")) {
        $response.StatusCode = 405
        $response.Headers.Add("Allow", "GET, HEAD")
      } else {
        $relativePath = [System.Uri]::UnescapeDataString($context.Request.Url.AbsolutePath.TrimStart("/"))
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
          $relativePath = "index.html"
        }
        $filePath = [System.IO.Path]::GetFullPath((Join-Path $root $relativePath))
        if (-not $filePath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
          $response.StatusCode = 403
        } elseif (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
          $response.StatusCode = 404
        } else {
          $extension = [System.IO.Path]::GetExtension($filePath).ToLowerInvariant()
          $contentType = switch ($extension) {
            ".html" { "text/html; charset=utf-8" }
            ".css" { "text/css; charset=utf-8" }
            ".js" { "text/javascript; charset=utf-8" }
            ".json" { "application/json; charset=utf-8" }
            default { "application/octet-stream" }
          }
          $bytes = [System.IO.File]::ReadAllBytes($filePath)
          $response.ContentType = $contentType
          $response.ContentLength64 = $bytes.Length
          if ($context.Request.HttpMethod -ne "HEAD") {
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
          }
        }
      }
    } catch {
      Write-Error ("Local preview request failed: " + $_.Exception.Message)
      $response.StatusCode = 500
    } finally {
      $response.Close()
    }
  }
} finally {
  $listener.Stop()
  $listener.Close()
}

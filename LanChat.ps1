#requires -Version 5.1
<#
LanChat - LAN Model Assistant
- Windows PowerShell 5.1+ / PowerShell 7+
- No compilation, no external modules, no internet access required
- Calls only the Base URL entered by the operator
- Does not save API keys, prompts, or responses to disk
#>

$ErrorActionPreference = "Stop"
$script:HttpTimeoutSeconds = 45

# Configure console text encoding. This does not replace explicit UTF-8 decoding
# of HTTP response bodies, which is handled separately in Read-ResponseUtf8.
# Keep this source file saved as UTF-8 with BOM for Windows PowerShell 5.1.
try {
    $script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [Console]::InputEncoding = $script:Utf8NoBom
    [Console]::OutputEncoding = $script:Utf8NoBom
    $OutputEncoding = $script:Utf8NoBom
} catch {
    # Some embedded hosts do not allow changing console encodings.
    # Continue running; the host's configured code page will be used.
}

function Write-Section([string]$Text) {
    Write-Host ""
    Write-Host ("=" * 68) -ForegroundColor DarkGray
    Write-Host " $Text" -ForegroundColor Cyan
    Write-Host ("=" * 68) -ForegroundColor DarkGray
}

function Read-NonEmpty([string]$Prompt, [string]$Default = "") {
    while ($true) {
        if ($Default) {
            $value = Read-Host "$Prompt [$Default]"
            if ([string]::IsNullOrWhiteSpace($value)) { $value = $Default }
        } else {
            $value = Read-Host $Prompt
        }
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
        Write-Host "此项不能为空。" -ForegroundColor Yellow
    }
}

function Read-ApiKey {
    $secure = Read-Host "API Key（服务不需要鉴权时直接回车）" -AsSecureString
    if ($secure.Length -eq 0) { return "" }
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Normalize-BaseUrl([string]$Url) {
    $u = $Url.Trim().TrimEnd("/")
    if ($u -match "/(chat/completions|models)$") {
        $u = $u -replace "/(chat/completions|models)$", ""
    }
    if ($u -match "/v1$") { return $u }
    return "$u/v1"
}

function New-RequestHeaders {
    $headers = @{ "Accept" = "application/json" }
    if (-not [string]::IsNullOrWhiteSpace($script:ApiKey)) {
        $headers["Authorization"] = "Bearer $script:ApiKey"
    }
    return $headers
}

function Read-ResponseUtf8 {
    param([Parameter(Mandatory=$true)][System.IO.Stream]$Stream)

    # Do not trust the server's charset header for JSON responses. Many
    # OpenAI-compatible gateways return UTF-8 JSON without a charset, or with
    # an incorrect charset. Read bytes and decode UTF-8 explicitly instead.
    $reader = [System.IO.StreamReader]::new($Stream, [System.Text.Encoding]::UTF8, $true)
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

function Invoke-ApiRequest {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("GET","POST")][string]$Method,
        [Parameter(Mandatory=$true)][string]$Url,
        [object]$Body = $null
    )

    $request = $null
    $response = $null
    try {
        # Use HttpWebRequest so response bytes can be decoded explicitly as UTF-8.
        # This avoids Windows PowerShell 5.1 Invoke-RestMethod charset-related
        # mojibake when the API returns Chinese text in JSON.
        $request = [System.Net.HttpWebRequest]([System.Net.WebRequest]::Create($Url))
        $request.Method = $Method
        $request.Timeout = $script:HttpTimeoutSeconds * 1000
        $request.ReadWriteTimeout = $script:HttpTimeoutSeconds * 1000
        $request.Accept = "application/json"
        $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

        if (-not [string]::IsNullOrWhiteSpace($script:ApiKey)) {
            $request.Headers["Authorization"] = "Bearer $script:ApiKey"
        }

        if ($Method -eq "POST") {
            $request.ContentType = "application/json; charset=utf-8"
            $json = ConvertTo-Json -InputObject $Body -Depth 20 -Compress
            $bytes = $script:Utf8NoBom.GetBytes([string]$json)
            $request.ContentLength = $bytes.Length
            $requestStream = $request.GetRequestStream()
            try {
                $requestStream.Write($bytes, 0, $bytes.Length)
            } finally {
                $requestStream.Dispose()
            }
        }

        $response = [System.Net.HttpWebResponse]$request.GetResponse()
        $responseText = Read-ResponseUtf8 -Stream $response.GetResponseStream()

        if ([string]::IsNullOrWhiteSpace($responseText)) { return $null }
        try {
            return (ConvertFrom-Json -InputObject $responseText -ErrorAction Stop)
        } catch {
            # Preserve non-JSON success bodies rather than hiding the response.
            return $responseText
        }
    } catch [System.Net.WebException] {
        $webException = $_.Exception
        $status = $null
        $responseText = $null
        $errorResponse = $webException.Response

        if ($errorResponse -is [System.Net.HttpWebResponse]) {
            try { $status = [int]$errorResponse.StatusCode } catch {}
            try {
                $errorStream = $errorResponse.GetResponseStream()
                if ($null -ne $errorStream) {
                    $responseText = Read-ResponseUtf8 -Stream $errorStream
                }
            } catch {
                # Retain the original HTTP/network error if the body cannot be read.
            } finally {
                try { $errorResponse.Dispose() } catch {}
            }
        }

        $message = $webException.Message
        if ($status) { $message = "HTTP $status - $message" }
        if (-not [string]::IsNullOrWhiteSpace($responseText)) {
            $message += "`n服务端错误详情：`n$responseText"
        }
        throw $message
    } catch {
        throw $_.Exception.Message
    } finally {
        if ($null -ne $response) {
            try { $response.Dispose() } catch {}
        }
    }
}

function Test-Models {
    Write-Section "测试模型列表接口"
    $url = "$($script:BaseUrl)/models"
    Write-Host "请求：GET $url"
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $result = Invoke-ApiRequest -Method GET -Url $url
        $sw.Stop()
        Write-Host "请求成功，耗时 $($sw.ElapsedMilliseconds) ms" -ForegroundColor Green
        $items = @()
        if ($null -ne $result.data) { $items = @($result.data) }
        if ($items.Count -gt 0) {
            Write-Host "服务返回的模型 ID：" -ForegroundColor Cyan
            $items | ForEach-Object {
                if ($_.id) { Write-Host "  - $($_.id)" }
            }
            $ids = @($items | ForEach-Object { $_.id } | Where-Object { $_ })
            if ($ids -contains $script:Model) {
                Write-Host "当前模型 ID 在列表中。" -ForegroundColor Green
            } else {
                Write-Host "提示：当前模型 ID 未出现在列表中；这不一定代表它不能调用。" -ForegroundColor Yellow
            }
        } else {
            Write-Host "接口有响应，但未发现标准的 data[].id 模型列表。" -ForegroundColor Yellow
            Write-Host ($result | ConvertTo-Json -Depth 8)
        }
        return $true
    } catch {
        $sw.Stop()
        Write-Host "模型列表测试失败：$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "注意：有些兼容服务不实现 /models；可继续测试聊天接口。" -ForegroundColor Yellow
        return $false
    }
}

function Test-Chat {
    Write-Section "测试 Chat Completions 接口"
    $url = "$($script:BaseUrl)/chat/completions"
    $prompt = "请只回复：内网模型连接成功。"
    $body = @{
        model = $script:Model
        messages = @(
            @{ role = "user"; content = $prompt }
        )
        stream = $false
        temperature = 0
        max_tokens = 64
    }
    Write-Host "请求：POST $url"
    Write-Host "模型：$($script:Model)"
    Write-Host "提示词：$prompt"
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $result = Invoke-ApiRequest -Method POST -Url $url -Body $body
        $sw.Stop()
        Write-Host "请求成功，耗时 $($sw.ElapsedMilliseconds) ms" -ForegroundColor Green
        $content = $null
        try { $content = $result.choices[0].message.content } catch {}
        if ($null -ne $content) {
            Write-Host "模型回复：" -ForegroundColor Cyan
            if ($content -is [string]) {
                Write-Host $content
            } else {
                Write-Host ($content | ConvertTo-Json -Depth 8)
            }
        } else {
            Write-Host "接口返回成功，但未找到 choices[0].message.content。" -ForegroundColor Yellow
            Write-Host ($result | ConvertTo-Json -Depth 10)
        }
        if ($result.usage) {
            Write-Host ("Token 用量：prompt={0}, completion={1}, total={2}" -f `
                $result.usage.prompt_tokens, $result.usage.completion_tokens, $result.usage.total_tokens)
        }
        return $true
    } catch {
        $sw.Stop()
        Write-Host "聊天接口测试失败：$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "耗时：$($sw.ElapsedMilliseconds) ms" -ForegroundColor DarkGray
        return $false
    }
}

function Start-Chat {
    Write-Section "交互聊天（输入 /exit 退出，/clear 清空本轮上下文）"
    $messages = New-Object System.Collections.Generic.List[object]
    while ($true) {
        $line = Read-Host "你"
        if ($line -eq "/exit") { break }
        if ($line -eq "/clear") {
            $messages.Clear()
            Write-Host "已清空本轮上下文。" -ForegroundColor Yellow
            continue
        }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $messages.Add(@{ role = "user"; content = $line })
        $body = @{
            model = $script:Model
            messages = @($messages.ToArray())
            stream = $false
        }
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $result = Invoke-ApiRequest -Method POST -Url "$($script:BaseUrl)/chat/completions" -Body $body
            $sw.Stop()
            $answer = $result.choices[0].message.content
            if ($null -eq $answer) {
                Write-Host "未找到标准回复字段，原始结果：" -ForegroundColor Yellow
                Write-Host ($result | ConvertTo-Json -Depth 10)
                continue
            }
            if ($answer -isnot [string]) { $answer = $answer | ConvertTo-Json -Depth 8 }
            Write-Host ""
            Write-Host "AI（$($sw.ElapsedMilliseconds) ms）：" -ForegroundColor Green
            Write-Host $answer
            Write-Host ""
            $messages.Add(@{ role = "assistant"; content = $answer })
        } catch {
            $sw.Stop()
            Write-Host "请求失败：$($_.Exception.Message)" -ForegroundColor Red
            # Remove the user message that failed so a retry won't duplicate it.
            if ($messages.Count -gt 0 -and $messages[$messages.Count - 1].role -eq "user") {
                $messages.RemoveAt($messages.Count - 1)
            }
        }
    }
}

Clear-Host
Write-Host "LanChat - 内网模型助手" -ForegroundColor Cyan
Write-Host "Windows PowerShell 5.1+；无需编译、无需安装模块、无需外网。"
Write-Host "本脚本不会把 API Key、问题或回答写入文件。"
Write-Host "只会向你输入的 Base URL 发起请求；请确认该地址属于获准的内网服务。"

Write-Section "配置服务"
$rawUrl = Read-NonEmpty "输入 API Base URL（例如 http://10.20.30.40:8000/v1）"
try {
    $parsed = [Uri]$rawUrl
    if ($parsed.Scheme -notin @("http","https")) {
        throw "只支持 http 或 https 地址。"
    }
    if (-not $parsed.IsAbsoluteUri) { throw "请输入完整地址，包含 http:// 或 https://" }
    if ($parsed.Host -in @("localhost","127.0.0.1","::1") -or $parsed.Host -match "^127\.") {
        Write-Host "当前地址指向本机回环地址；请确认这就是模型服务所在机器。" -ForegroundColor Yellow
    }
} catch {
    Write-Host "地址格式错误：$($_.Exception.Message)" -ForegroundColor Red
    exit 2
}
$script:BaseUrl = Normalize-BaseUrl $rawUrl
$script:ApiKey = Read-ApiKey
$script:Model = Read-NonEmpty "输入模型 ID（如需查询，可先尝试模型列表接口）"

Write-Host ""
Write-Host "规范化 API Base URL：$($script:BaseUrl)" -ForegroundColor Cyan
Write-Host "不会自动访问外部网站，也不会自动发现外部模型。"

while ($true) {
    Write-Section "主菜单"
    Write-Host "1. 测试模型列表 GET /models"
    Write-Host "2. 测试聊天接口 POST /chat/completions"
    Write-Host "3. 进入交互聊天"
    Write-Host "4. 连续执行模型列表 + 聊天测试"
    Write-Host "5. 显示当前目标地址（不显示 API Key）"
    Write-Host "0. 退出"
    $choice = Read-Host "请选择"
    switch ($choice) {
        "1" { [void](Test-Models) }
        "2" { [void](Test-Chat) }
        "3" { Start-Chat }
        "4" {
            [void](Test-Models)
            [void](Test-Chat)
        }
        "5" {
            Write-Host "Base URL：$($script:BaseUrl)"
            Write-Host "Model ID：$($script:Model)"
            Write-Host "API Key：$(if ($script:ApiKey) { '[已设置，内容隐藏]' } else { '[未设置]' })"
        }
        "0" { break }
        default { Write-Host "无效选项。" -ForegroundColor Yellow }
    }
    if ($choice -eq "0") { break }
}
# Best-effort cleanup of the in-memory key reference.
$script:ApiKey = $null
Write-Host "已退出。" -ForegroundColor DarkGray

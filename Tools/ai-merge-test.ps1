<#
.SYNOPSIS
    AI 助手的回归测试：把 AI 改过的文字合并回带格式的段落、解析模型输出。

.DESCRIPTION
    不碰 OneNote，只通过反射调用构建出来的 DLL 里 RenderDiagnostics 的入口，普通权限即可运行。
    重点验证 RichParagraph：改动落在加粗、链接、多个 one:T 里时格式不丢，实体写法不变，
    插入的空格不会跟进加粗，拼不回原样的段落会被认出来。另外覆盖思考强度参数和流式返回（SSE）的解析。

    加 -Live 会用本机 %APPDATA%\OneNoteCodeHelper\ai-settings.xml 真调一次接口，打印模型的修改结果。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Tools\ai-merge-test.ps1 -Live -Model deepseek-v4-pro -Effort high
#>
[CmdletBinding()]
param(
    [ValidateSet('Release', 'Debug')]
    [string]$Configuration = 'Release',

    # 直接指定要测的 DLL，优先于 -Configuration
    [string]$DllPath,

    # 真调一次 AI 接口
    [switch]$Live,

    # 模型 id，要在 ai-settings.xml 的 Models 里
    [string]$Model = 'deepseek-v4-flash',

    [ValidateSet('none', 'low', 'medium', 'high', 'max')]
    [string]$Effort = 'none'
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
if (-not $DllPath) {
    $DllPath = Join-Path $root "bin\$Configuration\net48\OneNoteCodeHelper.dll"
}

if (-not (Test-Path $DllPath)) {
    Write-Error "找不到 $DllPath，请先构建。"
    exit 1
}

# 按字节加载，不锁 bin 下的 DLL，免得影响下一次构建
$assembly = [System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes((Resolve-Path $DllPath)))
$flags = [System.Reflection.BindingFlags]'NonPublic, Public, Static'
$diagnostics = $assembly.GetType('OneNoteCodeHelper.Services.RenderDiagnostics', $true)

function Invoke-Diag([string]$name, [object[]]$arguments) {
    try {
        $diagnostics.GetMethod($name, $flags).Invoke($null, $arguments)
    }
    catch [System.Reflection.TargetInvocationException] {
        throw $_.Exception.InnerException
    }
}

# 拼一个 one:OE，每个参数是一个 one:T 里的 HTML
function New-Paragraph([string[]]$runs) {
    $ts = ($runs | ForEach-Object { "<one:T><![CDATA[$_]]></one:T>" }) -join ''
    "<one:OE xmlns:one=`"http://schemas.microsoft.com/office/onenote/2013/onenote`" objectID=`"{TEST}`">$ts</one:OE>"
}

$script:total = 0
$script:failed = 0

function Assert-Equal([string]$name, $actual, $expected) {
    $script:total++
    if ($actual -ceq $expected) {
        Write-Host "  ok    $name"
        return
    }

    $script:failed++
    Write-Host "  FAIL  $name" -ForegroundColor Red
    Write-Host "        期望：$expected" -ForegroundColor Red
    Write-Host "        实际：$actual" -ForegroundColor Red
}

function Test-Merge([string]$name, [string[]]$runs, [string]$newText, [string]$expected) {
    $xml = New-Paragraph $runs
    Assert-Equal $name (Invoke-Diag 'MergeParagraph' @($xml, $newText)) $expected
}

Write-Host '合并：'

Test-Merge '纯文本改错字' @('今天去公圆玩。') '今天去公园玩。' '今天去公园玩。'

Test-Merge '加粗里的错字，改完仍是加粗' `
    @("今天去<span style='font-weight:bold'>公圆</span>玩。") '今天去公园玩。' `
    "今天去<span style='font-weight:bold'>公园</span>玩。"

Test-Merge '链接保留，链接内外都改' `
    @('看<a href="http://x.com/a&amp;b">文挡</a>的说明，有错字在这理。') '看文档的说明，有错字在这里。' `
    '看<a href="http://x.com/a&amp;b">文档</a>的说明，有错字在这里。'

Test-Merge '跨两个 one:T' `
    @('第一部分有错字在这理，', "<span style='color:#FF0000'>第二部分</span>也有措字。") `
    '第一部分有错字在这里，第二部分也有错字。' `
    "第一部分有错字在这里，|<span style='color:#FF0000'>第二部分</span>也有错字。"

Test-Merge '实体写法不变' @('A&amp;B&nbsp;用法&lt;T&gt;很好') 'A&B 用法<T>很棒' 'A&amp;B&nbsp;用法&lt;T&gt;很棒'

Test-Merge '新插入的 & < > 要转义' @('A和B') 'A&B<C>' 'A&amp;B&lt;C&gt;'

Test-Merge '排版：中英文之间加空格' @('使用GitHub管理代码，共10个项目') '使用 GitHub 管理代码，共 10 个项目' `
    '使用 GitHub 管理代码，共 10 个项目'

Test-Merge '排版：加粗词两边的空格不进加粗' `
    @("我用<span style='font-weight:bold'>Git</span>管理") '我用 Git 管理' `
    "我用 <span style='font-weight:bold'>Git</span> 管理"

Test-Merge '嵌套标签里改字' `
    @("<span style='color:red'>红<span style='font-weight:bold'>粗措</span>字</span>") '红粗错字' `
    "<span style='color:red'>红<span style='font-weight:bold'>粗错</span>字</span>"

Test-Merge '没改动时原样拼回' `
    @("<span style='color:red'>红<span style='font-weight:bold'>粗</span>字</span>尾&nbsp;巴") '红粗字尾 巴' `
    "<span style='color:red'>红<span style='font-weight:bold'>粗</span>字</span>尾&nbsp;巴"

Test-Merge '段内换行 <br> 保留' @('第一行<br>第二行有措字') "第一行`n第二行有错字" '第一行<br>第二行有错字'

Test-Merge '字删光的 one:T 去掉，至少留一个' @('abc', 'def') 'xyz' 'xyz'

Write-Host ''
Write-Host '段落文本：'

Assert-Equal '实体解码、硬空格归一' (Invoke-Diag 'ParagraphText' @((New-Paragraph @('A&amp;B&nbsp;C<br>D')))) "A&B C`nD"
Assert-Equal '注释拼不回去' (Invoke-Diag 'ParagraphText' @((New-Paragraph @('<!-- x -->文字')))) 'UNSUPPORTED'
Assert-Equal 'img 拼不回去' (Invoke-Diag 'ParagraphText' @((New-Paragraph @("<img src='a.png'>文字")))) 'UNSUPPORTED'
Assert-Equal '闭标签对不上' (Invoke-Diag 'ParagraphText' @((New-Paragraph @('文字</span>')))) 'UNSUPPORTED'

Write-Host ''
Write-Host '相似度与模型输出：'

$similar = Invoke-Diag 'TextSimilarity' @('今天去公圆玩。', '今天去公园玩。')
Assert-Equal '改一个字的相似度高于安全阀' ($similar -ge 0.6) $true
$different = Invoke-Diag 'TextSimilarity' @('今天天气很好，适合出去走走。', '完全不同的另一句话')
Assert-Equal '整段改写的相似度低于安全阀' ($different -lt 0.6) $true

Assert-Equal '解析：去代码块、id 是字符串也认' `
    (Invoke-Diag 'ParseAiReply' @("``````json`n{`"paragraphs`":[{`"id`":2,`"text`":`"b`"},{`"id`":`"1`",`"text`":`"a`"}]}`n``````")) `
    "1=a`n2=b"
Assert-Equal '解析：空结果' (Invoke-Diag 'ParseAiReply' @('{"paragraphs":[]}')) ''

Assert-Equal '去换行：英文之间补空格' (Invoke-Diag 'CleanAiReplyText' @('abc', "hello`nworld")) 'hello world'
Assert-Equal '去换行：中文直接接上' (Invoke-Diag 'CleanAiReplyText' @('中文', "第一`r`n第二")) '第一第二'
Assert-Equal '去换行：原文本来就有换行时保留' (Invoke-Diag 'CleanAiReplyText' @("a`nb", "a`nc")) "a`nc"

Write-Host ''
Write-Host '思考强度参数（参照 opencode 的 deepseek variants）：'

$none = Invoke-Diag 'DescribeRequestBody' @('none')
Assert-Equal 'none：thinking 关闭' $none.Contains('"thinking":{"type":"disabled"}') $true
Assert-Equal 'none：不传 reasoning_effort' $none.Contains('reasoning_effort') $false

foreach ($level in @('low', 'medium', 'high', 'max')) {
    $body = Invoke-Diag 'DescribeRequestBody' @($level)
    Assert-Equal "${level}：thinking 打开，reasoning_effort=$level" `
        ($body.Contains('"thinking":{"type":"enabled"}') -and $body.Contains("`"reasoning_effort`":`"$level`"")) $true
}

Assert-Equal '不认识的值归到 high' ((Invoke-Diag 'DescribeRequestBody' @('ultra')).Contains('"reasoning_effort":"high"')) $true
Assert-Equal '模型参数是 id' ((Invoke-Diag 'DescribeRequestBody' @('high')).Contains('"model":"deepseek-v4-flash"')) $true

Write-Host ''
Write-Host '流式返回：'

$body = Invoke-Diag 'DescribeRequestBody' @('high')
Assert-Equal '请求走流式，并要求带上用量' `
    ($body.Contains('"stream":true') -and $body.Contains('"stream_options":{"include_usage":true}')) $true

# 仿照网关实际返回的格式：先是角色，中间夹一行保活注释，思考、正文各分几段，最后一段带 finish_reason 和用量
function New-Chunk([string]$delta, [string]$finish = 'null', [string]$usage = 'null') {
    "data: {`"choices`":[{`"index`":0,`"delta`":$delta,`"finish_reason`":$finish}],`"usage`":$usage}"
}
$stream = @(
    (New-Chunk '{"role":"assistant","content":null,"reasoning_content":""}'),
    '',
    ': keep-alive',
    '',
    (New-Chunk '{"content":null,"reasoning_content":"先看"}'),
    (New-Chunk '{"content":null,"reasoning_content":"一下"}'),
    (New-Chunk '{"content":"{\"paragraphs\":","reasoning_content":null}'),
    (New-Chunk '{"content":"[]}","reasoning_content":null}'),
    (New-Chunk '{"content":""}' '"stop"' '{"prompt_tokens":9,"completion_tokens":5}'),
    '',
    'data: [DONE]'
) -join "`r`n"
Assert-Equal '拼出正文，数出思考字数，保活行忽略' (Invoke-Diag 'ParseAiStream' @($stream)) '{"paragraphs":[]}|stop|4|complete'

$cut = @(
    (New-Chunk '{"content":"{\"paragraphs\":","reasoning_content":null}'),
    (New-Chunk '{"content":"[{\"id\":1","reasoning_content":null}')
) -join "`n"
Assert-Equal '没有结束标记：认出是断在半路' (Invoke-Diag 'ParseAiStream' @($cut)) '{"paragraphs":[{"id":1||0|incomplete'

Assert-Equal '只有 [DONE] 也算完整' (Invoke-Diag 'ParseAiStream' @("data: {`"choices`":[{`"delta`":{`"content`":`"x`"}}]}`ndata: [DONE]")) 'x||0|complete'

Assert-Equal '流里报错' (Invoke-Diag 'ParseAiStream' @('data: {"error":{"message":"upstream timeout","type":"new_api_error"}}')) `
    'ERROR: AI 接口返回错误：upstream timeout'

Assert-Equal '进度：还没收到东西' (Invoke-Diag 'DescribeAiLive' @(0, 0)) '等待 AI 响应'
Assert-Equal '进度：思考中' (Invoke-Diag 'DescribeAiLive' @(120, 0)) 'AI 正在处理：已思考 120 字'
Assert-Equal '进度：开始输出' (Invoke-Diag 'DescribeAiLive' @(120, 35)) 'AI 正在处理：已思考 120 字，已输出 35 字'

if ($Live) {
    Write-Host ''
    Write-Host '真调接口：'
    $typo = "今天天气很好，我们一起去公圆玩。`n这段没有错误。`n他的成积在班里名列前茅，大家都很佩服他。"
    $layout = "我们使用GitHub管理代码,一共有10个项目。`n今天学习了javascript的闭包。"

    foreach ($case in @(@('错别字修复', $typo), @('排版优化', $layout))) {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-Diag 'RunAiSample' @($case[0], $Model, $Effort, $case[1])
        Write-Host ("{0}（{1:0.0}s）" -f $result, $watch.Elapsed.TotalSeconds)
        Write-Host ''
    }
}

Write-Host ''
if ($script:failed -eq 0) {
    Write-Host "全部通过：$($script:total) / $($script:total)" -ForegroundColor Green
    exit 0
}

Write-Host ("通过 {0} / {1}，失败 {2}" -f ($script:total - $script:failed), $script:total, $script:failed) -ForegroundColor Red
exit 1

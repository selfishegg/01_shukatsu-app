@echo off
setlocal DisableDelayedExpansion
set "JREIT_SELF=%~f0"
set "JREIT_TARGET=%~1"
powershell.exe -NoProfile -STA -Command "$ErrorActionPreference='Stop'; try { $s=[IO.File]::ReadAllText($env:JREIT_SELF,[Text.Encoding]::UTF8); $a=[regex]::Split($s,'(?m)^# JREIT_POWERSHELL\r?$'); if($a.Length -ne 2){throw 'Script marker missing.'}; & ([scriptblock]::Create($a[1])) } catch { $m=('ERROR: '+$_.Exception.Message+[Environment]::NewLine+'TYPE: '+$_.Exception.GetType().FullName+[Environment]::NewLine+'AT: '+$_.InvocationInfo.PositionMessage+[Environment]::NewLine+'STACK: '+$_.ScriptStackTrace); Write-Host $m; try { $d=Join-Path (Split-Path $env:JREIT_SELF) '02_output'; [void][IO.Directory]::CreateDirectory($d); $lf=Join-Path $d ('JREIT_error_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.txt'); [IO.File]::WriteAllText($lf,$m,(New-Object Text.UTF8Encoding -ArgumentList $true)); Write-Host ('LOG: '+$lf) } catch {}; exit 1 }"
set "JREIT_EXIT=%ERRORLEVEL%"
if not "%JREIT_EXIT%"=="0" echo Excel update failed. Read the error above. A successful save was not confirmed.
pause
exit /b %JREIT_EXIT%
# JREIT_POWERSHELL
$ErrorActionPreference = 'Stop'
Write-Host ('BUILD 2026-09-15f  |  running: ' + $env:JREIT_SELF)
$base = [IO.Path]::GetDirectoryName($env:JREIT_SELF)
$outDir = Join-Path $base '02_output'
$culture = [Globalization.CultureInfo]::InvariantCulture
$utf8 = New-Object Text.UTF8Encoding -ArgumentList $true
$missing = [Type]::Missing
$moneyFields = @('price','acquisition','disposition','noi','profit')
$dateFields = @('disclosed','scheduled','acquiredDate','disposedDate','built')
$aliases = @{
    disclosed = @('開示日時','開示日','開示年月日','公表日時','公表日','発表日時','発表日')
    code = @('コード','証券コード','銘柄コード','reitコード')
    issuer = @('投資法人','投資法人名','銘柄名','法人名','reit名')
    transaction = @('取引','取引区分','売買区分','取得譲渡','取得or譲渡','取得売却','種別')
    property = @('物件名','物件名称','物件','資産名称','資産名')
    address = @('物件の住所','物件住所','住所','所在地','物件所在地','所在地住居表示')
    scheduled = @('取得譲渡予定日','譲渡取得予定日','取得譲渡予定日時','譲渡取得予定日時','取引予定日')
    acquiredDate = @('取得予定日','取得日','取得予定日時','取得年月日')
    disposedDate = @('譲渡予定日','譲渡日','譲渡予定日時','譲渡年月日','売却予定日','売却日')
    price = @('金額','価格','取引価格','取得譲渡価格','取得譲渡予定価格','取得売却価格')
    acquisition = @('取得価格','取得予定価格','取得金額','取得予定金額','取得価額','取得予定価額')
    disposition = @('譲渡価格','譲渡予定価格','売却価格','売却予定価格','譲渡金額','売却金額')
    area = @('延床面積','延べ床面積','延べ面積','床面積','地上面積','建物延床面積')
    rentable = @('賃貸可能面積','賃貸可能床面積','総賃貸可能面積')
    built = @('竣工日','竣工年月','竣工年月日','竣工','建築日','建築年月','建築年月日')
    profit = @('純利益','当期純利益','不動産純利益')
    noi = @('noi','年間noi','noi年額','運営純収益','年間運営純収益')
    yield = @('noi利回り','noiyield','取得noi利回り')
    basis = @('収益指標の定義期間','収益指標の定義対象期間','noiの定義期間','収益指標の定義','対象期間')
    note = @('注記','備考','注釈')
    pdf = @('出典pdf','出典pdf名','pdf名','資料名','出典資料')
    page = @('出典ページ','ページ','資料ページ')
    url = @('出典url','url','資料url','開示url')
}

function Release-Com($object) {
    if ($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) } catch {}
    }
}
function Normalize-Header([string]$value) {
    $value = $value.Normalize([Text.NormalizationForm]::FormKC).ToLowerInvariant().Replace('讓','譲')
    return [regex]::Replace($value,'[\s_・/／:：]','')
}
function Field-Of([string]$header) {
    $key = Normalize-Header $header
    $key = [regex]::Replace($key,'\([^)]*\)|\[[^\]]*\]','')
    $key = [regex]::Replace($key,'百万円|千万円|億円|万円|千円|円|m2|㎡|%|jst|年額$','')
    foreach ($field in $aliases.Keys) { if ($aliases[$field] -contains $key) { return $field } }
    return 'other'
}
function Money-Unit([string]$header, [double]$fallback) {
    $h = $header.Normalize([Text.NormalizationForm]::FormKC)
    if ($h -match '百万円') { return 1000000.0 }
    if ($h -match '千万円') { return 10000000.0 }
    if ($h -match '億円') { return 100000000.0 }
    if ($h -match '万円') { return 10000.0 }
    if ($h -match '千円') { return 1000.0 }
    if ($h -match '円') { return 1.0 }
    return $fallback
}
function Column-Name([int]$number) {
    $name = ''
    while ($number -gt 0) {
        $number--
        $name = ([string][char](65 + ($number % 26))) + $name
        $number = [int][math]::Floor($number / 26)
    }
    return $name
}
function Empty-Value($value) {
    if ($null -eq $value) { return $true }
    if ($value -is [string]) { return ($value.Trim() -match '^(?:|null|n/?a|未開示|非開示|不明|記載なし|[-–—－]+)$') }
    return $false
}
function Number-Value($value, [string]$label) {
    $s = [Convert]::ToString($value,$culture).Normalize([Text.NormalizationForm]::FormKC).Trim()
    if ($s.Contains(',')) {
        if ($s -notmatch '^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$') { throw ('Invalid comma placement: ' + $label) }
        $s = $s.Replace(',','')
    }
    $n = 0.0
    if (-not [double]::TryParse($s,[Globalization.NumberStyles]::Float,$culture,[ref]$n) -or [double]::IsNaN($n) -or [double]::IsInfinity($n)) { throw ('Expected a number without unit text, or null: ' + $label) }
    return $n
}
function Date-Value($value) {
    if ($value -is [datetime]) { return [pscustomobject]@{ Value=$value; Format='yyyy-mm-dd hh:mm' } }
    $text = ([string]$value).Trim()
    foreach ($f in @('yyyy-MM-dd HH:mm:ss','yyyy-MM-dd HH:mm','yyyy-MM-dd','yyyy/M/d H:mm','yyyy/M/d','yyyy年M月d日')) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($text,$f,$culture,[Globalization.DateTimeStyles]::None,[ref]$d)) {
            $display = 'yyyy-mm-dd'
            if ($f.Contains('H')) { $display = 'yyyy-mm-dd hh:mm' }
            return [pscustomobject]@{ Value=$d; Format=$display }
        }
    }
    return [pscustomobject]@{ Value=$text; Format='@' }
}
function Extract-Json([string]$text) {
    $text = $text.Trim().TrimStart([char]0xFEFF)
    $fence = [regex]::Match($text,'(?s)```[ \t]*(?:json)?[ \t]*\r?\n(.*?)```')
    if ($fence.Success) { $text = $fence.Groups[1].Value }
    $start = $text.IndexOf('{')
    $end = $text.LastIndexOf('}')
    if ($start -lt 0 -or $end -le $start) {
        $head = $text.Substring(0,[math]::Min(120,$text.Length)) -replace '\s+',' '
        throw ('No JSON object was found in the copied text. It begins with: ' + $head)
    }
    return $text.Substring($start, $end - $start + 1)
}
function Read-Answer([string]$text) {
    $json = Extract-Json $text
    try { $data = ConvertFrom-Json -InputObject $json -ErrorAction Stop }
    catch {
        $fixed = $json -replace '[\u201C\u201D\u201E\u201F\u2033]','"' -replace '[\u2018\u2019]',"'"
        $fixed = $fixed -replace ',\s*(\}|\])','$1'
        try { $data = ConvertFrom-Json -InputObject $fixed -ErrorAction Stop }
        catch {
            $head = $json.Substring(0,[math]::Min(300,$json.Length)) -replace '\s+',' '
            $hint = '  ||  Hint: JSON values must be plain literals. Check the saved text for Math.round(), arithmetic expressions, NaN or undefined.'
            throw ($_.Exception.Message + $hint + '  ||  Extracted JSON begins with: ' + $head)
        }
    }
    if ($null -eq $data -or $data.PSObject.Properties.Name -notcontains 'rows' -or $data.rows -isnot [Array]) { throw 'Expected JSON with a rows array. Use option 1 to prepare a compatible Copilot prompt.' }
    if (@($data.rows).Count -gt 5000) { throw 'Split the answer into batches of at most 5000 properties.' }
    foreach ($row in $data.rows) {
        if ($row -isnot [pscustomobject]) { throw 'Each rows item must be an object.' }
        foreach ($p in $row.PSObject.Properties) {
            if ($null -ne $p.Value -and $p.Value -isnot [string] -and $p.Value -isnot [ValueType]) { throw ('Nested data is not supported: ' + $p.Name) }
            if ($p.Value -is [bool]) { throw ('Unexpected true/false value: ' + $p.Name) }
        }
    }
    return $data
}
function Convert-Record($row, [double]$fallback) {
    $fields = @{}
    $exact = @{}
    foreach ($p in $row.PSObject.Properties) {
        if (Empty-Value $p.Value) { continue }
        $key = Normalize-Header $p.Name
        $exact[$key] = $p.Value
        $field = Field-Of $p.Name
        if ($field -in @('other','yield')) { continue }
        $v = $p.Value
        if ($field -eq 'noi' -and $p.Name -match '月間|月額|半期|半年|四半期|6か月|6ヶ月') { throw 'Annual NOI is required. A partial-period NOI cannot be used as annual NOI.' }
        if ($field -in $moneyFields) { $v = (Number-Value $v $p.Name) * (Money-Unit $p.Name $fallback) }
        elseif ($field -in @('area','rentable')) {
            if ($p.Name -match '坪') { throw 'Area must be in square metres.' }
            $v = Number-Value $v $p.Name
        } elseif ($field -in $dateFields) { $v = Date-Value $v }
        else { $v = [string]$v }
        if ($fields.ContainsKey($field)) {
            $prior = $fields[$field]
            $equal = $false
            if ($field -in $dateFields) { $equal = ($prior.Value -eq $v.Value -and $prior.Format -eq $v.Format) }
            elseif ($field -in $moneyFields -or $field -in @('area','rentable')) { $equal = ([math]::Abs($prior-$v) -le 1e-10 * [math]::Max(1.0,[math]::Abs([double]$v))) }
            else { $equal = ([string]$prior -ceq [string]$v) }
            if (-not $equal) { throw ('Conflicting source values for the same field: ' + $p.Name) }
            continue
        }
        $fields[$field] = $v
    }
    if (-not $fields.ContainsKey('property')) { throw 'A property name is required in each rows item.' }
    if ($fields.ContainsKey('transaction')) {
        if ($fields.transaction -in @('取得','購入','買入','買付','acquisition')) { $fields.transaction = '取得' }
        elseif ($fields.transaction -in @('譲渡','讓渡','売却','disposition')) { $fields.transaction = '譲渡' }
    }
    # Never substitute the sale price for the acquisition-price denominator.
    if (-not $fields.ContainsKey('acquisition') -and $fields.transaction -eq '取得' -and $fields.ContainsKey('price')) { $fields.acquisition = $fields.price }
    if (-not $fields.ContainsKey('disposition') -and $fields.transaction -eq '譲渡' -and $fields.ContainsKey('price')) { $fields.disposition = $fields.price }
    if (-not $fields.ContainsKey('price')) {
        if ($fields.transaction -eq '取得' -and $fields.ContainsKey('acquisition')) { $fields.price = $fields.acquisition }
        elseif ($fields.transaction -eq '譲渡' -and $fields.ContainsKey('disposition')) { $fields.price = $fields.disposition }
    }
    if ($fields.transaction -eq '取得' -and $fields.ContainsKey('price') -and $fields.ContainsKey('acquisition') -and [math]::Abs($fields.price-$fields.acquisition) -gt 1e-10 * [math]::Max(1.0,[math]::Abs([double]$fields.acquisition))) { throw 'The transaction price and acquisition price disagree for an acquisition. Check the JSON units and source.' }
    if (-not $fields.ContainsKey('acquiredDate') -and $fields.transaction -eq '取得' -and $fields.ContainsKey('scheduled')) { $fields.acquiredDate = $fields.scheduled }
    if (-not $fields.ContainsKey('disposedDate') -and $fields.transaction -eq '譲渡' -and $fields.ContainsKey('scheduled')) { $fields.disposedDate = $fields.scheduled }
    if (-not $fields.ContainsKey('scheduled')) {
        if ($fields.transaction -eq '取得' -and $fields.ContainsKey('acquiredDate')) { $fields.scheduled = $fields.acquiredDate }
        elseif ($fields.transaction -eq '譲渡' -and $fields.ContainsKey('disposedDate')) { $fields.scheduled = $fields.disposedDate }
    }
    $sortDate = [datetime]::MinValue
    if ($fields.ContainsKey('disclosed') -and $fields.disclosed.Value -is [datetime]) { $sortDate = $fields.disclosed.Value }
    return [pscustomobject]@{ Fields=$fields; Exact=$exact; SortDate=$sortDate }
}
function Yield-Formula([int]$noiCol, [int]$priceCol, [double]$noiUnit, [double]$priceUnit) {
    $n = 'RC' + $noiCol
    $p = 'RC' + $priceCol
    $ratio = $n + '/' + $p
    if ($noiUnit -ne $priceUnit) { $ratio = $n + '*' + $noiUnit.ToString('0',$culture) + '/(' + $p + '*' + $priceUnit.ToString('0',$culture) + ')' }
    return '=IF(OR(' + $n + '="",' + $p + '="",' + $p + '=0),"",' + $ratio + ')'
}

Add-Type -AssemblyName System.Windows.Forms
# Capture the answer before the user copies an Excel path to the clipboard.
$initialClipboard = [Windows.Forms.Clipboard]::GetText([Windows.Forms.TextDataFormat]::UnicodeText)
if ($null -eq [type]::GetTypeFromProgID('Excel.Application')) { throw 'Windows desktop Excel is required.' }
Write-Host '1 = Prepare a Copilot prompt, then import the answer'
Write-Host '2 = Import the JSON answer copied BEFORE starting this BAT (default)'
$mode = Read-Host 'Select 1 or 2; Enter = 2'
if ($mode -eq '') { $mode = '2' }
if ($mode -notin @('1','2')) { throw 'Select 1 or 2.' }
$target = $env:JREIT_TARGET
if ([string]::IsNullOrWhiteSpace($target)) { $target = Read-Host 'Excel filename or full local path (close the workbook first)' }
$target = [Environment]::ExpandEnvironmentVariables($target.Trim().Trim([char]34))
if ([string]::IsNullOrWhiteSpace($target)) { throw 'An Excel filename or local path is required.' }
if ($target -match '^https?://') { throw 'Use the local OneDrive-synced path, not a web sharing link.' }
if (-not [IO.Path]::IsPathRooted($target)) { $target = Join-Path $base $target }
$target = [IO.Path]::GetFullPath($target)
if (-not [IO.File]::Exists($target)) { throw ('File not found: ' + $target) }
if ([IO.Path]::GetExtension($target).ToLowerInvariant() -notin @('.xlsx','.xlsm')) { throw 'Use an existing .xlsx or .xlsm workbook.' }
Write-Host 'Money units for headers/JSON keys without a unit: 1=JPY, 2=thousand JPY, 3=million JPY, 4=hundred million JPY'
$unitChoice = Read-Host 'Select the unit; Enter = 3 (million JPY)'
if ($unitChoice -eq '') { $unitChoice = '3' }
$units = @{'1'=1.0;'2'=1000.0;'3'=1000000.0;'4'=100000000.0}
if (-not $units.ContainsKey($unitChoice)) { throw 'Select a unit from 1 to 4.' }
$fallbackUnit = $units[$unitChoice]
[void][IO.Directory]::CreateDirectory($outDir)
$excel = $books = $book = $sheets = $sheet = $cells = $tables = $table = $autoCorrect = $null
$saved = $false
$restoreAutoSave = $false
$restoreAutoFill = $null
$oldCalculation = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $books = $excel.Workbooks
    $book = $books.Open($target,0,$false,$missing,'','',$true)
    if ($book.ReadOnly) { throw 'Workbook is read-only or open elsewhere. Close it and retry.' }
    try { $restoreAutoSave = [bool]$book.AutoSaveOn } catch {}
    if ($restoreAutoSave) { $book.AutoSaveOn = $false }
    $oldCalculation = $excel.Calculation
    $excel.Calculation = -4135
    $autoCorrect = $excel.AutoCorrect
    $restoreAutoFill = $autoCorrect.AutoFillFormulasInLists
    $autoCorrect.AutoFillFormulasInLists = $false
    $sheets = $book.Worksheets
    for ($i=1; $i -le $sheets.Count; $i++) {
        $s = $sheets.Item($i)
        try { Write-Host ($i.ToString() + ' = ' + $s.Name) } finally { Release-Com $s }
    }
    $selection = Read-Host 'Sheet number; Enter = 1'
    if ($selection -eq '') { $selection = '1' }
    $sheetIndex = 0
    if (-not [int]::TryParse($selection,[ref]$sheetIndex) -or $sheetIndex -lt 1 -or $sheetIndex -gt $sheets.Count) { throw 'Invalid sheet number.' }
    $sheet = $sheets.Item($sheetIndex)
    if ($sheet.ProtectContents) { throw 'The selected sheet is protected.' }
    $cells = $sheet.Cells
    $end = $cells.Item(1,16384)
    $last = $end.End(-4159)
    try { $lastCol = [int]$last.Column } finally { Release-Com $last; Release-Com $end }
    $headerRange = $sheet.Range('A1:' + (Column-Name $lastCol) + '1')
    try { if ($headerRange.MergeCells -ne $false) { throw 'Use unmerged headers in row 1.' } } finally { Release-Com $headerRange }
    $columns = @()
    $seenHeaders = @{}
    for ($c=1; $c -le $lastCol; $c++) {
        $cell = $cells.Item(1,$c)
        try { $h = [string]$cell.Value2 } finally { Release-Com $cell }
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $key = Normalize-Header $h
        if ($seenHeaders.ContainsKey($key)) { throw ('Duplicate header: ' + $h) }
        $seenHeaders[$key] = $true
        $field = Field-Of $h
        if ($field -eq 'noi' -and $h -match '月間|月額|半期|半年|四半期|6か月|6ヶ月') { throw 'Use an annual NOI column for NOI yield.' }
        if ($field -in @('area','rentable') -and $h -match '坪') { throw 'Area columns must use square metres.' }
        $columns += [pscustomobject]@{ Header=$h; Key=$key; Field=$field; Col=$c; Unit=(Money-Unit $h $fallbackUnit); New=$false }
    }
    if (@($columns | Where-Object { $_.Field -eq 'property' }).Count -ne 1) { throw 'Row 1 must contain one property-name header, for example 物件名.' }
    $tables = $sheet.ListObjects
    if ($tables.Count -gt 1) { throw 'Use a sheet containing one property table, not several separate tables.' }
    if ($tables.Count -eq 1) {
        $table = $tables.Item(1)
        $hr = $table.HeaderRowRange
        $tc = $table.ListColumns
        try {
            if ($null -eq $hr -or $hr.Row -ne 1 -or $hr.Column -ne 1 -or $tc.Count -ne $lastCol) { throw 'The Excel Table must start at A1 and cover the row-1 headers.' }
        } finally { Release-Com $hr; Release-Com $tc }
    }
    # Plan missing calculation columns before changing the workbook.
    foreach ($spec in @(@('acquisition','取得価格（百万円）'),@('noi','NOI（百万円／年）'),@('yield','NOI利回り'))) {
        $fieldMatches = @($columns | Where-Object { $_.Field -eq $spec[0] })
        if ($fieldMatches.Count -gt 1) { throw ('Keep one calculation column for: ' + $spec[1]) }
        if ($fieldMatches.Count -eq 0) {
            $lastCol++
            if ($lastCol -gt 16384) { throw 'No free Excel column remains.' }
            $columns += [pscustomobject]@{ Header=$spec[1]; Key=(Normalize-Header $spec[1]); Field=$spec[0]; Col=$lastCol; Unit=1000000.0; New=$true }
        }
    }
    $priceColumn = @($columns | Where-Object { $_.Field -eq 'acquisition' })[0]
    $noiColumn = @($columns | Where-Object { $_.Field -eq 'noi' })[0]
    $yieldColumn = @($columns | Where-Object { $_.Field -eq 'yield' })[0]
    if ($mode -eq '1') {
        $link = Read-Host 'OneDrive PDF/folder link (optional)'
        $exampleRow = [ordered]@{}
        foreach ($column in $columns) { if ($column.Field -ne 'yield') { $exampleRow[$column.Header] = $null } }
        if (@($columns | Where-Object { $_.Field -eq 'transaction' }).Count -eq 0) { $exampleRow['取引'] = $null }
        if (@($columns | Where-Object { $_.Field -eq 'disclosed' }).Count -eq 0) { $exampleRow['開示日時（JST）'] = $null }
        $schema = ConvertTo-Json -InputObject ([ordered]@{rows=@($exampleRow);files_read=@();files_unread=@();notes=@()}) -Depth 5
        $prompt = @'
指定したOneDriveのPDF本文から、J-REITの取得・譲渡情報を抽出してください。
次のJSON形式だけで回答してください。rowsは1物件・1取引・1開示につき1オブジェクトです。
見出し名（JSONキー）はそのまま使い、未開示はnull。JSONの改行の有無は問いません。
金額・面積・NOIはキーに記した単位のJSON数値にし、カンマや単位文字を値に付けないでください。
値は確定した数値リテラルか文字列だけにし、Math.round() などの関数や 500*1.05 のような計算式を書かないでください。
取得価格には取得した対価だけを入れ、譲渡価格・売却代金・鑑定評価額・簿価を代入しないでください。
金額などの汎用列は取引区分に沿う価格です。取引は「取得」または「譲渡」、コードは文字列です。
一括価格しか開示されない場合は物件別に推計配分せず、nullとし注記に総額を残してください。
NOIは年間運営純収益です。半期・月次の数値から年額を推計しないでください。
純利益とNOIは別の指標です。純利益にNOIや売却益、投資法人全体の利益を入れないでください。
NOIの鑑定値（実績／予想）の区分と対象期間、持分・全体の違いを注記してください。
NOIと取得価格は同じ物件・持分にそろえ、開示値だけを使ってください。
NOI利回りはExcel側で計算するので回答に含めないでください。
開示日時はJSTのYYYY-MM-DD HH:mm。時刻不明なら日付だけ。PDF発行日から開示時刻を作らないでください。
日付は元資料の精度のままYYYY-MM-DDまたはYYYY-MM。不明な月日を補わないでください。
延床面積と賃貸可能面積は区別し、面積は平方メートル。複数棟や分割取得は注記してください。
読めたPDFだけを対象にし、出典名・ページ・URLを確認できる範囲で残してください。
files_readとfiles_unreadに対象PDF名、notesに制約を文字列配列で記録してください。
フォルダ内の全件を読み取れない場合は、その制約をnotesに明示してください。
対象物件がなければrowsは空配列。形式見本の架空物件は作らないでください。
'@
        $prompt += "`r`n単位の記載がない金額キーは、1単位=" + $fallbackUnit.ToString('0',$culture) + "円として出力。NOIは年額。`r`n" + $schema + "`r`n対象リンク：" + $link
        [Windows.Forms.Clipboard]::SetText($prompt)
        Start-Process 'https://m365.cloud.microsoft/chat/'
        Write-Host 'Prompt copied. Select the PDFs in Copilot, paste the prompt, and send.'
        [void](Read-Host 'Copy the complete JSON code block, return HERE, then press Enter')
        $raw = [Windows.Forms.Clipboard]::GetText([Windows.Forms.TextDataFormat]::UnicodeText)
    } else { $raw = $initialClipboard }
    if ([string]::IsNullOrWhiteSpace($raw)) { throw 'No answer was captured. Copy the JSON answer BEFORE starting the BAT, or use option 1.' }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $textPath = Join-Path $outDir ('JREIT_copied_' + $stamp + '.txt')
    [IO.File]::WriteAllText($textPath,$raw,$utf8)
    Write-Host ('Answer text: ' + $textPath)
    $data = Read-Answer $raw
    $records = @($data.rows | ForEach-Object { Convert-Record $_ $fallbackUnit } | Sort-Object SortDate -Descending)
    if ($records.Count -eq 0) { Write-Host 'No property rows. The workbook was not changed.'; exit 0 }
    $maxDate = 'unknown'
    if ($records[0].SortDate -ne [datetime]::MinValue) { $maxDate = $records[0].SortDate.ToString('yyyyMMdd') }
    $finalText = Join-Path $outDir ('JREIT_maxdisclosure_' + $maxDate + '_saved_' + $stamp + '.txt')
    [IO.File]::Move($textPath,$finalText)
    $count = $records.Count
    $endRow = $count + 1
    $dateOffset = 0
    if ($book.Date1904) { $dateOffset = 1462 }
    # Prepare every value before inserting rows.
    $writes = @()
    foreach ($column in $columns) {
        if ($column.Field -eq 'yield') { continue }
        $values = New-Object 'object[,]' $count,1
        $formats = @()
        $hasValue = $false
        for ($r=0; $r -lt $count; $r++) {
            $record = $records[$r]
            $v = $null
            if ($column.Field -eq 'other') {
                if ($record.Exact.ContainsKey($column.Key)) { $v = [string]$record.Exact[$column.Key] }
            } elseif ($record.Fields.ContainsKey($column.Field)) { $v = $record.Fields[$column.Field] }
            if ($null -eq $v) { continue }
            $hasValue = $true
            if ($column.Field -in $moneyFields) { $v = [double]$v / $column.Unit }
            elseif ($column.Field -in $dateFields) {
                $formats += ,@(($r+2),$v.Format)
                if ($v.Value -is [datetime]) { $v = $v.Value.ToOADate() - $dateOffset } else { $v = $v.Value }
            }
            if ($v -is [string] -and $v.Length -gt 32767) { throw ('Text exceeds one Excel cell: ' + $column.Header) }
            $values[$r,0] = $v
        }
        if ($column.Field -ne 'other' -or $hasValue) { $writes += [pscustomobject]@{ Column=$column; Values=$values; Formats=$formats } }
    }
    $backupDir = Join-Path $outDir 'backup'
    [void][IO.Directory]::CreateDirectory($backupDir)
    $backup = Join-Path $backupDir ([IO.Path]::GetFileNameWithoutExtension($target) + '_before_' + $stamp + [IO.Path]::GetExtension($target))
    [IO.File]::Copy($target,$backup,$false)
    Write-Host ('Backup: ' + $backup)
    foreach ($column in $columns) {
        if (-not $column.New) { continue }
        if ($null -ne $table) {
            $listColumns = $table.ListColumns
            $added = $listColumns.Add()
            try { $added.Name = $column.Header } finally { Release-Com $added; Release-Com $listColumns }
        } else {
            $cell = $cells.Item(1,$column.Col)
            try { $cell.Value2 = $column.Header } finally { Release-Com $cell }
        }
    }
    if ($null -ne $table) {
        $listRows = $table.ListRows
        try { for ($i=0; $i -lt $count; $i++) { $added = $listRows.Add(1,$true); Release-Com $added } } finally { Release-Com $listRows }
    } else {
        $insertRows = $sheet.Range('2:' + $endRow)
        try { [void]$insertRows.Insert(-4121,1) } finally { Release-Com $insertRows }
    }
    foreach ($write in $writes) {
        $column = $write.Column
        $letter = Column-Name $column.Col
        $where = 'column "' + $column.Header + '" (' + $letter + ', field=' + $column.Field + ')'
        $stage = 'open range'
        $range = $null
        try {
            $range = $sheet.Range($letter + '2:' + $letter + $endRow)
            $stage = 'set range number format'
            $range.NumberFormat = '@'
            if ($column.Field -in $moneyFields -or $column.Field -in @('area','rentable')) { $range.NumberFormat = '#,##0.########' }
            $stage = 'write values (' + $write.Values.GetType().Name + ', rows=' + $write.Values.GetLength(0) + ')'
            $range.Value2 = $write.Values
        } catch {
            throw ('Excel write failed at ' + $where + ' | stage: ' + $stage + ' | ' + $_.Exception.Message)
        } finally { Release-Com $range }
        foreach ($format in $write.Formats) {
            $cell = $cells.Item([int]$format[0],$column.Col)
            try { $cell.NumberFormat = [string]$format[1] }
            catch { throw ('Excel write failed at ' + $where + ' | stage: cell number format row ' + $format[0] + ' | ' + $_.Exception.Message) }
            finally { Release-Com $cell }
        }
    }
    $formula = Yield-Formula $noiColumn.Col $priceColumn.Col $noiColumn.Unit $priceColumn.Unit
    $yieldLetter = Column-Name $yieldColumn.Col
    $yieldRange = $sheet.Range($yieldLetter + '2:' + $yieldLetter + $endRow)
    try { $yieldRange.NumberFormat = '0.00%'; $yieldRange.FormulaR1C1 = $formula } finally { Release-Com $yieldRange }
    $excel.Calculate()
    # Check actual Excel results before saving the original file.
    for ($r=0; $r -lt $count; $r++) {
        $cell = $cells.Item(($r+2),$yieldColumn.Col)
        try {
            if (-not $cell.HasFormula) { throw 'The NOI yield formula was not written.' }
            $actual = $cell.Value2
        } finally { Release-Com $cell }
        $f = $records[$r].Fields
        if ($f.ContainsKey('noi') -and $f.ContainsKey('acquisition') -and $f.acquisition -ne 0) {
            $expected = $f.noi / $f.acquisition
            if ($null -eq $actual -or $actual -is [string] -or [math]::Abs(([double]$actual)-$expected) -gt 1e-10 * [math]::Max(1.0,[math]::Abs([double]$expected))) { throw 'NOI yield validation failed. The workbook was not saved.' }
        } elseif ($null -ne $actual -and [string]$actual -ne '') { throw 'Missing-input NOI yield should be blank. The workbook was not saved.' }
    }
    $excel.Calculation = $oldCalculation
    $book.Save()
    $saved = $true
    Write-Host ('Updated: ' + $target)
    Write-Host ('Inserted ' + $count + ' property rows at row 2. Existing rows were moved down.')
    Write-Host ('NOI yield uses acquisition price: ' + $priceColumn.Header)
    Write-Host ('Answer text: ' + $finalText)
    Write-Host 'maxdisclosure is the latest date in this answer, not proof of complete coverage.'
    Write-Host ('Unread PDFs: ' + ($data.files_unread -join ', '))
    Write-Host ('Notes: ' + ($data.notes -join ' / '))
} finally {
    if ($null -ne $autoCorrect -and $null -ne $restoreAutoFill) { try { $autoCorrect.AutoFillFormulasInLists = $restoreAutoFill } catch {} }
    if ($null -ne $excel -and $null -ne $oldCalculation) { try { $excel.Calculation = $oldCalculation } catch {} }
    if ($null -ne $book) {
        if ($saved -and $restoreAutoSave) { try { $book.AutoSaveOn = $true } catch {} }
        try { $book.Close($false) } catch {}
    }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-Com $table
    Release-Com $tables
    Release-Com $cells
    Release-Com $sheet
    Release-Com $sheets
    Release-Com $book
    Release-Com $books
    Release-Com $autoCorrect
    Release-Com $excel
}
if ($saved) { try { Start-Process $target } catch { Write-Host 'Saved. Open the workbook manually.' } }
exit 0

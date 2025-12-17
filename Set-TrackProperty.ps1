# <CAUTION!>
# このファイルは UTF-8 (BOM 付き) で保存すること。(ターミナルに日本語メッセージを表示するため)
# </CAUTION!>

<#
    .SYNOPSIS
    トラックに対するプロパティ情報を定義した .csv ファイルを読み込んで iTunes のトラックに反映する
    
    .PARAMETER PropertySpecifierFile
    Property 指定 CSV ファイル
    ファイルパスを表す [System.String] またはファイルを表す [System.IO.FileInfo] を指定
    Note
     # エンコーディング
       - 文字列エンコーディングは BOM 付き UTF-8 であること  
     # 書式
         - 1行目はプロパティ指定のためのタイトルとして定義  
         - 2行目以降にトラック毎のプロパティに指定する情報を記載  
        ## 必須となる列
         - FilePath  
           トラックのファイルパス
        ## オプションとなる列
         - Name  
           トラック名  
         - Artist  
           アーティスト名  
         - Album  
           アルバム名  
         - AlbumArtist  
           アルバムアーティスト名  
         - Year  
           年  
         - Compilation  
           コンピレーションかどうか TODO書式
         - DiscNumber  
           ディスク番号 (1Base)  
         - DiscCount  
           ディスク番号(全体) (1Base)  
         - TrackNumber  
           トラック番号 (1Base)  
         - TrackCount  
           トラック番号(全体) (1Base)  
         - AddArtworkFromFile  
           ファイルからアルバムアートワークを追加する  
           空文字の場合は無視  
           Note: すでにアートワークが設定済みでも、新たに追加する  
         - SortAlbum  
           アルバム名の '読み'  
         - SortAlbumArtist  
           アルバムアーティスト名の '読み'  
         - SortArtist  
           アーティスト名の '読み'  
#>
[CmdletBinding()]
Param(    
    [Parameter(Mandatory=$true)][ValidateScript({
        if (($_ -isnot [System.String]) -and ($_ -isnot [System.IO.FileInfo])){ # 型は文字列か [FileInfo](https://learn.microsoft.com/ja-jp/dotnet/api/system.io.fileinfo?view=net-8.0) でないといけない
            return $false
        } else {
            return $true
        }
    })]$PropertySpecifierFile
)

# <CSV ファイルから読み取った文字列の書式チェック & 型変換>---------------------

#Note
# Private 関数としては定義しない
#     理由: `Add-Member` した時の `-Name` の名称で関数コールした場合、以下エラーが発生する為
#          「"用語 '～～～' は、コマンドレット、関数、スクリプト ファイル、または操作可能なプログラムの名前として認識されません。」

# 
# Property 指定 CSV ファイル内における FileInfo 系変換 (必須パラメーター用)
function ConvertTo-FileInfo_m{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$FilePath
    )
    #todo トラックが iTunes に追加不能なフォーマットの場合 e.g. .flac
    if (-Not (Test-Path -LiteralPath $FilePath -PathType Leaf)){ # パスが存在しない場合
        throw [System.Exception]::new(
            "FilePath: `"" + $FilePath + "`"`n" +
            "No such file."
        )
    }
    return (Get-Item -LiteralPath $FilePath) # FileInfo を取得して返す
}

# 
# Property 指定 CSV ファイル内における FileInfo 系変換 (Artwork 用)
function ConvertTo-ArtworkSpecifier{
    [CmdletBinding()]
    Param(
        [System.String]$FilePath
    )
    $strarr_allowed_ext = @(
        '.jpg',
        '.jpeg',
        '.png',
        '.bmp'
    )
    if ($FilePath -eq ""){
        $out_fileinfo = $null
    }else{
        if (-Not (Test-Path -LiteralPath $FilePath -PathType Leaf)){ # パスが存在しない場合
            throw [System.Exception]::new(
                "FilePath: `"" + $FilePath + "`"`n" +
                "No such file."
            )
        }

        $str_ext = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
        if (-Not ($str_ext -in $strarr_allowed_ext)){
            throw [System.Exception]::new(
                "FilePath: `"" + $FilePath + "`"`n" +
                "File is not image. (Expected ext: " + ($strarr_allowed_ext -Join ", ") + ", but specified " + $str_ext + ")"
            )
        }
        $out_fileinfo = (Get-Item -LiteralPath $FilePath) # FileInfo を取得
    }
    return $out_fileinfo
}

#
# Property 指定 CSV ファイル内における文字列系変換
function ConvertTo-String{
    [CmdletBinding()]
    Param(
        [System.String]$CharcterString
    )
    return $CharcterString
}

#
# Property 指定 CSV ファイル内における boolean 系変換
function ConvertTo-Bool{
    [CmdletBinding()]
    Param(
        [System.String]$Bool_in
    )
    try{
        $Bool_out = [bool]::Parse($Bool_in)
    }catch{
        throw [System.Exception]::new("Expected boolean, but specified `"" + $Bool_in + "`".")
    }

    return $Bool_out
}

#
# Property 指定 CSV ファイル内における整数値系変換
function ConvertTo-UInt64{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$Uint64_in
    )
    try{
        $Uint64_out = [System.UInt64]::Parse($Uint64_in)
    }catch{
        throw [System.Exception]::new("Expected unsigned int, but specified `"" + $Uint64_in + "`".")
    }
    return $Uint64_out
}
# --------------------</CSV ファイルから読み取った文字列の書式チェック & 型変換>

$obj_converterUtil = [PSCustomObject]@{
    FilePath = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-FileInfo_m($arg) }
    Name = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    Artist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    Album = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    AlbumArtist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    Year = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-UInt64($arg) }
    Compilation = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-Bool($arg) }
    DiscNumber = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-UInt64($arg) }
    DiscCount = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-UInt64($arg) }
    TrackNumber = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-UInt64($arg) }
    TrackCount = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-UInt64($arg) }
    AddArtworkFromFile = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-ArtworkSpecifier($arg) }
    SortAlbum = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    SortAlbumArtist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    SortArtist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
}
# <引数チェック>------------------------------------------------------------------------------------------

if($PropertySpecifierFile -is [System.String]){ # 文字列の場合

    # .csv ファイルを指定するパス文字列を想定
    if (
        (-Not (Test-Path -LiteralPath $PropertySpecifierFile -PathType Leaf)) -or # パスが存在しない場合
        (-Not ([System.IO.Path]::GetExtension($PropertySpecifierFile) -ieq '.csv')) # .csv ファイルではない場合
    ){
        Write-Error (
            "PropertySpecifierFile: `"" + $PropertySpecifierFile + "`"`n" +
            "No such .csv file."
        )
        Exit 1 # 異常終了
    }

    $PropertySpecifierFile = Get-Item -LiteralPath $PropertySpecifierFile # FileInfo を取得
}

# -----------------------------------------------------------------------------------------</引数チェック>

Import-Module "$PSScriptRoot\Yakenohara.COM.iTunes.Track" -Force

# <CSV ファイルの読込>--------------------------------------------------------------------------------
try{
    # CSV ファイル読込
    # https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-7.5
    # Note
    # 1行のみの場合 -> 返却値は単一の System.Management.Automation.PSCustomObject
    # 2行以上の場合 -> 返却値はオブジェクトの配列 System.Object[]
    $objarr_rows = Import-Csv `
        -LiteralPath $PropertySpecifierFile.FullName `
        -WarningAction Stop `
        3> $null
}catch{
    Write-Error (
        "Cannot read property specifier file `"" + $PropertySpecifierFile.FullName + "`".`n" + 
        $_.ToString()
    )
    Exit 1 # 終了
}

# `Import-Csv` の返却値が単一の System.Management.Automation.PSCustomObject だった場合は、配列に変換
#     理由: 行毎のイテレーションで、配列要素のインデックス No を明示的に意識した for 文でアクセスしたい為
if ($objarr_rows -is [System.Management.Automation.PSCustomObject]) { # オブジェクトの場合 -> 1 行のみの場合
    [System.Object[]]$objarr_rows = @($objarr_rows[0]) # 配列に変換
}

$str_keys = $obj_converterUtil.PSObject.Properties.Name

# 未定義の列名が使用されていないかどうか確認
$strarr_tmp = $objarr_rows[0].PSObject.Properties.Name
for ($uint_idxOfTmp = 0 ; $uint_idxOfTmp -lt $strarr_tmp.Count ; $uint_idxOfTmp++){
    if (-Not ($strarr_tmp[$uint_idxOfTmp] -in $str_keys)) { # 未定義の列名を使用していた場合
        Write-Error (
            "Unkown property name `"" + $strarr_tmp[$uint_idxOfTmp] + "`" specified. (Column: " + ($uint_idxOfTmp + 1).ToString() + ")`n" +
            "Property specifier file: `"" + $PropertySpecifierFile.FullName + "`""
        )
        Exit 1
    }
}
# --------------------------------------------------------------------------------<CSV ファイルの読込>

# <Property の型チェック>-----------------------------------------------------------------------------
$hasharr_splatter = [System.Collections.Generic.List[System.Collections.Hashtable]]::new()

for ($j = 0 ; $j -lt $objarr_rows.Count ; $j++){ # 行毎ループ
    $hash_splatter = @{}
    for ($i = 0 ; $i -lt $strarr_tmp.Count ; $i++){ # 列毎ループ
        try{
            $hash_splatter[$strarr_tmp[$i]] = $obj_converterUtil.($strarr_tmp[$i]).fnc_converter($objarr_rows[$j].($strarr_tmp[$i]))
        }catch{
            throw [System.Exception]::new(
                "Following error occured while reading `"" + $PropertySpecifierFile.FullName + "`", Row: " + ($j + 2).ToString() + ", Column: " + ($i + 1).ToString() + " (Title: `"" + $strarr_tmp[$i] + "`")`n" +
                $_.Exception.InnerException.Message
            )
            exit 1
        }
    }
    $hasharr_splatter.Add($hash_splatter)
}
# -----------------------------------------------------------------------------<Property の型チェック>


# トラック毎プロパティ設定
for ($int_idxOfRow = 0 ; $int_idxOfRow -lt $hasharr_splatter.Count ; $int_idxOfRow++){

    # ハッシュのキー名 `FilePath` -> `Track` 変換
    $finfo_trackFile = $hasharr_splatter[$int_idxOfRow].FilePath
    $hasharr_splatter[$int_idxOfRow].Remove("FilePath")
    $hasharr_splatter[$int_idxOfRow].Track = $finfo_trackFile

    $hash_splatter = $hasharr_splatter[$int_idxOfRow]
    try{
        Set-IITTrackProperty @hash_splatter -ValidateOnly | Out-Null #
    }catch{
        throw [System.Exception]::new(
            "Following error occured while reading `"" + $PropertySpecifierFile.FullName + "`", Row: " + ($int_idxOfRow + 2).ToString() + "`n" + `
            $_.Exception.Message
        )
        exit 1
    }
    Set-IITTrackProperty @hash_splatter | Out-Null # トラックのプロパティを設定
}

Write-Host "Done!"

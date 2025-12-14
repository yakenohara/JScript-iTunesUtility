# <CAUTION!>
# このファイルは UTF-8 (BOM 付き) で保存すること。(ターミナルに日本語メッセージを表示するため)
# </CAUTION!>

#
# NOTE
#  - .csv ファイル内のエンコーディングが BOM 付き UTF-8 となっているかどうかの確認はしない (実装が大変な為)  
#  - .csv ファイルの書式
#

<#
    .SYNOPSIS
    
    
    .PARAMETER PropertySpecifierFile
    

    

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
function ConvertTo-FileInfo_m{
<#
    .SYNOPSIS
    Track の Property 指定 CSV ファイル内におけるファイルパス指定用チェック & 型変換 (必須指定)

    .DESCRIPTION
    以下の場合は例外 [System.Exception] を throw する
     - 引数 `FilePath` がファイルを表さない
    
    .PARAMETER FilePath
    トラックのファイルパス

    .OUTPUTS
    [System.IO.FileInfo] トラックのファイル
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$FilePath
    )
    
    if (-Not (Test-Path -LiteralPath $FilePath -PathType Leaf)){ # パスが存在しない場合
        throw [System.Exception]::new(
            "FilePath: `"" + $FilePath + "`"`n" +
            "No such file."
        )
    }

    # FileInfo を取得して返す
    return (Get-Item -LiteralPath $FilePath)
}

function ConvertTo-String{
<#
    .SYNOPSIS
    Track の Property 指定 CSV ファイル内における文字列指定用チェック (必須指定)
    
    .PARAMETER CharcterString
    文字列

    .OUTPUTS
    [System.IO.FileInfo] or $null 文字列。引数 `CharcterString` が指定されない場合は、 $null
#>
    [CmdletBinding()]
    Param(
        [System.String]$CharcterString
    )
    
    # FileInfo を取得して返す
    return $CharcterString
}

function ConvertTo-1to999{
<#
    .SYNOPSIS
    Track の Property 指定 CSV ファイル内におけるトラック No 指定用チェック & 型変換

    .DESCRIPTION
    以下の場合は例外 [System.Exception] を throw する
     - 引数 `TrackNo` が 1-999 範囲内の整数ではない
    
    .PARAMETER TrackNo
    トラック No
    
    .OUTPUTS
    [System.UInt64] トラック No
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$TrackNo
    )

    try{
        $TrackNo_casted = [System.UInt64]::Parse($TrackNo)
    }catch{ # キャストエラー (-> 1 以上の整数値でない)
        throw [System.Exception]::new(
            "Expected unsigned int, but specified `"" + $TrackNo + "`""
        )
    }
    if(
        ($TrackNo_casted -lt 1) -or
        (999 -lt $TrackNo_casted)
    ){
        throw [System.Exception]::new(
            "Expected Range 1 to 999, but specified " + $TrackNo
        )
    }

    return $TrackNo_casted
}

function ConvertTo-TrackNo{
<#
    .SYNOPSIS
    Track の Property 指定 CSV ファイル内におけるトラック No 指定用チェック & 型変換

    .DESCRIPTION
    以下の場合は例外 [System.Exception] を throw する
     - 引数 `TrackNo` が 1-999 範囲内の整数ではない
    
    .PARAMETER TrackNo
    トラック No
    
    .OUTPUTS
    [System.UInt64] トラック No
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$TrackNo
    )

    try{
        $TrackNo_casted = [System.UInt64]::Parse($TrackNo)
    }catch{ # キャストエラー (-> 1 以上の整数値でない)
        throw [System.Exception]::new(
            "Expected unsigned int, but specified `"" + $TrackNo + "`""
        )
    }
    if(
        ($TrackNo_casted -lt 1) -or
        (999 -lt $TrackNo_casted)
    ){
        throw [System.Exception]::new(
            "Expected Range 1 to 999, but specified " + $TrackNo
        )
    }

    return $TrackNo_casted
}
# --------------------</CSV ファイルから読み取った文字列の書式チェック & 型変換>

$obj_converterUtil = [PSCustomObject]@{
    FilePath = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-FileInfo($arg) }
    Name = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-String($arg) }
    Artist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    Album = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    AlbumArtist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    Year = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    Compilation = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    DiscNumber = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    DiscCount = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    TrackNumber = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    TrackCount = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    AddArtworkFromFile = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    SortAlbum = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    SortAlbumArtist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
    SortArtist = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_converter -Value { Param($arg) ConvertTo-1to999($arg) }
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

function Private:Set-IITTrackProperty{
<#
    .SYNOPSIS
    指定された名称がプレイリスト名として存在しないことを確認する

    .PARAMETER Track
    プロパティ指定対象のトラック

    .PARAMETER Name
    トラック名を設定する
    使用する Property or Method : `HRESULT IITObject::Name  (  [in] BSTR  name   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITObject.html#z49_1</SDKREF>
    .PARAMETER Artist
    アーティスト名を設定する
    使用する Property or Method : `HRESULT IITTrack::Artist  (  [in] BSTR  artist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_5</SDKREF>
    .PARAMETER Album
    アルバム名を設定する
    使用する Property or Method : `HRESULT IITTrack::Album  (  [in] BSTR  album   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_3</SDKREF>
    .PARAMETER AlbumArtist
    アルバムアーティスト名を設定する
    使用する Property or Method : `HRESULT IITFileOrCDTrack::AlbumArtist  (  [in] BSTR  albumArtist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_25</SDKREF>
    .PARAMETER Year
    年を設定する
    使用する Property or Method : `HRESULT IITTrack::Year  (  [in] long  year   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_52</SDKREF>
    .PARAMETER Compilation
    コンピレーションかどうかを設定する true : 'コンピレーション' , false : 'コンピレーションではない'
    使用する Property or Method : `HRESULT IITTrack::Compilation  (  [in] VARIANT_BOOL  shouldBeCompilation   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_12</SDKREF>
    .PARAMETER DiscNumber
    ディスク番号を設定する (1Base)
    使用する Property or Method : `HRESULT IITTrack::DiscNumber  (  [in] long  discNumber   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_19</SDKREF>
    .PARAMETER DiscCount
    ディスク番号(全体)を設定する (1Base)
    使用する Property or Method : `HRESULT IITTrack::DiscCount  (  [in] long  discCount   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_17</SDKREF>
    .PARAMETER TrackNumber
    トラック番号を設定する (1Base)
    使用する Property or Method : `HRESULT IITTrack::TrackNumber  (  [in] long  trackNumber   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_48</SDKREF>
    .PARAMETER TrackCount
    トラック番号(全体)を設定する (1Base)
    使用する Property or Method : `HRESULT IITTrack::TrackCount  (  [in] long  trackCount   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z77_46</SDKREF>
    .PARAMETER AddArtworkFromFile
    ファイルからアルバムアートワークを設定する
    使用する Property or Method : `HRESULT IITTrack::AddArtworkFromFile  (  [in] BSTR  filePath,    [out, retval] IITArtwork **  iArtwork  )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html#z75_2</SDKREF>
    .PARAMETER SortAlbum
    アルバム名の '読み' を設定する
    使用する Property or Method : `HRESULT IITFileOrCDTrack::SortAlbum  (  [in] BSTR  album   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_39</SDKREF>
    .PARAMETER SortAlbumArtist
    アルバムアーティスト名の '読み' を設定する
    使用する Property or Method : `HRESULT IITFileOrCDTrack::SortAlbumArtist  (  [in] BSTR  albumArtist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_41</SDKREF>
    .PARAMETER SortArtist
    アーティスト名の '読み' を設定する
    使用する Property or Method : `HRESULT IITFileOrCDTrack::SortArtist  (  [in] BSTR  artist   )`
                                 <SDKREF>iTunesCOM.chm::/interfaceIITFileOrCDTrack.html#z81_43</SDKREF>

    .OUTPUTS
    [System.Boolean] 指定された名称がプレイリスト名として存在しない場合: $true, 存在する場合: $false
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.__ComObject]$Track,
        [System.String]$Name
    )
    $Track.Name = $Name

}

#Note
# 使っていない関数。
# すでに存在するプレイリスト名を使って新しくプレイリストを作成しても、作成は成功するので、本関数は不要。
function Private:Test-IsInvalidPlayListName{
<#
    .SYNOPSIS
    指定された名称がプレイリスト名として存在しないことを確認する

    .PARAMETER Suggestion
    確認したい名称

    .PARAMETER iTunesObj
    "iTunes.Application" の COM オブジェクト

    .OUTPUTS
    [System.Boolean] 指定された名称がプレイリスト名として存在しない場合: $true, 存在する場合: $false
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$Suggestion,
        [Parameter(Mandatory=$true)][System.__ComObject]$iTunesObj
    )

    $IITPlaylistCollection_playLists = $iTunesObj.LibrarySource.Playlists # プレイリストの取得
                                                                            # <SDKREF>iTunesCOM.chm::/interfaceIITPlaylistCollection.html</SDKREF>

    $bl_isInvalidPlayListName = $true

    # プレイリスト毎ループ
    for($int_idxOfPlayLists = 1 ; $int_idxOfPlayLists -le $IITPlaylistCollection_playLists.Count; $int_idxOfPlayLists++ ){
        $IITPlaylist_playList = $IITPlaylistCollection_playLists.Item($int_idxOfPlayLists)  # プレイリストオブジェクトを取得
                                                                                            # <SDKREF>iTunesCOM.chm::/interfaceIITPlaylist.html</SDKREF>

        if ($IITPlaylist_playList.Name -eq $Suggestion){ # プレイリスト名が一致した場合
            return $false 
        }
    }

    # すべてのプレイリスト名と '一致しない' 事を確認したので true を返却
    return $true
}

# iTunesObject生成
$str_comobjName_iTunes = "iTunes.Application"
try{
    $comobj_iTunes = New-Object -ComObject $str_comobjName_iTunes
}catch{
    Write-Error (
        "Cannot create COM Object `"" + $str_comobjName_iTunes + "`".`n" + 
        $_.ToString()
    )
    Exit 1 # 終了
}

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

# 作業用一時プレイリストの作成
$str_playListName = "0_temporary"
$IITPlaylist_playList = $comobj_iTunes.CreatePlaylist($str_playListName)   # プレイリストの作成
                                                                           # <SDKREF>iTunesCOM.chm::/interfaceIiTunes.html#z5_2</SDKREF>

# トラック毎プロパティ設定
for ($int_idxOfRow = 0 ; $int_idxOfRow -lt $hasharr_splatter.Count ; $int_idxOfRow++){

    $finfo_trackFile = $hasharr_splatter[$int_idxOfRow].FilePath
    $hasharr_splatter[$int_idxOfRow].Remove("FilePath") # `Set-IITTrackProperty` をコールする際には不要なので削除

    $IITOperationStatus_status_trackAdding = $comobj_iTunes.LibraryPlaylist.AddFile($finfo_trackFile) # トラックを追加
    while ($IITOperationStatus_status_trackAdding.InProgress) { # 追加完了まで待機
        Start-Sleep -Milliseconds 10
    }
    # Write-Host $IITOperationStatus_status_trackAdding.Tracks.Item(1).Name
    $IITTrack_track_added = $IITOperationStatus_status_trackAdding.Tracks.Item(1) # 追加したトラックの `IITTrack` オブジェクトを取得
    $hasharr_splatter[$int_idxOfRow].Track = $IITTrack_track_added # `Set-IITTrackProperty` をコールする際にのスプラッターに追加

    $hash_splatter = $hasharr_splatter[$int_idxOfRow]
    Set-IITTrackProperty @hash_splatter # トラックのプロパティを設定

}

Write-Host "Done!"

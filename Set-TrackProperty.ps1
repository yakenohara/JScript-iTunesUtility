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

[System.String[]]$methodNames = @(
    "FilePath",
    "Name",
    "Artist",
    "Album",
    "AlbumArtist",
    "Year",
    "Compilation",
    "DiscNumber",
    "DiscCount",
    "TrackNumber",
    "TrackCount",
    "AddArtworkFromFile",
    "SortAlbum",
    "SortAlbumArtist",
    "SortArtist"
)

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
# --------------------------------------------------------------------------------<CSV ファイルの読込>

write-host $objarr_rows[0].PSObject.Properties.Name

write-host "ok"
return

# 作業用一時プレイリストの作成
$str_playListName = "0_temporary"
$str_playListName = "Drive"
$IITPlaylist_playList = $comobj_iTunes.CreatePlaylist($str_playListName)   # プレイリストの作成
                                                                           # <SDKREF>iTunesCOM.chm::/interfaceIiTunes.html#z5_2</SDKREF>

# トラック毎プロパティ設定
for ($int_idxOfRow = 0 ; $int_idxOfRow -le $objarr_rows.Count ; $int_idxOfRow++){
    write-host $objarr_rows[$int_idxOfRow]

}

# 作業用一時プレイリストの削除
$IITPlaylist_playList.Delete() # <SDKREF>iTunesCOM.chm::/interfaceIITPlaylist.html#z51_0</SDKREF>

Write-Host "Done!"

# <CSV ファイルから読み取った文字列の書式チェック & 型変換>---------------------

#Note
# Private 関数としては定義しない
#     理由: `Add-Member` した時の `-Name` の名称で関数コールした場合、以下エラーが発生する為
#          「"用語 'ConvertTo-1to999' は、コマンドレット、関数、スクリプト ファイル、または操作可能なプログラムの名前として認識されません。」

function ConvertTo-1to999{
<#
    .SYNOPSIS
    
    
    .PARAMETER PropertySpecifierFile
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)][System.String]$x
    )
    try{
        $y = [System.UInt64]::Parse($x)
    }catch{ # キャストエラー (-> 1 以上の整数値でない)
        throw [System.Exception]::new(
            "Expected unsigned int, but specified `"" + $x + "`""
        )
    }

    if(
        ($y -lt 1) -or
        (999 -lt $y)
    ){
        throw [System.Exception]::new(
            "Expected Range 1 to 999, but specified " + $x
        )
    }

    return $y
}
# --------------------</CSV ファイルから読み取った文字列の書式チェック & 型変換>


function fnc_b{
    Param(
        $x
    )
    return ($x * 2)
}

function fnc_x{
    Param(
        $FilePath,
        $Name
    )
    write-host ("FilePath: " + $FilePath + ", Name: " + $Name)
}



$obj = [PSCustomObject]@{
    FilePath = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_checker -Value { Param($arg) ConvertTo-1to999($arg) }
    Name = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_checker -Value { Param($arg) fnc_b($arg) }
}

$path = "E:\git\yakenohara\0-Whole-Projects\projects\JScript-iTunesUtility\test_utf8.csv"

$hash = @{}

$names = $obj.PSObject.Properties.Name
$values = $obj.PSObject.Properties.Value

try{
    for ($i = 0 ; $i -lt $names.Count ; $i++){
        $hash[$names[$i]] = $values[$i].fnc_checker("0")
    }
}catch{
    throw [System.Exception]::new(
        "Following error occured while reading `"" + $path + "`", Row: " + (0 + 2).ToString() + ", Column: " + ($i + 1).ToString() + " (Title: `"" + $names[$i] + "`")`n" +
        $_.Exception.InnerException.Message
    )
    exit 1
}

fnc_x @hash

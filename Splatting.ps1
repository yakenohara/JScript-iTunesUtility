function fnc_a{
    Param(
        $x
    )
    return ($x + 1)
}
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
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_checker -Value { Param($arg) fnc_a($arg) }
    Name = `
        [PSCustomObject]@{} | Add-Member -PassThru ScriptMethod -Name fnc_checker -Value { Param($arg) fnc_b($arg) }
}

$hash = @{}

$names = $obj.PSObject.Properties.Name
$values = $obj.PSObject.Properties.Value

for ($i = 0 ; $i -lt $names.Count ; $i++){
    $hash[$names[$i]] = $values[$i].fnc_checker(2)
}

fnc_x @hash

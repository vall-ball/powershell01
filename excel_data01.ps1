#функция нахождения даты получения паспорта
function Get-PassportDate {
    param (
        [string]$BirthDateString,
        [string]$FirstStudyingDayString
    )
    $BirthDate = Get-Date -Date $BirthDateString 
    $FirstStudyingDay = Get-Date -Date $FirstStudyingDayString 
    #случайная задержка от 10 до 25 дней
    $Delay =  Get-Random -Maximum 25 -Minimum 10
    $Difference = $FirstStudyingDay - $BirthDate
    $PassportDate = $null
    
    if (($Difference.Days + $Delay) / 365 -lt 20) {
        $PassportDate = $BirthDate.AddYears(14).AddDays($Delay)
    } elseif (($Difference.Days + $Delay) / 365 -lt 45) {
        $PassportDate = $BirthDate.AddYears(20).AddDays($Delay)
    } else {
        $PassportDate = $BirthDate.AddYears(45).AddDays($Delay)
    }
    while (Is-Holiday($PassportDate)) {
        $PassportDate = $PassportDate.AddDays(1)
    }
    return $PassportDate   
}

#функция проверки даты на праздичность
function Is-Holiday{
    param (
        [DateTime]$PassportDate
    )
    #список праздничных дат
    $Holidays = @(
    (Get-Date -Date '23.02.2022'),
    (Get-Date -Date '08.03.2022'),
    (Get-Date -Date '12.06.2022'),
    (Get-Date -Date '04.11.2022'),
    (Get-Date -Date '31.12.2022')
)
    if ($PassportDate.Month -eq 1 -and $PassportDate.Day -ge 1 -and $PassportDate.Day -le 14) {
        return $true   
    }
    if ($PassportDate.Month -eq 5 -and $PassportDate.Day -ge 1 -and $PassportDate.Day -le 10) {
        return $true   
    }
    if ($PassportDate.DayOfWeek -eq [DayOfWeek].GetEnumNames()[0] -or $PassportDate.DayOfWeek -eq [DayOfWeek].GetEnumNames()[6]) {
        return $true
    }
    foreach ($d in $Holidays) {
        if ($d.Day -eq $PassportDate.Day -and $d.Month -eq $PassportDate.Month) {
            return $true
        }
    }
    return $false
}


#функция нахождения последних 2-х цифр серии
function Get-LastNumbersOfSeria{
    param (
        [DateTime]$PassportDate
    )
  return ([string]($PassportDate.Year)).Substring(2)
}


#функция нахождения первых 2-х цифр серии
function Get-FirstNumbersOfSeria{
    param (
        [String]$City,
        [Hashtable]$hash
    )
    Write-Output $City
    Write-Output $hash
    $answer = $hash[$City]
    return $answer
}


#функция создания словаря регионов и их номеров
function Create-MapOfRegions{
    $csv=Import-CSV -path D:\coding\workplace\pshell\mapa.csv
    $hash = @{}
    foreach ($o in $csv) {
        $hash.add($o.region, $o.number)
    }
    return $hash

}

#функция генерации случайного номера от 000101 до 999999
function Create-RandomNumber{
    $number =  Get-Random -Maximum 999999 -Minimum 101
    $answer = [String]$number
    if ($answer.Length -eq 4) {
        return '00' + $answer
    }
    if ($answer.Length -eq 5) {
        return '0' + $answer
    }
    return $answer
}


#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$WorkBook = $Excel.Workbooks.Open("D:\coding\workplace\pshell\dlavala.xlsx")
$WorkSheet = $WorkBook.Sheets('1ч) для договоров для Вал')
$map = Create-MapOfRegions


#Внесение измеений в файл
for ($i = 1; $i -le $WorkSheet.UsedRange.Columns['B'].rows.Count; $i++) {
    #запись даты выдачи паспорта
    if ($WorkSheet.Columns['K'].rows[$i].Text.Length -eq 0) {
        $PassportDate = Get-PassportDate $WorkSheet.UsedRange.Columns['E'].rows[$i].text $WorkSheet.UsedRange.Columns['M'].rows[$i].text
        $date = $PassportDate.Date.ToString().Substring(0, 10)
        $WorkSheet.Columns['K'].rows[$i] = ([string]$date)
    }
    #запись серии паспорта
    if ($WorkSheet.Columns['H'].rows[$i].Text.Length -eq 0) {
        $s1 = Get-FirstNumbersOfSeria $WorkSheet.UsedRange.Columns['G'].rows[$i].text $map
        $s2 = Get-LastNumbersOfSeria $PassportDate
        $Seria = $s1[2] + $s2
        $WorkSheet.Columns['H'].rows[$i] = ([String]$Seria)
    }
    #запись номера паспорта
    if ($WorkSheet.Columns['I'].rows[$i].Text.Length -eq 0) {
        $WorkSheet.Columns['I'].rows[$i] = Create-RandomNumber
    }

}

#Выход из Excel
$WorkBook.Save()
$WorkBook.close($true)
$Excel.Quit()
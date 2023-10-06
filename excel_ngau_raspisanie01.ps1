#функция генерации списка дат обучения (теория, практика, консультация и экзамен БЕЗ выходных)
function Get-ListOfDates {
    param (
        [string]$BeginDate,
        [int]$QuantityOfDaysOfTeory,
        [int]$QuantityOfDaysOfPractice
    )
    $answer = [System.Collections.ArrayList]::new()
    $FirstDay = Get-Date -Date $BeginDate
    if (-not (Is-Holiday($FirstDay))) {
        $QuantityOfDaysOfTeory--
        $null = $answer.Add(@($FirstDay, 'T'))
    }
    while ($QuantityOfDaysOfTeory -ne 0) {
        $FirstDay = $FirstDay.AddDays(1)
        if (-not (Is-Holiday($FirstDay))) {  
            $QuantityOfDaysOfTeory--
            $null = $answer.Add(@($FirstDay, 'T'))
        }
    }
    while ($QuantityOfDaysOfPractice -ne 0) {
        $FirstDay = $FirstDay.AddDays(1)
        if (-not (Is-Holiday($FirstDay))) {  
            $QuantityOfDaysOfPractice--
            $null = $answer.Add(@($FirstDay, 'P'))
        }
    }
    #поиск даты консультации
    $DateOfConsultation = $answer[$answer.Count - 1][0].addDays(1)
    while (Is-Holiday($DateOfConsultation)) {
        $DateOfConsultation = $DateOfConsultation.AddDays(1)
    }
    $null = $answer.Add(@($DateOfConsultation, 'к'))
    #поиск даты экзамена
    $DateOfExam = $answer[$answer.Count - 1][0].addDays(1)
    while (Is-Holiday($DateOfExam)) {
        $DateOfExam = $DateOfExam.AddDays(1)
    }
    $null = $answer.Add(@($DateOfExam, 'э'))
    return $answer
}


#функция проверки даты на праздичность и выходные
function Is-Holiday{
    param (
        [DateTime]$MyDate
    )
    #список праздничных дат
    $Holidays = @(
    (Get-Date -Date '22.11.2022'),
    (Get-Date -Date '08.03.2022'),
    (Get-Date -Date '12.06.2022'),
    (Get-Date -Date '04.11.2022'),
    (Get-Date -Date '31.12.2022')
)
    foreach ($d in $Holidays) {
        if ($d.Day -eq $MyDate.Day -and $d.Month -eq $MyDate.Month) {
            return $true
        }
    }
    return $MyDate.DayOfWeek -eq 'Sunday' -or $MyDate.DayOfWeek -eq 'Saturday'
}


#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$WorkBook = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\obrazec_raspisanie.xlsx")
$WorkSheet = $WorkBook.Sheets('1')
#создание списка дат
Write-Output "Введите дату начала обучения в формате `"dd.mm.yyyy`""
$DateOfBegin = Read-Host
Write-Output "Введите количество часов теории"
$QuantityOfDaysOfTeory = ([int](Read-Host)) / 8
Write-Output "Введите количество часов практики"
$QuantityOfDaysOfPractice = ([int](Read-Host)) / 8
$ListOfDates = Get-ListOfDates $DateOfBegin $QuantityOfDaysOfTeory $QuantityOfDaysOfPractice


#заполение таблицы и проход по списку
$index = 0
$row = 1
$rowOfDays = 1
$hours = 8
$daysShift = 9
while ($true) {
    $a = $WorkSheet.UsedRange.Columns['I'].rows[$row].Value()
    $b = $WorkSheet.UsedRange.Columns['G'].rows[$row].Value()
    if ($a -like '*Дни занятий*') {
        $rowOfDays = $row + 2
        $daysShift = 9
        $WorkSheet.UsedRange.Columns['I'].rows[$row + 1].Value() = $ListOfDates[$index][0].Month
    }
    if ($b -match "\w+ \w\.\w\.") {
        $time = [int]$WorkSheet.UsedRange.Columns['F'].rows[$row].Value()
        if ($hours - $time -gt 0) {
            $hours -= $time
            $WorkSheetSroki.Cells.Item($row, $daysShift) = $time
            $WorkSheetSroki.Cells.Item($rowOfDays, $daysShift) = $ListOfDates[$index][0].Day
        } elseif ($hours - $time -eq 0) {
            $hours = 8 
            $WorkSheetSroki.Cells.Item($row, $daysShift) = $time
            $WorkSheetSroki.Cells.Item($rowOfDays, $daysShift) = $ListOfDates[$index][0].Day
            $index++  
            $daysShift++  
        } else {
            $ost = $time - $hours
            $hours = 8 - $ost
            $range = $WorkSheet.Range('A' + $row,'AE' + $row)
            $range.copy()
            $range.insert(-4121)
            $WorkSheetSroki.Cells.Item($row, $daysShift) = $hours
        }
    }
    if ($ListOfDates[$index][0].Day -eq 1) {
       $WorkSheetSroki.Cells.Item($rowOfDays - 1, $daysShift) = $ListOfDates[$index][0].Month
    }
    $row++
}


#сохранение, закрытие файлов, закрытие Excel
$WorkBook.SaveAs('D:\coding\workplace\pshell\test\raspisanie.xlsx')
$WorkBook.close($true)
$Excel.Quit()
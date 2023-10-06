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
$WorkBook = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\obrazec_zurnal.xlsx")
$WorkSheet = $WorkBook.Sheets('1')
#переменная для учёта количества строк
Write-Output 'Введите количество строк'
$LimitOfRows = [int](Read-Host)
#создание списка дат
Write-Output "Введите дату начала обучения в формате `"dd.mm.yyyy`""
$DateOfBegin = Read-Host
Write-Output "Введите количество часов теории"
$QuantityOfDaysOfTeory = ([int](Read-Host)) / 8
Write-Output "Введите количество часов практики"
$QuantityOfDaysOfPractice = ([int](Read-Host)) / 8
$ListOfDates = Get-ListOfDates $DateOfBegin $QuantityOfDaysOfTeory $QuantityOfDaysOfPractice
#итерация по списку дат и заполнение расписания теоретических занятий
$index = 0
$row = 1
$hours = 8
while ($row -le $LimitOfRows) {
    if ($WorkSheet.UsedRange.Columns['G'].rows[$row].Value() -match "\w+ \w\.\w\.") {
         $ostatok = [int]$WorkSheet.UsedRange.Columns['B'].rows[$row].Value()
         $WorkSheet.UsedRange.Columns['A'].rows[$row].Value() = $ListOfDates[$index][0]
         if (($hours - $ostatok) -eq 0) {
            $index++
            $hours = 8
         } elseif (($hours - $ostatok) -gt 0) {
            $hours -= $ostatok
         } else {
            $WorkSheet.UsedRange.Columns['B'].rows[$row].Value() = $hours
            $a = $ostatok - $hours
            $hours = 8 + $hours - $ostatok
            $range = $WorkSheet.Range('A' + $row,'I' + $row)
            $range.copy()
            $range.insert(-4121)
            $WorkSheet.UsedRange.Columns['B'].rows[$row + 1].Value() = $a
            $LimitOfRows++
            $index++
         }
    }
    $row++
}
#поиск первого и последнего дня практики, даты консультации и экзамена
$FirstPracticeDay = $null
$LastPracticeDay = $null
$DayOfExam = $null
$DayOfConsultation = $null
while ($true) {
    if ($ListOfDates[$index][1] -eq 'P') {
        $FirstPracticeDay = $ListOfDates[$index][0]
        break
    }
}
$index = $ListOfDates.Length - 1
while ($true) {
    if ($ListOfDates[$index][1] -eq 'э') {
        $DayOfExam = $ListOfDates[$index][0]
    }
    if ($ListOfDates[$index][1] -eq 'к') {
        $DayOfConsultation = $ListOfDates[$index][0]
    }
    if ($ListOfDates[$index][1] -eq 'P') {
        $LastPracticeDay = $ListOfDates[$index][0]
        break
    }
    $index--
}
     
#поиск и заполнение дат практики, консультации и экзамена
$row--
while ($true) {
    $a = $WorkSheet.UsedRange.Columns['D'].rows[$row].Value()
    if ($a -like '*Производственное обучение*') {
        $WorkSheet.UsedRange.Columns['A'].rows[$row].Value() = $FirstPracticeDay.Date.ToString("dd.MM.yyyy") + '-' + $LastPracticeDay.Date.ToString("dd.MM.yyyy")    
    } 
    if ($a -like '*Консультации*') {
        $WorkSheet.UsedRange.Columns['A'].rows[$row].Value() = $DayOfConsultation.Date.ToString("dd.MM.yyyy")   
    }
    if ($a -like '*Квалификационный экзамен*') {
        $WorkSheet.UsedRange.Columns['A'].rows[$row].Value() = $DayOfExam.Date.ToString("dd.MM.yyyy")   
        break
    }       
    $row++
}


#сохранение, закрытие файлов, закрытие Excel
$WorkBook.SaveAs('D:\coding\workplace\pshell\test\zhurnal.xlsx')
$WorkBook.close($true)
$Excel.Quit()
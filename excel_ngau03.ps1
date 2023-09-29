#функция генерации списка дат обучения (теория и практика)
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
    } else {
        $null = $answer.Add(@($FirstDay, 'В'))
    }
    while ($QuantityOfDaysOfTeory -ne 0) {
        $FirstDay = $FirstDay.AddDays(1)
        $null = $answer.Add(@($FirstDay, 'T'))
        if (-not (Is-Holiday($FirstDay))) {  
            $QuantityOfDaysOfTeory--
        }
    }
    while ($QuantityOfDaysOfPractice -ne 0) {
        $FirstDay = $FirstDay.AddDays(1)
        if (-not (Is-Holiday($FirstDay))) {  
            $QuantityOfDaysOfPractice--
            $null = $answer.Add(@($FirstDay, 'P'))
        } else {
            $null = $answer.Add(@($FirstDay, 'В'))
        }
    }
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
    return $MyDate.DayOfWeek -eq 'Sunday'
}


#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$WorkBookSource = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\ishodnik.xlsx")
$WorkSheetSource = $WorkBookSource.Sheets('1')
#дата начала занятий
$BeginDate = ($WorkSheetSource.UsedRange.Columns['A'].rows[1].text -split ": ")[1]
#профессия
$Profession = $WorkSheetSource.UsedRange.Columns['A'].rows[3].text
#поиск количества часов теории
$i = 1
while ($true) {
    $teoria = $WorkSheetSource.UsedRange.Columns['B'].rows[$i].text
    if ($teoria -like '*Теоретическое обучение*') {
        break
    }
    $i++
}
$TimeOfTeory = $WorkSheetSource.UsedRange.Columns['C'].rows[$i].text
#поиск количества часов практики
$i = 1
while ($true) {
    $practika = $WorkSheetSource.UsedRange.Columns['B'].rows[$i].text
    if ($practika -like '*Производственное обучение*') {
        break
    }
    $i++
}
$TimeOfPractice = $WorkSheetSource.UsedRange.Columns['C'].rows[$i].text
#поиск дат занятия теорий и практикой
$ListOfDates = Get-ListOfDates $BeginDate ([int]$TimeOfTeory / 8) ([int]$TimeOfPractice / 8)
Write-Output $ListOfDates
#поиск даты консультации
$DateOfConsultation = $ListOfDates[$ListOfDates.Length - 1][0].AddDays(1)
while (Is-Holiday($DateOfConsultation)) {
    $DateOfConsultation = $DateOfConsultation.AddDays(1)
}
Write-Output $DateOfConsultation
#поиск даты экзамена
$DateOfExam = $DateOfConsultation.AddDays(1)
while (Is-Holiday($DateOfExam)) {
    $DateOfExam = $DateOfExam.AddDays(1)
}
Write-Output $DateOfExam
#поиск последнего дня теории и первого дня практики
$LastTeoryDay = $null
$FirstPracticeDay = $null
$i = 0
while ($i -lt $ListOfDates.Length) {
    if ($ListOfDates[$i][1] -eq 'P') {
        $FirstPracticeDay = $ListOfDates[$i][0]
        while ($true) {
            if ($ListOfDates[$i][1] -eq 'T') {
                $LastTeoryDay = $ListOfDates[$i][0]
                break
            }
        $i--
        }
        break
    }
    $i++
}


#Запись данных в файл-образец
#Открытие файла
$WorkBookSroki = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\obrazec.xlsx")
$WorkSheetSroki = $WorkBookSroki.Sheets('1')
#первая строка
$WorkSheetSroki.Cells.Item(1, 1) = 'Сроки обучения ' + $ListOfDates[0].Date.ToString("dd.MM.yyyy") + ' г. по ' + $DateOfExam.Date.ToString("dd.MM.yyyy") + ' г.'
$WorkSheetSroki.Cells.Item(1, 1).Font.Bold = $true
#вторая строка
$WorkSheetSroki.Cells.Item(2, 1) = 'Учебное заведение ИДПО ФГБОУ ВО Новосибирский ГАУ'
#третья строка
$WorkSheetSroki.Cells.Item(3, 1) = 'Профессия ' + $Profession
#время теории
$WorkSheetSroki.Cells.Item(4, 1) = 'теория = ' + $TimeOfTeory + ' часов c ' + $ListOfDates[0].Date.ToString("dd.MM.yyyy") + ' г. - по ' + $LastTeoryDay.Date.ToString("dd.MM.yyyy") + ' г.'
#время практики
$WorkSheetSroki.Cells.Item(5, 1) = 'практика = ' + $TimeOfPractice + ' часов c ' + $FirstPracticeDay.Date.ToString("dd.MM.yyyy") + ' г. - по ' + $ListOfDates[$ListOfDates.length - 1].Date.ToString("dd.MM.yyyy") + ' г.'
#время консультации
$WorkSheetSroki.Cells.Item(4, 'Y') = 'консультация =8 часов ' + $DateOfConsultation.Date.ToString("dd.MM.yyyy") + ' г.'
#время экзамена
$WorkSheetSroki.Cells.Item(5, 'Y') = 'экзамен = 8 часов ' + $DateOfExam.Date.ToString("dd.MM.yyyy") + ' г.'
#6-я строка
$WorkSheetSroki.Cells.Item(6, 'A') = 'всего =' + ([int]$TimeOfTeory + [int]$TimeOfPractice + 16) + ' часов'
#7-я строка
$WorkSheetSroki.Cells.Item(7, 'AF') = 'Час'
$WorkSheetSroki.Cells.Item(7, 'AG') = 'Дни'
$WorkSheetSroki.Cells.Item(7, 'AH') = 'Прожив.'

$WorkBookSource.close($true)
$WorkBookSroki.SaveAs('D:\coding\workplace\pshell\test\sroki.xlsx')
$WorkBookSroki.close($true)
$Excel.Quit()

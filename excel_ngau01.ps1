#функция генерации списка дат обучения (теория, практика, консультация и экзамен)
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
        if (-not (Is-Holiday($FirstDay))) {  
            $QuantityOfDaysOfTeory--
            $null = $answer.Add(@($FirstDay, 'T'))
        } else {
            $null = $answer.Add(@($FirstDay, 'В'))
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
    #поиск даты консультации
    $DateOfConsultation = $answer[$answer.Count - 1][0].addDays(1)
    while (Is-Holiday($DateOfConsultation)) {
        $null = $answer.Add(@($DateOfConsultation, 'В'))
        $DateOfConsultation = $DateOfConsultation.AddDays(1)
    }
    $null = $answer.Add(@($DateOfConsultation, 'к'))
    #поиск даты экзамена
    $DateOfExam = $answer[$answer.Count - 1][0].addDays(1)
    while (Is-Holiday($DateOfExam)) {
        $null = $answer.Add(@($DateOfExam, 'В'))
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
    return $MyDate.DayOfWeek -eq 'Sunday'
}

#Создание словаря для месяцев
$months_map = @{}
$months_map[1] = 'Месяц январь'
$months_map[2] = 'Месяц февраль'
$months_map[3] = 'Месяц март'
$months_map[4] = 'Месяц апрель'
$months_map[5] = 'Месяц май'
$months_map[6] = 'Месяц июнь'
$months_map[7] = 'Месяц июль'
$months_map[8] = 'Месяц август'
$months_map[9] = 'Месяц сентябрь'
$months_map[10] = 'Месяц октябрь'
$months_map[11] = 'Месяц ноябрь'
$months_map[12] = 'Месяц декабрь'


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
#поиск дат занятия
$ListOfDates = Get-ListOfDates $BeginDate ([int]$TimeOfTeory / 8) ([int]$TimeOfPractice / 8)
#поиск последнего дня теории, первого и последнего дней практики, дня экзамена, дня консультации
$LastTeoryDay = $null
$FirstPracticeDay = $null
$LastPracticeDay = $null
$DayOfExam = $null
$DayOfConsultation = $null
Write-Output $ListOfDates
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
$i = 0
$flag = $true
while ($flag) {
    if ($ListOfDates[$i][1] -eq 'к') {
        $DayOfConsultation = $ListOfDates[$i][0]
        while ($true) {
            if ($ListOfDates[$i][1] -eq 'P') {
                $LastPracticeDay = $ListOfDates[$i][0]
                $flag = $false
                break
            }
            $i--
        }
    }
    $i++
}
$i = 0
while ($i -lt $ListOfDates.Length) {
    if ($ListOfDates[$i][1] -eq 'э') {
        $DayOfExam = $ListOfDates[$i][0]
        break
    }
    $i++
}



#Запись данных в файл-образец
#Открытие файла
$WorkBookSroki = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\obrazec.xlsx")
$WorkSheetSroki = $WorkBookSroki.Sheets('1')
#первая строка
$WorkSheetSroki.Cells.Item(1, 1) = 'Сроки обучения ' + $ListOfDates[0][0].Date.ToString("dd.MM.yyyy") + ' г. по ' + $DayOfExam.Date.ToString("dd.MM.yyyy") + ' г.'
$WorkSheetSroki.Cells.Item(1, 1).Font.Bold = $true
#вторая строка
$WorkSheetSroki.Cells.Item(2, 1) = 'Учебное заведение ИДПО ФГБОУ ВО Новосибирский ГАУ'
#третья строка
$WorkSheetSroki.Cells.Item(3, 1) = 'Профессия ' + $Profession
#время теории
$WorkSheetSroki.Cells.Item(4, 1) = 'теория = ' + $TimeOfTeory + ' часов c ' + $ListOfDates[0][0].Date.ToString("dd.MM.yyyy") + ' г. - по ' + $LastTeoryDay.Date.ToString("dd.MM.yyyy") + ' г.'
#время практики
$WorkSheetSroki.Cells.Item(5, 1) = 'практика = ' + $TimeOfPractice + ' часов c ' + $FirstPracticeDay.Date.ToString("dd.MM.yyyy") + ' г. - по ' + $LastPracticeDay.Date.ToString("dd.MM.yyyy") + ' г.'
#время консультации
$WorkSheetSroki.Cells.Item(4, 'Y') = 'консультация =8 часов ' + $DayOfConsultation.Date.ToString("dd.MM.yyyy") + ' г.'
#время экзамена
$WorkSheetSroki.Cells.Item(5, 'Y') = 'экзамен = 8 часов ' + $DayOfExam.Date.ToString("dd.MM.yyyy") + ' г.'
#6-я строка
$WorkSheetSroki.Cells.Item(6, 'A') = 'всего =' + ([int]$TimeOfTeory + [int]$TimeOfPractice + 16) + ' часов'
#7-я строка
$WorkSheetSroki.Cells.Item(7, 'AF') = 'Час'
$WorkSheetSroki.Cells.Item(7, 'AG') = 'Дни'
$WorkSheetSroki.Cells.Item(7, 'AH') = 'Прожив.'
#цикл построения графиков занятий по месяцам
$number_of_row = 13
$index = 0
$count_of_days = 0
$days_of_studies = 0
$hours_of_practice = 0
$hours_of_teory = 0
while ($index -ne $ListOfDates.Length) {
    $day = $ListOfDates[$index]
    if ($index -eq 0 -or $day[0].Day -eq 1) {
        #очистка счётчиков
        $count_of_days = 0
        $hours_of_practice = 0
        $hours_of_teory = 0
        #копирование ячеек и очистка их содержимого
        $range1 = $WorkSheetSroki.Range('A8','AH12')
        $range1.copy()
        $range2 = $WorkSheetSroki.Range('A' + $number_of_row, 'AH' + ($number_of_row + 4))
        $WorkSheetSroki.paste($range2)
        $range2.ClearContents()
        #записываем имя месяца
        $WorkSheetSroki.Cells.Item($number_of_row, 'A') = $months_map[$day[0].Month]
        #заполняем строку номерами дней месяца
        $i = 1;
        $quantity_of_days = [DateTime]::DaysInMonth($day[0].Year, $day[0].Month)
        while ($i -le $quantity_of_days) {
            $WorkSheetSroki.Cells.Item($number_of_row + 1, $i) = $i
            $i++
        }  
    }
    #заполняем строки днями, часами и выходными
    $WorkSheetSroki.Cells.Item($number_of_row + 2, $day[0].Day) = '1'
    $count_of_days++
    if ($day[1] -eq 'T') {
        $WorkSheetSroki.Cells.Item($number_of_row + 3, $day[0].Day) = '8'
        $hours_of_teory += 8
    } elseif ($day[1] -eq 'P') {
        $WorkSheetSroki.Cells.Item($number_of_row + 4, $day[0].Day) = '8'
        $hours_of_practice += 8
    } elseif ($day[1] -eq 'к') {
        $WorkSheetSroki.Cells.Item($number_of_row + 4, $day[0].Day) = 'к'
        $hours_of_practice += 8
    } elseif ($day[1] -eq 'э') {
        $WorkSheetSroki.Cells.Item($number_of_row + 4, $day[0].Day) = 'э'
        $hours_of_practice += 8
    }
    else {
        $WorkSheetSroki.Cells.Item($number_of_row + 4, $day[0].Day) = 'В'
    }
    if (($day[0].Day -eq [DateTime]::DaysInMonth($day[0].Year, $day[0].Month) -or ($index -eq ($ListOfDates.Length - 1)))) {
        $WorkSheetSroki.Cells.Item($number_of_row + 2, 'AG') = $count_of_days
        $WorkSheetSroki.Cells.Item($number_of_row + 3, 'AF') = $hours_of_teory
        $WorkSheetSroki.Cells.Item($number_of_row + 4, 'AF') = $hours_of_practice
        $days_of_studies += $count_of_days
        $number_of_row += 5
    }
    $index++
}
$range1 = $WorkSheetSroki.Range('A8','AH12')
$range1.delete()
#запись суммарного количества дней и часов
$WorkSheetSroki.Cells.Item($number_of_row - 5, 'AF') = ([int]$TimeOfTeory + [int]$TimeOfPractice + 16)
$WorkSheetSroki.Cells.Item($number_of_row - 5, 'AG') = $days_of_studies


#сохранение, закрытие файлов, закрытие Excel
$WorkBookSource.close($true)
$WorkBookSroki.SaveAs('D:\coding\workplace\pshell\test\sroki.xlsx')
$WorkBookSroki.close($true)
$Excel.Quit()

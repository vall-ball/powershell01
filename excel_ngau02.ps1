#функция генерации списка дат обучения
function Get-ListOfDates {
    param (
        [string]$BeginDate,
        [int]$QuantityOfDays
    )
    $answer = [System.Collections.ArrayList]::new()
    $FirstDay = Get-Date -Date $BeginDate
    $null = $answer.Add($FirstDay)
    if ($FirstDay.DayOfWeek -ne 'Sunday' -and -not (Is-Holiday($FirstDay))) {
        $QuantityOfDays--
    }
    while ($QuantityOfDays -ne 0) {
        $FirstDay = $FirstDay.AddDays(1)
        $null = $answer.Add($FirstDay)
        if ($FirstDay.DayOfWeek -ne 'Sunday' -and -not (Is-Holiday($FirstDay))) {  
            $QuantityOfDays--
        }
    }
    return $answer
}


#функция проверки даты на праздичность
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
    return $false
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
#$Excel.Visible = $true
$WorkBookSource = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\ishodnik.xlsx")
$WorkSheetSource = $WorkBookSource.Sheets('1')
#дата начала занятий
$BeginDate = ($WorkSheetSource.UsedRange.Columns['A'].rows[1].text -split ": ")[1]
#количество дней обучения
$QuantityOfDays = [int]($WorkSheetSource.UsedRange.Columns['A'].rows[2].text -split ": ")[1] / 8
#список дат для обучения
$ListOfDates = Get-ListOfDates $BeginDate $QuantityOfDays
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
#поиск дат занятия теорией
$ListOfDatesOfTeory = Get-ListOfDates $BeginDate ([int]$TimeOfTeory / 8)
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
#поиск дат занятия практикой
$ListOfDatesOfPractice = Get-ListOfDates ($ListOfDatesOfTeory[$ListOfDatesOfTeory.Length - 1].AddDays(1)).toString() ([int]$TimeOfPractice / 8)
#поиск даты консультации
$DateOfConsultation = $ListOfDatesOfPractice[$ListOfDatesOfPractice.Length - 1].AddDays(1)
while ((Is-Holiday($DateOfConsultation)) -or ($DateOfConsultation.DayOfWeek -eq 'Sunday')) {
    $DateOfConsultation = $DateOfConsultation.AddDays(1)
}
#поиск даты экзамена
$DateOfExam = $DateOfConsultation.AddDays(1)
while (($DateOfExam.DayOfWeek -eq 'Sunday')  -or (Is-Holiday($DateOfExam))) {
    $DateOfExam = $DateOfExam.AddDays(1)
}



#Запись данных в файл-образец
#Открытие файла
#$Excel = New-Object -ComObject Excel.Application
#$Excel.Visible = $true
$WorkBookSroki = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\obrazec.xlsx")
$WorkSheetSroki = $WorkBookSroki.Sheets('1')
#первая строка
$WorkSheetSroki.Cells.Item(1, 1) = 'Сроки обучения ' + $ListOfDates[0].Date.ToString("dd.MM.yyyy") + ' г. по ' + $ListOfDates[$ListOfDates.Length - 1].Date.ToString("dd.MM.yyyy") + ' г.'
$WorkSheetSroki.Cells.Item(1, 1).Font.Bold = $true
#вторая строка
$WorkSheetSroki.Cells.Item(2, 1) = 'Учебное заведение ИДПО ФГБОУ ВО Новосибирский ГАУ'
#третья строка
$WorkSheetSroki.Cells.Item(3, 1) = 'Профессия ' + $Profession
#время теории
$WorkSheetSroki.Cells.Item(4, 1) = 'теория = ' + $TimeOfTeory + ' часов c ' + $ListOfDatesOfTeory[0].Date.ToString("dd.MM.yyyy") + ' г. - по ' + $ListOfDatesOfTeory[$ListOfDatesOfTeory.length - 1].Date.ToString("dd.MM.yyyy") + ' г.'
#время практики
$WorkSheetSroki.Cells.Item(5, 1) = 'практика = ' + $TimeOfPractice + ' часов c ' + $ListOfDatesOfPractice[0].Date.ToString("dd.MM.yyyy") + ' г. - по ' + $ListOfDatesOfPractice[$ListOfDatesOfPractice.length - 1].Date.ToString("dd.MM.yyyy") + ' г.'
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
#Пострение расписания
$number_of_row = 8
while ($true) {
    $WorkSheetSroki.Cells.Item($number_of_row, 'A') = $months_map[$ListOfDatesOfTeory[0].Month]
    $number_of_row++
    $i = 1;
    while ($i -le [DateTime]::DaysInMonth($ListOfDatesOfTeory[0].Year, $ListOfDatesOfTeory[0].Month)) {
    }
    break
}


#test копирование ячеек и очистка их соедержимого
#$Range = $WorkSheetSroki.Range('A8','AH13')
#$Range.copy()
#$range2 = $WorkSheetSroki.Range("A30:AH35")
#$WorkSheetSroki.paste($range2)
#$range2.ClearContents()




$WorkBookSource.close($true)
$WorkBookSroki.SaveAs('D:\coding\workplace\pshell\test\sroki.xlsx')
$WorkBookSroki.close($true)
$Excel.Quit()

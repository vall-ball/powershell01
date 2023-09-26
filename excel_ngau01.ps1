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


#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
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
    if ($teoria -like '*1.Теоретическое обучение *') {
        break
    }
    $i++
}
$TimeOfTeory = $WorkSheetSource.UsedRange.Columns['C'].rows[$i].text
#поиск дат занятия теорией
$ListOfDatesOfTeory = Get-ListOfDates $BeginDate [int]$TimeOfTeory / 8


#Запись данных в файл
# Добавить рабочую книгу
$WorkBookSroki = $Excel.Workbooks.Add()
$WorkSheetSroki = $WorkBookSroki.Worksheets.Item(1)
# Переименовывать лист
$WorkSheetSroki.Name = '1'
#первая строка
$WorkSheetSroki.Cells.Item(1, 1) = 'Сроки обучения ' + $ListOfDates[0].Date.ToString("dd.MM.yyyy") + ' г. по ' + $ListOfDates[$ListOfDates.Length - 1].Date.ToString("dd.MM.yyyy") + ' г.'
$WorkSheetSroki.Cells.Item(1, 1).Font.Bold = $true
$Range = $WorkSheetSroki.Range('A1','AG1')
$Range.Merge()
$Range.HorizontalAlignment = -4108
#вторая строка
$WorkSheetSroki.Cells.Item(2, 1) = 'Учебное заведение ИДПО ФГБОУ ВО Новосибирский ГАУ'
$Range = $WorkSheetSroki.Range('A2','AG2')
$Range.Merge()
$Range.HorizontalAlignment = -4108
#третья строка
$WorkSheetSroki.Cells.Item(3, 1) = 'Профессия ' + $Profession
$Range = $WorkSheetSroki.Range('A3','AG3')
$Range.Merge()
$Range.HorizontalAlignment = -4108
#время теории
$WorkSheetSroki.Cells.Item(4, 1) = 'теория =' + $TimeOfTeory + ' c ' + $ListOfDatesOfTeory[0].Date.ToString("dd.MM.yyyy") + ' г. - по ' + $ListOfDatesOfTeory[$ListOfDatesOfTeory.length - 1].Date.ToString("dd.MM.yyyy") + ' г.'






$WorkBookSource.close($true)
$WorkBookSroki.SaveAs('D:\coding\workplace\pshell\test\sroki.xlsx')
$WorkBookSroki.close($true)
$Excel.Quit()

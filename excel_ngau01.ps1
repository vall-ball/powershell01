#функция генерации списка дат обучения
function Get-PassportDate {
    param (
        [string]$BeginDate,
        [int]$QuantityOfDays
    )
    $answer = [System.Collections.ArrayList]::new()
    $FirstDay = Get-Date -Date $BeginDate
    $answer.Add($FirstDay)
    if ($FirstDay.DayOfWeek -ne 7) {
        $QuantityOfDays--
    }
    while ($QuantityOfDays -ne 0) {
        $FirstDay = $FirstDay.AddDays(1)
        $answer.Add($FirstDay)
        if ($FirstDay.DayOfWeek -ne 7) {
            $QuantityOfDays--
        }
    }
    return $answer
}



#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$WorkBookSource = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\ishodnik.xlsx")
$WorkSheetSource = $WorkBookSource.Sheets('1')


# Добавить рабочую книгу
$WorkBookSroki = $Excel.Workbooks.Add()
$WorkSheetSroki = $WorkBookSroki.Worksheets.Item(1)
# Переименовывать лист
$WorkSheetSroki.Name = '1'
#
$WorkSheetSroki.Cells.Item(1, 1) = 'Сроки обучения'
$WorkSheetSroki.Cells.Item(1, 1).Font.Bold = $true
$Range = $WorkSheetSroki.Range('A1','AG1')
$Range.Merge()
$Range.HorizontalAlignment = -4108
#
$WorkSheetSroki.Cells.Item(2, 1) = 'Учебное заведение ИДПО ФГБОУ ВО Новосибирский ГАУ'
$Range = $WorkSheetSroki.Range('A2','AG2')
$Range.Merge()
$Range.HorizontalAlignment = -4108






$WorkBookSource.close($true)
$WorkBookSroki.SaveAs('D:\coding\workplace\pshell\test\sroki.xlsx')
$WorkBookSroki.close($true)
$Excel.Quit()

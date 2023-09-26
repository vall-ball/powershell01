#функция сравнения двух массивов
function compare-arrays{
    param (
        $arr1,
        $arr2
    )
  
    return ($arr1[0] -eq $arr2[0] -and $arr1[1] -eq $arr2[1] -and $arr1[2] -eq $arr2[2])
 }


 #функция поиска массива в списке
 function is-array-in-list{
    param (
        $arr,
        $list
    )
    for ($i = 0; $i -lt $list.count; $i++) {
    if (compare-arrays $arr $list[$i]) {
            return $true
        }
    }
    return $false
 }

#Получения списка ФИО, программ и времени обучения из файла


#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$WorkBook = $Excel.Workbooks.Open("D:\coding\workplace\pshell\test\1.xlsx")
$WorkSheet1 = $WorkBook.Sheets('1')

#создание списка пользователей
$list = [System.Collections.ArrayList]::new()
for ($i = 2; $i -le $WorkSheet1.UsedRange.Columns['D'].rows.Count; $i++) {
    
        $u = @($WorkSheet1.UsedRange.Columns['D'].rows[$i].text, $WorkSheet1.UsedRange.Columns['I'].rows[$i].text, $WorkSheet1.UsedRange.Columns['J'].rows[$i].text)
        $list.Add($u)
        
}


#поиск совпадений в другой вкладке и окраска ячеек
$WorkSheet2 = $WorkBook.Sheets('2022')
for ($i = 2; $i -le $WorkSheet2.UsedRange.Columns['D'].rows.Count; $i++) {
    $a = @($WorkSheet2.UsedRange.Columns['D'].rows[$i].text, $WorkSheet2.UsedRange.Columns['I'].rows[$i].text, $WorkSheet2.UsedRange.Columns['J'].rows[$i].text)   
    if (is-array-in-list $a  $list) {     
        $WorkSheet2.UsedRange.Columns['D'].rows[$i].Interior.ColorIndex = 6
        $WorkSheet2.Columns['E'].rows[$i] = $WorkSheet1.UsedRange.Columns['E'].rows[$i].text
        $WorkSheet2.Columns['F'].rows[$i] = $WorkSheet1.UsedRange.Columns['F'].rows[$i].text
        $WorkSheet2.Columns['G'].rows[$i] = $WorkSheet1.UsedRange.Columns['G'].rows[$i].text
        $WorkSheet2.Columns['H'].rows[$i] = $WorkSheet1.UsedRange.Columns['H'].rows[$i].text
    }     
}


#Выход из Excel
$WorkBook.Save()
$WorkBook.close($true)
$Excel.Quit()
#функция поиска элемента с таким именем в списке
function Get-UserByName{
    param (
        $users,
        [String] $name
    )
    foreach ($u in $users) {
        if ($u.Name -eq $name) {
            return $u
        }
    }
    return $null
 }




#Чтение данных из файла
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true
$WorkBook = $Excel.Workbooks.Open("D:\coding\workplace\pshell\dlavala.xlsx")
$WorkSheet = $WorkBook.Sheets('Лист1')

#создание списка пользователей
$frdo = [System.Collections.ArrayList]::new()

for ($i = 1; $i -le $WorkSheet.UsedRange.Columns['F'].rows.Count; $i++) {
    if ($WorkSheet.Columns['F'].rows[$i].Text.Length -ne 0) {
        $u = New-Object  –TypeName PSCustomObject -Property @{Name = $WorkSheet.UsedRange.Columns['F'].rows[$i].text; 
                                          BirthDate = $WorkSheet.UsedRange.Columns['G'].rows[$i].text;
                                          Snils = $WorkSheet.UsedRange.Columns['H'].rows[$i].text}
        $frdo.Add($u)
        }      
}
   
$WorkSheet = $WorkBook.Sheets('для договоров')
#поиск людей в списке и изменение 
for ($i = 1; $i -le $WorkSheet.UsedRange.Columns['B'].rows.Count; $i++) {
    $user = Get-UserByName $frdo $WorkSheet.UsedRange.Columns['B'].rows[$i].text
    if ($user -ne $null) {
        if ($user.BirthDate -ne $WorkSheet.UsedRange.Columns['E'].rows[$i].text) {
            $WorkSheet.Columns['F'].rows[$i] = ([String]$user.BirthDate)
        }
        if ($WorkSheet.UsedRange.Columns['M'].rows[$i].text.Length -eq 0) {
            $WorkSheet.Columns['M'].rows[$i] = ([String]$user.Snils)
        }
        
        }      
}

#Выход из Excel
$WorkBook.Save()
$WorkBook.close($true)
$Excel.Quit()
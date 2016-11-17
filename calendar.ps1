# Создает календарь за заданный год в Excel
# (c) 2015
#----------------------------------------------------
$year=2017 # за какой год календарь
#----------------------------------------------------
# Праздники в формате ММ-ДД
$holidays="01-01","01-02","01-03","01-04","01-05","01-06","01-07","01-08",
            "02-23","03-08","05-01","05-09","06-12","11-04"
#----------------------------------------------------
# Начальная строка и столбец календаря в книге
$startRow=1
$startCol=1

# Константы Excel
$xlCenter=-4108 # Горизонтальное выравнивание по центру
$xlRed=3 # Красный цвет

$Culture = [System.Globalization.CultureInfo]::CurrentCulture
# Имена дней недели
$WeekDayNames= $culture.DateTimeFormat.AbbreviatedDayNames 
# Имена выходных дней недели
$RedWeekDays=$WeekDayNames[0],$WeekDayNames[6] # Суббота и воскресенье выходные
$shift=[int]$Culture.DateTimeFormat.FirstDayOfWeek # сдвиг в зависимости от начального дня недели
$newdays=$WeekDayNames[$shift..($WeekDayNames.count-1)]+$WeekDayNames[0..($shift-1)] # Правильно сдвинутые имена дней недели
# $dayweek=((7-$shift)+$Date.dayofweek)%7 # Так получать номер дня недели для сдвинутых имен

$XL=New-Object -ComObject Excel.Application # Создать приложение
$XL.Visible=$True # Сделать окно видимым
$workbook=$XL.WorkBooks.Add() > $Null # Создать книгу

#Remove other worksheets
#1..2 | ForEach {
#    $XL.worksheets.item($_).Delete()
#}

#Connect to first worksheet to rename and make active
$serverInfoSheet = $XL.Worksheets.Item(1)
$serverInfoSheet.Name = "$year"
$serverInfoSheet.Activate() | Out-Null


#--------------------------------------------------------
# Печать месяца начиная с указанного столбца и колонки
function PrintMonth ($month,$year,$rowMonth,$colMonth)
{

# Последний день месяца
$LastDay = [System.DateTime]::DaysInMonth($year, $month)
 
# Начальная и конечная даты месяца
$DateStart=Get-Date -Day 01 -month $month -year $year 
$DateEnd=Get-Date -Day $LastDay -month $month -year $year

# Номер недели с которой начинается месяц
$startweek=$Culture.Calendar.GetWeekOfYear($DateStart, $Culture.DateTimeFormat.CalendarWeekRule, $Culture.DateTimeFormat.FirstDayOfWeek) #Номер первой недели месяца

# Задать ширину колонок
for ($i=$colMonth;$i -le $colMonth+8;$i++)
    {
    $XL.Columns.Item($i).ColumnWidth=3
    }

# Название месяца
    $XL.Cells.Item($rowMonth,$colMonth)=$DateStart.tostring("MMMM")
    # Объединение ячеек для красоты
    $ran=Get-Range ($rowMonth) ($colMonth) ($rowMonth) ($colMonth+6) 
    $XL.Range($ran).Select() > $nul
    $XL.Selection.Merge()> $nul
    $XL.Selection.HorizontalAlignment = $xlCenter
    $XL.Selection.Font.Bold=$true
# Дни недели
  PrintWeekDays ($rowMonth+1) ($colMonth)

# Перебор всех дней в месяце
for ($d=$DateStart;$d -le $dateEnd;$d=$d.Adddays(1))
{

# день недели
$dayweek=((7-$shift)+$d.dayofweek)%7

# номер недели в году
$week=$Culture.Calendar.GetWeekOfYear($d, $Culture.DateTimeFormat.CalendarWeekRule, $Culture.DateTimeFormat.FirstDayOfWeek)

$row=$rowMonth+2+$week-$startweek
$col=$colMonth+$dayweek

# Выводим номер дня
$XL.Cells.Item($row,$col)=$d.day
border $row $col
$XL.Cells.Item($row,$col).HorizontalAlignment = $xlCenter

$strDay=$d.tostring("MM-dd")

# Раскраска выходных дней
$dayweekOld=$d.DayOfWeek.value__ #0 воскресенье
if (($dayweekOld -eq 0) -or ($dayweekOld -eq 6) -or ($holidays -contains $strDay)) #Сб, Вск или праздник
    {
    $XL.Cells.Item($row,$col).Font.ColorIndex=$xlRed #Red
    }
}
}
#---------------------------------------------------------------


# Рисует границы у заданной ячейки
function Border ($row,$col)
{
$ran=Get-Range $row $col $row $col
$dataRange = $XL.Range($ran) 
7..12 | ForEach {
    $dataRange.Borders.Item($_).LineStyle = 1
    $dataRange.Borders.Item($_).Weight = 2
    }
}
# Возвращает строковое значение диапазона по координатам области
function Get-Range ($row1,$col1,$row2,$col2)
{
$let1=[char]($col1+64)
$let2=[char]($col2+64)
#$range="$let1$y1:$let2$y2"
$range="{0}{1}:{2}{3}" -f $let1,$row1,$let2,$row2
$range
}
function Get-Letter ($col)
{
$let1=[char]($col1+64)
}

# Печатает названия дней недели, горизонтально начиная с указнной ячейки
function PrintWeekDays($row,$col)
{
for ($i=0;$i -lt 7;$i++)
    {
    $XL.Cells.Item($row,$col+$i)=$newdays[$i]
    $XL.Cells.Item($row,$col+$i).Font.Bold=$true # Жирный
    $XL.Cells.Item($row,$col+$i).HorizontalAlignment = $xlCenter # Выровнять по центру
    if ($RedWeekDays -Contains $newdays[$i]) # Красный день календаря
    #if ($i -gt 4)
        {
        $XL.Cells.Item($row,$col+$i).Font.ColorIndex=$xlRed #Red
        }
    }

} # end function
#--------------------------------------------------------

# Высота и ширина месяца в ячейках
$MonthHeight=8
$MonthWidth=8

# Заголовок календаря
$XL.Cells.Item($startRow,$startCol)="Календарь на $year год"
$ran=Get-Range ($startRow) ($startCol) ($startRow) ($startCol+$MonthWidth*3-2) 
$XL.Range($ran).Select() > $nul
    $XL.Selection.Merge()> $nul
    $XL.Selection.HorizontalAlignment = $xlCenter
    $XL.Selection.Font.Bold=$true


# Идекс колонки и строки текущего месяца
$colIndex=0
$rowIndex=0

for ($i=1;$i -le 12;$i++) # Перебор всех месяцев
    {
    # Вывод месяцев слева на право
    PrintMonth $i $year (($rowIndex*$MonthHeight)+$startRow+2) (($colIndex*$MonthHeight)+$startCol)
    $colIndex++
    if ($i%3 -eq 0) # Каждый третий месяц, опускаемся вниз
        {
        $rowIndex++
        $colIndex=0
        }
       
    }


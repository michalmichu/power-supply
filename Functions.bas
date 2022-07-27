Attribute VB_Name = "Functions"
Function nb_hours_month(year, month)
    
    var_month = month
    var_year = year
    
    'Calculation for the first day of the following month
    date_next_month = DateSerial(var_year, var_month + 1, 1)
    
    last_day_month = date_next_month - 1
    
    nb_hours_m = Day(last_day_month) * 24
    
    If month = 3 Then
        nb_hours_m = nb_hours_m - 1
    ElseIf month = 10 Then
        nb_hours_m = nb_hours_m + 1
    End If
    
    nb_hours_month = nb_hours_m
    
End Function

Function priceBaseYearMonth(year, month)

    var_month = month
    var_year = year
    Set DaneBook = Workbooks("SprawdzanieWycen.xlsm").Sheets("Rynek_dane")
    Dim get_year As Integer
    Dim get_month As Integer
    
    
    For i = 31 To 36
        get_year = DaneBook.Cells(i, 14).Value
        If get_year = var_year Then
        Exit For
        End If
    Next
    
    For j = 15 To 28
        get_month = DaneBook.Cells(30, j).Value
        If get_month = var_month Then
        Exit For
        End If
    Next
    
    priceBaseYearMonth = DaneBook.Cells(i, j).Value

End Function
Function pricePeakYearMonth(year, month)

    var_month = month
    var_year = year
    Set DaneBook = Workbooks("SprawdzanieWycen.xlsm").Sheets("Rynek_dane")
    Dim get_year As Integer
    Dim get_month As Integer
    
    For i = 20 To 25
        get_year = DaneBook.Cells(i, 14).Value
        If get_year = var_year Then
        Exit For
        End If
    Next
    
    For j = 15 To 28
        get_month = DaneBook.Cells(19, j).Value
        If get_month = var_month Then
        Exit For
        End If
    Next
    
    
    pricePeakYearMonth = DaneBook.Cells(i, j).Value

End Function


Function PriceBaseWeightedAverage(year, monthStart, monthStop)
        
    Dim month As Long
    Dim avg As Double
    Dim sum As Double
    avg = 0
    sum = 0
    
    For month = monthStart To monthStop
        nb_hour_m = nb_hours_month(year, month)
        pr_base_ym = priceBaseYearMonth(year, month)
        avg = avg + nb_hour_m * pr_base_ym
        sum = sum + nb_hour_m
    
    Next month
    
    PriceBaseWeightedAverage = Round(avg / sum, 2)

End Function

Function peak_hours_month(year, month)
    Set HolidaySheet = Workbooks("SprawdzanieWycen.xlsm").Sheets("Break")
    'Month / Year of the date
    var_month = month
    var_year = year
    TotalDaysSunday = 0
    TotalDaysSaturday = 0
    TotalBreakDays = 0
    
    'Calculation for the first day of the following month
    date_next_month = DateSerial(var_year, var_month + 1, 1)
    
    'Date of the last day
    last_day_month = date_next_month - 1
    Days = Day(last_day_month)
    
    'Number of the Sundays
    xDay = 1
     For X = 1 To Days
        If Weekday(DateSerial(var_year, var_month, X)) = xDay Then
            TotalDaysSunday = TotalDaysSunday + 1
        End If
    Next X
    
    'Number of the Saturdays
    xDay = 7
     For X = 1 To Days
        If Weekday(DateSerial(var_year, var_month, X)) = xDay Then
            TotalDaysSaturday = TotalDaysSaturday + 1
        End If
    Next X
    
    'Number of the Break
     Break_year_position = (var_year - 2022) * 5
     Break_month_position = Break_year_position + 4
     Break_day_position = Break_year_position + 2
     Break_m = month_name(var_month)
     For X = 2 To 14
        Break_month = HolidaySheet.Cells(X, Break_month_position).Value
        Break_day_name = HolidaySheet.Cells(X, Break_day_position).Value
        If Break_month = Break_m And Break_day_name <> "Sob." And Break_day_name <> "Niedz." Then
            TotalBreakDays = TotalBreakDays + 1
        End If
        If TotalBreakDays > 1 Then Exit For
        
    Next X
    
    'Number hours from 1 for the last day of month (= last day)
    peak_hours_m = (Days - TotalDaysSunday - TotalDaysSaturday - TotalBreakDays) * 15
    
    
    peak_hours_month = peak_hours_m
    
End Function

Function PricePeakWeightedAverage(year, monthStart, monthStop)
        
    Dim month As Long
    Dim avg As Double
    Dim sum As Double
    avg = 0
    sum = 0
    
    For month = monthStart To monthStop
    
        peak_hour_m = peak_hours_month(year, month)
        pr_peak_ym = pricePeakYearMonth(year, month)
        avg = avg + peak_hour_m * pr_peak_ym
        sum = sum + peak_hour_m
    
    Next month
    
    PricePeakWeightedAverage = Round(avg / sum, 2)

End Function

Function month_name(month)
Select Case month
Case 1:    month_n = "Sty"
Case 2:    month_n = "Lut"
Case 3:    month_n = "Mar"
Case 4:    month_n = "Kwi"
Case 5:    month_n = "Maj"
Case 6:    month_n = "Cze"
Case 7:    month_n = "Lip"
Case 8:    month_n = "Sie"
Case 9:    month_n = "Wrz"
Case 10:   month_n = "Pa�"
Case 11:   month_n = "Lis"
Case 12:   month_n = "Gru"
End Select
month_name = month_n

End Function

Function pricePMYearMonth(year, monthStart, monthStop)

    var_year = year
    Price = 0
    monthNumber = 0
    Set DaneBook = Workbooks("SprawdzanieWycen.xlsm").Sheets("Rynek_dane")
    Dim get_month As Integer
    Dim get_year As Integer
    
    
    For i = 15 To 74
        get_year = DaneBook.Cells(5, i).Value
        If get_year = var_year Then
        Exit For
        End If
    Next
    For j = i + monthStart - 1 To i + monthStop - 1
        Price = Price + DaneBook.Cells(16, j).Value
        monthNumber = monthNumber + 1
    Next
    
    pricePMYearMonth = Round(Price / monthNumber, 2)

End Function

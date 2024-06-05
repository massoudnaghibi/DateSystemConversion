Module Module2

    Dim weekdays() As String = New String(6) {"یک شنبه", "دوشنبه", "سه شنبه", "چهارشنبه", "پنج شنبه", "جمعه", "شنبه"}
    Dim yearmonths(,) As String = New String(2, 11) {{"فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"}, _
                                                     {"ژانویه", "فوریه", "مارس", "آوریل", "مه", "ژوئن", "جولای", "آگوست", "سپتامبر", "اکتبر", "نوامبر", "دسامبر"}, _
                                                     {"محرم", "صفر", "ربیع الاول", "ربیع الثانی", "جمادی الاول", "جمادی الثانی", "رجب", "شعبان", "رمضان", "شوال", "ذی القعده", "ذی الحجه"}}
    ' The "stDateCovert Function()" accepts two parameters:
    '  1- The "stDate" (Date String) which Should Be converted from one system to another in the format "YYYY/MM/DD" e.g. "1393/15/08"
    '  2- The "DateConvType" (Conversion Type) Constant: "PI","PJ","IP","IJ","JP","JI", while P:Persian, I:Islamic, J:Julian(Christian)
    Function stDateConvert(ByVal stDate As String, ByVal DateConvType As DateConvType2) As String
        Dim iY, iM, iD As Integer
        Dim sY, sM, sD As String
        stDate = stDate.Trim()

        iY = CInt(stDate.Substring(0, 4))
        iM = CInt(stDate.Substring(5, 2))
        iD = CInt(stDate.Substring(8, 2))
        Select Case DateConvType
            Case "PI"
                persian_islamic(iY, iM, iD)
            Case "IP"
                islamic_persian(iY, iM, iD)
            Case "PJ"
                persian_civil(iY, iM, iD)
            Case "JP"
                civil_persian(iY, iM, iD)
            Case "IJ"
                islamic_civil(iY, iM, iD)
            Case "JI"
                civil_islamic(iY, iM, iD)
        End Select
        sY = iY.ToString("D4")
        sM = iM.ToString("D2")
        sD = iD.ToString("D2")
        stDateConvert = sY + "/" + sM + "/" + sD
    End Function

    ' The "stDateDesc Function()" accepts two parameters:
    '  1- The "stDate" (Date String) which DayOfWeek and Month Names should be generated for it in the format "YYYY/MM/DD" e.g. "1393/15/08"
    '  2- The "DateConvType" (Conversion Type) Constant: "P","I","J", while P:Persian, I:Islamic, J:Julian(Christian)
    '  3- The "WDMY" (Selected Date elements to be in Function Output) any Combination of "W","D","M","Y" , while "W":DayOfWeek, "D":DayOfMonth, "M":MonthOfYear, "Y":YearOfDate 
    Function stDateDesc(ByVal stDate As String, ByVal DateConvType As DateConvType1, ByVal WDMY As String)
        Dim stDateDescIn As String = ""
        Dim JDate As Date
        Dim DayOfWeek As String = ""
        Dim DayOfMonth As String = ""
        Dim MonthOfYear As String = ""
        Dim YearOfDate As String = ""
        Dim MonthRow As Integer = 0
        stDate = stDate.Trim()

        If InStr(WDMY, "W") > 0 Then DayOfWeek = "true"
        If InStr(WDMY, "D") > 0 Then DayOfMonth = "true"
        If InStr(WDMY, "M") > 0 Then MonthOfYear = "true"
        If InStr(WDMY, "Y") > 0 Then YearOfDate = "true"
        If DateConvType = DateConvType1.Julian_J Then
            stDateDescIn = stDate
            MonthRow = 1
        Else
            Select Case DateConvType
                Case DateConvType1.Persian_J
                    stDateDescIn = stDateConvert(stDate, DateConvType2.Persian_Julian)
                    MonthRow = 0
                Case DateConvType1.Islamic_J
                    stDateDescIn = stDateConvert(stDate, DateConvType2.Islamic_Julian)
                    MonthRow = 2
            End Select
        End If
        JDate = stDateDescIn
        If DayOfWeek = "true" Then DayOfWeek = weekdays(CInt(JDate.DayOfWeek))
        If DayOfMonth = "true" Then DayOfMonth = CInt(stDate.Substring(8, 2)).ToString
        If MonthOfYear = "true" Then MonthOfYear = yearmonths(MonthRow, (CInt(stDate.Substring(5, 2)) - 1))
        If YearOfDate = "true" Then YearOfDate = stDate.Substring(0, 4)
        stDateDesc = DayOfWeek & " " & DayOfMonth & " " & MonthOfYear & " " & YearOfDate
    End Function
End Module

Module Module1
    Public Const ISO_8601 = 1
    Public Const Gregorian = ISO_8601
    Sub islamic_persian(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_persian(islamic_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub islamic_julian(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_julian(islamic_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub islamic_civil(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_civil(islamic_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub persian_islamic(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_islamic(persian_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub persian_julian(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_julian(persian_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub persian_civil(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_civil(persian_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub civil_persian(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_persian(civil_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub civil_islamic(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_islamic(civil_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub julian_persian(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Call jdn_persian(julian_jdn(iYear, iMonth, iDay), iYear, iMonth, iDay)
    End Sub
    Sub jdn_civil2(ByRef jdn As Integer, ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)

        Dim l As Integer

        Dim n As Integer
        Dim i As Integer
        Dim j As Integer

        If (jdn > 2299160) Then
            l = jdn + 68569
            n = ((4 * l) \ 146097)
            l = l - ((146097 * n + 3) \ 4)
            i = ((4000 * (l + 1)) \ 1461001)
            l = l - ((1461 * i) \ 4) + 31
            j = ((80 * l) \ 2447)
            iDay = l - ((2447 * j) \ 80)
            l = (j \ 11)
            iMonth = j + 2 - 12 * l
            iYear = 100 * (n - 49) + i + l
        Else
            Call jdn_julian(jdn, iYear, iMonth, iDay)
        End If

    End Sub

    ' Given a julian day number, compute corresponding Hijri date.
    ' As a reference point, the routine uses the fact that the iYear
    ' 1405 A.H. started immediatly after lunar conjunction number 1048
    ' which occured on September 1984 25d 3h 10m UT.
    '
    Sub jdn_islamic(ByRef jd As Integer, ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Dim mjd As Double
        Dim k As Integer
        Dim hm As Integer

        Call jdn_civil(jd, iYear, iMonth, iDay)
        k = Int(0.6 + (iYear + (CShort(iMonth - 0.5)) / 12.0# + iDay / 365.0# - 1900) * 12.3685)
        Do
            mjd = visibility(k)
            k = k - 1
        Loop While (mjd > (jd - 0.5))
        k = k + 1
        hm = k - 1048
        iYear = 1405 + Fix(hm / 12)
        'iYear = 1405 + Int(hm / 12)

        iMonth = (hm Mod 12) + 1
        If (hm <> 0 And iMonth <= 0) Then
            iMonth = iMonth + 12
            iYear = iYear - 1
        End If
        If iYear <= 0 Then iYear = iYear - 1
        iDay = Int(jd - mjd + 0.5)
    End Sub
    Sub jdn_persian(ByRef jdn As Integer, ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Dim depoch As Object
        Dim cycle As Object
        Dim cyear As Object
        Dim ycycle As Object
        Dim aux1, aux2 As Object
        Dim yday As Object
        'UPGRADE_WARNING: Couldn't resolve default property of object depoch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        depoch = jdn - persian_jdn(475, 1, 1)
        'UPGRADE_WARNING: Couldn't resolve default property of object depoch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        cycle = Fix(depoch / 1029983)
        'UPGRADE_WARNING: Couldn't resolve default property of object depoch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'UPGRADE_WARNING: Couldn't resolve default property of object cyear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        cyear = depoch Mod 1029983
        'UPGRADE_WARNING: Couldn't resolve default property of object cyear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If cyear = 1029982 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ycycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ycycle = 2820
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object cyear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object aux1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            aux1 = Fix(cyear / 366)
            'UPGRADE_WARNING: Couldn't resolve default property of object cyear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object aux2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            aux2 = cyear Mod 366
            'UPGRADE_WARNING: Couldn't resolve default property of object aux1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object aux2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object ycycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ycycle = Int(((2134 * aux1) + (2816 * aux2) + 2815) / 1028522) + aux1 + 1
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object cycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object ycycle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iYear = ycycle + (2820 * cycle) + 474
        If iYear <= 0 Then
            iYear = iYear - 1
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object yday. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yday = (jdn - persian_jdn(iYear, 1, 1)) + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object yday. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If yday <= 186 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object yday. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            iMonth = Ceil(yday / 31)
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object yday. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            iMonth = Ceil((yday - 6) / 30)
        End If
        iDay = (jdn - persian_jdn(iYear, iMonth, 1)) + 1
    End Sub
    Sub jdn_julian(ByRef jdn As Integer, ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)
        Dim l As Integer
        Dim k As Integer
        Dim n As Integer
        Dim i As Integer
        Dim j As Integer

        j = jdn + 1402
        k = ((j - 1) \ 1461)
        l = j - 1461 * k
        n = ((l - 1) \ 365) - (l \ 1461)
        i = l - 365 * n + 30
        j = ((80 * i) \ 2447)
        iDay = i - ((2447 * j) \ 80)
        i = (j \ 11)
        iMonth = j + 2 - 12 * i
        iYear = 4 * k + n + i - 4716

    End Sub
    Sub jdn_civil(ByRef jdn As Integer, ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short)

        Dim l As Integer
        Dim k As Integer
        Dim n As Integer
        Dim i As Integer
        Dim j As Integer

        If (jdn > 2299160) Then
            l = jdn + 68569
            n = ((4 * l) \ 146097)
            l = l - ((146097 * n + 3) \ 4)
            i = ((4000 * (l + 1)) \ 1461001)
            l = l - ((1461 * i) \ 4) + 31
            j = ((80 * l) \ 2447)
            iDay = l - ((2447 * j) \ 80)
            l = (j \ 11)
            iMonth = j + 2 - 12 * l
            iYear = 100 * (n - 49) + i + l
        Else
            Call jdn_julian(jdn, iYear, iMonth, iDay)
        End If

    End Sub

    Function islamic_jdn(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short) As Integer

        ' NMONTH is the number of months between julian day number 1 and
        ' the iYear 1405 A.H. which started immediatly after lunar
        ' conjunction number 1048 which occured on September 1984 25d
        ' 3h 10m UT.
        Const NMONTHS As Integer = (1405 * 12 + 1)

        Dim k As Integer

        If (iYear < 0) Then iYear = iYear + 1
        k = iMonth + iYear * 12 - NMONTHS ' nunber of months since 1/1/1405
        islamic_jdn = Int(visibility(k + CInt(1048)) + iDay + 0.5)
    End Function

    ' Determine Julian day from Persian date
    Function persian_jdn(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short) As Integer
        Const PERSIAN_EPOCH As Integer = 1948321 ' The JDN of 1 Farvardin 1
        Dim epbase As Integer
        Dim epyear As Integer
        Dim mdays As Integer
        If iYear >= 0 Then
            epbase = iYear - 474
        Else
            epbase = iYear - 473
        End If
        epyear = 474 + (epbase Mod 2820)
        If iMonth <= 7 Then
            mdays = (CInt(iMonth) - 1) * 31
        Else
            mdays = (CInt(iMonth) - 1) * 30 + 6
        End If
        persian_jdn = CInt(iDay) + mdays + Fix(((epyear * 682) - 110) / 2816) + (epyear - 1) * 365 + Fix(epbase / 2820) * 1029983 + (PERSIAN_EPOCH - 1)
    End Function
    Function civil_jdn(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short, Optional ByRef CalendarType As Short = Gregorian) As Integer
        Dim lYear As Integer
        Dim lMonth As Integer
        Dim lDay As Integer

        If CalendarType = Gregorian And ((iYear > 1582) Or ((iYear = 1582) And (iMonth > 10)) Or ((iYear = 1582) And (iMonth = 10) And (iDay > 14))) Then
            lYear = CInt(iYear)
            lMonth = CInt(iMonth)
            lDay = CInt(iDay)
            civil_jdn = ((1461 * (lYear + 4800 + ((lMonth - 14) \ 12))) \ 4) + ((367 * (lMonth - 2 - 12 * (((lMonth - 14) \ 12)))) \ 12) - ((3 * (((lYear + 4900 + ((lMonth - 14) \ 12)) \ 100))) \ 4) + lDay - 32075
        Else
            civil_jdn = julian_jdn(iYear, iMonth, iDay)
        End If

    End Function
    Function julian_jdn(ByRef iYear As Short, ByRef iMonth As Short, ByRef iDay As Short) As Integer
        Dim lYear As Integer
        Dim lMonth As Integer
        Dim lDay As Integer

        lYear = CInt(iYear)
        lMonth = CInt(iMonth)
        lDay = CInt(iDay)

        julian_jdn = 367 * lYear - ((7 * (lYear + 5001 + ((lMonth - 9) \ 7))) \ 4) + ((275 * lMonth) \ 9) + lDay + 1729777

    End Function
    ' parameters for Makkah: for a new moon to be visible after sunset on
    ' a the same day in which it started, it has to have started before
    ' (SUNSET-MINAGE)-TIMZ=3 A.M. local time.
    Function visibility(ByRef n As Integer) As Double

        ' parameters for Makkah: for a new moon to be visible after sunset on
        ' a the same day in which it started, it has to have started before
        ' (SUNSET-MINAGE)-TIMZ=3 A.M. local time.
        Const TIMZ As Double = 3.0#
        Const MINAGE As Double = 13.5
        Const SUNSET As Double = 19.5 ' approximate
        Const TIMDIF As Double = (SUNSET - MINAGE)

        Dim jd As Double
        Dim tf As Single
        Dim d As Integer

        jd = tmoonphase(n, 0)
        d = Int(jd)
        tf = (jd - d)
        If (tf <= 0.5) Then ' new moon starts in the afternoon
            visibility = (jd + 1.0#)
        Else ' new moon starts before noon
            tf = (tf - 0.5) * 24 + TIMZ ' local time
            If (tf > TIMDIF) Then
                visibility = (jd + 1.0#) ' age at sunset < min for visiblity
            Else
                visibility = (jd)
            End If
        End If
    End Function
    Private Function Ceil(ByRef number As Single) As Integer
        Ceil = -System.Math.Sign(number) * Int(-System.Math.Abs(number))
        ' or
        'Ceil = CInt(number + (Sgn(number) * 0.5))
    End Function
    ' Given an integer _n_ and a phase selector (nph=0,1,2,3 for
    ' new,first,full,last quarters respectively, function returns the
    ' Julian date/time (integer part is the julian day number,
    ' fraction is the time) of the Nth such phase since January 1900.
    ' Adapted from "Astronomical  Formulae for Calculators" by
    ' Jean Meeus, Third Edition, Willmann-Bell, 1985.
    Function tmoonphase(ByRef n As Integer, ByRef nph As Short) As Double

        Const RPD As Double = (0.0174532925199433) ' radians per degree (pi/180)

        Dim jd As Double
        Dim t As Double
        Dim t2 As Double
        Dim t3 As Double
        Dim k As Double
        Dim ma As Double
        Dim sa As Double
        Dim tf As Double
        Dim xtra As Double

        k = n + nph / 4.0#
        t = k / 1236.85
        t2 = t * t
        t3 = t2 * t
        jd = 2415020.75933 + 29.53058868 * k - 0.0001178 * t2 - 0.000000155 * t3 + 0.00033 * System.Math.Sin(RPD * (166.56 + 132.87 * t - 0.009173 * t2))
        '
        '   Sun's mean anomaly
        sa = RPD * (359.2242 + 29.10535608 * k - 0.0000333 * t2 - 0.00000347 * t3)
        '
        '   Moon's mean anomaly
        ma = RPD * (306.0253 + 385.81691806 * k + 0.0107306 * t2 + 0.00001236 * t3)

        '
        '   Moon's argument of latitude
        tf = RPD * 2.0# * (21.2964 + 390.67050646 * k - 0.0016528 * t2 - 0.00000239 * t3)
        '
        '   should reduce to interval 0-1.0 before calculating further
        Select Case nph
            Case 0, 2
                xtra = (0.1734 - 0.000393 * t) * System.Math.Sin(sa) + 0.0021 * System.Math.Sin(sa * 2) - 0.4068 * System.Math.Sin(ma) + 0.0161 * System.Math.Sin(2 * ma) - 0.0004 * System.Math.Sin(3 * ma) + 0.0104 * System.Math.Sin(tf) - 0.0051 * System.Math.Sin(sa + ma) - 0.0074 * System.Math.Sin(sa - ma) + 0.0004 * System.Math.Sin(tf + sa) - 0.0004 * System.Math.Sin(tf - sa) - 0.0006 * System.Math.Sin(tf + ma) + 0.001 * System.Math.Sin(tf - ma) + 0.0005 * System.Math.Sin(sa + 2 * ma)
            Case 1, 3
                xtra = (0.1721 - 0.0004 * t) * System.Math.Sin(sa) + 0.0021 * System.Math.Sin(sa * 2) - 0.628 * System.Math.Sin(ma) + 0.0089 * System.Math.Sin(2 * ma) - 0.0004 * System.Math.Sin(3 * ma) + 0.0079 * System.Math.Sin(tf) - 0.0119 * System.Math.Sin(sa + ma) - 0.0047 * System.Math.Sin(sa - ma) + 0.0003 * System.Math.Sin(tf + sa) - 0.0004 * System.Math.Sin(tf - sa) - 0.0006 * System.Math.Sin(tf + ma) + 0.0021 * System.Math.Sin(tf - ma) + 0.0003 * System.Math.Sin(sa + 2 * ma) + 0.0004 * System.Math.Sin(sa - 2 * ma) - 0.0003 * System.Math.Sin(2 * sa + ma)
                If (nph = 1) Then
                    xtra = xtra + 0.0028 - 0.0004 * System.Math.Cos(sa) + 0.0003 * System.Math.Cos(ma)
                Else
                    xtra = xtra - 0.0028 + 0.0004 * System.Math.Cos(sa) - 0.0003 * System.Math.Cos(ma)
                End If
            Case Else
                tmoonphase = 0
                Exit Function
        End Select
        '   convert from Ephemeris Time (ET) to (approximate)Universal Time (UT)
        tmoonphase = jd + xtra - (0.41 + 1.2053 * t + 0.4992 * t2) / 1440
    End Function

End Module


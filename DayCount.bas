Attribute VB_Name = "DayCount"
'
'
'   Copyright (C) 2003  Prof. Jayanth R. Varma, jrvarma@iimahd.ernet.in,
'   Indian Institute of Management, Ahmedabad 380 015, INDIA
'
'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program (see file COPYING); if not, write to the
'   Free Software Foundation, Inc., 59 Temple Place, Suite 330,
'   Boston, MA  02111-1307  USA
'

'Force all variables to be declared with a dim statement
Option Explicit

'This function and the related function day_count implement
'the various day count conventions in the fixed income markets
'Several variants of the ACT/ACT and 30/360 convention are implemented
'
'
'
'

Function year_fraction(from_date As Date, to_date As Date, _
                     basis As String, Optional cpn_date As Date, _
                    Optional freq As Integer)
Dim yto As Integer, yfrom As Integer, new_year_day As Date, leapday As Date
Dim c1 As Date, c2 As Date, dt As Date
If (from_date = to_date) Then
    year_fraction = 0
    Exit Function
End If
If (InStr(basis, "360")) Then
  'This includes all the 30/360 conventions and the ACT/360 conventions
  'In all of these the denominator is simply 360
  'The computation of the numerator is documented in the function day_count
  year_fraction = day_count(from_date, to_date, basis, cpn_date, freq) / 360
ElseIf (InStr(basis, "365")) Then
  'This includes only the ACT/365 conventions
  'In all of these the denominator is simply 365
  'The computation of the numerator is simply a subtraction of dates
  'But for design modularity, we call the day_count function to do this subtraction
  year_fraction = day_count(from_date, to_date, basis, cpn_date, freq) / 365
Else
' That leaves us with the ACT/ACT conventions. There are three
' variants of this convention as documented in
' http://www.isda.org/c_and_a/pdf/mktc1198.pdf
' They different on the denominator daycount but
' the numerator is the actual number of days between
' from_date and to_date
  Select Case basis
    Case "ACT/ACT-ISMA"
        ' In this method the denominator is the number of days
        ' in the coupon period. This division gives us a fraction
        ' of the coupon period. Dividing by coupon frequency gives
        ' the fraction of a year
        ' This convention is also known as ACT/ACT Bonds
        If (cpn_date > to_date) Then
            c1 = from_date
            c2 = cpn_date
            dt = to_date
        Else
            c1 = cpn_date
            c2 = to_date
            dt = from_date
        End If
        year_fraction = (to_date - from_date) / (c2 - c1) / freq
    Case "ACT/ACT-ISDA"
        'Here the denominator is the number of days in the given calendar year
        'If the interval between from_date and to_date spans two calendar year
        'it is broken up into two subperiods and for each of these the
        'appropriate denominator is used
        'This convention is also known as ACT/ACT Historical
        yfrom = year(from_date)
        yto = year(to_date)
        If (yfrom = yto) Then
            year_fraction = (to_date - from_date) / year_length(yto)
        Else
            new_year_day = DateSerial(yto, 1, 1)
            year_fraction = (new_year_day - from_date) / year_length(yfrom) + _
                            (to_date - new_year_day) / year_length(yto)
        End If
    Case "ACT/ACT-EURO"
        'In this the denominator depends on whether the time interval
        'includes a leap day (denominator=366) or not (denominator=365)
        'This convention is also known as ACT/ACT AFB
        yfrom = year(from_date)
        yto = year(to_date)
        If Not (is_leap_year(yfrom) Or is_leap_year(yto)) Then
            year_fraction = (to_date - from_date) / 365
        Else
            If (is_leap_year(yfrom)) Then
                leapday = DateSerial(yfrom, 2, 29)
            Else
                leapday = DateSerial(yto, 2, 29)
            End If
            If ((from_date - leapday) * (to_date - leapday) < 0) Then
            'from_date and to_date are on opposite sides of the leapday
                year_fraction = (to_date - from_date) / 366
            Else
                year_fraction = (to_date - from_date) / 365
            End If
        End If
    Case Else
        year_fraction = "Invalid Basis " & basis
  End Select
End If
End Function
Function day_count(from_date As Date, to_date As Date, _
                     basis As String, Optional cpn_date As Date, _
                    Optional freq As Integer)
Dim y1 As Integer, y2 As Integer, m1 As Integer, m2 As Integer, d1 As Integer, _
    d2 As Integer
If (from_date = to_date) Then
    day_count = 0
    Exit Function
End If
If (InStr(basis, "ACT/")) Then
  day_count = to_date - from_date
Else
  Select Case basis
    Case "30/360-US"
    'We use the ISDA defintion reproduced in
    'http://www.fpml.org/spec/2003/tr-fpml-3-0-2003-06-27/html/schemeDefinitions.html#dayCountFractionScheme
    'Per Annex to the 2000 ISDA Definitions (June 2000 Version), Section 4.16.
    'Day Count Fraction, paragraph (e), i.e.
    'if "30/360", "360/360" or "Bond Basis" is specified, the number of days in the
    'Calculation Period or Compounding Period in respect of which payment is being made
    'divided by 360 (the number of days to be calculated on the basis of a year of 360
    'days with 12 30-day months (unless (i) the last day of the Calculation Period or
    'Compounding Period is the 31st day of a month but the first day of the Calculation
    'Period or Compounding Period is a day other than the 30th or 31st day of a month,
    'in which case the month that includes that last day shall not be considered to be
    'shortened to a 30-day month, or (ii) the last day of the Calculation Period or
    'Compounding Period is the last day of the month of February, in which case the
    'month of February shall not be considered to be lengthened to a 30-day month)).
    '
    '
    ' We supplement these with the Wikepedia rules that deal with February
    ' http://en.wikipedia.org/wiki/Day_count_convention
    ' EOM refers to bonds that pay interest on the last day of the month
    ' for example 28 February and 31 August rather than 28 Feb and 28 Aug
    ' EOM is assumed in the code below
    ' If the investment is EOM and (D1 is the last day of February)
    ' and (D2 is the last day of February), then change D2 to 30.
    ' If the investment is EOM and (D1 is the last day of February),
    ' then change D1 to 30.
    ' If D2 is 31 and D1 is 30 or 31, then change D2 to 30.
    ' If D1 is 31, then change D1 to 30.

        y1 = year(from_date)
        d1 = Day(from_date)
        m1 = Month(from_date)
        y2 = year(to_date)
        d2 = Day(to_date)
        m2 = Month(to_date)
        ' If (D1 is the last day of February)
        ' and (D2 is the last day of February), then change D2 to 30.
        If (m1 = 2 And is_end_of_month(from_date) _
            And m2 = 2 And is_end_of_month(to_date)) Then
            d2 = 30
        End If
        ' If (D1 is the last day of February), then change D1 to 30.
        If (m1 = 2 And is_end_of_month(from_date)) Then
            d1 = 30
        End If
        ' If D2 is 31 and D1 is 30 or 31, then change D2 to 30.
        If (d2 = 31 And d1 >= 30) Then
            d2 = 30
        End If
        ' If D1 is 31, then change D1 to 30.
        If (d1 = 31) Then
            d1 = 30
        End If
        'Below is old code that is commented out
            'Following 3 lines ensure that the start month is treated as 30 days and it
            'changes a start date of February 28 into 30
            'If (is_end_of_month(from_date)) Then
            '    d1 = 30
            'End If
        'Above is old code that is commented out
        day_count = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    Case "30/360-MSRB"
    'This is rule G-33 of the Municipal Securities Rulemaking Board
    'http://ww1.msrb.org/msrb1/rules/ruleg33.htm
    'Number of Days = (Y2 - Y1) 360 + (M2 - M1) 30 + (D2 - D1)
    'For purposes of this formula the symbols shall be defined as follows:
    '"M1" is the month of the date on which the computation period begins;
    '"D1" is the day of the date on which the computation period begins;
    '"Y1" is the year of the date on which the computation period begins;
    '"M2" is the month of the date on which the computation period ends;
    '"D2" is the day of the date on which the computation period ends; and
    '"Y2" is the year of the date on which the computation period ends.
    'For purposes of this formula, if the symbol "D2" has a value of "31," and the
    'symbol "D1" has a value of "30" or "31," the value of the symbol "D2" shall be
    'changed to "30." If the symbol "D1" has a value of "31," the value of the symbol
    '"D1" shall be changed to "30." For purposes of this rule time periods shall be
    'computed to include the day specified in the rule for the beginning of the period
    'but not to include the day specified for the end of the period.
    'This gives a different answer from 30/360-US if the start date is the last day of
    'February. From the 28th Feb  to 5th March there are 7 days by this method,
    'but only 5 days by the standard 30/360
        y1 = year(from_date)
        d1 = Day(from_date)
        m1 = Month(from_date)
        y2 = year(to_date)
        d2 = Day(to_date)
        m2 = Month(to_date)
        If (d1 = 31) Then
            d1 = 30
        End If
        If (d2 = 31 And d1 = 30) Then
            d2 = 30
        End If
        day_count = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    Case "30E/360"
    'We use the ISDA defintion reproduced in
    'http://www.fpml.org/spec/2003/tr-fpml-3-0-2003-06-27/html/schemeDefinitions.html#dayCountFractionScheme
    'Per Annex to the 2000 ISDA Definitions (June 2000 Version), Section 4.16.
    'Day Count Fraction, paragraph (f), i.e.
    'if "30E/360" or "Eurobond Basis" is specified, the number of days in the Calculation
    'Period or Compounding Period in respect of which payment is being made divided by
    '360 (the number of days to be calculated on the basis of a year of 360 days with 12
    '30-day months, without regard to the date of the first day or last day of the
    'Calculation Period or Compounding Period unless, in the case of the final Calculation
    'Period or Compounding Period, the Termination Date is the last day of the month of
    'February, in which case the month of February shall not be considered to be lengthened
    'to a 30-day month).
        y1 = year(from_date)
        d1 = Day(from_date)
        m1 = Month(from_date)
        y2 = year(to_date)
        d2 = Day(to_date)
        m2 = Month(to_date)
        If (d1 = 31) Then
            d1 = 30
        End If
        If (d2 = 31) Then
            d2 = 30
        End If
        day_count = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    Case "30E+/360"
        'This differs from 30E/360 in that if the last day of the calculation period
        'is the 31st of a month, it is changed to the 1st of the next month
        y1 = year(from_date)
        d1 = Day(from_date)
        m1 = Month(from_date)
        If (d1 = 31) Then
            d1 = 30
        End If
        d2 = Day(to_date)
        If (d2 = 31) Then
            d2 = 1
            y2 = year(to_date + 1)
            m2 = Month(to_date + 1)
        Else
            y2 = year(to_date)
            m2 = Month(to_date)
        End If
        day_count = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    Case "30/360-Strict"
        'This is a strict implementation of the principle that all months have 30 days
        'While other methods exclude February from this rule when the calculation period
        'ends on the last day of February, this method makes no such exception
        'The major financial markets do not use this method
        'Perhaps Italy does?
        y1 = year(from_date)
        d1 = Day(from_date)
        m1 = Month(from_date)
        y2 = year(to_date)
        d2 = Day(to_date)
        m2 = Month(to_date)
        If (is_end_of_month(from_date)) Then
            d1 = 30
        End If
        If (is_end_of_month(to_date)) Then
            d2 = 30
        End If
        day_count = (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
    Case Else
        day_count = "Invalid Basis " & basis
  End Select
End If
End Function
Function is_leap_year(year As Integer) As Boolean
'It is a leap year if it has 366 days
is_leap_year = (year_length(year) = 366)
End Function

Function year_length(year As Integer) As Integer
'Count the number of days from the beginning of the year to the beginning of the next year
year_length = DateSerial(year + 1, 1, 1) - DateSerial(year, 1, 1)
End Function
Function is_end_of_month(dt As Date) As Boolean
'It is end of month if tomorrow is a different month
is_end_of_month = (Month(dt + 1) <> Month(dt))
End Function
Function month_length(y As Integer, m As Integer) As Integer
Dim bom As Date, bonm As Date
'Count the number of days from the beginning of the month to the beginning of the next month
bom = DateSerial(y, m, 1)
If m = 12 Then
    bonm = DateSerial(y + 1, 1, 1)
Else
    bonm = DateSerial(y, m + 1, 1)
End If
month_length = bonm - bom
End Function



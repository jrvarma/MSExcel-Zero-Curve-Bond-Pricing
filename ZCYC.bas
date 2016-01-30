Attribute VB_Name = "ZCYC"
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
' This function computes the price of a bond using a zero coupon yield
' curve and a spread to be added to that yield curve
' For example, we might take a risk free yield curve and add a spread
' to value a corporate bond
' Apart from the bond characteristics (maturity, coupon, redemption value
' and coupon frequency), the function needs the settlement date of the
' trade, the day count convention to be used as well as the yield curve
' and spread
' The day count conventions are documented under the function day_count
'
Function bond_price(settle As Date, mature As Date, _
    cpn_rate As Double, yld_curve As Range, spread As Double, _
    yield_type As String, redemption As Double, _
    freq As Integer, basis As String, _
    Optional Excel_Compatible As Boolean = False)
Dim next_coupon As Date, prev_coupon As Date, fraction As Double
Dim accrue_fraction As Double, accrint As Double, annual_rate As Double
Dim ncoup As Integer, coupon As Double, i As Integer, pvthis As Double
Dim yield As Double
'fraction of year to next coupon
next_coupon = next_cpn(settle, mature, freq)
prev_coupon = prev_cpn(settle, mature, freq)
'compute accrued interest
accrue_fraction = year_fraction(prev_coupon, settle, basis, next_coupon, freq)
accrint = 100 * cpn_rate * accrue_fraction
If Excel_Compatible Then
    'Excel does not compute time to next coupon
    'It computes 1 - accrual fraction and adjusts for frequency
    fraction = 1 / freq - accrue_fraction
Else
    fraction = year_fraction(settle, next_coupon, basis, prev_coupon, freq)
End If
' find the number of coupons excluding the one at maturity
ncoup = number_coupons(settle, mature, freq)
coupon = 100 * cpn_rate / freq
bond_price = 0
' discount and add each stand alone coupon
' Discounting requires two steps
' 1. Interpolate the zero yield curve to get the zero rate to use
' 2. Convert this zero rate into a discount factor
' The function interpol_yld does the first job
' The function discount_factor does the second job
For i = 1 To ncoup
    yield = interpol_yld(yld_curve, fraction + (i - 1) / freq) + spread
    pvthis = coupon * discount_factor(yield, fraction + (i - 1) / freq, yield_type, freq)
    bond_price = bond_price + pvthis
                    
Next i
' discount and add redemption value and final coupon
If (ncoup = 0) Then
    ' use simple interest (short governments convention = money market convention)
    ' we interpolate the yield curve to get the rate
    ' convert it into an equivalent annual rate if the yield curve
    ' is based on continuous compounding
    annual_rate = annual_simple_interest_rate(interpol_yld(yld_curve, 1) + spread, _
                                            yield_type)
    bond_price = (redemption + coupon) / (1 + annual_rate * fraction)
Else
    ' if there is a coupon to go before maturity
    ' use compound interest rates for discounting
    yield = interpol_yld(yld_curve, fraction + ncoup / freq) + spread
    bond_price = bond_price + (redemption + coupon) * _
                 discount_factor(yield, fraction + ncoup / freq, yield_type, freq)
End If
' We now have the dirty price. Subtract accrued interest to get clean price
bond_price = bond_price - accrint
End Function

Function bond_price_like_excel(settle As Date, mature As Date, _
    cpn_rate As Double, yld_curve As Range, redemption As Double, _
    freq As Integer, basis As Integer)
Dim basis_string As String
basis_string = Choose(basis + 1, "30/360-US", "ACT/ACT-ISMA", "ACT/360", _
                    "ACT/365", "30E/360")
bond_price_like_excel = bond_price(settle, mature, cpn_rate, yld_curve, 0, "custom-freq", _
            redemption, freq, basis_string, True)

End Function
Function number_coupons(settle As Date, mature As Date, freq As Integer) As Integer
'construct an approximation (upper bound )
number_coupons = Int(freq * (mature - settle) / 365) + 1
'and then refine it to get the correct number
While (settle >= nth_cpn(number_coupons, settle, mature, freq))
    number_coupons = number_coupons - 1
Wend
End Function
Function next_cpn(settle As Date, mature As Date, freq As Integer)
Dim y As Integer, m As Integer, d As Integer, years As Integer, months As Integer
Dim n As Integer
n = number_coupons(settle, mature, freq)
next_cpn = nth_cpn(n, settle, mature, freq)
End Function
Function prev_cpn(settle As Date, mature As Date, freq As Integer)
Dim y As Integer, m As Integer, d As Integer, years As Integer, months As Integer
Dim n As Integer, nxt As Date
n = number_coupons(settle, mature, freq)
prev_cpn = nth_cpn(n + 1, settle, mature, freq)
End Function
Function nth_cpn(n As Integer, settle As Date, mature As Date, freq As Integer)
'this is the n'th coupon counting backward from maturity and excluding
'the coupon at maturity
Dim y As Integer, m As Integer, d As Integer, years As Integer, months As Integer
years = Int(n / freq)
months = Int(n * 12 / freq) - 12 * years
y = year(mature)
y = y - years
m = Month(mature)
If (m > months) Then
    m = m - months
Else
    m = m + 12 - months
    y = y - 1
End If
d = Day(mature)
If (is_end_of_month(mature) Or d > month_length(y, m)) Then
    nth_cpn = DateSerial(y, m, month_length(y, m))
Else
    nth_cpn = DateSerial(y, m, d)
End If
End Function
Function interpol_yld(yld_curve As Range, t As Double) As Double
Dim nrows As Integer, n As Integer
Dim beta0 As Double, beta1 As Double, beta2 As Double, tau As Double
If (yld_curve.Columns.Count = 1 And yld_curve.Rows.Count = 1) Then
    interpol_yld = yld_curve.Value
    Exit Function
End If
If (yld_curve.Columns.Count <> 2) Then
    MsgBox ("Yield Curve must have two columns for maturity/parameter and  yield")
    Exit Function
End If
'=B$3+(B$4+B$5)*(1-EXP(-calc!$A7/B$6))/(calc!$A7/B$6)-B$5*(EXP(-calc!$A7/B$6))
If (LCase(yld_curve.Range("A1").Value) = "beta 0" And _
    LCase(yld_curve.Range("A4").Value) = "tau") Then
    'Nelson Seigel Parameters for Yield Curve
    beta0 = yld_curve.Range("B1").Value
    beta1 = yld_curve.Range("B2").Value
    beta2 = yld_curve.Range("B3").Value
    tau = yld_curve.Range("B4").Value
    interpol_yld = beta0 + (beta1 + beta2) * (1 - Exp(-t / tau)) / (t / tau) - _
                   beta2 * Exp(-t / tau)
    interpol_yld = interpol_yld / 100
    Exit Function
End If
nrows = yld_curve.Rows.Count
n = 1
While (Application.Index(yld_curve, n, 1) < t And n < nrows)
    n = n + 1
Wend
If n = 1 Or Application.Index(yld_curve, nrows, 1) < t Then
    interpol_yld = Application.Index(yld_curve, n, 2)
Else
    interpol_yld = Application.Index(yld_curve, n - 1, 2) + _
    (Application.Index(yld_curve, n, 2) - Application.Index(yld_curve, n - 1, 2)) * _
    (t - Application.Index(yld_curve, n - 1, 1)) / _
    (Application.Index(yld_curve, n, 1) - Application.Index(yld_curve, n - 1, 1))
End If
End Function
Function annual_simple_interest_rate(yield As Double, yield_type As String) As Double
If (yield_type = "continuous") Then
        annual_simple_interest_rate = Exp(yield) - 1
Else
        annual_simple_interest_rate = yield
End If
End Function

Function discount_factor(yield As Double, t As Double, yield_type As String, _
                            Optional freq As Integer = 1) As Double
Select Case LCase(yield_type)
    Case "continuous"
        discount_factor = Exp(-yield * t)
    Case "annual"
        discount_factor = Exp(-Log(1 + yield) * t)
    Case "semi-annual"
        discount_factor = Exp(-Log(1 + yield / 2) * 2 * t)
    Case "custom-freq"
        discount_factor = Exp(-Log(1 + yield / freq) * freq * t)
End Select
End Function


Function accr_interest(settle As Date, mature As Date, _
    cpn_rate As Double, freq As Integer, basis As String)
Dim next_coupon As Date, prev_coupon As Date
Dim accrue_fraction As Double
next_coupon = next_cpn(settle, mature, freq)
prev_coupon = prev_cpn(settle, mature, freq)
'compute accrued interest
accrue_fraction = year_fraction(prev_coupon, settle, basis, next_coupon, freq)
accr_interest = 100 * cpn_rate * accrue_fraction
End Function

Public Function ns_zero(beta0 As Double, beta1 As Double, _
                beta2 As Double, tau As Double, t As Double, _
                Optional percent As Boolean = True) As Double
' Compute the zero yield for maturity t from the
' Nelson-Seigel parameters for the zero curve
ns_zero = beta0 + (beta1 + beta2) * (1 - Exp(-t / tau)) / (t / tau) _
        - beta2 * Exp(-t / tau)
If percent Then ns_zero = ns_zero / 100
End Function

Public Function ns_par(beta0 As Double, beta1 As Double, _
                beta2 As Double, tau As Double, t As Double, _
                Optional freq As Integer = 1) _
                As Double
' Compute the par yield for maturity t from the
' Nelson-Seigel parameters for the zero curve
' Compounding  frequency is the same as coupon frequency freq
' We do not require that t is an integer number of coupon periods
Dim delta As Double, t0 As Double, sumdf As Double, yld As Double
Dim df As Double, pvcf As Double
delta = 1# / freq
t0 = 0#
sumdf = 0#
' To compute the present value factor for coupons,
' we sum the zero df's for all coupon dates except the last
While t0 + delta <= t - delta
    t0 = t0 + delta
    ' Compute yield and convert into a df
    yld = ns_zero(beta0, beta1, beta2, tau, t0)
    sumdf = sumdf + discount_factor(yld, t0, "continuous")
Wend
' Compute yield and convert into a df
yld = ns_zero(beta0, beta1, beta2, tau, t)
df = discount_factor(yld, t, "continuous")
' Let the coupon rate be c
' Each coupon except the last is c/freq
' The last possibly short coupon is c*(t-t0)
' We compute the pv of coupon factor such that
'           pv of coupons = c*pvcf
pvcf = sumdf / freq + df * (t - t0)
' We solve for c such that c*pvcf = 1-df
' where df is the pv of the principal redemption
ns_par = (1 - df) / pvcf
' If t < delta, the above yield assumes a compounding
' frequency of t. To correct this to a frequency of delta
' uncomment the following lines
' If (t < delta) Then
'     ns_par = (1 + ns_par) ^ (delta / t) - 1
' End If
' If t is > delta and is not an integer multiple of delta
' ns_par is the coupon rate and not the yield.
' It is difficult to correct this error
End Function

Function BootStrap(Par As Range, Optional freq As Integer = 1, _
        Optional Compounding As String = "custom-freq", _
        Optional percent As Boolean = False)
Dim zero() As Double, n As Integer, df As Double, sumdf As Double
Dim period As Double, i As Integer, par_yld As Double, zyld As Double
Dim r As Integer, c As Integer
n = Par.Cells.Count
r = Par.Rows.Count
c = Par.Columns.Count
ReDim zero(r, c)
period = 1# / freq
sumdf = 0
For i = 1 To n
    par_yld = Par.Cells(i)
    If percent Then par_yld = par_yld / 100
    df = (1# - par_yld * period * sumdf) / (1 + par_yld * period)
    zyld = -Log(df) / (period * i)
    zyld = converted_yield(zyld, Compounding, freq)
    If percent Then zyld = zyld * 100
    If (r = 1) Then
        zero(0, i - 1) = zyld
    Else
        zero(i - 1, 0) = zyld
    End If
    sumdf = sumdf + df
Next i
BootStrap = zero
End Function
Function BootStrapInterpolated(Par As Range, _
        Mat As Range, _
        Optional freq As Integer = 1, _
        Optional Compounding As String = "custom-freq", _
        Optional percent As Boolean = False)
Dim zero() As Double, n As Integer, df As Double, sumdf As Double
Dim period As Double, i As Integer, par_yld As Double, zyld As Double
Dim r As Integer, c As Integer
Dim outarray() As Double, tmax As Double, t As Double
Dim int_t As Integer, frac As Double
If (Par.Columns.Count <> 2) Then
    MsgBox ("Yield Curve must have two columns for maturity and  yield")
    Exit Function
End If
tmax = Application.max(Par.Cells(Par.Cells.Count / 2, 1), _
                       Mat.Cells(Mat.Cells.Count))
n = tmax * freq
ReDim zero(n)
period = 1# / freq
sumdf = 0
For i = 1 To n
    par_yld = interpol_yld(Par, i * period)
    If percent Then par_yld = par_yld / 100
    df = (1# - par_yld * period * sumdf) / (1 + par_yld * period)
    zyld = -Log(df) / (period * i)
    zyld = converted_yield(zyld, Compounding, freq)
    If percent Then zyld = zyld * 100
    zero(i - 1) = zyld
    sumdf = sumdf + df
Next i
zero(n) = zero(n - 1)
r = Mat.Rows.Count
c = Mat.Columns.Count
ReDim outarray(r - 1, c - 1)
For i = 1 To Mat.Cells.Count
    t = Mat.Cells(i)
    int_t = Int(t * freq)
    frac = (t * freq - int_t)
    zyld = zero(int_t - 1) + frac * (zero(int_t) - zero(int_t - 1))
    If (r = 1) Then
        outarray(0, i - 1) = zyld
    Else
        outarray(i - 1, 0) = zyld
    End If
Next i
BootStrapInterpolated = outarray
End Function
Function converted_yield(yield As Double, yield_type As String, _
                            Optional freq As Integer = 1) As Double
Select Case LCase(yield_type)
    Case "continuous"
        converted_yield = yield
    Case "annual"
        converted_yield = Exp(yield) - 1
    Case "semi-annual"
        converted_yield = (Exp(yield / 2) - 1) * 2
    Case "custom-freq"
        converted_yield = (Exp(yield / freq) - 1) * freq
End Select
End Function



Attribute VB_Name = "modOptionPricing"
Option Explicit

Public Function BinomialOptionPrice(S As Double, K As Double, r As Double, q As Double, _
    sigma As Double, T As Double, n As Integer, _
    isCall As Boolean, isAmerican As Boolean) As Double

    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    Dim i As Integer, j As Integer
    Dim OptVal() As Double, Stock() As Double

    dt = T / n
    
    ' === Your New u and d factors ===
    u = Exp((r - q) * dt + (sigma * Sqr(dt)))
    d = Exp((r - q) * dt - (sigma * Sqr(dt)))
    
    ' === p formula kept as before ===
    p = (Exp((r - q) * dt) - d) / (u - d)
    disc = Exp(-r * dt)

    ReDim Stock(0 To n, 0 To n)
    ReDim OptVal(0 To n, 0 To n)

    ' Build Stock Tree
    For j = 0 To n
        For i = 0 To j
            Stock(i, j) = S * (u ^ (j - i)) * (d ^ i)
        Next i
    Next j

    ' Terminal Payoffs
    For i = 0 To n
        If isCall Then
            OptVal(i, n) = Application.Max(0, Stock(i, n) - K)
        Else
            OptVal(i, n) = Application.Max(0, K - Stock(i, n))
        End If
    Next i

    ' Backward Induction
    For j = n - 1 To 0 Step -1
        For i = 0 To j
            Dim HoldVal As Double: HoldVal = disc * (p * OptVal(i, j + 1) + (1 - p) * OptVal(i + 1, j + 1))
            Dim ExVal As Double
            If isCall Then ExVal = Stock(i, j) - K Else ExVal = K - Stock(i, j)
            
            If isAmerican Then
                OptVal(i, j) = Application.Max(HoldVal, ExVal)
            Else
                OptVal(i, j) = HoldVal
            End If
        Next i
    Next j
    
    BinomialOptionPrice = OptVal(0, 0)
End Function

Public Function BlackScholes(S As Double, K As Double, r As Double, q As Double, _
    sigma As Double, T As Double, isCall As Boolean) As Double
    Dim d1 As Double, d2 As Double
    d1 = (Log(S / K) + (r - q + 0.5 * sigma ^ 2) * T) / (sigma * Sqr(T))
    d2 = d1 - sigma * Sqr(T)
    If isCall Then
        BlackScholes = S * Exp(-q * T) * Application.WorksheetFunction.Norm_S_Dist(d1, True) - _
                       K * Exp(-r * T) * Application.WorksheetFunction.Norm_S_Dist(d2, True)
    Else
        BlackScholes = K * Exp(-r * T) * Application.WorksheetFunction.Norm_S_Dist(-d2, True) - _
                       S * Exp(-q * T) * Application.WorksheetFunction.Norm_S_Dist(-d1, True)
    End If
End Function

Public Sub OpenMyForm()
    ' This line opens the UserForm
    ' Make sure "frmoptionpricing" matches the (Name) in your Properties window
    frmOptionPricing.Show
End Sub


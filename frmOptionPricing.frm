VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionPricing 
   Caption         =   "Option Pricing Model"
   ClientHeight    =   8460.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14016
   OleObjectBlob   =   "frmOptionPricing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptionPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub lbl_2_0_Click()

End Sub

Private Sub lbl_3_2_Click()

End Sub

Private Sub lbl_4_2_Click()

End Sub

Private Sub optBS_Click()

End Sub

Private Sub txtS0_Change()

End Sub

Private Sub UserForm_Click()
Option Explicit

Private Sub UserForm_Initialize()
    cboOptionType.AddItem "European Call"
    cboOptionType.AddItem "European Put"
    cboOptionType.AddItem "American Call"
    cboOptionType.AddItem "American Put"
    cboOptionType.ListIndex = 0
End Sub

Private Sub cmdCalculate_Click()
    Dim S0 As Double, K As Double, r As Double, q As Double, sigma As Double
    Dim T As Double, n As Integer
    Dim priceBin As Double, priceBS As Double
    Dim isCall As Boolean, isAmerican As Boolean
    Dim diff As Double
    Dim resultMsg As String

    On Error GoTo ErrorHandler
    
    ' 1. Load values
    S0 = CDbl(txtS0.Value): K = CDbl(txtK.Value): r = CDbl(txtR.Value)
    sigma = CDbl(txtSigma.Value): T = CDbl(txtT.Value): n = CInt(txtN.Value): q = CDbl(txtQ.Value)

    isCall = (InStr(1, cboOptionType.Value, "Call", vbTextCompare) > 0)
    isAmerican = (InStr(1, cboOptionType.Value, "American", vbTextCompare) > 0)

    ' 2. Conditional Calculation
    ' Clear results first
    resultMsg = ""
    
    ' Only calculate Binomial if Binomial or Both is selected
    If optBinomial.Value Or optBoth.Value Then
        priceBin = BinomialOptionPrice(S0, K, r, q, sigma, T, n, isCall, isAmerican)
        resultMsg = "Binomial Price: " & Format(priceBin, "0.0000")
    End If

    ' Only calculate BS if BS or Both is selected
    If optBS.Value Or optBoth.Value Then
        priceBS = BlackScholes(S0, K, r, q, sigma, T, isCall)
        ' Add a line break if we already have text from Binomial
        If resultMsg <> "" Then resultMsg = resultMsg & vbCrLf
        resultMsg = resultMsg & "Black-Scholes Price: " & Format(priceBS, "0.0000")
    End If

    ' 3. Convergence Analysis (Only if "Both" is selected)
    If optBoth.Value Then
        diff = Abs(priceBin - priceBS)
        resultMsg = resultMsg & vbCrLf & "----------------------------" & vbCrLf
        resultMsg = resultMsg & "Price Difference: " & Format(diff, "0.0000") & vbCrLf & vbCrLf
        
        If isAmerican Then
            resultMsg = resultMsg & "Note: American options may never exactly equal Black-Scholes " & _
                        "due to the early exercise premium."
        Else
            ' Use Chr(236) or "8" for the infinity symbol
            resultMsg = resultMsg & "Convergence: As n -> 8, the Binomial price will equal Black-Scholes." & vbCrLf & _
                        "At n=" & n & ", the model is " & Format(1 - (diff / priceBS), "0.0%") & " accurate."
        End If
    End If

    ' 4. Update UI
    LblResult.Caption = resultMsg

    ' 5. Draw Tree (Usually only for Binomial or Both)
    If optBinomial.Value Or optBoth.Value Then
        Call DrawBinomialTree(S0, K, r, q, sigma, T, n, isCall, isAmerican)
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Check your inputs! Error: " & Err.Description
End Sub
Private Sub DrawBinomialTree(S0 As Double, K As Double, r As Double, q As Double, _
                             sigma As Double, T As Double, n As Integer, _
                             isCall As Boolean, isAmerican As Boolean)
    Dim dt As Double, u As Double, d As Double, p As Double, disc As Double
    Dim i As Integer, j As Integer
    Dim ws As Worksheet
    
    ' Set Worksheet
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("BinomialTree")
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add
        ws.Name = "BinomialTree"
    End If
    On Error GoTo 0

    ' 1. Math Setup
    dt = T / n
    u = Exp((r - q) * dt + (sigma * Sqr(dt)))
    d = Exp((r - q) * dt - (sigma * Sqr(dt)))
    p = (Exp((r - q) * dt) - d) / (u - d)
    disc = Exp(-r * dt)

    Dim S_Tree() As Double, V_Tree() As Double
    ReDim S_Tree(0 To n, 0 To n): ReDim V_Tree(0 To n, 0 To n)

    ' Calculate Trees
    For i = 0 To n: For j = 0 To i
        S_Tree(i, j) = S0 * (u ^ (i - j)) * (d ^ j)
    Next j: Next i
    
    For j = 0 To n
        If isCall Then V_Tree(n, j) = Application.Max(S_Tree(n, j) - K, 0) Else V_Tree(n, j) = Application.Max(K - S_Tree(n, j), 0)
    Next j
    
    For i = n - 1 To 0 Step -1: For j = 0 To i
        V_Tree(i, j) = disc * (p * V_Tree(i + 1, j) + (1 - p) * V_Tree(i + 1, j + 1))
        If isAmerican Then
            Dim ExVal As Double
            If isCall Then ExVal = S_Tree(i, j) - K Else ExVal = K - S_Tree(i, j)
            V_Tree(i, j) = Application.Max(V_Tree(i, j), ExVal)
        End If
    Next j: Next i

    ' 2. Output and Visual Formatting
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone ' Reset background

    For i = 0 To n: For j = 0 To i
        With ws.Cells(4 + j * 2, 2 + i * 3)
            .Value = "S=" & Format(S_Tree(i, j), "0.00") & vbCrLf & "V=" & Format(V_Tree(i, j), "0.00")
            .Font.Size = 10
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            
            ' --- MULTI-COLOR HEAT MAP LOGIC ---
            
            ' 1. GREEN: Early Exercise Optimal (American Only)
            If isAmerican And i < n Then
                Dim hVal As Double, eVal As Double
                hVal = disc * (p * V_Tree(i + 1, j) + (1 - p) * V_Tree(i + 1, j + 1))
                If isCall Then eVal = S_Tree(i, j) - K Else eVal = K - S_Tree(i, j)
                
                If Round(eVal, 5) > Round(hVal, 5) And eVal > 0 Then
                    .Interior.Color = RGB(144, 238, 144) ' Bright Green
                    .Font.Bold = True
                    .Borders.Weight = xlThick
                End If
            End If

            ' 2. PINK: Out-of-the-Money (Option is worthless)
            ' Only color pink if not already colored green
            If V_Tree(i, j) = 0 And .Interior.ColorIndex = xlNone Then
                .Interior.Color = RGB(255, 204, 204) ' Light Pink
            End If

            ' 3. LIGHT BLUE: Near Strike Price (At-the-Money)
            ' Highlights nodes where stock is within 2% of Strike
            If Abs(S_Tree(i, j) - K) / K < 0.02 And .Interior.ColorIndex = xlNone Then
                .Interior.Color = RGB(204, 229, 255) ' Light Blue
            End If
            
        End With
    Next j: Next i
    
    ' 3. Professional Legend
    ws.Range("A1").Value = "BINOMIAL TREE ANALYSIS"
    ws.Range("A1").Font.Bold = True
    
    ws.Range("A2").Value = "GREEN: Early Exercise"
    ws.Range("A2").Interior.Color = RGB(144, 238, 144)
    
    ws.Range("A3").Value = "PINK: Worthless (OTM)"
    ws.Range("A3").Interior.Color = RGB(255, 204, 204)
    
    ws.Range("A4").Value = "BLUE: Near Strike (ATM)"
    ws.Range("A4").Interior.Color = RGB(204, 229, 255)
    
    ws.Columns.AutoFit
    ws.Cells.WrapText = True
    ws.Activate
End Sub


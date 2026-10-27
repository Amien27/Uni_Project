Attribute VB_Name = "modMath"
Option Explicit

Public Function Ncdf(x As Double) As Double
    Ncdf = WorksheetFunction.Norm_S_Dist(x, True)
End Function


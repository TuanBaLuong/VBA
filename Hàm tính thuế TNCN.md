# VBA
Code VBA
Function ThueTNCN2020(THUNHAP As Double, cachtinh As Boolean, Optional phuthuoc As Byte = 0, Optional GiamTru As Double = 0) As Double
    Application.Volatile (False)
    THUNHAP = Round(THUNHAP, 0)
    If (cachtinh = True) Then
        THUNHAP = THUNHAP - 9000000 - (phuthuoc * 3600000) - GiamTru
    Else
        THUNHAP = THUNHAP
    End If
    Select Case THUNHAP
        Case 1 To 5000000: ThueTNCN2020 = THUNHAP * 0.05
        Case 5000001 To 10000000: ThueTNCN2020 = THUNHAP * 0.1 - 250000
        Case 10000001 To 18000000: ThueTNCN2020 = THUNHAP * 0.15 - 750000
        Case 18000001 To 32000000: ThueTNCN2020 = THUNHAP * 0.2 - 1650000
        Case 32000001 To 52000000: ThueTNCN2020 = THUNHAP * 0.25 - 3250000
        Case 52000001 To 80000000: ThueTNCN2020 = THUNHAP * 0.3 - 5850000
        Case Is > 80000000: ThueTNCN2020 = THUNHAP * 0.35 - 9850000
    End Select
End Function

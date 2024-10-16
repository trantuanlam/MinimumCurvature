Attribute VB_Name = "Ragland"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
' Dogleg / Toolface related routines
'--------------------------------

Option Explicit

' Nhung thu tuc lien quan toi tinh xien

Function getDoglegRad(I1 As Double, I2 As Double, DeltaAz As Double) As Double
    ' cac tinh toan deu tren rad
    Dim t As Double
    If Abs(DeltaAz) < Zero * Zero Then
'        getDoglegRad = I2 - I1
        getDoglegRad = Abs(I2 - I1)
        Exit Function
    End If
    
    ' Bieu thuc sau doi khi sinh loi khi I1=I2 va DeltaAz =0
    
    t = Cos(I1) * Cos(I2) + Sin(I1) * Sin(I2) * Cos(DeltaAz)
    t = WorksheetFunction.Acos(t)
'    If I2 < I1 Then t = -t
    getDoglegRad = t
End Function

Function getDoglegDeg(ByVal I1 As Double, ByVal I2 As Double, ByVal DeltaAz As Double) As Double
    ' i1, i2, dogleg: degree
    I1 = WorksheetFunction.Radians(I1)
    I2 = WorksheetFunction.Radians(I2)
    DeltaAz = WorksheetFunction.Radians(DeltaAz)
    getDoglegDeg = WorksheetFunction.Degrees(getDoglegRad(I1, I2, DeltaAz))
End Function

Function DLS100Deg(ByVal I1 As Double, ByVal I2 As Double, ByVal DeltaAz As Double, ByVal DL As Double) As Double
  ' DoglegSeverity
  ' Cac gia tri goc tinh bang Degree
  ' Ket qua: Degree/100m
    DLS100Deg = getDoglegDeg(I1, I2, DeltaAz) * 100 / DL
End Function

Function DLS30Deg(ByVal I1 As Double, ByVal I2 As Double, ByVal DeltaAz As Double, ByVal DL As Double) As Double
  ' DoglegSeverity
  ' Cac gia tri goc tinh bang Degree
  ' Ket qua: Degree/100m
    DLS30Deg = getDoglegDeg(I1, I2, DeltaAz) * 30 / DL
End Function

Function getToolFaceRad(I1 As Double, I2 As Double, Dogleg As Double) As Double
    ' Tinh toolface (theo cong thuc tinh goc dat bo khoan cu)
    ' Cho truoc goc nghieng dau hiep khoan I1,
    '           goc nghieng cuoi hiep khoan I2 va Dogleg
    ' ham nay xac dinh goc dat bo khoan cu de thu duoc cac thong so tren sau hiep khoan
    
    ' Cac tham so va ket qua: radian
    
    Dim t As Double
    t = (Cos(I1) * Cos(Dogleg) - Cos(I2)) / (Sin(I1) * Sin(Dogleg))
    If Abs(Abs(t) - 1) < Zero * Zero Then
        t = 0
    Else
        t = WorksheetFunction.Acos(t)
    End If
    getToolFaceRad = t
End Function

Function getToolFaceDeg(ByVal I1 As Double, ByVal I2 As Double, ByVal Dogleg As Double) As Double
    ' cac tham so va ket qua tinh bang Degree
    I1 = WorksheetFunction.Radians(I1)
    I2 = WorksheetFunction.Radians(I2)
    Dogleg = WorksheetFunction.Radians(Dogleg)
    getToolFaceDeg = WorksheetFunction.Degrees(getToolFaceRad(I1, I2, Dogleg))
End Function

Function getI2Rad(I1 As Double, Dogleg As Double, toolface As Double)
    ' Tinh goc cuoi cung I2 neu biet
    ' Goc dau tien I1, Dogleg va toolface
    ' Cong thuc tinh suy ra tu cong thuc tinh Toolface
    ' cac gia tri tham so va ket qua tinh bang Rad
    
    Dim t As Double
    t = Cos(I1) * Cos(Dogleg) - Cos(toolface) * Sin(I1) * Sin(Dogleg)
    If Abs(Abs(t) - 1) < Zero * Zero Then
        t = 0
    Else
        t = WorksheetFunction.Acos(t)
    End If
    getI2Rad = t
End Function

Function getI2Deg(I1 As Double, Dogleg As Double, toolface As Double)
    I1 = WorksheetFunction.Radians(I1)
    Dogleg = WorksheetFunction.Radians(Dogleg)
    toolface = WorksheetFunction.Radians(toolface)
    getI2Deg = WorksheetFunction.Degrees(getI2Rad(I1, Dogleg, toolface))
End Function

Function getDeltaAz(I1 As Double, I2 As Double, Dogleg As Double) As Double
    ' tinh Delta Azumuth tu cac goc nghieng va Dogleg
    ' Cong thuc tinh suy ra tu cong thuc tinh Dogleg
    Dim t As Double
    If (I1 < Zero) Or (I2 < Zero) Then
        getDeltaAz = 0
        Exit Function
    End If
    t = (Cos(Dogleg) - Cos(I1) * Cos(I2)) / (Sin(I1) * Sin(I2))
    If Abs(Abs(t) - 1) < Zero Then
        ' xu ly truong hop gan 0, cac phep tinh co the cho ket qua > 1
        t = 0
    Else
        t = WorksheetFunction.Acos(t)
    End If
    getDeltaAz = t
End Function

Function getDeltaAzDeg(ByVal I1 As Double, ByVal I2 As Double, ByVal Dogleg As Double) As Double
    ' tinh Delta Azumuth tu cac goc nghieng va Dogleg
    ' Cong thuc tinh suy ra tu cong thuc tinh Dogleg
    I1 = WorksheetFunction.Radians(I1)
    I2 = WorksheetFunction.Radians(I2)
    Dogleg = WorksheetFunction.Radians(Dogleg)
    getDeltaAzDeg = WorksheetFunction.Degrees(getDeltaAz(I1, I2, Dogleg))
End Function


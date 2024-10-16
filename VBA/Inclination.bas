Attribute VB_Name = "Inclination"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
' General Survey specific routines
'--------------------------------

Option Explicit
'
'' Trong nay co nhieu thu tuc thua. Nen xem xet bo di
'
'Function InclVectorAtLen1(ByVal I1 As Double, ByVal A1 As Double, _
'           ByVal I2 As Double, ByVal A2 As Double, _
'           ByVal DLen As Double, ByVal AtLen As Double) As TRCoord
'   ' Thu tuc nay co the bo vi khong dung
'   ' tat ca tinh bang radian
'   ' chung minh?
'    Dim Alpha As Double, aDogleg As Double
'    Dim x As Double, y As Double
'    Dim Va As TRCoord, Vb As TRCoord
'    'DLen As Double,
'    ' DLen = L2 - L1
'
'    ' aDogleg and Alpha in rad
'    aDogleg = getDoglegRad(I1, I2, A2 - A1)
'    Alpha = aDogleg * AtLen / DLen
'    y = Sin(Alpha) / Sin(aDogleg)
'    x = Cos(Alpha) - y * Cos(aDogleg)
'    Va = InclRadToVector(I1, A1)
'    Vb = InclRadToVector(I2, A2)
'    InclVectorAtLen1 = LinerCombine(x, Va, y, Vb)
'End Function
'
'Function NormAzimuthDeg(ByVal Azimuth As Double) As Double
'    ' chuan hoa gia tri Azimuth 0..360 degree
'    ' chua xu ly truong hop goc nho hon -360 hoac lon hon 720 degree
'    If Azimuth > 360 Then
'        Azimuth = Azimuth - 360
'    ElseIf Azimuth < 0 Then
'        Azimuth = Azimuth + 360
'    End If
'
'    NormAzimuthDeg = Azimuth
'End Function
'
'Function MeanAz(ByVal Az1 As Double, ByVal Az2 As Double) As Double
'    ' Tinh gia tri Azimuth trung binh
'    ' Chu y la khong the dung trung binh cong
'    Dim tmp As Double
'    tmp = 0.5 * (Az2 + Az1)
'
'    If Abs(Az2 - Az1) > 180 Then
'        tmp = 180 + tmp
'        If tmp > 360 Then
'            tmp = tmp - 360
'        End If
'    End If
'    MeanAz = tmp
'End Function
'
'Function AngleFromZXY(ByVal z As Double, ByVal x As Double, ByVal y As Double)
'    Dim A As TRCoord, B As TRCoord
'    A = MakeCoord(z, x, y)
'    B = UnitVector(A)
'    AngleFromZXY = AngleFromVectorDeg(A)
'End Function
'
'Function AzimuthFromZXY(ByVal z As Double, ByVal x As Double, ByVal y As Double)
'    Dim A As TRCoord, B As TRCoord
'    A = MakeCoord(z, x, y)
'    B = UnitVector(A)
'    AzimuthFromZXY = AzimuthFromVectorDeg(A)
'End Function
'
'Function AngleFromVectorDeg(A As TRCoord) As Double
'    ' a: unit vector
''    AngleFromVectorDeg = WorksheetFunction.Degrees(ArcCos(a.z))
'    AngleFromVectorDeg = WorksheetFunction.Degrees(WorksheetFunction.Acos(A.z))
'End Function
'
Function AzimuthFromVectorDeg(A As TRCoord) As Double
    ' A unit vector
    Dim t As Double
    t = WorksheetFunction.Atan2(A.x, A.y) ' Y: east, X - north
    ' |
    ' | X: north
    ' |
    '--------------
    ' |    Y: east

    If t < 0 Then t = t + 2 * WorksheetFunction.Pi()
    t = WorksheetFunction.Degrees(t)
    AzimuthFromVectorDeg = t
End Function

Sub VectorToInclDeg(A1 As TRCoord, ByRef Angle As Double, Az As Double)
    Dim A As TRCoord
    A = UnitVector(A1)
'    Angle = WorksheetFunction.Degrees(ArcCos(a.z))
    Angle = WorksheetFunction.Degrees(WorksheetFunction.Acos(A.z))
    Az = AzimuthFromVectorDeg(A)
End Sub



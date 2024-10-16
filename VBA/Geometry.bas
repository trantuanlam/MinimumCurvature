Attribute VB_Name = "Geometry"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
' Geometry types and routines
'--------------------------------

Option Explicit

Public Const Zero = 0.0000001

Public Type TRCoord
  z As Double  ' TVD
  x As Double  ' north
  y As Double  ' east
End Type

Function NegZ(C As TRCoord) As TRCoord
    NegZ.x = C.x
    NegZ.y = C.y
    NegZ.z = -C.z
End Function

Function SumCoord(C1 As TRCoord, C2 As TRCoord) As TRCoord
    SumCoord.z = C1.z + C2.z
    SumCoord.x = C1.x + C2.x
    SumCoord.y = C1.y + C2.y
End Function

Function MakeCoord(ByVal z As Double, ByVal x As Double, ByVal y As Double) As TRCoord
    MakeCoord.z = z
    MakeCoord.x = x
    MakeCoord.y = y
End Function

Function MakeVector(ByVal x As Double, ByVal y As Double, ByVal z As Double) As TRCoord
' khong khac gi MakeCoord, chi khac thu tu tham so
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.z = z
End Function

Function SizeVector(C As TRCoord, factor As Double) As TRCoord
    SizeVector.z = C.z * factor
    SizeVector.x = C.x * factor
    SizeVector.y = C.y * factor
End Function

Function UnitVector(C As TRCoord) As TRCoord
    Dim L As Double
    With C
        L = Sqr(.z * .z + .x * .x + .y * .y)
    End With
    UnitVector.z = C.z / L
    UnitVector.x = C.x / L
    UnitVector.y = C.y / L
End Function

Function Scalar(V1 As TRCoord, V2 As TRCoord) As Double
    Scalar = V1.x * V2.x + V1.y * V2.y + V1.z * V2.z
End Function

Function InclDegToVector(ByVal Angle As Double, ByVal Az As Double) As TRCoord
    ' Angle, Az in Degree
    Dim i As Double, A As Double
    i = WorksheetFunction.Radians(Angle)
    A = WorksheetFunction.Radians(Az)
    InclDegToVector.z = Cos(i)
    InclDegToVector.x = Sin(i) * Cos(A)
    InclDegToVector.y = Sin(i) * Sin(A)
End Function

Function InclRadToVector(ByVal Angle As Double, ByVal Az As Double) As TRCoord
    ' Angle, Az in radian
    InclRadToVector.z = Cos(Angle)
    InclRadToVector.x = Sin(Angle) * Cos(Az)
    InclRadToVector.y = Sin(Angle) * Sin(Az)
End Function

Function LinerCombine(ByVal x As Double, C1 As TRCoord, ByVal y As Double, C2 As TRCoord) As TRCoord
    ' return vector x.c1 +  y.c2
    Dim A As TRCoord, B As TRCoord
    A = SizeVector(C1, x)
    B = SizeVector(C2, y)
    LinerCombine = SumCoord(A, B)
End Function

Function Interpolate(A As Double, B As Double, x As Double) As Double
    Interpolate = A + (B - A) * x
End Function

Function FloatEqual(D1 As Double, D2 As Double) As Boolean
    ' so sanh hai so voi do chinh xac Zero
    FloatEqual = Abs(D1 - D2) < Zero
End Function


Function Mul3dVector(V1 As TRCoord, V2 As TRCoord) As TRCoord
    ' Tich vector cua hai vector
    '                |i  j  k |
    ' Result = a*b = |ax ay az|
    '                |bx by bz|
    
    Mul3dVector.x = V1.y * V2.z - V1.z * V2.y
    Mul3dVector.y = V1.z * V2.x - V1.x * V2.z
    Mul3dVector.z = V1.x * V2.y - V1.y * V2.x
End Function

Function IsZeroVector(V As TRCoord) As Boolean
    IsZeroVector = Abs(V.x) < Zero And Abs(V.y) < Zero And Abs(V.z) < Zero
End Function

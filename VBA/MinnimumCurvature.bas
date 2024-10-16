Attribute VB_Name = "MinnimumCurvature"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
' Routines for survey calculation using Minimum Curvature Method.
'--------------------------------

Option Explicit

'minimum curvature method

Function mcGetCoordRad(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As TRCoord
    ' I1, I2, A1, A2 tinh bang Radian
    
    If Abs(I1) < Zero And Abs(I2) < Zero Then
        ' thang dung
        mcGetCoordRad.z = DL
        mcGetCoordRad.x = 0
        mcGetCoordRad.y = 0
    Else
        If (Abs(I1 - I2) < Zero) And (Abs(A1 - A2) < Zero) Then
           ' on dinh goc
            Dim dL1 As Double
            dL1 = DL * Sin((I1 + I2) / 2)
            
            mcGetCoordRad.z = DL * Cos((I1 + I2) / 2)
            mcGetCoordRad.x = dL1 * Cos((A1 + A2) / 2)
            mcGetCoordRad.y = dL1 * Sin((A1 + A2) / 2)
        Else
            Dim sinI1 As Double, sinI2 As Double, cosI1 As Double, cosI2 As Double
            Dim sinA1 As Double, sinA2 As Double, cosA1 As Double, cosA2 As Double
            Dim tmp As Double, RF As Double, aDogleg As Double
            sinI1 = Sin(I1)
            sinI2 = Sin(I2)
            cosI1 = Cos(I1)
            cosI2 = Cos(I2)
            
            sinA1 = Sin(A1)
            sinA2 = Sin(A2)
            cosA1 = Cos(A1)
            cosA2 = Cos(A2)

            ' aDogleg = dolegRad(I1, I2, A1, A2)
            aDogleg = WorksheetFunction.Acos(cosI1 * cosI2 + sinI1 * sinI2 * Cos(A2 - A1))
            
            RF = 2 / (aDogleg) * Tan(aDogleg / 2)
            tmp = DL / 2 * RF
            
            mcGetCoordRad.z = tmp * (cosI1 + cosI2)
            mcGetCoordRad.x = tmp * (sinI1 * cosA1 + sinI2 * cosA2)
            mcGetCoordRad.y = tmp * (sinI1 * sinA1 + sinI2 * sinA2)
            
        End If
    End If
End Function

Function mcGetCoordDeg(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As TRCoord
    ' I1, I2, A1, A2 in Degree
    
    If Abs(I1) < Zero And Abs(I2) < Zero Then
        ' thang dung
        mcGetCoordDeg.z = DL
        mcGetCoordDeg.x = 0
        mcGetCoordDeg.y = 0
    Else
        I1 = WorksheetFunction.Radians(I1)
        I2 = WorksheetFunction.Radians(I2)
        A1 = WorksheetFunction.Radians(A1)
        A2 = WorksheetFunction.Radians(A2)
        If (Abs(I1 - I2) < Zero) And (Abs(A1 - A2) < Zero) Then
           ' on dinh goc
            Dim dL1 As Double
            dL1 = DL * Sin((I1 + I2) / 2)
            
            mcGetCoordDeg.z = DL * Cos((I1 + I2) / 2)
            mcGetCoordDeg.x = dL1 * Cos((A1 + A2) / 2)
            mcGetCoordDeg.y = dL1 * Sin((A1 + A2) / 2)
        Else
            Dim sinI1 As Double, sinI2 As Double, cosI1 As Double, cosI2 As Double
            Dim sinA1 As Double, sinA2 As Double, cosA1 As Double, cosA2 As Double
            Dim tmp As Double, RF As Double, aDogleg As Double
            sinI1 = Sin(I1)
            sinI2 = Sin(I2)
            cosI1 = Cos(I1)
            cosI2 = Cos(I2)
            
            sinA1 = Sin(A1)
            sinA2 = Sin(A2)
            cosA1 = Cos(A1)
            cosA2 = Cos(A2)

            ' aDogleg = dolegRad(I1, I2, A1, A2)
            aDogleg = WorksheetFunction.Acos(cosI1 * cosI2 + sinI1 * sinI2 * Cos(A2 - A1))
            
            RF = 2 / (aDogleg) * Tan(aDogleg / 2)
            tmp = DL / 2 * RF
            
            mcGetCoordDeg.z = tmp * (cosI1 + cosI2)
            mcGetCoordDeg.x = tmp * (sinI1 * cosA1 + sinI2 * cosA2)
            mcGetCoordDeg.y = tmp * (sinI1 * sinA1 + sinI2 * sinA2)
            
        End If
    End If
End Function

'--------- North -------------

Function mcNorth(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As Double
    Dim C As TRCoord
    C = mcGetCoordDeg(I1, I2, A1, A2, DL)
    mcNorth = C.x
End Function

Function mcNorthRad(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As Double
    Dim C As TRCoord
    C = mcGetCoordRad(I1, I2, A1, A2, DL)
    mcNorthRad = C.x
End Function

'--------- East -------------

Function mcEast(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As Double
    Dim C As TRCoord
    C = mcGetCoordDeg(I1, I2, A1, A2, DL)
    mcEast = C.y
End Function

Function mcEastRad(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As Double
    Dim C As TRCoord
    C = mcGetCoordRad(I1, I2, A1, A2, DL)
    mcEastRad = C.y
End Function

'--------- Vertical -------------

Function mcVertical(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As Double
    Dim C As TRCoord
    C = mcGetCoordDeg(I1, I2, A1, A2, DL)
    mcVertical = C.z
End Function

Function mcVerticalRad(ByVal I1 As Double, ByVal I2 As Double, ByVal A1 As Double, ByVal A2 As Double, ByVal DL As Double) As Double
    Dim C As TRCoord
    C = mcGetCoordRad(I1, I2, A1, A2, DL)
    mcVerticalRad = C.z
End Function

'--------- DirAngle -------------

Function DirAngleDeg(aNorth As Double, aEast As Double) As Double
    Dim t As Double
    If Abs(aEast) > Zero Then
        t = WorksheetFunction.Degrees(WorksheetFunction.Atan2(aNorth, aEast))
        If t < 0 Then t = t + 360
    Else
        t = 0
    End If
    
    DirAngleDeg = t
End Function

Function DirAngleRad(aNorth As Double, aEast As Double) As Double
    Dim t As Double
    If Abs(aEast) > Zero Then
        t = WorksheetFunction.Atan2(aNorth, aEast)
        If t < 0 Then t = t + 2 * WorksheetFunction.Pi()
    Else
        t = 0
    End If
    
    DirAngleRad = t
End Function

'-----------------------------------------------------------------------------------
Sub InterpolateAngles(ByVal D1 As Double, ByVal Angle1 As Double, ByVal Az1 As Double, _
                      ByVal D2 As Double, ByVal Angle2 As Double, ByVal Az2 As Double, _
                      ByVal D As Double, ByRef Angle As Double, ByRef Az As Double)

' Day la thu tuc chinh dung de Interpolate Angle va Azimuth
' theo phuong phap Min curvature
' Angle in Degree

    Dim DL As Double, Alpha As Double, x As Double, y As Double
    Dim A As TRCoord, B As TRCoord, t As TRCoord

    If FloatEqual(D, D1) Or _
    (FloatEqual(Angle1, Angle2) And FloatEqual(Angle1, 0)) Or _
    (FloatEqual(Angle1, Angle2) And FloatEqual(Az1, Az2)) Then
        Angle = Angle1
        Az = Az1
    Else
        ' chuyen quan Rad de tinh cho nhanh
        Angle1 = WorksheetFunction.Radians(Angle1)
        Angle2 = WorksheetFunction.Radians(Angle2)
        Az1 = WorksheetFunction.Radians(Az1)
        Az2 = WorksheetFunction.Radians(Az2)
        
        ' DL and Alpha: in radians
        DL = getDoglegRad(Angle1, Angle2, Az2 - Az1)
        Alpha = DL * (D - D1) / (D2 - D1)
        A = InclRadToVector(Angle1, Az1)
        B = InclRadToVector(Angle2, Az2)
        y = Sin(Alpha) / Sin(DL)
        x = Cos(Alpha) - Cos(DL) * y
        t = LinerCombine(x, A, y, B)
        Call VectorToInclDeg(t, Angle, Az)
  End If
End Sub




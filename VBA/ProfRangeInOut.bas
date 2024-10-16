Attribute VB_Name = "ProfRangeInOut"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
' In/Out routines for survey record
'--------------------------------

Option Explicit


Function CountToNonBlank(R As Range) As Integer
    Dim i As Integer
    i = 1
    While Trim(R(i, 1)) <> ""
        i = i + 1
    Wend
    CountToNonBlank = i - 1
End Function

Function ReadProfileFromRange(R As Range, Map As TRProfToRangeMap, Prof() As TRProfile) As Long
    ' Return number of data points
    
    ReadProfileFromRange = 0
    Dim i As Long
    Dim n As Long
    
    n = CountToNonBlank(R)
    If n = 0 Then Exit Function
    
    ReDim Prof(0 To n - 1)
    Dim pr As TRProfile

    For i = 1 To n
        With Map
            If .iTD > 0 Then pr.TD = R(i, .iTD)
            If .iTVD > 0 Then pr.TVD = R(i, .iTVD)
            If .iAngle > 0 Then pr.Angle = R(i, .iAngle)
            If .iAzimuth > 0 Then pr.Azimuth = R(i, .iAzimuth)
            If .iDirection > 0 Then pr.Direction = R(i, .iDirection)
            If .iDisplacement > 0 Then pr.Displacement = R(i, .iDisplacement)
            If .iDLS100 > 0 Then pr.DLS100 = R(i, .iDLS100)
            If .iEast > 0 Then pr.East = R(i, .iEast)
            If .iNorth > 0 Then pr.North = R(i, .iNorth)
            If .iShortenLen > 0 Then pr.ShortenLen = R(i, .iShortenLen)
        End With
        Prof(i - 1) = pr
    Next i
    ReadProfileFromRange = n
End Function

Sub WriteProfileToRange(Prof() As TRProfile, ByRef RDest As Range, Map As TRProfToRangeMap)
    Dim j As Long
    Dim pr As TRProfile
    
    Dim tmp
    tmp = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    For j = LBound(Prof) To UBound(Prof)
        pr = Prof(j)
        With Map
            If .iTD > 0 Then RDest.Cells(j, .iTD) = pr.TD
            If .iTVD > 0 Then RDest.Cells(j, .iTVD) = pr.TVD
            If .iAngle > 0 Then RDest.Cells(j, .iAngle) = pr.Angle
            If .iAzimuth > 0 Then RDest.Cells(j, .iAzimuth) = pr.Azimuth
            If .iDirection > 0 Then RDest.Cells(j, .iDirection) = pr.Direction
            If .iDisplacement > 0 Then RDest.Cells(j, .iDisplacement) = pr.Displacement
            If .iDLS100 > 0 Then RDest.Cells(j, .iDLS100) = pr.DLS100
            If .iEast > 0 Then RDest.Cells(j, .iEast) = pr.East
            If .iNorth > 0 Then RDest.Cells(j, .iNorth) = pr.North
            If .iShortenLen > 0 Then RDest.Cells(j, .iShortenLen) = pr.ShortenLen
        End With
    Next j
    
    Application.Calculation = tmp
End Sub

'Sub ReadProfileStd1FromRangeold(ProfR As Range, ByRef Prof() As TRProfile)
'    Dim Map As TRProfToRangeMap
'    Map = stdProfToRangeMap(stdMap1)
'
'    Call ReadProfileFromRange(ProfR, Map, Prof)
'End Sub

Sub ReadProfileStd1FromRange(ProfR As Range, ByRef Prof() As TRProfile)
    Dim Map As TRProfToRangeMap
    Map = stdProfToRangeMap(stdMapForACADApp)
    
    Call ReadProfileFromRange(ProfR, Map, Prof)
    Call profCalculateValues(Prof)
End Sub

Sub ReadProfileStd1(ProfName As String, ByRef Prof() As TRProfile)
    Dim ProfR As Range
    Set ProfR = Range(ProfName)
    Call ReadProfileStd1FromRange(ProfR, Prof)
End Sub

Sub ReadCalcWriteProfMap1(R As Range)
    ' read prof : MD Angle Azimuth
    ' read prof : MD Angle Azimuth TVD, DriAngle, NS/EW, DLS
    Dim Map As TRProfToRangeMap
    Dim Prof() As TRProfile
    ReDim Prof(100)
    Map = stdProfToRangeMap(stdMap0)
    Call ReadProfileFromRange(R, Map, Prof)
    Call profCalculateValues(Prof())
    
    Map = stdProfToRangeMap(stdMap1)
   
    Dim RDest As Range
    Set RDest = R
    
    Dim j As Long, K As Long
    Dim pr As TRProfile
    K = 1
    For j = LBound(Prof) To UBound(Prof)
        pr = Prof(j)
        With Map
            If .iTVD > 0 Then RDest.Cells(K, .iTVD) = pr.TVD
            If .iDirection > 0 Then RDest.Cells(K, .iDirection) = pr.Direction
            If .iDisplacement > 0 Then RDest.Cells(K, .iDisplacement) = pr.Displacement
            If .iDLS100 > 0 Then RDest.Cells(K, .iDLS100) = pr.DLS100
            If .iEast > 0 Then RDest.Cells(K, .iEast) = pr.East
            If .iNorth > 0 Then RDest.Cells(K, .iNorth) = pr.North
        End With
        K = K + 1
    Next j
    
End Sub


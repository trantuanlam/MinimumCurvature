Attribute VB_Name = "ProfTypes"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
' Types and routines to handle well survey record
'--------------------------------

Option Explicit

' Survey record
Public Type TRProfile
    TD As Double ' measured depth
    Angle As Double
    Azimuth As Double
    TVD As Double
    Direction As Double
    Displacement As Double
    North As Double
    East As Double
    DLS100 As Double
    ShortenLen As Double ' hieu so giua TD va TVD
End Type

Type TRProfToRangeMap
    ' gia tri trong cac truong nay la cac Index chi toi cac
    ' cot trong Range tuong ung voi cac truong cua TRProfile
    ' Neu gia tri nay = 0 thi Truong tuong ung khong co trong Range
    ' Cac thu tuc read/write se dung record nay de thuc hien vao/ra tren worksheet
    iTD As Integer
    iAngle As Integer
    iAzimuth As Integer
    iTVD As Integer
    iDirection As Integer
    iDisplacement As Integer
    iNorth As Integer
    iEast As Integer
    iDLS100 As Integer
    iShortenLen As Integer ' hieu so giua TD va TVD
End Type

' standard ProfileToRangeMap

Public Const stdMap0 = 0
Public Const stdMap1 = 1
Public Const stdMap2 = 2
Public Const stdMapForACADApp = 3

Function stdProfToRangeMap(ByVal i As Integer) As TRProfToRangeMap
    With stdProfToRangeMap
        Select Case i
            Case stdMap0
                .iTD = 1
                .iAngle = 2
                .iAzimuth = 3
            Case stdMap1
                .iTD = 1
                .iAngle = 2
                .iAzimuth = 3
                .iTVD = 4
                .iDirection = 5
                .iDisplacement = 6
                .iNorth = 7
                .iEast = 8
                .iDLS100 = 9
            Case stdMap2
                .iTD = 1
                .iAngle = 2
                .iAzimuth = 3
                .iTVD = 4
                .iShortenLen = 5
                .iDirection = 6
                .iDisplacement = 7
                .iNorth = 8
                .iEast = 9
                .iDLS100 = 10
            Case stdMapForACADApp
                .iTD = 1
                .iAngle = 2
                .iAzimuth = 3
                .iTVD = 4
                .iNorth = 5
                .iEast = 6
                .iDLS100 = 7
            Case Else
                Err.Raise vbObjectError + 1, "stdProfToRangeMap", "Tham so ham khong hop le", " ", 0
        End Select
    End With
End Function

Sub ProfToPts(Prof() As TRProfile, Pts() As Double)
'   Sub ProfToPts(Prof() As TRProfile, Pts() As Double, cDepthScaleFactor As Double)
    ' Chuyen doi cac toa do East, North, TVD, range thanh mang 3 chieu, co the dung de ve poly 3d trong ACAD
    Dim i As Integer, j As Integer, n As Integer
    i = 1
    j = 0
    n = UBound(Prof) - LBound(Prof) + 1
    ReDim Pts(0 To 3 * n - 1)
    For i = LBound(Prof) To UBound(Prof)
        Pts(j) = Prof(i).East
        Pts(j + 1) = Prof(i).North
        Pts(j + 2) = -Prof(i).TVD '* cDepthScaleFactor
        j = j + 3
    Next i
End Sub

Function ProfVector(Prof() As TRProfile, ByVal i As Integer) As TRCoord
    ' No range checking
    ProfVector.x = Prof(i + 1).East - Prof(i).East
    ProfVector.y = Prof(i + 1).North - Prof(i).North
    ProfVector.z = Prof(i + 1).TVD - Prof(i).TVD
End Function

Function NegZProfVector(Prof() As TRProfile, ByVal i As Integer) As TRCoord
    ' No range checking
    NegZProfVector.x = Prof(i + 1).East - Prof(i).East
    NegZProfVector.y = Prof(i + 1).North - Prof(i).North
    NegZProfVector.z = -(Prof(i + 1).TVD - Prof(i).TVD)
End Function

Function UnitProfVector(Prof() As TRProfile, ByVal i As Integer) As TRCoord
    ' No range checking
    UnitProfVector = UnitVector(ProfVector(Prof, i))
End Function

Function NegZUnitProfVector(Prof() As TRProfile, ByVal i As Integer) As TRCoord
    ' No range checking
    NegZUnitProfVector = UnitVector(NegZProfVector(Prof, i))
End Function


Sub profCalculateValues(Prof() As TRProfile)
    ' tinh toan cac gia tri profile theo 3 gia tri co ban: Depth, Angle, Azimuth
    ' Gia tri cua Prof(0).TVD, Prof(0).North... phai duoc gan san
    
    Dim P0 As TRProfile, profRec As TRProfile
    Dim i As Long
    
    P0 = Prof(LBound(Prof))
    For i = LBound(Prof) + 1 To UBound(Prof)
        profRec = Prof(i)
        
        profRec.TVD = P0.TVD + mcVertical(P0.Angle, profRec.Angle, P0.Azimuth, profRec.Azimuth, profRec.TD - P0.TD)
        profRec.North = P0.North + mcNorth(P0.Angle, profRec.Angle, P0.Azimuth, profRec.Azimuth, profRec.TD - P0.TD)
        profRec.East = P0.East + mcEast(P0.Angle, profRec.Angle, P0.Azimuth, profRec.Azimuth, profRec.TD - P0.TD)
        profRec.Displacement = Sqr(profRec.North ^ 2 + profRec.East ^ 2)
        profRec.Direction = DirAngleDeg(profRec.North, profRec.East)
        profRec.DLS100 = DLS100Deg(P0.Angle, profRec.Angle, profRec.Azimuth - P0.Azimuth, profRec.TD - P0.TD)
        
        Prof(i) = profRec
        
        P0 = profRec
    Next i
End Sub

'Function locFindLoMD(Prof() As TRProfile, ByVal MD As Double, ByRef N As Long) As Boolean
'    Dim NN As Long
'    NN = UBound(Prof)
'    Dim i As Long
'    locFindLoMD = False
'    While i < N
'        If Prof(i).TD >= MD Then
'            N = i - 1
'            If N < 0 Then N = 0
'            locFindLoMD = True
'            Exit Function
'        End If
'        i = i + 1
'    Wend
'End Function
'
'Sub profPartOfProf(Prof() As TRProfile, MDin As Double, Mout As Double, subProf() As TRProfile)
'    ' tao ra profile la 1 phan cua profile tu khoang MD1 toi MD2
'    ' MD1, MD2 phai nam trong khoang gia tri cua Profile
'    Dim N As Long
'    N = UBound(Prof)
'    ReDim subProf(0, UBound(Prof) + 2)
'    Dim i As Long, j As Long
'    Dim N1 As Long, N2 As Long
'    i = 0: j = 0
'
'    If locFindLoMD(Prof, MDin, N1) And locFindLoMD(Prof, MDout, N2) Then
'        If Abs(MDin - Prof(N1).TD) > 0.01 Then
'           ' Interpolate
'            subProf(j) = Prof(N1)
'        End If
'    End If
'End Sub

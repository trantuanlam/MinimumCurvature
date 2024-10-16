Attribute VB_Name = "ProfileInterpolateNew"
'-------------------------------
' Directional Survey Calculation
' Minimum Curvature Method
'
' By Tran Tuan Lam
' trantuanlam@hotmail.com
'--------------------------------
Option Explicit

Type TRGenParam
    StepBuildDrop As Double
    StepHold As Double
    DLSEpsilon As Double 'khi DLS nho hon gia tri nay thi coi nhu la Hold
End Type

Function LoadGenParam(RName As String) As TRGenParam
    Dim R As Range
    Set R = Range(RName)
    LoadGenParam = LoadGenParamFromRange(R)
End Function

Function LoadGenParamFromRange(R As Range) As TRGenParam
    With LoadGenParamFromRange
        .StepBuildDrop = R(1, 1)
        .StepHold = R(2, 1)
        .DLSEpsilon = R(3, 1)
    End With
End Function

Function DefaultGenParam() As TRGenParam
    DefaultGenParam.DLSEpsilon = 0.05
    DefaultGenParam.StepBuildDrop = 25
    DefaultGenParam.StepHold = 5000
End Function

'======================================================================
'  Sub...
'======================================================================

Private Function InterpolateProf(P0 As TRProfile, P1 As TRProfile, ByVal aTD As Double) As TRProfile
    Dim profRec As TRProfile
        
    If Abs(P0.TD - aTD) < Zero Then
        InterpolateProf = P0
        Exit Function
    End If
    
    profRec.TD = aTD
       
'    If Abs(P1.DLS100) < Zero Then
    If FloatEqual(P0.Angle, P1.Angle) And FloatEqual(P0.Azimuth, P1.Azimuth) Then
        Dim x As Double
        x = (aTD - P0.TD) / (P1.TD - P0.TD)
        profRec.Angle = Interpolate(P0.Angle, P1.Angle, x) ' chua chinh xac lam
        profRec.Azimuth = Interpolate(P0.Azimuth, P1.Azimuth, x)
        profRec.TVD = Interpolate(P0.TVD, P1.TVD, x)
    Else
        Call InterpolateAngles(P0.TD, P0.Angle, P0.Azimuth, P1.TD, P1.Angle, P1.Azimuth, profRec.TD, profRec.Angle, profRec.Azimuth)
        profRec.TVD = P0.TVD + mcVertical(P0.Angle, profRec.Angle, P0.Azimuth, profRec.Azimuth, aTD - P0.TD)
    End If
    
    profRec.North = P0.North + mcNorth(P0.Angle, profRec.Angle, P0.Azimuth, profRec.Azimuth, profRec.TD - P0.TD)
    profRec.East = P0.East + mcEast(P0.Angle, profRec.Angle, P0.Azimuth, profRec.Azimuth, profRec.TD - P0.TD)
    profRec.Displacement = Sqr(profRec.North ^ 2 + profRec.East ^ 2)
    profRec.Direction = DirAngleDeg(profRec.North, profRec.East)
    profRec.DLS100 = DLS100Deg(P0.Angle, profRec.Angle, profRec.Azimuth - P0.Azimuth, profRec.TD - P0.TD)
    InterpolateProf = profRec
End Function

'========================

Private Function findNearestTD(Prof() As TRProfile, ByVal aValue As Double, aInxed As Long) As Boolean
    ' Tim gia tri TVD gan nhat nho hon val
    ' ta gia su mang Prof duoc sap tang dan theo TD
    Dim i As Long, n As Long
    findNearestTD = False
    
  '--------16 Aug 2011
    i = LBound(Prof)
    If Prof(i).TD > aValue Then
        'MD out of range (MD less than MD of first point)
        Exit Function
    End If
    '-------------------
    
    For i = LBound(Prof) To UBound(Prof) - 1
        If Prof(i + 1).TD > aValue - Zero Then
            findNearestTD = True
            aInxed = i
            Exit Function
        End If
    Next
End Function

Private Function findNearestTVD(Prof() As TRProfile, ByVal aValue As Double, aInxed As Long) As Boolean
    ' Tim gia tri TVD gan nhat nho hon val
    ' ta gia su mang Prof duoc sap tang dan theo MD
    Dim i As Long, n As Long
    findNearestTVD = False
    
    '--------16 Aug 2011
    i = LBound(Prof)
    If Prof(i).TVD > aValue Then
        'tvd out of range - TVD is less than TVD of
        Exit Function
    End If
    '-------------------
    
    For i = LBound(Prof) To UBound(Prof) - 1
        If Prof(i + 1).TVD > aValue - Zero Then
            findNearestTVD = True
            aInxed = i
            Exit Function
        End If
    Next
End Function

Function GetProfileReccordfromTD(Prof() As TRProfile, aTD As Double, profRec As TRProfile, ErrMsgStr As String) As Boolean
    Dim i As Long
    Dim P0 As TRProfile, P1 As TRProfile
    
    If findNearestTD(Prof, aTD, i) Then
        P0 = Prof(i)
        P1 = Prof(i + 1)
        profRec = InterpolateProf(P0, P1, aTD)
        GetProfileReccordfromTD = True
    Else
        ' neu co tham so se raise error
        If ErrMsgStr <> "" Then
            Err.Raise vbObjectError + 27, "Profile Lib", ErrMsgStr + " do khong tim thay chieu sau >" & aTD & " trong profile", "", 0
        Else
            GetProfileReccordfromTD = False
        End If
    End If
End Function

' 25/04/2005: sua lai
'Function IsHolding(ByRef P0 As TRProfile, ByRef P1 As TRProfile) As Boolean
'    If Abs(P0.Angle - P1.Angle) < Zero And Abs(P0.Azimuth - P1.Azimuth) < Zero Then
'        IsHolding = True
'    Else
'        IsHolding = False
'    End If
'End Function

' 04/Dec/2011: sua lai
Function IsHolding(ByRef P0 As TRProfile, ByRef P1 As TRProfile) As Boolean
    IsHolding = False
    If Abs(P0.Angle - P1.Angle) < Zero Then
        If (Abs(P0.Angle) < Zero) Or Abs(P0.Azimuth - P1.Azimuth) < Zero Then
            IsHolding = True
        End If
    End If
End Function


Function locCalcForHolding(ByRef P0 As TRProfile, ByRef P1 As TRProfile, DeltaTVD As Double) As Double
    locCalcForHolding = DeltaTVD * (Abs(P1.TD - P0.TD)) / Abs(P1.TVD - P0.TVD)
End Function


Function GetProfileReccordfromTVD(Prof() As TRProfile, aTVD As Double, profRec As TRProfile, ErrMsgStr As String) As Boolean
    Dim i As Long
    Dim P0 As TRProfile, P1 As TRProfile
    Dim OK As Boolean
    
    OK = findNearestTVD(Prof, aTVD, i)

    If OK Then
        Dim aMD As Double
        P0 = Prof(i)
        P1 = Prof(i + 1)
        If IsHolding(P0, P1) Then
            aMD = P0.TD + locCalcForHolding(P0, P1, aTVD - P0.TVD)
        Else
            OK = findTDFromTVD(P0.Angle, P0.Azimuth, P1.Angle, P1.Azimuth, P1.TD - P0.TD, aTVD - P0.TVD, aMD)
            If OK Then
                aMD = aMD + P0.TD
            End If
        End If
    End If
    
    If OK Then
        profRec = InterpolateProf(P0, P1, aMD)
        GetProfileReccordfromTVD = True
    Else
        ' neu co tham so se raise error
        If ErrMsgStr <> "" Then
            Err.Raise vbObjectError + 27, "Profile Lib", ErrMsgStr + " do khong tim thay chieu sau >" & aTVD & " trong profile", "", 0
        Else
            GetProfileReccordfromTVD = False
        End If
    End If
End Function

'========================================================
Function profGetTVDfromTD(Prof() As TRProfile, aTD As Double) As Double
    Dim P As TRProfile
    If GetProfileReccordfromTD(Prof(), aTD, P, "TVD:Khong tinh duoc") Then
        profGetTVDfromTD = P.TVD
    End If
End Function

Function GetTVDfromTD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetTVDfromTD = profGetTVDfromTD(Prof(), aTD)
End Function

Function mcGetTVDfromMD(SurveyRange As Range, aTD As Double) As Double
' another name for GetTVDfromTD
' this name look better
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(SurveyRange, Prof)
    mcGetTVDfromMD = profGetTVDfromTD(Prof(), aTD)
End Function

'======================================================
Function profGetAnglefromTD(Prof() As TRProfile, aTD As Double) As Double
    Dim P As TRProfile
    If GetProfileReccordfromTD(Prof(), aTD, P, "Angle:Khong tinh duoc") Then
        profGetAnglefromTD = P.Angle
    End If
End Function

Function GetAnglefromTD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetAnglefromTD = profGetAnglefromTD(Prof(), aTD)
End Function

Function mcGetInclinationfromMD(SurveyRange As Range, aTD As Double) As Double
' another name for GetAnglefromTD
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(SurveyRange, Prof)
    mcGetInclinationfromMD = profGetAnglefromTD(Prof(), aTD)
End Function

'============================================================
Function profGetAzimuthfromTD(Prof() As TRProfile, aTD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTD(Prof(), aTD, P, "Azimuth:Khong tinh duoc") Then
        profGetAzimuthfromTD = P.Azimuth
    End If
End Function

Function GetAzimuthfromTD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetAzimuthfromTD = profGetAzimuthfromTD(Prof(), aTD)
End Function

Function mcGetAzimuthfromMD(SurveyRange As Range, aTD As Double) As Double
    ' another name for GetAzimuthfromTD
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(SurveyRange, Prof)
    mcGetAzimuthfromMD = profGetAzimuthfromTD(Prof(), aTD)
End Function
'============================================================

Function profGetNorthfromTD(Prof() As TRProfile, aTD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTD(Prof(), aTD, P, "North:Khong tinh duoc") Then
        profGetNorthfromTD = P.North
    End If
End Function

Function GetNorthfromTD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetNorthfromTD = profGetNorthfromTD(Prof(), aTD)
End Function

Function mcGetNorthfromMD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetNorthfromMD = profGetNorthfromTD(Prof(), aTD)
End Function

'============================================================

Function profGetEastfromTD(Prof() As TRProfile, aTD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTD(Prof(), aTD, P, "East:Khong tinh duoc") Then
        profGetEastfromTD = P.East
    End If
End Function

Function GetEastfromTD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetEastfromTD = profGetEastfromTD(Prof(), aTD)
End Function

Function mcGetEastfromMD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetEastfromMD = profGetEastfromTD(Prof(), aTD)
End Function

'============================================================

Function profGetDLS100fromTD(Prof() As TRProfile, aTD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTD(Prof(), aTD, P, "DLS100:Khong tinh duoc") Then
        profGetDLS100fromTD = P.DLS100
    End If
End Function

Function GetDLS100fromTD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetDLS100fromTD = profGetDLS100fromTD(Prof(), aTD)
End Function

Function mcGetDLS100fromMD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetDLS100fromMD = profGetDLS100fromTD(Prof(), aTD)
End Function

Function mcGetDLS30fromMD(R As Range, aTD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetDLS30fromMD = profGetDLS100fromTD(Prof(), aTD) * 30 / 100
End Function

'========================= From TVD -> Other parmeters ===================

Private Sub GetMiddleParamRad(A As TRCoord, B As TRCoord, ByRef C As TRCoord, ByRef AngleRad As Double, ByRef AzimuthRad As Double)
    ' Angle in radian
    C = SumCoord(A, B)
    C = UnitVector(C)
    
    AngleRad = WorksheetFunction.Acos(C.z)
    
    Dim t As Double
    t = WorksheetFunction.Atan2(C.x, C.y)
    If t < 0 Then t = t + 2 * WorksheetFunction.Pi()
    AzimuthRad = t
End Sub


Function findTDFromTVD(ByVal I1 As Double, ByVal A1 As Double, ByVal I2 As Double, ByVal A2 As Double, ByVal DL As Double, _
                      ByVal aTVD As Double, ByRef aMD As Double) As Boolean
' aTVD >=0
' ta co cac tham so profile la I1, A1, I2, A2 tai cac diem thu nhat va thu hai va Delta MD
' ta can tinh chieu sau theo than gieng tai diem co chieu sau thang dung la aTVD

'    Dim Dogleg As Double, RF As Double
'    aDogleg = WorksheetFunction.Acos(cosI1 * cosI2 + sinI1 * sinI2 * Cos(A2 - A1))
'    RF = 2 / (aDogleg) * Tan(aDogleg / 2)

    I1 = WorksheetFunction.Radians(I1)
    I2 = WorksheetFunction.Radians(I2)
    A1 = WorksheetFunction.Radians(A1)
    A2 = WorksheetFunction.Radians(A2)
    
    Dim TVD1 As Double, TVD2 As Double, TVDx As Double
    Dim MDepth As Double
    
    Dim orgI1 As Double, orgA1 As Double, orgDL As Double
    orgI1 = I1: orgA1 = A1: orgDL = DL
    
    TVD1 = 0
    TVD2 = mcVerticalRad(I1, I2, A1, A2, DL)
    
    If Abs(TVD1 - aTVD) < Zero Then
         findTDFromTVD = True
         aMD = 0
         Exit Function
    ElseIf Abs(TVD2 - aTVD) < Zero Then
         findTDFromTVD = True
         aMD = DL
         Exit Function
    End If
    
    If (TVD1 - aTVD) * (TVD2 - aTVD) > 0 Then
        findTDFromTVD = False
        Exit Function
    End If
    
    Dim x1 As TRCoord, x2 As TRCoord, x As TRCoord
    Dim Angle As Double, Az As Double
    
    x1 = InclRadToVector(I1, A1)
    x2 = InclRadToVector(I2, A2)
    Call GetMiddleParamRad(x1, x2, x, Angle, Az)
    TVDx = mcVerticalRad(I1, Angle, A1, Az, DL / 2)
    aMD = 0
    
    Do While Abs(DL) > Zero
        If (TVD1 - aTVD) * (TVDx - aTVD) < 0 Then
            I2 = Angle
            A2 = Az
            x2 = x
            DL = DL / 2
            TVD2 = TVDx
        Else
            I1 = Angle
            A1 = Az
            x1 = x
            DL = DL / 2
            aMD = aMD + DL
            TVD1 = TVDx
        End If
        
        Call GetMiddleParamRad(x1, x2, x, Angle, Az)
        TVDx = mcVerticalRad(orgI1, Angle, orgA1, Az, aMD + DL / 2)
    Loop
 
    findTDFromTVD = True
End Function

'------------
'==================================
Function profGetTDfromTVD(Prof() As TRProfile, aTVD As Double) As Double
    Dim P As TRProfile
    If GetProfileReccordfromTVD(Prof(), aTVD, P, "TD:Khong tinh duoc (from TVD)") Then
        profGetTDfromTVD = P.TD
    End If
End Function

Function GetTDfromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetTDfromTVD = profGetTDfromTVD(Prof(), aTVD)
End Function

Function mcGetMDfromTVD(SurveyRange As Range, aTVD As Double) As Double
   ' another name for GetTDfromTVD
   ' this name is better
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(SurveyRange, Prof)
    mcGetMDfromTVD = profGetTDfromTVD(Prof(), aTVD)
End Function
'=================================

Function profGetAnglefromTVD(Prof() As TRProfile, aTVD As Double) As Double
    Dim P As TRProfile
    If GetProfileReccordfromTVD(Prof(), aTVD, P, "Angle:Khong tinh duoc (from TVD)") Then
        profGetAnglefromTVD = P.Angle
    End If
End Function

Function GetAnglefromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetAnglefromTVD = profGetAnglefromTVD(Prof(), aTVD)
End Function

Function mcGetInclinationfromTVD(SurveyRange As Range, aTVD As Double) As Double
'another name for GetAnglefromTVD
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(SurveyRange, Prof)
    mcGetInclinationfromTVD = profGetAnglefromTVD(Prof(), aTVD)
End Function
'=================================

Function profGetAzimuthfromTVD(Prof() As TRProfile, aTVD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTVD(Prof(), aTVD, P, "Azimuth:Khong tinh duoc (from TVD)") Then
        profGetAzimuthfromTVD = P.Azimuth
    End If
End Function

Function GetAzimuthfromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetAzimuthfromTVD = profGetAzimuthfromTVD(Prof(), aTVD)
End Function

Function mcGetAzimuthfromTVD(SurveyRange As Range, aTVD As Double) As Double
' another name for GetAzimuthfromTVD
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(SurveyRange, Prof)
    mcGetAzimuthfromTVD = profGetAzimuthfromTVD(Prof(), aTVD)
End Function

'=================================
Function profGetNorthfromTVD(Prof() As TRProfile, aTVD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTVD(Prof(), aTVD, P, "North:Khong tinh duoc (from TVD)") Then
        profGetNorthfromTVD = P.North
    End If
End Function

Function GetNorthfromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetNorthfromTVD = profGetNorthfromTVD(Prof(), aTVD)
End Function

Function mcGetNorthfromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetNorthfromTVD = profGetNorthfromTVD(Prof(), aTVD)
End Function

'=================================
Function profGetEastfromTVD(Prof() As TRProfile, aTVD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTVD(Prof(), aTVD, P, "East:Khong tinh duoc (from TVD)") Then
        profGetEastfromTVD = P.East
    End If
End Function

Function GetEastfromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetEastfromTVD = profGetEastfromTVD(Prof(), aTVD)
End Function

Function mcGetEastfromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetEastfromTVD = profGetEastfromTVD(Prof(), aTVD)
End Function

'=================================
Function profGetDLS100fromTVD(Prof() As TRProfile, aTVD As Double)
    Dim P As TRProfile
    If GetProfileReccordfromTVD(Prof(), aTVD, P, "DLS100:Khong tinh duoc (from TVD)") Then
        profGetDLS100fromTVD = P.DLS100
    End If
End Function

Function GetDLS100fromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    GetDLS100fromTVD = profGetDLS100fromTVD(Prof(), aTVD)
End Function

Function mcGetDLS100fromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetDLS100fromTVD = profGetDLS100fromTVD(Prof(), aTVD)
End Function

Function mcGetDLS30fromTVD(R As Range, aTVD As Double) As Double
    Dim Prof() As TRProfile
    Call ReadProfileStd1FromRange(R, Prof)
    mcGetDLS30fromTVD = profGetDLS100fromTVD(Prof(), aTVD) * 30 / 100
End Function

'==========================================
'  Generate detail profile
'==========================================

Private Sub UpdateProfReccord(outProf() As TRProfile, j As Long, pr As TRProfile)
    outProf(j) = pr
    j = j + 1
End Sub

Private Sub UpdateInterval(outProf() As TRProfile, j As Long, prPrevious As TRProfile, pr As TRProfile, Step As Double)
    Dim TD As Double
    If (pr.TD - prPrevious.TD) > Step Then
        TD = (prPrevious.TD \ Step) * Step
        TD = TD + Step
        While (prPrevious.TD < TD) And (TD < pr.TD)
            Call UpdateProfReccord(outProf(), j, InterpolateProf(prPrevious, pr, TD))
            TD = TD + Step
        Wend
        Call UpdateProfReccord(outProf, j, pr)
    Else
        Call UpdateProfReccord(outProf, j, pr)
    End If
End Sub

Sub GenProfArr(inProf() As TRProfile, outProf() As TRProfile, GenParam As TRGenParam)
    Dim i As Long, j As Long
    Dim pr As TRProfile
    Dim prPrevious As TRProfile
  
    Dim L As Long, U As Long
    L = LBound(inProf)
    U = UBound(inProf)
    
    ReDim outProf(0 To 3000)
    outProf(0) = inProf(L)
    prPrevious = inProf(L)
    
    j = 1
    i = i + 1
    For i = L + 1 To U
        pr = inProf(i)
        If Abs(pr.DLS100) > GenParam.DLSEpsilon Then
            ' Build or drop
            Call UpdateInterval(outProf, j, prPrevious, pr, GenParam.StepBuildDrop)
        Else
           ' Hold or vertical
           Call UpdateInterval(outProf, j, prPrevious, pr, GenParam.StepHold)
        End If
        prPrevious = pr
    Next i
    
    outProf(j) = inProf(U)
    ReDim Preserve outProf(0 To j)
End Sub

Function GetProfInterval(Prof() As TRProfile, TD1 As Double, TD2 As Double, TntervalOfProf() As TRProfile) As Boolean
    ' chua lam xong
    ' lay 1 doan profile tu TD1 toi TD2.
    ' TD1 va2 TD2 phai nam trong khoan gioi han cua Prof
    Dim I1 As Long, I2 As Long
    
    Dim j As Long, i As Long
    j = 0
    If findNearestTD(Prof, TD1, I1) And findNearestTD(Prof, TD2, I2) Then
        ReDim TntervalOfProf(0 To UBound(Prof) + 2)
        Dim P0 As TRProfile, P1 As TRProfile
        If Abs(Prof(I1).TD - TD1) > 10 * Zero Then
            P0 = Prof(I1)
            P1 = Prof(I1 + 1)
            TntervalOfProf(j) = InterpolateProf(P0, P1, TD1)
        Else
            TntervalOfProf(j) = Prof(I1)
        End If
        j = j + 1
        
        For i = I1 + 1 To I2
            TntervalOfProf(j) = Prof(i)
            j = j + 1
        Next
        
        If Abs(Prof(I2).TD - TD2) > 0.01 Then
            P0 = Prof(I2)
            P1 = Prof(I2 + 1)
            TntervalOfProf(j) = InterpolateProf(P0, P1, TD2)
            j = j + 1
        End If
        
        ReDim Preserve TntervalOfProf(0 To j - 1)
        GetProfInterval = True
    End If
    
End Function

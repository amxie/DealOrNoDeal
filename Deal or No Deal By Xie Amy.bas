Attribute VB_Name = "Module1"
Option Explicit

Public Function BankOffer(ByRef A() As Long, ByVal J As Integer, ByVal R As Integer) As Single
    Dim Average As Single
    Dim Amount As Single
    Dim Total As Long
    Dim Percentage As Single
    Dim Counter As Integer
    Dim X As Integer
    
    Total = 0
    Counter = 0
    Percentage = 0.1
    
    For X = 1 To J
        If frmMain.lblMoney(X).BackStyle = 1 Then
            Total = Total + A(X)
            Counter = Counter + 1
        End If
    Next X
    Average = Total / Counter
    Amount = Average * R * Percentage
    frmMain.fraBank.Visible = True
    frmMain.fraMain.Visible = False
    
    BankOffer = Amount
End Function

Public Sub Display(ByVal Offer As Boolean, ByVal R As Integer, ByVal COpen As Integer, ByVal CRem As Integer, ByVal CLeft)
'   Checks to not overwrite the load message and not display the round during bank offerings.
    If R > 0 Then
'       Updates Counters after every briefcase.
        frmMain.lblOpened.Caption = Str$(COpen)
        frmMain.lblRemaining.Caption = Str$(CRem)
        If Offer = False Then
            frmMain.lblSpeaker.Caption = "Round " & Str$(R) & vbCrLf & "Please select " & Str$(CLeft) & " more cases to elminate."
        End If
    End If
End Sub

Public Sub LastCase(ByVal M As Integer)
    Dim X As Integer
    
'   makes the player's briefcase visible in the frame and disables the remaining briefcase
    For X = 1 To M
        If frmMain.cmdBriefcase(X).Visible = True Then
            frmMain.cmdBriefcase(X).Enabled = False
        End If
        If frmMain.cmdBriefcase(0).Caption = frmMain.cmdBriefcase(X).Caption Then
            frmMain.cmdBriefcase(0).Visible = False
            frmMain.cmdBriefcase(X).Visible = True
        End If
    Next X
End Sub

Public Sub LoadMoney(ByRef M() As Long, ByRef RB() As Long, ByVal F As String)
    Dim X As Integer
    Dim Y As Integer
    
    X = 0
    
'   Obtains the file data
    Open F For Input As #1
    Do While Not EOF(1)
        X = X + 1
        Input #1, M(X)
    Loop
    Close #1

'   Displays all money amount to the user in format
    For Y = 1 To X
        frmMain.lblMoney(Y).Caption = Format$(M(Y), "$##,###,###")
    Next Y
    RandomizeArray M(), RB(), X
End Sub

Public Sub RandomizeArray(ByRef A() As Long, ByRef B() As Long, ByVal M As Integer)
    Dim RandomInt As Integer
    Dim Temp As Long
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    
    For X = 1 To M
        B(X) = A(X)
    Next X
    For Y = 1 To 100
        For Z = 1 To M
            RandomInt = Int(Rnd * (M) + 1)
            Temp = B(Z)
            B(Z) = B(RandomInt)
            B(RandomInt) = Temp
        Next Z
    Next Y
End Sub

Public Sub ResetForm(ByRef A As Boolean, ByRef B As Integer, ByRef C As Integer, ByRef D As Integer, ByVal M As Integer)
    Dim X As Integer
    
    A = True
    B = M
    C = 0
    D = 0
    
    For X = 1 To M
        frmMain.lblMoney(X).BackStyle = 1
        frmMain.lblMoney(X).ForeColor = vbBlack
        frmMain.cmdBriefcase(X).Visible = True
        frmMain.cmdBriefcase(X).Enabled = True
    Next X
    
    frmMain.fraBank.Visible = False
    frmMain.fraMain.Visible = True
    frmMain.cmdBriefcase(0).Visible = False
    frmMain.imgNo.Visible = False
    frmMain.imgYes.Visible = False
    frmMain.imgDeal.Visible = True
    frmMain.imgNoDeal.Visible = True
    frmMain.lblRemaining.Caption = Str$(B)
    frmMain.lblOpened.Caption = Str$(C)
    frmMain.lblSpeaker.Caption = "Please select your briefcase!"
End Sub

Public Sub Search(ByRef A() As Long, ByVal B As Long, ByVal J As Integer)
    Dim X As Integer
    
    For X = 1 To J
        If B = A(X) Then
            frmMain.lblMoney(X).BackStyle = 0
            If frmMain.mnuNight.Checked = True Then
                frmMain.lblMoney(X).ForeColor = vbWhite
            End If
        End If
    Next X
End Sub



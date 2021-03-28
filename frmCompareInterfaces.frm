VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCompareInterfaces 
   Caption         =   "Compare controls interfaces"
   ClientHeight    =   6684
   ClientLeft      =   2880
   ClientTop       =   2112
   ClientWidth     =   9336
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6684
   ScaleWidth      =   9336
   Begin VB.CheckBox chkLongForInteger 
      Caption         =   "Allow Long for Integer"
      Height          =   252
      Left            =   408
      TabIndex        =   20
      Top             =   6150
      Width           =   3000
   End
   Begin VB.CheckBox chkMembersHelpStrings 
      Caption         =   "Check members help strings"
      Height          =   252
      Left            =   3500
      TabIndex        =   19
      Top             =   5900
      Width           =   3000
   End
   Begin VB.CheckBox chkMembersIDs 
      Caption         =   "Check members IDs"
      Height          =   252
      Left            =   3500
      TabIndex        =   18
      Top             =   5650
      Width           =   3000
   End
   Begin VB.CheckBox chkMembersFlags 
      Caption         =   "Check members flags"
      Height          =   252
      Left            =   3500
      TabIndex        =   17
      Top             =   5400
      Value           =   1  'Checked
      Width           =   3000
   End
   Begin VB.CheckBox chkByRefByVal 
      Caption         =   "Check parameters ByRef/ByVal"
      Height          =   252
      Left            =   408
      TabIndex        =   16
      Top             =   5400
      Value           =   1  'Checked
      Width           =   3000
   End
   Begin VB.CheckBox chkEnumTypes 
      Caption         =   "Check Enumerations types"
      Height          =   252
      Left            =   408
      TabIndex        =   15
      Top             =   5900
      Width           =   3000
   End
   Begin VB.CheckBox chkParamNames 
      Caption         =   "Check parameters names."
      Height          =   252
      Left            =   408
      TabIndex        =   14
      Top             =   5650
      Width           =   3000
   End
   Begin TabDlg.SSTab sst1 
      Height          =   1308
      Left            =   24
      TabIndex        =   3
      Top             =   7344
      Visible         =   0   'False
      Width           =   6204
      _ExtentX        =   10964
      _ExtentY        =   2307
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Original properties"
      TabPicture(0)   =   "frmCompareInterfaces.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtOriginalProperties"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Original methods"
      TabPicture(1)   =   "frmCompareInterfaces.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtOriginalMethods"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Original events"
      TabPicture(2)   =   "frmCompareInterfaces.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtOriginalEvents"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "New properties"
      TabPicture(3)   =   "frmCompareInterfaces.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtNewProperties"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "New methods"
      TabPicture(4)   =   "frmCompareInterfaces.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtNewMethods"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "New events"
      TabPicture(5)   =   "frmCompareInterfaces.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "txtNewEvents"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin RichTextLib.RichTextBox txtNewEvents 
         Height          =   372
         Left            =   336
         TabIndex        =   11
         Top             =   720
         Width           =   1092
         _ExtentX        =   2117
         _ExtentY        =   1058
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompareInterfaces.frx":00A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtNewMethods 
         Height          =   372
         Left            =   -74832
         TabIndex        =   10
         Top             =   720
         Width           =   1092
         _ExtentX        =   2117
         _ExtentY        =   1058
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompareInterfaces.frx":0126
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtNewProperties 
         Height          =   372
         Left            =   -74568
         TabIndex        =   9
         Top             =   696
         Width           =   1092
         _ExtentX        =   2117
         _ExtentY        =   1058
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompareInterfaces.frx":01A4
      End
      Begin RichTextLib.RichTextBox txtOriginalEvents 
         Height          =   372
         Left            =   -74832
         TabIndex        =   6
         Top             =   708
         Width           =   1092
         _ExtentX        =   2117
         _ExtentY        =   1058
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompareInterfaces.frx":0222
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtOriginalMethods 
         Height          =   372
         Left            =   -74736
         TabIndex        =   5
         Top             =   708
         Width           =   1092
         _ExtentX        =   2117
         _ExtentY        =   1058
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompareInterfaces.frx":02A0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox txtOriginalProperties 
         Height          =   372
         Left            =   -74856
         TabIndex        =   4
         Top             =   708
         Width           =   1092
         _ExtentX        =   2117
         _ExtentY        =   1058
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCompareInterfaces.frx":031E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer tmrCompare 
      Interval        =   1
      Left            =   72
      Top             =   6720
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   444
      Left            =   4512
      TabIndex        =   1
      Top             =   6648
      Width           =   1308
   End
   Begin CompCtrls.ControlInterface CI1 
      Height          =   4072
      Left            =   216
      TabIndex        =   0
      Top             =   672
      Width           =   4344
      _ExtentX        =   7684
      _ExtentY        =   7176
   End
   Begin CompCtrls.ControlInterface CI2 
      Height          =   4072
      Left            =   4740
      TabIndex        =   2
      Top             =   672
      Width           =   4344
      _ExtentX        =   7684
      _ExtentY        =   7176
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy results to the clipboard"
      Height          =   444
      Left            =   708
      TabIndex        =   12
      Top             =   6648
      Width           =   3300
   End
   Begin VB.Label Label1 
      Caption         =   "Note: both controls need to be compiled."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   396
      Left            =   312
      TabIndex        =   13
      Top             =   4896
      Visible         =   0   'False
      Width           =   5004
   End
   Begin VB.Label lblControl2 
      Alignment       =   2  'Center
      Caption         =   "Replacement control:"
      Height          =   276
      Left            =   6102
      TabIndex        =   8
      Top             =   240
      Width           =   1816
   End
   Begin VB.Label lblControl1 
      Alignment       =   2  'Center
      Caption         =   "Original control:"
      Height          =   276
      Left            =   1540
      TabIndex        =   7
      Top             =   240
      Width           =   1816
   End
End
Attribute VB_Name = "frmCompareInterfaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkByRefByVal_Click()
    Compare
End Sub

Private Sub chkEnumTypes_Click()
    Compare
End Sub

Private Sub chkLongForInteger_Click()
    Compare
End Sub

Private Sub chkMembersFlags_Click()
    Compare
End Sub

Private Sub chkMembersHelpStrings_Click()
    Compare
End Sub

Private Sub chkMembersIDs_Click()
    Compare
End Sub

Private Sub chkParamNames_Click()
    Compare
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    Dim rtb As RichTextBox
    
    Select Case sst1.Tab
        Case 0
            Set rtb = txtOriginalProperties
        Case 1
            Set rtb = txtOriginalMethods
        Case 2
            Set rtb = txtOriginalEvents
        Case 3
            Set rtb = txtNewProperties
        Case 4
            Set rtb = txtNewMethods
        Case 5
            Set rtb = txtNewEvents
    End Select
    
    rtb.SelStart = 0
    rtb.SelLength = Len(txtNewProperties.Text)
    Clipboard.Clear
    Clipboard.SetText rtb.Text
    Clipboard.SetText rtb.TextRTF, vbCFRTF
    rtb.SelStart = 0
    
End Sub

Private Sub Form_Load()
    If Not InIDE Then
        MsgBox "This project needs to be run in the IDE (uncompiled).", vbCritical
        End
    End If
    Me.Height = 7100
    If CI1.GetControlTypeName <> "" Then
        CI1.Visible = False
        lblControl1.Visible = False
    Else
        CI1.TipNoteColor = vbRed
        sst1.Visible = False
    End If
    If CI2.GetControlTypeName <> "" Then
        CI2.Visible = False
        lblControl2.Visible = False
    Else
        CI2.TipNoteColor = vbRed
        sst1.Visible = False
    End If
    sst1.Tab = 0
    If (CI1.GetControlTypeName <> "") And (CI2.GetControlTypeName <> "") Then
        Me.WindowState = vbMaximized
    End If
End Sub

Private Function InIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number Then
        InIDE = True
    End If
End Function

Private Sub Form_Resize()
    Dim iTab As Long
    
    If Me.WindowState = vbNormal Then
        If Me.Width < 6400 Then
            Me.Width = 6400
        End If
        If Me.Height < 4000 Then
            Me.Height = 4000
        End If
    End If
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 300, Me.ScaleHeight - cmdClose.Height - 180
    cmdCopy.Top = cmdClose.Top
    If (CI1.TipNoteColor <> vbRed) And (CI2.TipNoteColor <> vbRed) Then
        sst1.Visible = False
        sst1.Move 30, 130, Me.ScaleWidth - 60, cmdClose.Top - 1400
        chkByRefByVal.Top = sst1.Top + sst1.Height + 120
        chkParamNames.Top = chkByRefByVal.Top + chkParamNames.Height '+ 60
        chkEnumTypes.Top = chkParamNames.Top + chkParamNames.Height '+ 60
        chkLongForInteger.Top = chkEnumTypes.Top + chkParamNames.Height
        chkMembersFlags.Top = chkByRefByVal.Top
        chkMembersIDs.Top = chkParamNames.Top
        chkMembersHelpStrings.Top = chkEnumTypes.Top
        iTab = sst1.Tab
        sst1.Tab = 0
        txtOriginalProperties.Move 30, sst1.TabHeight * 2 + 30, sst1.Width - 180, sst1.Height - sst1.TabHeight * 2 - 60
        sst1.Tab = 1
        txtOriginalMethods.Move 30, sst1.TabHeight * 2 + 30, sst1.Width - 180, sst1.Height - sst1.TabHeight * 2 - 60
        sst1.Tab = 2
        txtOriginalEvents.Move 30, sst1.TabHeight * 2 + 30, sst1.Width - 180, sst1.Height - sst1.TabHeight * 2 - 60
        sst1.Tab = 3
        txtNewProperties.Move 30, sst1.TabHeight * 2 + 30, sst1.Width - 180, sst1.Height - sst1.TabHeight * 2 - 60
        sst1.Tab = 4
        txtNewMethods.Move 30, sst1.TabHeight * 2 + 30, sst1.Width - 180, sst1.Height - sst1.TabHeight * 2 - 60
        sst1.Tab = 5
        txtNewEvents.Move 30, sst1.TabHeight * 2 + 30, sst1.Width - 180, sst1.Height - sst1.TabHeight * 2 - 60
        sst1.Tab = 0
        sst1.Visible = True
'        txtOriginalProperties.Text = "Original Properties"
'        txtOriginalMethods.Text = "Original Methods"
'        txtOriginalEvents.Text = "Original Events"
'        txtNewProperties.Text = "New Properties"
'        txtNewMethods.Text = "New Methods"
'        txtNewEvents.Text = "New Events"
        sst1.Tab = iTab
        sst1.Visible = True
    End If
End Sub

Private Sub tmrCompare_Timer()
    tmrCompare.Enabled = False
    If CI1.GetControlTypeName = "" Then
        MsgBox "Please add control 1 inside the box at design time.", vbExclamation
        Exit Sub
    End If
    CI1.Encapsulate = True
    If CI2.GetControlTypeName = "" Then
        MsgBox "Please add control 2 inside the box at design time.", vbExclamation
        Exit Sub
    End If
    CI2.Encapsulate = True
    
    Compare
End Sub

Private Sub Compare()
    Dim iProperties1 As Collection
    Dim iMethods1 As Collection
    Dim iEvents1 As Collection
    Dim iProperties2 As Collection
    Dim iMethods2 As Collection
    Dim iEvents2 As Collection
    
    Set iProperties1 = CI1.Properties
    If Not CI1.Error Then
        Set iMethods1 = CI1.Methods
        Set iEvents1 = CI1.Events
    End If
    Set iProperties2 = CI2.Properties
    If Not CI2.Error Then
        Set iMethods2 = CI2.Methods
        Set iEvents2 = CI2.Events
    End If
    
    ' Original properties
    SetOriginalMembers "Original properties", iProperties1, txtOriginalProperties, iProperties2
    ' Original methods
    SetOriginalMembers "Original methods", iMethods1, txtOriginalMethods, iMethods2
    ' Original events
    SetOriginalMembers "Original events", iEvents1, txtOriginalEvents, iEvents2
    
    ' New properties
    SetNewMembers "New properties", iProperties1, txtNewProperties, iProperties2
    ' New methods
    SetNewMembers "New methods", iMethods1, txtNewMethods, iMethods2
    ' New events
    SetNewMembers "New events", iEvents1, txtNewEvents, iEvents2
End Sub

Private Sub SetNewMembers(nTitle As String, iMembers1 As Collection, txt As RichTextBox, iMembers2 As Collection)
    Dim iMem As cMember
    Dim iMembers() As String
    Dim iMembersObj() As cMember
    Dim c As Long
    Dim iParam As cParameter
    Dim p As Long
    Dim iNoReturnValue As Boolean
    Dim iDefaultValue As String
    
    ReDim iMembers(0)
    ReDim iMembersObj(0)
    
    For Each iMem In iMembers2
        If Not MethodExists(iMembers1, iMem.Name) Then
            AddToList iMembers, iMem.Name
            AddObjectToList iMembersObj, iMem
        End If
    Next
    OrderVector iMembers, iMembersObj
    txt.Text = ""
    txt.SelFontSize = 12
    txt.SelColor = vbBlue
    txt.SelText = nTitle & ":" & vbCrLf & vbCrLf
    txt.SelColor = vbBlack
    txt.SelFontSize = 9
    For c = 1 To UBound(iMembers)
        Set iMem = iMembersObj(c)
        iNoReturnValue = False
        If (iMem.ReturnTypeName = "HRESULT") Or (iMem.ReturnTypeName = "Any") Then
            iNoReturnValue = True
        End If
        txt.SelBold = True
        If (iMem.MemberFlags And FUNCFLAG_FHIDDEN) = FUNCFLAG_FHIDDEN Then
            txt.SelColor = &HA0A0A0
        End If
        txt.SelText = iMembers(c)
        txt.SelBold = False
        txt.SelColor = vbBlack
        If (iMem.MemberFlags And FUNCFLAG_FHIDDEN) = FUNCFLAG_FHIDDEN Then
            txt.SelText = " [Hidden]"
        End If
        If iMem.ParamCount > 0 Then
            txt.SelText = " ("
            txt.SelItalic = True
            For p = 1 To iMem.ParamCount
                Set iParam = iMem.Parameters(p)
                iDefaultValue = ""
                If iParam.IsOptional And (iParam.DefaultValue <> "Undefined") Then
                    iDefaultValue = iParam.DefaultValue
                    Select Case iParam.TypeName
                        Case "Boolean"
                            If iDefaultValue = "False" Then iDefaultValue = ""
                        Case "String"
                            If iDefaultValue = """""" Then iDefaultValue = ""
                        Case "Integer", "Long", "Single", "Double", "Byte"
                            If iDefaultValue = "0" Then iDefaultValue = ""
                    End Select
                    'Debug.Print iParam.TypeName, iParam.DefaultValue
                End If
                txt.SelText = IIf(iParam.IsOptional, "Optional ", "") & IIf(iParam.IsByVal, "ByVal ", "") & iParam.Name & " As " & iParam.TypeName & IIf(iDefaultValue <> "", " = " & iDefaultValue, "")
                If p < iMem.ParamCount Then
                    txt.SelText = ", "
                End If
            Next
            txt.SelItalic = False
            txt.SelText = ")"
        End If
        If Not iNoReturnValue Then
            txt.SelText = " As " & iMem.ReturnTypeName
        End If
        txt.SelText = vbCrLf
    Next
    txt.SelStart = 0
    
End Sub
    
Private Sub SetOriginalMembers(nTitle As String, nMembers1 As Collection, txt As RichTextBox, nMembers2 As Collection)
    Dim iMem1 As cMember
    Dim iMem2 As cMember
    Dim iMembers() As String
    Dim iMembersObj() As cMember
    Dim c As Long
    Dim iParam1 As cParameter
    Dim iParam2 As cParameter
    Dim p As Long
    Dim iNoReturnValue As Boolean
    Dim iDefaultValue As String
    Dim iParamDiff As Boolean
    Dim iDiffRV As Boolean
    Dim iWarningText As String
    Dim iNewControlDetails As String
    Dim iRet1 As String
    Dim iRet2 As String
    
    ReDim iMembers(0)
    ReDim iMembersObj(0)
    
    For Each iMem1 In nMembers1
        AddToList iMembers, iMem1.Name
        AddObjectToList iMembersObj, iMem1
    Next
    OrderVector iMembers, iMembersObj
    txt.Text = ""
    txt.SelFontSize = 12
    txt.SelColor = vbBlue
    txt.SelText = nTitle & ":" & vbCrLf & vbCrLf
    txt.SelColor = vbBlack
    txt.SelFontSize = 9
    For c = 1 To UBound(iMembers)
        iWarningText = ""
        iNewControlDetails = ""
        Set iMem1 = iMembersObj(c)
        Set iMem2 = Nothing
        If MethodExists(nMembers2, iMem1.Name) Then
            Set iMem2 = nMembers2(iMem1.Name)
        End If
        iNoReturnValue = False
        If (iMem1.ReturnTypeName = "HRESULT") Or (iMem1.ReturnTypeName = "Any") Then
            iNoReturnValue = True
        End If
        If iMem2 Is Nothing Then
            txt.SelColor = vbRed
        End If
        If (iMem1.MemberFlags And FUNCFLAG_FHIDDEN) = FUNCFLAG_FHIDDEN Then
            txt.SelColor = &HA0A0A0
        End If
        txt.SelBold = True
        txt.SelText = iMembers(c)
        txt.SelBold = False
        txt.SelColor = vbBlack
        If (iMem1.MemberFlags And FUNCFLAG_FHIDDEN) = FUNCFLAG_FHIDDEN Then
            txt.SelText = " [Hidden]"
        End If
        If iMem2 Is Nothing Then txt.SelColor = vbBlack
        iParamDiff = False
        If iMem1.ParamCount > 0 Then
            txt.SelText = " ("
            txt.SelItalic = True
            For p = 1 To iMem1.ParamCount
                Set iParam1 = iMem1.Parameters(p)
                Set iParam2 = Nothing
                If Not iMem2 Is Nothing Then
                    If p <= iMem2.ParamCount Then
                        Set iParam2 = iMem2.Parameters(p)
                    End If
                    If Not iParam2 Is Nothing Then
                        If Not ParamsASreEqual(iParam1, iParam2) Then
                            iParamDiff = True
                        End If
                    Else
                        iParamDiff = True
                    End If
                End If
                iDefaultValue = ""
                If iParam1.IsOptional And (iParam1.DefaultValue <> "Undefined") Then
                    iDefaultValue = iParam1.DefaultValue
                    Select Case iParam1.TypeName
                        Case "Boolean"
                            If iDefaultValue = "False" Then iDefaultValue = ""
                        Case "String"
                            If iDefaultValue = """""" Then iDefaultValue = ""
                        Case "Integer", "Long", "Single", "Double", "Byte"
                            If iDefaultValue = "0" Then iDefaultValue = ""
                    End Select
                    'Debug.Print iParam1.TypeName, iParam1.DefaultValue
                End If
                txt.SelText = IIf(iParam1.IsOptional, "Optional ", "") & IIf(iParam1.IsByVal, "ByVal ", "") & iParam1.Name & " As " & iParam1.TypeName & IIf(iDefaultValue <> "", " = " & iDefaultValue, "")
                If p < iMem1.ParamCount Then
                    txt.SelText = ", "
                End If
            Next
            txt.SelItalic = False
            txt.SelText = ")"
        End If
        iDiffRV = False
        If Not iNoReturnValue Then
            If Not iMem2 Is Nothing Then
                iRet1 = iMem1.ReturnTypeName
                iRet2 = iMem2.ReturnTypeName
                If chkLongForInteger.Value = 1 Then
                    If iRet1 = "Integer" Then
                        If iMem2.ReturnTypeLong Then
                            iRet2 = "Integer"
                        End If
                    End If
                End If
                If chkEnumTypes.Value = 1 Then
                    If iRet1 <> iRet2 Then
                        iDiffRV = True
                    End If
                Else
                    If iMem1.ReturnTypeLong Or iMem2.ReturnTypeLong Then
                        If iMem1.ReturnTypeLong <> iMem2.ReturnTypeLong Then
                            If Not ((iRet1 = "Integer") And (iRet2 = "Integer")) Then
                                iDiffRV = True
                            End If
                        End If
                    Else
                        If iRet1 <> iRet2 Then
                            iDiffRV = True
                        End If
                    End If
                End If
            End If
            If iDiffRV Then
                txt.SelColor = vbRed
                txt.SelText = " As " & iMem1.ReturnTypeName
                iWarningText = iWarningText & " [Different return type]"
                If (iMem2.ReturnTypeName = "HRESULT") Or (iMem2.ReturnTypeName = "Any") Then
                    iNewControlDetails = iNewControlDetails & "New control: it does not return a value (it is not a function)" & vbCrLf
                Else
                    iNewControlDetails = iNewControlDetails & "New control return type: As " & iMem2.ReturnTypeName & vbCrLf
                End If
                txt.SelColor = vbBlack
            Else
                txt.SelText = " As " & iMem1.ReturnTypeName
            End If
        Else
            If Not iMem2 Is Nothing Then
                If (iMem2.ReturnTypeName <> "HRESULT") And (iMem2.ReturnTypeName <> "Any") Then
                    iDiffRV = True
                End If
            End If
            If iDiffRV Then
                iWarningText = iWarningText & " [Different return type]"
                iNewControlDetails = iNewControlDetails & "New control return type: As " & iMem2.ReturnTypeName
            End If
        End If
        If Not iMem2 Is Nothing Then
            If chkMembersFlags.Value = 1 Then
                If iMem1.MemberFlags <> iMem2.MemberFlags Then
                    iWarningText = iWarningText & " [Different flags]"
                End If
            End If
            If chkMembersIDs.Value = 1 Then
                If iMem1.MemberId <> iMem2.MemberId Then
                    iWarningText = iWarningText & " [Different MemberId]"
                End If
            End If
            If chkMembersHelpStrings.Value = 1 Then
                If iMem1.HelpString <> iMem2.HelpString Then
                    iWarningText = iWarningText & " [Different help string]"
                End If
            End If
        End If
        If iMem2 Is Nothing Then
            iWarningText = iWarningText & " [Missing in new control]"
        ElseIf iParamDiff Then
            
            iWarningText = iWarningText & " [Difference in params]"
            iNewControlDetails = iNewControlDetails & "New control params: "
            For p = 1 To iMem2.ParamCount
                Set iParam2 = iMem2.Parameters(p)
                iDefaultValue = ""
                If iParam2.IsOptional And (iParam2.DefaultValue <> "Undefined") Then
                    iDefaultValue = iParam2.DefaultValue
                    Select Case iParam2.TypeName
                        Case "Boolean"
                            If iDefaultValue = "False" Then iDefaultValue = ""
                        Case "String"
                            If iDefaultValue = """""" Then iDefaultValue = ""
                        Case "Integer", "Long", "Single", "Double", "Byte"
                            If iDefaultValue = "0" Then iDefaultValue = ""
                    End Select
                    'Debug.Print iParam2.TypeName, iParam2.DefaultValue
                End If
                iNewControlDetails = iNewControlDetails & IIf(iParam2.IsOptional, "Optional ", "") & IIf(iParam2.IsByVal, "ByVal ", "") & iParam2.Name & " As " & iParam2.TypeName & IIf(iDefaultValue <> "", " = " & iDefaultValue, "")
                If p < iMem1.ParamCount Then
                    iNewControlDetails = iNewControlDetails & ", "
                End If
            Next
            iNewControlDetails = iNewControlDetails & vbCrLf
        End If
        
        If iWarningText <> "" Then
            txt.SelColor = vbRed
            txt.SelText = iWarningText
            txt.SelColor = vbBlack
        Else
            txt.SelText = " [OK]"
        End If
        txt.SelText = vbCrLf
        If iNewControlDetails <> "" Then
            txt.SelColor = vbBlue
            txt.SelText = iNewControlDetails
            txt.SelColor = vbBlack
        End If
    Next
    txt.SelStart = 0
    
End Sub
    
Private Function MethodExists(nCol As Collection, nMethodName As String) As Boolean
    On Error GoTo ErrorExit
    Call nCol(nMethodName)
    MethodExists = True
    Exit Function
    
ErrorExit:
End Function

Private Function ParamsASreEqual(nPar1 As cParameter, nPar2 As cParameter) As Boolean
    Dim iDef1 As String
    Dim iDef2 As String
    Dim iTypeName1 As String
    Dim iTypeName2 As String
    
    iDef1 = nPar1.DefaultValue
    iDef2 = nPar2.DefaultValue
    iTypeName1 = nPar1.TypeName
    iTypeName2 = nPar2.TypeName
    
    If chkLongForInteger.Value = 1 Then
        If iTypeName1 = "Integer" Then
            If nPar2.TypeLong Then
                iTypeName2 = "Integer"
            End If
        End If
    End If
    
    If (iTypeName1 = "String") Or (iTypeName1 = "Variant") Then
        If iDef1 = "Undefined" Then iDef1 = ""
        If iDef2 = "Undefined" Then iDef2 = ""
    ElseIf IsTypeNumeric(iTypeName1) Then
        If iDef1 = "Undefined" Then iDef1 = "0"
        If iDef2 = "Undefined" Then iDef2 = "0"
    End If
    
    If iDef1 <> iDef2 Then Exit Function
    If chkByRefByVal.Value = 1 Then
        If nPar1.IsByVal <> nPar2.IsByVal Then Exit Function
    End If
    If nPar1.IsOptional <> nPar2.IsOptional Then Exit Function
    If chkEnumTypes.Value = 1 Then
        If iTypeName1 <> iTypeName2 Then Exit Function
    Else
        If nPar1.TypeLong Or nPar2.TypeLong Then
            If Not ((iTypeName1 = "Integer") And (iTypeName2 = "Integer")) Then
                If nPar1.TypeLong <> nPar2.TypeLong Then Exit Function
            End If
        Else
            If iTypeName1 <> iTypeName2 Then Exit Function
        End If
    End If
    If (chkParamNames.Value = 1) Then
        If nPar1.Name <> nPar2.Name Then Exit Function
    End If
    ParamsASreEqual = True
End Function

Private Function IsTypeNumeric(nTypeName As String) As Boolean
    Select Case nTypeName
        Case "Integer", "Decimal", "Long", "Single", "Double", "Byte", "Currency"
            IsTypeNumeric = True
    End Select
End Function

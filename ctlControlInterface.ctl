VERSION 5.00
Begin VB.UserControl ControlInterface 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7344
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2880
   ScaleWidth      =   7344
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add inside here the control that you want to replicate and then run the project"
      Height          =   2376
      Left            =   24
      TabIndex        =   0
      Top             =   120
      Width           =   2724
   End
End
Attribute VB_Name = "ControlInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const cTipNoteDefault As String = "Please add a control inside the box"

Private mTypeLibFile As String
Private mEncapsulate As Boolean

Private mHasMoveMethod As Boolean
Private mHasVisibleProperty As Boolean
Private mHasHwndProperty As Boolean
Private mError As Boolean

Private mLines() As String
Private mLineCount As Long
Private mUB As Long

Private mExtenderProperties() As cMember
Private mExtenderMethods() As cMember
Private mExtenderEvents() As cMember

Private mProperties() As cMember
Private mMethods() As cMember
Private mEvents() As cMember

Private mColProperties As Collection
Private mColMethods As Collection
Private mColEvents As Collection

Private mDataPrepared As Boolean

Public Property Let Encapsulate(nValue As Boolean)
    mEncapsulate = nValue
End Property
    
Public Property Get Encapsulate() As Boolean
    Encapsulate = mEncapsulate
End Property

'Private Sub AddLine(nText As String)
'    mLineCount = mLineCount + 1
'    If (mLineCount - 1) > mUB Then
'        mUB = mUB + 100
'        ReDim Preserve mLines(mUB)
'    End If
'    mLines(mLineCount - 1) = nText
'End Sub
'
'Public Function GetText() As String
'    mUB = 100
'    ReDim mLines(mUB)
'    mLineCount = 0
'
'    GenerateText
'    If mLineCount > 0 Then
'        ReDim Preserve mLines(mLineCount - 1)
'        mUB = mLineCount - 1
'        GetText = Join(mLines, vbCrLf)
'    End If
'End Function

Private Sub UserControl_InitProperties()
    lblNote.Caption = cTipNoteDefault
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblNote.Caption = PropBag.ReadProperty("TipNote", cTipNoteDefault)
End Sub

Private Sub UserControl_Resize()
    If UserControl.ScaleWidth < 2700 Then
        lblNote.Width = UserControl.Width
    Else
        lblNote.Width = 2700
    End If
    lblNote.Move (UserControl.ScaleWidth - lblNote.Width) / 2
End Sub

Public Function GetControlTypeName() As String
    If UserControl.ContainedControls.Count = 0 Then
        'MsgBox "You need to add the source control at design time inside the box (UserControl) before running the program.", vbExclamation
    Else
        GetControlTypeName = TypeName(UserControl.ContainedControls(0))
    End If
End Function

Private Sub ReadMembers()
    Dim iMem As cMember
    Dim i As Long
    Dim iTLB As InterfaceInfo
    
    Set iTLB = TLI.InterfaceInfoFromObject(UserControl.ContainedControls(0))
    On Error Resume Next
    mTypeLibFile = iTLB.Parent.ContainingFile
    If Err.Number Then
        MsgBox "The control must be contained in a compiled OCX/DLL", vbCritical
        mError = True
        Exit Sub
    End If
    On Error GoTo 0
    
    mHasMoveMethod = False
    mHasVisibleProperty = False
    mHasHwndProperty = False
    
    Erase mExtenderProperties
    Erase mExtenderMethods
    Erase mExtenderEvents
    Erase mProperties
    Erase mMethods
    Erase mEvents
    Set mColProperties = Nothing
    Set mColMethods = Nothing
    Set mColEvents = Nothing
    
    StoreMemberInfo UserControl.Extender, mExtenderProperties, INVOKE_PROPERTYGET, Array("Properties", "Methods", "Events")
    StoreMemberInfo UserControl.Extender, mExtenderProperties, INVOKE_PROPERTYPUT
    StoreMemberInfo UserControl.Extender, mExtenderProperties, INVOKE_PROPERTYPUTREF
    StoreMemberInfo UserControl.Extender, mExtenderMethods, INVOKE_FUNC
    StoreMemberInfo UserControl.Extender, mExtenderEvents, INVOKE_EVENTFUNC
    
    i = UBound(mExtenderEvents) + 1
    ReDim Preserve mExtenderEvents(UBound(mExtenderEvents) + 4)
    Set iMem = New cMember
    iMem.Name = "LinkClose"
    Set mExtenderEvents(i) = iMem
    i = i + 1
    Set iMem = New cMember
    iMem.Name = "LinkError"
    Set mExtenderEvents(i) = iMem
    i = i + 1
    Set iMem = New cMember
    iMem.Name = "LinkOpen"
    Set mExtenderEvents(i) = iMem
    i = i + 1
    Set iMem = New cMember
    iMem.Name = "LinkNotify"
    Set mExtenderEvents(i) = iMem

    StoreMemberInfo UserControl.ContainedControls(0), mProperties, INVOKE_PROPERTYGET, mExtenderProperties, True
    If mError Then Exit Sub
    StoreMemberInfo UserControl.ContainedControls(0), mProperties, INVOKE_PROPERTYPUT, mExtenderProperties, True
    If mError Then Exit Sub
    StoreMemberInfo UserControl.ContainedControls(0), mProperties, INVOKE_PROPERTYPUTREF, mExtenderProperties, True
    If mError Then Exit Sub
    StoreMemberInfo UserControl.ContainedControls(0), mMethods, INVOKE_FUNC, mExtenderMethods, True
    If mError Then Exit Sub
    StoreMemberInfo UserControl.ContainedControls(0), mEvents, INVOKE_EVENTFUNC, mExtenderEvents, True
    If mError Then Exit Sub
    
    PutAppearancePropertyFirst
End Sub

Private Sub StoreMemberInfo(ByVal nObject As Object, nVariable As Variant, nMemberType As InvokeKinds, Optional nSkipList As Variant, Optional UseMethod2 As Boolean)
    Dim m As Long
    Dim iMem As cMember
    Dim p As Long
    Dim iParamInfo As ParameterInfo
    Dim iTypeName As String
    Dim iTLI As TypeLibInfo
    Dim iTI As TypeInfo
    Dim iMembers As Members
    Dim t As Long
    Dim iControlTypeName As String
    Dim iSkip As Boolean
    Dim s As Long
    Dim iMembers2 As Members
    Dim iParamName As String
    Dim c As Long
    Dim iDefaultValue As String
    Dim iVar As Variant
    
    If UseMethod2 Then
        iControlTypeName = TypeName(nObject)
        If mTypeLibFile = "" Then
            MsgBox "TypeLib file not set", vbCritical
            Exit Sub
        End If
        Set iTLI = TLI.TypeLibInfoFromFile(mTypeLibFile)
        For t = 1 To iTLI.TypeInfos.Count
            If Not ((iTLI.TypeInfos(t).AttributeMask And TYPEFLAG_FHIDDEN) = TYPEFLAG_FHIDDEN) Then
                If LCase(iTLI.TypeInfos(t).TypeKindString) = "coclass" Then
                    If iTLI.TypeInfos(t).Name = iControlTypeName Then
                        If nMemberType = INVOKE_EVENTFUNC Then
                            Set iMembers = iTLI.TypeInfos(iTLI.TypeInfos(t).DefaultEventInterface.TypeInfoNumber + 1).Members     ' iTLI.TypeInfos(t).ITypeInfo
                        Else
                            Set iMembers = iTLI.TypeInfos(iTLI.TypeInfos(t).DefaultInterface.TypeInfoNumber + 1).Members     ' iTLI.TypeInfos(t).ITypeInfo
                        End If
                        Exit For
                    End If
                End If
            End If
        Next t
    Else
        Set iTI = TLI.ClassInfoFromObject(nObject)
        If nMemberType = INVOKE_EVENTFUNC Then
            Set iMembers = iTI.DefaultEventInterface.Members
        Else
            Set iMembers = iTI.DefaultInterface.Members
        End If
        Set iTLI = TLI.TypeLibInfoFromFile(mTypeLibFile)
    End If
    
    If iMembers Is Nothing Then
        mError = True
        MsgBox "The File you pointed does not correspond to the Control that you added.", vbCritical
        Exit Sub
    End If
    For m = 1 To iMembers.Count
        If (iMembers(m).AttributeMask And FUNCFLAG_FRESTRICTED) = 0 Then  ' Not restricted
            If Left$(iMembers(m).Name, 1) <> "_" Then
                If (iMembers(m).InvokeKind = nMemberType) Or (nMemberType = INVOKE_EVENTFUNC) Then
                    If UseMethod2 Then
                        If iMembers(m).InvokeKind = INVOKE_FUNC Then
                            If iMembers(m).Name = "Move" Then
                                mHasMoveMethod = True
                            End If
                        End If
                        If iMembers(m).InvokeKind = INVOKE_PROPERTYGET Then
                            If iMembers(m).Name = "Visible" Then
                                mHasVisibleProperty = True
                            ElseIf iMembers(m).Name = "hWnd" Then
                                mHasHwndProperty = True
                            End If
                        End If
                    End If
                    iSkip = False
                    If Not IsMissing(nSkipList) Then
                        For s = LBound(nSkipList) To UBound(nSkipList)
                            If IsObject(nSkipList(s)) Then
                                If Not nSkipList(s) Is Nothing Then
                                    If nSkipList(s).Name = iMembers(m).Name Then
                                        iSkip = True
                                        Exit For
                                    End If
                                End If
                            Else
                                If nSkipList(s) = iMembers(m).Name Then
                                    iSkip = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    If Not iSkip Then
                        Set iMem = GetMember(nVariable, iMembers(m).Name)
                        If iMem Is Nothing Then
                            ReDim Preserve nVariable(UBound(nVariable) + 1)
                            Set iMem = New cMember
                            iMem.Name = iMembers(m).Name
                            If iMembers(m).MemberId < 1 Then
                                iMem.MemberId = iMembers(m).MemberId
                                If iMem.MemberId < -10000 Then
                                    If iMem.Name = "Enabled" Then
                                        iMem.MemberId = -514
                                    End If
                                End If
                            Else
                                iMem.MemberId = 1
                            End If
                            Set nVariable(UBound(nVariable)) = iMem
                        End If
                        iMem.MemberFlags = iMem.MemberFlags Or iMembers(m).AttributeMask
                        If nMemberType = INVOKE_PROPERTYGET Then
                            iMem.HasGet = True
                        End If
                        If nMemberType = INVOKE_PROPERTYPUT Then
                            If iMembers(m).Name <> "hWnd" Then
                                iMem.HasLet = True
                            End If
                        End If
                        If nMemberType = INVOKE_PROPERTYPUTREF Then
                            iMem.HasSet = True
                        End If
                        If iMem.HelpString = "" Then
                            If Trim$(iMembers(m).HelpString) <> "" Then
                                iMem.HelpString = iMembers(m).HelpString
                            End If
                        End If
                        
                        iMem.ReturnTypeName = GetTypeName(iMembers(m).ReturnType)
                        iMem.ReturnTypeName2 = GetTypeName(iMembers(m).ReturnType, True)
                        iMem.ReturnTypeObject = (iMembers(m).ReturnType.VarType = 0)
                        iMem.ReturnTypeLong = (iMem.ReturnTypeName2 = "Long")
                        If Not iMembers(m).ReturnType.TypeInfo Is Nothing Then
                            If iMem.ReturnTypeObject Then
                                iMem.ReturnTypeObject = iMem.ReturnTypeObject And (Not (LCase$(iMembers(m).ReturnType.TypeInfo.TypeKindString) = "enum"))
                            End If
                        End If
                        iMem.ReturnTypeIsNumeric = IsVarNumeric(iMembers(m).ReturnType)
                        If iMem.ReturnTypeName = "Long" Then
                            If InStr(iMem.Name, "Color") > 0 Then
                                iMem.ReturnTypeName = "OLE_COLOR"
                            End If
                        End If
                        If iMem.ParamCount = 0 Then
                            If (nMemberType = INVOKE_PROPERTYPUT) Or (nMemberType = INVOKE_PROPERTYGET) Or (nMemberType = INVOKE_PROPERTYPUTREF) Or (nMemberType = INVOKE_EVENTFUNC) Or (nMemberType = INVOKE_FUNC) Then
                                ' parameters
                                For p = 1 To iMembers(m).Parameters.Count
                                    Set iParamInfo = iMembers(m).Parameters(p)
                                    iParamName = iParamInfo.Name
                                    If iParamName = "" Then
                                        c = 1
                                        iParamName = "Param" & c
                                        Do Until Not iMem.ParamExists(iParamName)
                                            c = c + 1
                                            iParamName = "Param" & c
                                        Loop
                                    End If
                                    iDefaultValue = "Undefined"
                                    If (iParamInfo.Flags And PARAMFLAG_FOPT) <> 0 Then
                                        On Error Resume Next
                                        iVar = Empty
                                        iVar = iParamInfo.DefaultValue
                                        If Not IsEmpty(iDefaultValue) Then
                                            iDefaultValue = iVar
                                            If VarType(iVar) = vbString Then
                                                iDefaultValue = """" & iDefaultValue & """"
                                            ElseIf VarType(iVar) = vbBoolean Then
                                                iDefaultValue = Replace(iDefaultValue, CStr(False), "False")
                                                iDefaultValue = Replace(iDefaultValue, CStr(True), "True")
                                            End If
                                        End If
                                        On Error GoTo 0
                                    End If
                                    iMem.AddParameter iParamName, GetTypeName(iParamInfo.VarTypeInfo), GetTypeName(iParamInfo.VarTypeInfo, True), GetTypeName(iParamInfo.VarTypeInfo, True) = "Long", (iParamInfo.Flags And PARAMFLAG_FOUT) = 0, (iParamInfo.Flags And PARAMFLAG_FOPT) <> 0, iDefaultValue, iTLI.Name
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next

End Sub

Private Function GetMember(nMembers As Variant, nName As String) As cMember
    Dim c As Long
    
    For c = LBound(nMembers) To UBound(nMembers)
        If Not nMembers(c) Is Nothing Then
            If nMembers(c).Name = nName Then
                Set GetMember = nMembers(c)
                Exit Function
            End If
        End If
    Next
    
End Function

Private Function GetTypeName(ByVal nVarTypeInfo As VarTypeInfo, Optional nGenericType As Boolean = False) As String
    Dim iStr As String
    Dim iVarType As Long
    Dim iKnownObjectType As Boolean
    
    iVarType = nVarTypeInfo.VarType
    If iVarType <> 0 Then
        Select Case (iVarType And &HFF&)
            Case VT_BOOL
                iStr = "Boolean"
            Case VT_BSTR, VT_LPSTR, VT_LPWSTR
                iStr = "String"
            Case VT_DATE
                iStr = "Date"
            Case VT_INT
                iStr = "Integer"
            Case VT_VARIANT
                iStr = "Variant"
            Case VT_DECIMAL
                iStr = "Decimal"
            Case VT_I4
                iStr = "Long"
            Case VT_I2
                iStr = "Integer"
            Case VT_I8
                iStr = "Unknown"
            Case VT_SAFEARRAY
                iStr = "SafeArray"
            Case VT_CLSID
                iStr = "CLSID"
            Case VT_UINT
                iStr = "UInt"
            Case VT_UI4
'                iStr = "ULong"
                iStr = "Long"
            Case VT_UNKNOWN
                iStr = "Unknown"
            Case VT_VECTOR
                iStr = "Vector"
            Case VT_R4
                iStr = "Single"
            Case VT_R8
                iStr = "Double"
            Case VT_DISPATCH
                iStr = "Object"
            Case VT_UI1
                iStr = "Byte"
            Case VT_CY
                iStr = "Currency"
            Case VT_HRESULT
                iStr = "HRESULT" ' note if this was a function it should be a sub
            Case VT_VOID
                iStr = "Any"
            Case VT_ERROR
                iStr = "Long"
            Case Else
                iStr = "<Unsupported Variant Type"
                Select Case (iVarType And &HFF&)
                    Case VT_UI1
                        iStr = iStr & "(VT_UI1)"
                    Case VT_UI2
                        iStr = iStr & "(VT_UI2)"
                    Case VT_UI4
                        iStr = iStr & "(VT_UI4)"
                    Case VT_UI8
                        iStr = iStr & "(VT_UI8)"
                    Case VT_USERDEFINED
                        iStr = iStr & "(VT_USERDEFINED)"
                End Select
                iStr = iStr & ">"
        End Select
        If (iVarType And VT_ARRAY) = VT_ARRAY Then
            iStr = iStr & "()"
        End If
        
        GetTypeName = iStr
    Else
        On Error Resume Next
        iStr = ""
        iStr = nVarTypeInfo.TypeInfo.Name
        If Left(iStr, 1) = "_" Then
            iStr = Mid$(iStr, 2)
        End If
        iKnownObjectType = False
        Select Case iStr
            Case "Picture", "Font", "Collection", "ContainedControls", "DataObject"
                iKnownObjectType = True
        End Select
        
        If nVarTypeInfo.TypeLibInfoExternal Is Nothing Then
            On Error GoTo 0
            If nGenericType Then
                If Not iKnownObjectType Then
                    GetTypeName = "Object"
                Else
                    GetTypeName = nVarTypeInfo.TypeInfo.Name
                End If
            Else
                GetTypeName = nVarTypeInfo.TypeInfo.Name
            End If
        Else
            If (LCase$(nVarTypeInfo.TypeLibInfoExternal) = "stdole") Then
                On Error GoTo 0
                If nGenericType Then
                    If Not iKnownObjectType Then
                        GetTypeName = "Object"
                    Else
                        GetTypeName = nVarTypeInfo.TypeInfo.Name
                    End If
                Else
                    GetTypeName = nVarTypeInfo.TypeInfo.Name
                End If
            Else
                On Error GoTo 0
                If nGenericType Then
                    If Not iKnownObjectType Then
                        GetTypeName = "Object"
                    Else
                        GetTypeName = nVarTypeInfo.TypeInfo.Name
                    End If
                Else
                    GetTypeName = nVarTypeInfo.TypeInfo.Name
                End If
            End If
        End If
    End If
    If Left(GetTypeName, 1) = "_" Then
        GetTypeName = Mid$(GetTypeName, 2)
    End If
    If nGenericType Then
        GetTypeName = Replace$(GetTypeName, "OLE_COLOR", "Long")
    
        If Not nVarTypeInfo.TypeInfo Is Nothing Then
            If (LCase$(nVarTypeInfo.TypeInfo.TypeKindString) = "enum") Then
                GetTypeName = "Long"
            End If
        End If
    End If

End Function

Private Function IsVarNumeric(ByVal nVarTypeInfo As VarTypeInfo) As Boolean
    Dim iStr As String
    Dim iVarType As Long
    
    iVarType = nVarTypeInfo.VarType
    If iVarType <> 0 Then
        Select Case (iVarType And &HFF&)
            Case VT_BOOL
                
            Case VT_BSTR, VT_LPSTR, VT_LPWSTR
                
            Case VT_DATE
                IsVarNumeric = True
            Case VT_INT
                IsVarNumeric = True
            Case VT_VARIANT
                
            Case VT_DECIMAL
                IsVarNumeric = True
            Case VT_I4
                IsVarNumeric = True
            Case VT_I2
                IsVarNumeric = True
            Case VT_I8
                
            Case VT_SAFEARRAY
                
            Case VT_CLSID
                
            Case VT_UINT
                IsVarNumeric = True
            Case VT_UI4
                IsVarNumeric = True
            Case VT_UNKNOWN
                
            Case VT_VECTOR
                
            Case VT_R4
                IsVarNumeric = True
            Case VT_R8
                IsVarNumeric = True
            Case VT_DISPATCH
                
            Case VT_UI1
                IsVarNumeric = True
            Case VT_CY
                IsVarNumeric = True
            Case VT_HRESULT
                
            Case VT_VOID
                
            Case VT_ERROR
                IsVarNumeric = True
            Case Else
        
        End Select
    End If
End Function

Private Sub PutAppearancePropertyFirst()
    Dim c As Long
    Dim iMem As cMember
    Dim iIndex As Long
    
    iIndex = -1
    For c = LBound(mProperties) To UBound(mProperties)
        Set iMem = mProperties(c)
        If iMem.Name = "Appearance" Then
            iIndex = c
            Exit For
        End If
    Next
    If iIndex > 0 Then
        Set iMem = mProperties(iIndex)
        Set mProperties(iIndex) = mProperties(0)
        Set mProperties(0) = iMem
    End If
End Sub

Private Function IsInList(nList, nValue, Optional nFirstElement As Long = 0, Optional nLastElement As Long = -1) As Boolean
    Dim c As Long
    
    If nLastElement = -1 Then
        nLastElement = UBound(nList)
    Else
        If nLastElement > UBound(nList) Then
            nLastElement = UBound(nList)
        End If
    End If
    
    For c = nFirstElement To nLastElement
        If nList(c) = nValue Then
            IsInList = True
            Exit For
        End If
    Next c
End Function


Public Property Let TipNote(nText As String)
    lblNote.Caption = nText
    PropertyChanged "TipNote"
End Property

Public Property Get TipNote() As String
    TipNote = lblNote.Caption
End Property

Public Property Let TipNoteColor(nColor As OLE_COLOR)
    lblNote.ForeColor = nColor
End Property

Public Property Get TipNoteColor() As OLE_COLOR
    TipNoteColor = lblNote.ForeColor
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TipNote", lblNote.Caption, cTipNoteDefault
End Sub

Public Property Get Properties() As Collection
    Dim c As Long
    
    If mColProperties Is Nothing Then
        EnsureDataPrepared
        Set mColProperties = New Collection
        If Not mError Then
            For c = LBound(mProperties) To UBound(mProperties)
                mColProperties.Add mProperties(c).Clone, mProperties(c).Name
            Next
        End If
    End If
    Set Properties = mColProperties
End Property

Public Property Get Methods() As Collection
    Dim c As Long
    
    If mColMethods Is Nothing Then
        EnsureDataPrepared
        Set mColMethods = New Collection
        If Not mError Then
            For c = LBound(mMethods) To UBound(mMethods)
                mColMethods.Add mMethods(c).Clone, mMethods(c).Name
            Next
        End If
    End If
    Set Methods = mColMethods
End Property

Public Property Get Events() As Collection
    Dim c As Long
    
    If mColEvents Is Nothing Then
        EnsureDataPrepared
        Set mColEvents = New Collection
        If Not mError Then
            For c = LBound(mEvents) To UBound(mEvents)
                mColEvents.Add mEvents(c).Clone, mEvents(c).Name
            Next
        End If
    End If
    Set Events = mColEvents
End Property

Private Sub EnsureDataPrepared()
    If Not mDataPrepared Then
        If UserControl.ContainedControls.Count = 0 Then
            MsgBox "You need to add the control to replicate at design time inside the box (UserControl) before running the program.", vbExclamation
            Exit Sub
        End If
    '
        mError = False
        ReadMembers
        If Not mError Then
            mDataPrepared = True
        Else
            MsgBox "Error reading interface of " & ContainedControls(0).Name & ".", vbCritical
        End If
    End If
End Sub

Public Property Get Error() As Boolean
    Error = mError
End Property

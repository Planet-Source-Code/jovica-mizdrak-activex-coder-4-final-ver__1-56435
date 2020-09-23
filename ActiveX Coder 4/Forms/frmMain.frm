VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveX Coder 4"
   ClientHeight    =   7725
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lstmain 
      Height          =   3135
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   327682
      SmallIcons      =   "imglst1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove All"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Invert Selection"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UnSelect All"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select All"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   4680
      List            =   "frmMain.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Generate && Copy"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   7
      ToolTipText     =   "Copy generated code"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtmain 
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   4200
   End
   Begin VB.TextBox txtmain 
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   4080
   End
   Begin VB.TextBox txtmain 
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   1320
      Width           =   4200
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Add"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "Add values"
      Top             =   120
      Width           =   1560
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Remove Selected"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      ToolTipText     =   "Remove selected entrie"
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CDL1 
      Left            =   7440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox Rich 
      Height          =   2535
      Left            =   0
      TabIndex        =   13
      Top             =   5175
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   4471
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":0004
   End
   Begin VB.Label lbllstcount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Property name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of variable:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Container variable:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4695
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin ComctlLib.ImageList imgSmall 
      Left            =   8040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0086
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":05D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0E7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileS 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSS 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImp 
         Caption         =   "Import..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Co&py"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "A&dd"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "Re&move Selected"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditRemoveAll 
         Caption         =   "Remove &All..."
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditGen 
         Caption         =   "Generate &Code"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuLstMain 
      Caption         =   "lstmain"
      Visible         =   0   'False
      Begin VB.Menu mnuLstMainEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuLstMainSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLstMainRR 
         Caption         =   "Remove Selected"
      End
      Begin VB.Menu mnuLstMainDelR 
         Caption         =   "Remove &All..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private tHt As LVHITTESTINFO
Private list_item As ListItem
Private X As New clsList
Private i As Integer
Private DocChanged As Boolean
Private docname As String
Private xx As Integer

Const LVM_FIRST = &H1000&
Const LVM_HITTEST = LVM_FIRST + 18

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Dim TT As CTooltip
Dim m_lCurItemIndex As Long

Private Sub ClearTxt()
txtmain(0).Text = ""
Combo1.Text = ""
txtmain(2).Text = ""
txtmain(3).Text = ""
Rich.Text = ""
Rich.Text = ""

End Sub


Private Sub cmdmain_Click(index As Integer)
On Error Resume Next
Select Case index
    Case 0 'Add

    Set list_item = lstmain.ListItems.Add(, , lstmain.ListItems.Count + 1)
    list_item.SmallIcon = SetListIcon(Combo1)
    list_item.SubItems(1) = txtmain(0)
    list_item.SubItems(2) = Combo1
    list_item.SubItems(3) = txtmain(2)
    list_item.SubItems(4) = txtmain(3)
    
    DocChanged = True
        listcount
        checklst
        txtmain(0) = Empty
        Combo1 = Empty
        txtmain(2) = Empty
        txtmain(3) = Empty
        Rich.Text = Empty
    
    Case 1 'Remove
    Dim i As Integer
    Dim srtx As String
    Do
    For i = 0 To FindLVCHKED - 1
    srtx = Get_After_Comma(i, countComa)
            lstmain.ListItems.Remove lstmain.ListItems.Item(srtx + 1).index
        listcount
        checklst
        Rich.Text = Empty
    Next i
    Loop Until FindLVCHKED <= 0
        
    Case 2 'Generate
    If FindLVCHKED <= 0 Then Exit Sub
    Dim xitem As Integer
        Rich.Text = Empty
        Rich.Text = Rich.Text & "Option Explicit" & vbNewLine & vbNewLine
        'Generate Private Decleractions
    For i = 0 To FindLVCHKED - 1
    xitem = Get_After_Comma(i, countComa)
            With lstmain.ListItems.Item(xitem + 1)
                Rich.Text = Rich.Text & "Private " & .SubItems(3) & " As " & .SubItems(2) & vbNewLine
            End With
        Next
        Rich.Text = Rich.Text & vbNewLine
        'Generate Get, Let properties
        For i = 0 To FindLVCHKED - 1
        xitem = Get_After_Comma(i, countComa)
            With lstmain.ListItems.Item(xitem + 1)
                Rich.Text = Rich.Text & generate(.SubItems(1), .SubItems(2), .SubItems(3)) & vbNewLine & vbNewLine
            End With
        Next
        'Generate UserControl_ReadProperties
        Rich.Text = Rich.Text & vbNewLine & "Private Sub UserControl_ReadProperties(PropBag As PropertyBag)" & vbNewLine
        For i = 0 To FindLVCHKED - 1
        xitem = Get_After_Comma(i, countComa)
            With lstmain.ListItems.Item(xitem + 1)
                If .SubItems(2) = "String" Then .SubItems(4) = """" & .SubItems(4) & """"
                Rich.Text = Rich.Text & vbTab & .SubItems(1) & " = PropBag.ReadProperty(" & """" & .SubItems(1) & """" & ", " & .SubItems(4) & ")" & vbNewLine
                .SubItems(4) = Replace(.SubItems(4), """", "")
            End With
        Next
        'Generate UserControl_WriteProperties
        Rich.Text = Rich.Text & "End Sub" & vbNewLine & vbNewLine & "Private Sub UserControl_WriteProperties(PropBag As PropertyBag)" & vbNewLine
        Rich.Text = Rich.Text & "   With PropBag" & vbNewLine
        For i = 0 To FindLVCHKED - 1
        xitem = Get_After_Comma(i, countComa)
            With lstmain.ListItems.Item(xitem + 1)
                If .SubItems(2) = "String" Then .SubItems(4) = """" & .SubItems(4) & """"
                Rich.Text = Rich.Text & vbTab & "Call .WriteProperty (" & """" & .SubItems(1) & """" & ", " & .SubItems(3) & ", " & .SubItems(4) & ")" & vbNewLine
                .SubItems(4) = Replace(.SubItems(4), """", "")
            End With
        Next
        Rich.Text = Rich.Text & "   End With" & vbNewLine
        Rich.Text = Rich.Text & "End Sub"
    Case 3, 5 'Copy
        cmdmain_Click 2
        Clipboard.Clear
        Clipboard.SetText Rich.Text
    Case 4 'Exit
        Unload Me
        End
    Case 6 ' Clears the listview:
    If MsgBox("Are you sure you want to delete all entries in the List?", _
    vbCritical + vbYesNo, App.Title) = vbNo _
    Then Exit Sub
    
    lstmain.ListItems.Clear
    listcount
    checklst
    
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
ComboKeyPress Combo1, KeyAscii
End Sub


Private Sub Combo1_LostFocus()
    Combo1.SelLength = 0
End Sub

Private Function countComa()
   Dim i As Long
   Dim r As Long
   Dim LV As LV_ITEM
   
  'a string to build the msgbox text with
   Dim b As String

  'iterate through each item, checking its item state
   For i = 0 To lstmain.ListItems.Count
      r = SendMessage(lstmain.hwnd, LVM_GETITEMSTATE, i, ByVal LVIS_STATEIMAGEMASK)
     'when an item is checked, the LVM_GETITEMSTATE call
     'returns 8192 (&H2000&).
      If (r And &H2000&) Then
         'it is checked, so pad the LV_ITEM string members
         With LV
            .cchTextMax = MAX_PATH
            .pszText = Space$(MAX_PATH)
         End With
        'and retrieve the value (text) of the checked item
         Call SendMessage(lstmain.hwnd, LVM_GETITEMTEXT, i, LV)
         b = b & CStr(i) & ","
      End If
   Next
   countComa = b
End Function

Private Function FindLVCHKED()
Dim CharCount As String
Dim Char As String
Char = ","
    ' returns 5 but 6 if +1
    CharCount = Occurs(countComa, Char) '+ 1
If CharCount <= 0 Then
CharCount = 0
FindLVCHKED = "0"
Else
FindLVCHKED = CharCount
End If
End Function

Private Sub Command1_Click()
EnhLitView_CheckAllItems lstmain
listcount
End Sub

Private Sub Command2_Click()
EnhLitView_UnCheckAllItems lstmain
listcount
End Sub

Private Sub Command3_Click()
EnhListView_InvertAllChecks lstmain
listcount
End Sub

Private Sub Command4_Click()
cmdmain_Click (6)
End Sub

Private Sub Form_Load()
Dim list_item As ListItem

    Set X.list = lstmain
    X.addcolumn "ID", "id", 700, True, False
    X.addcolumn "Property Name", "pname", 1640, True, False
    X.addcolumn "Type Var", "tvar", 1600, False, True
    X.addcolumn "Container Var", "cvar", 1600, False, True
    X.addcolumn "Default Value", "defvalue", 1600, False, False
    lstmain.SmallIcons = imgSmall
    
    If Combo1.listcount = 0 Then LoadCombo Combo1
    
    ShowHeaderIcon 0, 0, True
    
    tHt.iItem = -1
    ' set lvVSS to set nodes for project.
    Call ListView_FullRowSelect(lstmain)
    Call ListView_GridLines(lstmain)
   
   lstmain.Refresh
   checklst
   listcount
   
    Call SendMessage(lstmain.hwnd, _
                    LVM_SETEXTENDEDLISTVIEWSTYLE, _
                    LVS_EX_CHECKBOXES, ByVal True)

   Set TT = New CTooltip
   TT.Style = TTBalloon
   TT.Icon = TTIconInfo
    lstmain.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
If DocChanged = True Then
    
    Select Case MsgBox( _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, "ActiveX Coder 3")
    
    Case vbYes
        mnuFileS_Click
    Case vbNo
        Unload frmMain
    Case vbCancel
        Cancel = True
    
    End Select

End If
End Sub

Private Sub lstmain_Click()
listcount
End Sub

Private Sub lstmain_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

   Dim i As Long
   Static sOrder
   
   sOrder = Not sOrder
   
  'Use default sorting to sort the items in the list
   lstmain.SortKey = ColumnHeader.index - 1
   lstmain.SortOrder = Abs(sOrder)
   lstmain.Sorted = True
   
  'clear the image from the headers not
  'currently selected, and update the
  'header clicked
   For i = 0 To 4
      
     'if this is the index of the header clicked
      If i = lstmain.SortKey Then
      
           'ShowHeaderIcon colNo, imgIndex, showFlag
            ShowHeaderIcon lstmain.SortKey, _
                           lstmain.SortOrder, _
                           True
                           
      Else: ShowHeaderIcon i, 0, False
      End If
   
   Next
   
End Sub

Private Sub lstmain_DblClick()
Call mnuLstMainEdit_Click
End Sub

Private Sub lstmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

      
    tHt = ListView_HitTest(lstmain, X, Y)
        
    If Button <> 2 Then Exit Sub
    
    If tHt.iItem = -1 Then
        mnuLstMainRR.Enabled = False
        mnuLstMainEdit.Enabled = False
        mnuLstMainDelR.Enabled = False
    Else
        mnuLstMainRR.Enabled = True
        mnuLstMainEdit.Enabled = True
        mnuLstMainDelR.Enabled = True
        lstmain.ListItems(tHt.iItem + 1).Selected = True
    End If
    
    PopupMenu mnuLstMain
End Sub

Private Sub lstmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Dim lvhti As LVHITTESTINFO
   Dim lItemIndex As Long
   Dim lvs As ListItem
   lvhti.pt.X = X / Screen.TwipsPerPixelX
   lvhti.pt.Y = Y / Screen.TwipsPerPixelY
   lItemIndex = SendMessage(lstmain.hwnd, LVM_HITTEST, 0, lvhti) + 1
   
   If m_lCurItemIndex <> lItemIndex Then
      m_lCurItemIndex = lItemIndex
      If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
         TT.Destroy
      Else
      Set lvs = lstmain.ListItems(m_lCurItemIndex)
         TT.Title = "Property Info "
         TT.TipText = lstmain.ColumnHeaders.Item(2) & ": " & lvs.SubItems(1) _
         & vbCrLf & lstmain.ColumnHeaders.Item(3) & ": " & lvs.SubItems(2) _
         & vbCrLf & lstmain.ColumnHeaders.Item(4) & ": " & lvs.SubItems(3) _
         & vbCrLf & lstmain.ColumnHeaders.Item(5) & ": " & lvs.SubItems(4)
         TT.Create lstmain.hwnd
      End If
   End If
End Sub

Private Sub lstmain_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo Err
Dim LstIcon As Integer
Dim the_array() As String
Dim list_item As ListItem
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim r As Long
Dim C As Long

    ' Load the file.
    file_name = Data.Files(1)
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ",")
    num_cols = UBound(one_line)
    ReDim the_array(num_rows, num_cols)

    ' Copy the data into the array.
    For r = 0 To num_rows
        one_line = Split(lines(r), ",")
        For C = 0 To num_cols
            the_array(r, C) = one_line(C)
        Next C
    Next r
    
    ' Prove we have the data loaded.
For i = 1 To r
        If xx >= r Then xx = 0
Dim cb As String
    cb = the_array(xx, 2)

If cb = "Boolean" Then
LstIcon = 3
ElseIf cb = "Byte" Then
LstIcon = 3
ElseIf cb = "Currency" Then
LstIcon = 3
ElseIf cb = "Date" Then
LstIcon = 3
ElseIf cb = "Double" Then
LstIcon = 3
ElseIf cb = "Integer" Then
LstIcon = 3
ElseIf cb = "Long" Then
LstIcon = 3
ElseIf cb = "New" Then
LstIcon = 3
ElseIf cb = "OLE_CANCELBOOL" Then
LstIcon = 3
ElseIf cb = "OLE_COLOR" Then
LstIcon = 3
ElseIf cb = "OLE_HANDLE" Then
LstIcon = 3
ElseIf cb = "OLE_OPTEXCLUSIVE" Then
LstIcon = 3
ElseIf cb = "Single" Then
LstIcon = 3
ElseIf cb = "StdFont" Then
LstIcon = 4
ElseIf cb = "StdPicture" Then
LstIcon = 4
ElseIf cb = "String" Then
LstIcon = 3
ElseIf cb = "Variant" Then
LstIcon = 3
End If
    Set list_item = lstmain.ListItems.Add(, , lstmain.ListItems.Count + 1)
    list_item.SmallIcon = LstIcon
    list_item.SubItems(1) = the_array(xx, 1)
    list_item.SubItems(2) = the_array(xx, 2)
    list_item.SubItems(3) = the_array(xx, 3)
    list_item.SubItems(4) = the_array(xx, 4)
xx = xx + 1
Next i

listcount
checklst
DocChanged = True
    Exit Sub
Err:
    MsgBox "The File could not be loaded", vbExclamation
End Sub

Private Sub mnuEditAdd_Click()
cmdmain_Click 0
End Sub

Private Sub mnuEditCopy_Click()
cmdmain_Click 3
End Sub

Private Sub mnuEditGen_Click()
cmdmain_Click 2
End Sub

Private Sub mnuEditRemove_Click()
cmdmain_Click 1
End Sub

Private Sub mnuEditRemoveAll_Click()
cmdmain_Click 6
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileImp_Click()
LoadArray
listcount
checklst
End Sub

Private Sub mnuFileLoad_Click()
mnuFileNew_Click
LoadArray
listcount
checklst
End Sub

Private Sub mnuFileNew_Click()
Dim Cancel As Integer

If DocChanged = False Then
    DocChanged = False
    lstmain.ListItems.Clear
    Rich.Text = ""
    listcount
    checklst
    ClearTxt
Else
    Select Case MsgBox("The file has changed." & vbCr & vbCr & _
            "Do you wish to save your changes?", _
            vbExclamation + vbYesNoCancel, "ActiveX Coder 3")
    
    Case vbYes
        mnuFileS_Click
    Case vbNo
        DocChanged = False
        lstmain.ListItems.Clear
        Rich.Text = ""
        listcount
        checklst
        ClearTxt
    Case vbCancel
        Cancel = True
    
    End Select
End If

End Sub

Private Sub SaveNow()
On Error Resume Next
Rich.Text = ""
For i = 1 To lstmain.ListItems.Count
Rich.Text = Rich.Text & lstmain.ListItems.Item(i)
Rich.Text = Rich.Text & "," & lstmain.ListItems.Item(i).SubItems(1)
Rich.Text = Rich.Text & "," & lstmain.ListItems.Item(i).SubItems(2)
Rich.Text = Rich.Text & "," & lstmain.ListItems.Item(i).SubItems(3)
Rich.Text = Rich.Text & "," & lstmain.ListItems.Item(i).SubItems(4)
If i = lstmain.ListItems.Count Then GoTo save:
Rich.Text = Rich.Text & vbNewLine
Next

save:

End Sub


Private Sub mnuFileS_Click()
Call SaveNow
If docname = "" Then
    mnuFileSS_Click
Else
Rich.SaveFile docname, rtfText
DocChanged = False
End If

End Sub

Private Sub mnuFileSS_Click()
Call SaveNow
Dim Cancel As Boolean
On Error GoTo errorhandler
Cancel = False

CDL1.DefaultExt = ".txt"
CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
CDL1.CancelError = True
CDL1.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt

CDL1.ShowSave

If Not Cancel Then
    If UCase(Right(CDL1.FileName, 3)) = "RTF" Then
        Rich.SaveFile CDL1.FileName, rtfRTF
    Else
        Rich.SaveFile CDL1.FileName, rtfText
    End If
    Rich.FileName = CDL1.FileName
    docname = CDL1.FileName
    Me.Caption = App.Title & " " & docname
    DocChanged = False
End If

Exit Sub

errorhandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If

End Sub

Private Sub mnuHelpAbout_Click()
 ShellAbout Me.hwnd, App.Title, "Coded by: Jovica Mizdrak" _
        & vbNewLine & "E-mail: j3d_jovica@hotmail.com", ByVal 0&
End Sub

Private Sub listcount()
If lstmain.ListItems.Count >= 1 Then
lbllstcount.Caption = FindLVCHKED & " of " & lstmain.ListItems.Count & " are selected"
Else
lbllstcount.Caption = "List is empty"
End If

End Sub

Private Sub checklst()
        If lstmain.ListItems.Count = 0 Then
        cmdmain(5).Enabled = False
        mnuEditGen.Enabled = False
        mnuEditCopy.Enabled = False
        cmdmain(1).Enabled = False
        mnuEditRemove.Enabled = False
        mnuEditRemoveAll.Enabled = False
        Command4.Enabled = False
        End If
        
        If Not lstmain.ListItems.Count = 0 Then
        cmdmain(5).Enabled = True
        mnuEditGen.Enabled = True
        mnuEditCopy.Enabled = True
        cmdmain(1).Enabled = True
        mnuEditRemove.Enabled = True
        mnuEditRemoveAll.Enabled = True
        Command4.Enabled = True
        End If
        
        If DocChanged = True Then
        mnuFileS.Enabled = True
        End If
        
        If DocChanged = False Then
        mnuFileS.Enabled = False
        End If
        
        If lstmain.ListItems.Count = 0 Then
        mnuFileS.Enabled = False
        Else
        mnuFileS.Enabled = True
        End If
         
End Sub

Private Sub mnuLstMainDelR_Click()
cmdmain_Click 6
End Sub

Private Sub mnuLstMainEdit_Click()
On Error Resume Next
frmRow.open_dlg lstmain.SelectedItem.SubItems(1), _
                lstmain.SelectedItem.SubItems(2), _
                lstmain.SelectedItem.SubItems(3), _
                lstmain.SelectedItem.SubItems(4)
End Sub

Private Sub mnuLstMainRR_Click()
cmdmain_Click 1
End Sub

Private Sub LoadArray()
On Error GoTo exits:
Dim LstIcon As Integer
Dim the_array() As String
Dim list_item As ListItem
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Long
Dim num_cols As Long
Dim r As Long
Dim C As Long

Dim Cancel As Boolean
On Error GoTo errorhandler
Cancel = False

CDL1.Filter = "Text Files (*.txt)|*.txt|RichText Files (*.rtf)|*.rtf|All Files|*.*"
CDL1.CancelError = True
CDL1.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
CDL1.ShowOpen

If Not Cancel Then
        file_name = CDL1.FileName
        docname = file_name
        Me.Caption = App.Title & " " & docname
        DocChanged = False
End If
GoTo loadl:
' -------------------
errorhandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If

loadl:
On Error GoTo exits:
    ' Load the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ",")
    num_cols = UBound(one_line)
    ReDim the_array(num_rows, num_cols)

    ' Copy the data into the array.
    For r = 0 To num_rows
        one_line = Split(lines(r), ",")
        For C = 0 To num_cols
            the_array(r, C) = one_line(C)
        Next C
    Next r
    
    ' Prove we have the data loaded.

For i = 1 To r
    If xx >= r Then xx = 0
Dim cb As String
    cb = the_array(xx, 2)

If cb = "Boolean" Then
LstIcon = 3
ElseIf cb = "Byte" Then
LstIcon = 3
ElseIf cb = "Currency" Then
LstIcon = 3
ElseIf cb = "Date" Then
LstIcon = 3
ElseIf cb = "Double" Then
LstIcon = 3
ElseIf cb = "Integer" Then
LstIcon = 3
ElseIf cb = "Long" Then
LstIcon = 3
ElseIf cb = "New" Then
LstIcon = 3
ElseIf cb = "OLE_CANCELBOOL" Then
LstIcon = 3
ElseIf cb = "OLE_COLOR" Then
LstIcon = 3
ElseIf cb = "OLE_HANDLE" Then
LstIcon = 3
ElseIf cb = "OLE_OPTEXCLUSIVE" Then
LstIcon = 3
ElseIf cb = "Single" Then
LstIcon = 3
ElseIf cb = "StdFont" Then
LstIcon = 4
ElseIf cb = "StdPicture" Then
LstIcon = 4
ElseIf cb = "String" Then
LstIcon = 3
ElseIf cb = "Variant" Then
LstIcon = 3
End If
    Set list_item = lstmain.ListItems.Add(, , lstmain.ListItems.Count + 1)
    list_item.SmallIcon = LstIcon
    list_item.SubItems(1) = the_array(xx, 1)
    list_item.SubItems(2) = the_array(xx, 2)
    list_item.SubItems(3) = the_array(xx, 3)
    list_item.SubItems(4) = the_array(xx, 4)
xx = xx + 1
Next i

exits:
DocChanged = True
End Sub

Private Sub Rich_KeyUp(KeyCode As Integer, Shift As Integer)
'setcolors
End Sub

Private Sub Rich_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'setcolors
End Sub

Private Sub Rich_Change()
'setcolors
End Sub

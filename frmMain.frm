VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "qbd Database Code Creation Add-in"
   ClientHeight    =   4425
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   7185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert Code"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cmData 
      Left            =   3480
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   6975
      TabIndex        =   2
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox picInstructions 
         BackColor       =   &H00C0FFFF&
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3795
         ScaleWidth      =   6915
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   6975
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3855
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   14
            Text            =   "frmMain.frx":014A
            Top             =   0
            Width           =   6855
         End
      End
      Begin VB.PictureBox picData 
         AutoRedraw      =   -1  'True
         ClipControls    =   0   'False
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2475
         ScaleWidth      =   1995
         TabIndex        =   7
         Top             =   120
         Width           =   2055
         Begin VB.PictureBox picDataBar 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   8
            Top             =   0
            Width           =   1935
            Begin VB.Label lblData 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               BackStyle       =   0  'Transparent
               Caption         =   "Database"
               ForeColor       =   &H8000000E&
               Height          =   225
               Left            =   120
               TabIndex        =   12
               Top             =   0
               Width           =   1695
            End
         End
         Begin ComctlLib.TreeView tvData 
            Height          =   1455
            Left            =   0
            TabIndex        =   9
            ToolTipText     =   "The layout of the Database"
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   2566
            _Version        =   327682
            HideSelection   =   0   'False
            Indentation     =   459
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   0
         End
      End
      Begin VB.PictureBox picDetails 
         AutoRedraw      =   -1  'True
         ClipControls    =   0   'False
         Height          =   2535
         Left            =   2280
         ScaleHeight     =   2475
         ScaleWidth      =   2235
         TabIndex        =   3
         Top             =   120
         Width           =   2295
         Begin VB.PictureBox picDetailsBar 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1935
            TabIndex        =   4
            Top             =   0
            Width           =   1935
            Begin VB.Label lblDetails 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               BackStyle       =   0  'Transparent
               Caption         =   "Information"
               ForeColor       =   &H8000000E&
               Height          =   225
               Left            =   120
               TabIndex        =   5
               Top             =   0
               Width           =   1695
            End
         End
         Begin ComctlLib.ListView lvDetails 
            Height          =   1455
            Left            =   0
            TabIndex        =   6
            ToolTipText     =   "Details properties of individual items"
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Attribute"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Details"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2160
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image imgSplit 
         Height          =   2535
         Left            =   2160
         MousePointer    =   9  'Size W E
         Top             =   120
         Width           =   60
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   6000
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin ComctlLib.ImageList imlTree 
      Left            =   4080
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0457
            Key             =   "Field"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0631
            Key             =   "dbOpen"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":080B
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":09E5
            Key             =   "Relation"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0BBF
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D99
            Key             =   "Query"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      DatabaseCoder 1.3.1
' Type:         Add-In
' Author:       edward moth
' Copyright:    © 1999-2000 qbd software ltd
'
' FACTORS:
'   Fun Factor:   Nil
'   Usefulness:   Fair to Middling - Dull code automation
'   CoolCode:     Minimal to Non-existent although we like the
'                 Information_SQL routine - simple but effective
'
' What exactly does it do edward?
' Looks at an Access97 database, yawns for a second or two.
' Shows a pretty TreeView of the database structure and allows
' you to click on items and get the exciting (not) lowdown on
' them.  Then it churns out dull code to create the database
' from scratch and puts it in a new module in your project.
'
' See the README.TXT file under RELATED DOCUMENTS for further
' Information
'
' ==============================================================
' Module:       frmMain
' Purpose:      Front-end/Does the work
' ==============================================================
Option Explicit

Private Type qtSplitterMove
    sLeft As Single
    sRight As Single
    bMove As Boolean
End Type

Private sText As String
Private sQuery As String
Private qSplit As qtSplitterMove ' Properties for the splitters
Private eWindowState As FormWindowStateConstants
Public VBInstance As VBIDE.VBE
Public Connect As Connect




Private Sub cmdExit_Click()
    Connect.Hide
End Sub




Private Sub cmdInsert_Click()

Dim iCode As Long
Dim vProject As VBProject
Dim vMod As VBComponent
Dim sName As String
Dim bModule As Boolean
Dim eResponse As VbMsgBoxResult

On Error GoTo InsertErr

If VBInstance.VBProjects.Count = 0 Then
    MsgBox "There are no projects to Add the Code to."
    Exit Sub
End If
Me.MousePointer = vbHourglass
sName = VBInstance.ActiveVBProject.Name
Set vProject = VBInstance.VBProjects.Item(sName)

sName = "mdbCreator"

For iCode = 1 To vProject.VBComponents.Count
If vProject.VBComponents(iCode).Name = sName Then
    bModule = True
End If
Next
If bModule Then
    eResponse = MsgBox("A module named 'mdbCreator' already exists." & vbCrLf _
                    & "Add code to this module?", vbYesNo, "Warning: Module Exists")
    If eResponse = vbNo Then
        Exit Sub
    Else
        Set vMod = vProject.VBComponents.Item(sName)
    End If
Else
    Set vMod = vProject.VBComponents.Add(vbext_ct_StdModule)
    vMod.Name = sName
End If

With vMod
    .CodeModule.AddFromString sText
    If qDB.Queries > 0 Then
    .CodeModule.AddFromString sQuery
    End If
    .Description = "qbd Database Code Creator"
End With



Set vProject = Nothing
Set vMod = Nothing
Me.MousePointer = vbDefault
Exit Sub

InsertErr:
MsgBox "An error occured while trying to create code module." & vbCrLf & "Error: " & Err.Description
cmdInsert.Enabled = False
Me.MousePointer = vbDefault
End Sub

Private Sub cmdOpen_Click()

Dim bProgress As Boolean
Information_Clear
bProgress = Database_Open
If Not bProgress Then
    cmdInsert.Enabled = False
    Exit Sub
End If
DoEvents
Me.MousePointer = vbHourglass

lblData.Caption = qDB.Name
Me.Refresh
Information_Update
DoEvents
Me.Refresh
bProgress = Database_Compile

cmdInsert.Enabled = bProgress
Me.MousePointer = vbDefault


End Sub





Private Sub imgSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

qSplit.bMove = True
Main_SplitterMove X

End Sub

Private Sub imgSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If qSplit.bMove Then
Main_SplitterMove X
End If

End Sub

Private Sub imgSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With qSplit
If .bMove Then
.bMove = False
picSplit.Visible = False
Main_Resize
End If
End With

End Sub

Public Sub Main_SplitterMove(ByVal X As Single)


With qSplit
    X = X + imgSplit.Left
    If X < .sLeft Then
        X = .sLeft
    ElseIf X > .sRight Then
        X = .sRight
    End If
End With

imgSplit.Move X
picSplit.Move X

picDetails.Left = X + 60
picData.Width = X

picSplit.Visible = True

End Sub


Private Sub Form_Load()

'Check to see if a Project is open

If VBInstance.VBProjects.Count = 0 Then
MsgBox "There are no Projects available.  Open a project and restart this Add-In."
Connect.Hide
Exit Sub
End If


tvData.ImageList = imlTree
Information_FieldType
Information_Clear

End Sub


Private Sub Form_Resize()

Main_Resize

End Sub


Private Sub Main_Resize()

Dim sngTemp As Single

If eWindowState = vbMinimized Then
    eWindowState = Me.WindowState
    Exit Sub
End If
eWindowState = Me.WindowState
If eWindowState = vbMinimized Then
    Exit Sub
End If

With picMain
    qSplit.sRight = .ScaleWidth \ 2 + 30
    qSplit.sLeft = .ScaleWidth \ 4 + 30
    imgSplit.Height = .ScaleHeight
End With
    If imgSplit.Left < qSplit.sLeft Then
        imgSplit.Left = qSplit.sLeft
    End If

    If imgSplit.Left > qSplit.sRight Then
        imgSplit.Left = qSplit.sRight
    End If


    
With imgSplit
    picSplit.Move .Left, .Top, .Width, .Height
    picData.Move 0, 0, .Left, .Height
    sngTemp = .Left + 60
    picDetails.Move sngTemp, 0, picMain.ScaleWidth - sngTemp, picMain.ScaleHeight
End With

picDataBar.Move 0, 0, picData.ScaleWidth
lblData.Width = picDataBar.ScaleWidth - 120

picDetailsBar.Move 0, 0, picDetails.ScaleWidth
lblDetails.Width = picDetailsBar.ScaleWidth - 120

sngTemp = picDataBar.Height
tvData.Move 0, sngTemp, picData.ScaleWidth, picData.ScaleHeight - sngTemp
lvDetails.Move 0, sngTemp, picDetails.ScaleWidth, picDetails.ScaleHeight - sngTemp
If lvDetails.Width > 5340 Then
    lvDetails.ColumnHeaders(1).Width = (lvDetails.Width - 840) / 3
    lvDetails.ColumnHeaders(2).Width = (lvDetails.Width - 840) / 3 * 2
Else
    lvDetails.ColumnHeaders(1).Width = 1500
    lvDetails.ColumnHeaders(2).Width = 3000
End If


End Sub

Public Sub Information_Clear()
Dim tvNode As Node
lblData.Caption = "Database"
lblData.Caption = "Information"
cmdInsert.Enabled = False
tvData.Nodes.Clear
Set tvNode = tvData.Nodes.Add(, , "Main", "No database loaded")
lvDetails.ListItems.Clear
End Sub

Public Function Database_Open() As Boolean

On Local Error GoTo Database_Open_Error

cmData.Filter = "Access Database (*.mdb)|*.mdb|All files (*.*)|*.*"
cmData.FilterIndex = 0
cmData.DialogTitle = "Open File..."
cmData.ShowOpen
Set qData = Nothing
Set qData = DBEngine.OpenDatabase(cmData.FileName, True, True)
qDB.Name = cmData.FileTitle
Database_Open = True
Exit Function

Database_Open_Error:
Database_Open = False
If Err.Number = cdlCancel Or Err.Number = 3059 Then
    Exit Function
End If
MsgBox "An error occured while trying to open " & cmData.FileName & vbCrLf & "Error: " & Err.Description

End Function

Public Sub Information_Update()

Dim iTable As Integer
Dim iRelate As Integer
Dim iIndex As Integer
Dim iField As Integer
Dim iCount As Integer
Dim qTable As TableDef
Dim sTableNode As String
Dim qField As Field
Dim qIndex As Index
Dim qRelation As Relation
Dim qQuery As QueryDef
Dim sSQLQueryText As String
Dim qNode As Node
Dim iNode As Integer

ReDim qlNode(0)
ReDim qlTable(0)
ReDim qlRelation(0)
ReDim qlField(0)
ReDim qlIndex(0)
ReDim qlQuery(0)

With qDB
    .Relations = qData.Relations.Count
    .Tables = qData.TableDefs.Count
    .Queries = qData.QueryDefs.Count
    .Fields = 0
    .Indexes = 0
    If .Relations > 1 Or .Tables > 1 Then
    .ItemCount = True
    Else
    .ItemCount = False
    End If
End With

tvData.Nodes.Clear
Set qNode = tvData.Nodes.Add(, tvwFirst, "D0", "Database: " & qDB.Name, "dbOpen")
qNode.Tag = 0
qlNode(0).Name = qDB.Name
qlNode(0).Reference = 0
qlNode(0).Type = qdDatabase
iNode = 1
ReDim qlTable(0 To qDB.Tables)
iTable = 0

Do While iTable <= qDB.Tables - 1

Set qTable = qData.TableDefs(iTable)
If CBool(qTable.Attributes And TableDefAttributeEnum.dbSystemObject) Then
qlTable(iTable).Name = "#"
GoTo IU_Table_Complete
End If

With qlTable(iTable)
    .Name = qTable.Name
    .Fields = qTable.Fields.Count
    If .Fields > 1 Or .Indexes > 1 Then
    qDB.ItemCount = True
    End If
    qDB.Fields = qDB.Fields + .Fields
'    qDB.Indexes = qDB.Indexes + .Indexes
sTableNode = "T" & iNode
Set qNode = tvData.Nodes.Add("D0", tvwChild, sTableNode, "Table: " & .Name, "Table")
qNode.Tag = iNode
ReDim Preserve qlNode(iNode)
qlNode(iNode).Name = .Name
qlNode(iNode).Reference = iTable
qlNode(iNode).Type = qdTable
iNode = iNode + 1
' Get table attributes
If CBool(qTable.Attributes And TableDefAttributeEnum.dbAttachedODBC) Then
.Attributes = Attributes_Add(.Attributes, "dbAttachedODBC")
End If
If CBool(qTable.Attributes And TableDefAttributeEnum.dbAttachedTable) Then
.Attributes = Attributes_Add(.Attributes, "dbAttachedTable")
End If
If CBool(qTable.Attributes And TableDefAttributeEnum.dbAttachExclusive) Then
.Attributes = Attributes_Add(.Attributes, "dbAttachExclusive")
End If
If CBool(qTable.Attributes And TableDefAttributeEnum.dbAttachSavePWD) Then
.Attributes = Attributes_Add(.Attributes, "dbAttachSavePWD")
End If
If CBool(qTable.Attributes And TableDefAttributeEnum.dbHiddenObject) Then
.Attributes = Attributes_Add(.Attributes, "dbHiddenObject")
End If
If CBool(qTable.Attributes And TableDefAttributeEnum.dbSystemObject) Then
.Attributes = Attributes_Add(.Attributes, "dbSystemObject")
End If
End With



' Get Field information

iCount = 0
Do While iCount <= qlTable(iTable).Fields - 1

Set qField = qTable.Fields(iCount)
ReDim Preserve qlField(0 To iField)

With qlField(iField)
    .Name = qField.Name
    .DefaultValue = qField.DefaultValue
    .Required = qField.Required
    .Size = qField.Size
    .Type = qField.Type
    .Table = iTable
    .Index = False
    If CBool(qField.Attributes And FieldAttributeEnum.dbAutoIncrField) Then
    .Attributes = Attributes_Add(.Attributes, "dbAutoIncrField")
    End If
    If CBool(qField.Attributes And FieldAttributeEnum.dbFixedField) Then
    .Attributes = Attributes_Add(.Attributes, "dbFixedField")
    End If
    If CBool(qField.Attributes And FieldAttributeEnum.dbHyperlinkField) Then
    .Attributes = Attributes_Add(.Attributes, "dbHyperlinkField")
    End If
    If CBool(qField.Attributes And FieldAttributeEnum.dbSystemField) Then
    .Attributes = Attributes_Add(.Attributes, "dbSystemField")
    End If
    If CBool(qField.Attributes And FieldAttributeEnum.dbUpdatableField) Then
    .Attributes = Attributes_Add(.Attributes, "dbUpdatableField")
    End If
    If CBool(qField.Attributes And FieldAttributeEnum.dbVariableField) Then
    .Attributes = Attributes_Add(.Attributes, "dbVariableField")
    End If
    Set qNode = tvData.Nodes.Add(sTableNode, tvwChild, "F" & iNode, "Field: " & .Name, "Field")
    qNode.Tag = iNode
    ReDim Preserve qlNode(iNode)
    qlNode(iNode).Name = .Name
    qlNode(iNode).Reference = iField
    qlNode(iNode).Type = qdField
    iNode = iNode + 1
End With
iField = iField + 1

iCount = iCount + 1
Loop


'Find Index information
iCount = 0
Do While iCount <= qTable.Indexes.Count - 1 ' qlTable(iTable).Indexes - 1
Set qIndex = qTable.Indexes(iCount)
If Not qTable.Indexes(iCount).Foreign Then
qlTable(iTable).Indexes = qlTable(iTable).Indexes + 1
ReDim Preserve qlIndex(0 To iIndex)

' Get Index information
With qlIndex(iIndex)

    .Name = qIndex.Name
    .FieldIndex = Information_Index_Get(qIndex.Fields(0).Name, qdField, iTable)
    qlField(.FieldIndex).Index = True
    .Sort = CBool(qIndex.Fields(0).Attributes And dbDescending)
    .Table = iTable
    .Primary = qIndex.Primary
    .Required = qIndex.Required
    .Unique = qIndex.Unique
    Set qNode = tvData.Nodes.Add(sTableNode, tvwChild, "I" & iNode, "Index: " & .Name, "Index")
    qNode.Tag = iNode
    ReDim Preserve qlNode(iNode)
    qlNode(iNode).Name = .Name
    qlNode(iNode).Reference = iIndex
    qlNode(iNode).Type = qdIndex
    iNode = iNode + 1
End With
iIndex = iIndex + 1
End If
iCount = iCount + 1
Loop

qDB.Indexes = qDB.Indexes + qlTable(iTable).Indexes


IU_Table_Complete:

iTable = iTable + 1
Loop

' Query Information
If qDB.Queries > 0 Then
Set qNode = tvData.Nodes.Add("D0", tvwChild, "QUERY", "Queries", "Query")
qNode.Tag = iNode
ReDim Preserve qlNode(iNode)
qlNode(iNode).Name = "Queries"
qlNode(iNode).Reference = 0
qlNode(iNode).Type = qdQueries
iNode = iNode + 1
End If
iCount = 0
Do While iCount <= qDB.Queries - 1
Set qQuery = qData.QueryDefs(iCount)
ReDim Preserve qlQuery(0 To iCount)
With qlQuery(iCount)
.Name = qQuery.Name
.Fields = qQuery.Fields.Count
.Type = qQuery.Type

Select Case .Type
Case QueryDefTypeEnum.dbQAction
.TypeText = "Action"
Case QueryDefTypeEnum.dbQAppend
.TypeText = "Append"
Case QueryDefTypeEnum.dbQCompound
.TypeText = "Compound"
Case QueryDefTypeEnum.dbQCrosstab
.TypeText = "Crosstab"
Case QueryDefTypeEnum.dbQDDL
.TypeText = "DDL"
Case QueryDefTypeEnum.dbQDelete
.TypeText = "Delete"
Case QueryDefTypeEnum.dbQMakeTable
.TypeText = "Make Table"
Case QueryDefTypeEnum.dbQProcedure
.TypeText = "Procedure"
Case QueryDefTypeEnum.dbQSelect
.TypeText = "Select"
Case QueryDefTypeEnum.dbQSetOperation
.TypeText = "Set Operation"
Case QueryDefTypeEnum.dbQSPTBulk
.TypeText = "SPT Bulk"
Case QueryDefTypeEnum.dbQSQLPassThrough
.TypeText = "SQL Pass Through"
Case QueryDefTypeEnum.dbQUpdate
.TypeText = "Update"
Case Else
.TypeText = .Type
End Select
.SQLText = Information_SQL(qQuery.SQL)

Set qNode = tvData.Nodes.Add("QUERY", tvwChild, "Q" & iNode, "Query: " & .Name, "Query")
qNode.Tag = iNode
ReDim Preserve qlNode(iNode)
qlNode(iNode).Name = .Name
qlNode(iNode).Reference = iCount
qlNode(iNode).Type = qdQuery
iNode = iNode + 1

iCount = iCount + 1
End With
Loop


Do While iRelate <= qDB.Relations - 1
Set qRelation = qData.Relations(iRelate)
ReDim Preserve qlRelation(0 To iRelate)
With qlRelation(iRelate)
    .Name = qRelation.Name
    .Table = Information_Index_Get(qRelation.Table, qdTable, 0)
    .ForeignTable = Information_Index_Get(qRelation.ForeignTable, qdTable, 0)
    .Field = Information_Index_Get(qRelation.Fields(0).Name, qdField, .Table)
    .ForeignField = Information_Index_Get(qRelation.Fields(0).ForeignName, qdField, .ForeignTable)

    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationDeleteCascade) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationDeleteCascade")
    End If
    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationDontEnforce) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationDontEnforce")
    End If
    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationInherited) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationInherited")
    End If
    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationLeft) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationLeft")
    End If
    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationRight) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationRight")
    End If
    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationUnique) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationUnique")
    End If
    If CBool(qRelation.Attributes And RelationAttributeEnum.dbRelationUpdateCascade) Then
    .Attributes = Attributes_Add(.Attributes, "dbRelationUpdateCascade")
    End If
    Set qNode = tvData.Nodes.Add("D0", tvwChild, "R" & iNode, "Relation: " & .Name, "Relation")
    qNode.Tag = iNode
    ReDim Preserve qlNode(iNode)
    qlNode(iNode).Name = .Name
    qlNode(iNode).Reference = iRelate
    qlNode(iNode).Type = qdRelation
    iNode = iNode + 1
End With
iRelate = iRelate + 1
Loop
    

tvData.Nodes("D0").Selected = True
Information_Item_Get 0

Set qData = Nothing
Set qTable = Nothing
Set qRelation = Nothing
Set qField = Nothing
Set qIndex = Nothing

End Sub

Private Function Attributes_Add(ByVal sText As String _
                                , ByVal sNew As String) As String

If sText <> "" Then
sText = sText & " + "
End If
sText = sText & sNew
Attributes_Add = sText

End Function



Private Sub tvData_NodeClick(ByVal Node As ComctlLib.Node)
Node.EnsureVisible

If Node.Key = "Main" Then
Exit Sub
End If

Information_Item_Get Node.Tag

End Sub

Private Sub Information_Item_Get(ByVal iNode As Integer)

Dim iRef As Integer
Dim lvItem As ListItem


lvDetails.ListItems.Clear
iRef = qlNode(iNode).Reference

Select Case qlNode(iNode).Type
Case qDatabaseObjectEnum.qdDatabase
With qDB
lblDetails.Caption = "Database: " & .Name
Set lvItem = lvDetails.ListItems.Add(1, , "Name")
lvItem.SubItems(1) = .Name
Set lvItem = lvDetails.ListItems.Add(2, , "Object")
lvItem.SubItems(1) = "Database"
Set lvItem = lvDetails.ListItems.Add(3, , "Tables")
lvItem.SubItems(1) = .Tables
Set lvItem = lvDetails.ListItems.Add(4, , "Tables")
lvItem.SubItems(1) = .Tables
Set lvItem = lvDetails.ListItems.Add(5, , "Queries")
lvItem.SubItems(1) = .Queries
Set lvItem = lvDetails.ListItems.Add(6, , "Indexes")
lvItem.SubItems(1) = .Indexes
Set lvItem = lvDetails.ListItems.Add(7, , "Fields")
lvItem.SubItems(1) = .Fields
End With

Case qDatabaseObjectEnum.qdTable
With qlTable(iRef)
lblDetails.Caption = "Table: " & .Name
Set lvItem = lvDetails.ListItems.Add(1, , "Name")
lvItem.SubItems(1) = .Name
Set lvItem = lvDetails.ListItems.Add(2, , "Object")
lvItem.SubItems(1) = "Table"
Set lvItem = lvDetails.ListItems.Add(3, , "Attributes")
lvItem.SubItems(1) = .Attributes
Set lvItem = lvDetails.ListItems.Add(4, , "Indexes")
lvItem.SubItems(1) = .Indexes
Set lvItem = lvDetails.ListItems.Add(5, , "Fields")
lvItem.SubItems(1) = .Fields
End With

Case qDatabaseObjectEnum.qdIndex
With qlIndex(iRef)
lblDetails.Caption = "Index: " & .Name
Set lvItem = lvDetails.ListItems.Add(1, , "Name")
lvItem.SubItems(1) = .Name
Set lvItem = lvDetails.ListItems.Add(2, , "Object")
lvItem.SubItems(1) = "Index"
Set lvItem = lvDetails.ListItems.Add(3, , "Field")
lvItem.SubItems(1) = qlField(.FieldIndex).Name
Set lvItem = lvDetails.ListItems.Add(4, , "Table")
lvItem.SubItems(1) = qlTable(.Table).Name
Set lvItem = lvDetails.ListItems.Add(5, , "Primary")
lvItem.SubItems(1) = .Primary
Set lvItem = lvDetails.ListItems.Add(6, , "Required")
lvItem.SubItems(1) = .Required
Set lvItem = lvDetails.ListItems.Add(7, , "Unique")
lvItem.SubItems(1) = .Unique
Set lvItem = lvDetails.ListItems.Add(8, , "Sort")
If .Sort Then
lvItem.SubItems(1) = "Descending"
Else
lvItem.SubItems(1) = "Ascending"
End If
End With

Case qDatabaseObjectEnum.qdField
With qlField(iRef)
lblDetails.Caption = "Field: " & .Name
Set lvItem = lvDetails.ListItems.Add(1, , "Name")
lvItem.SubItems(1) = .Name
Set lvItem = lvDetails.ListItems.Add(2, , "Object")
lvItem.SubItems(1) = "Field"
Set lvItem = lvDetails.ListItems.Add(3, , "Attributes")
lvItem.SubItems(1) = .Attributes
Set lvItem = lvDetails.ListItems.Add(4, , "Table")
lvItem.SubItems(1) = qlTable(.Table).Name
Set lvItem = lvDetails.ListItems.Add(5, , "Required")
lvItem.SubItems(1) = .Required
Set lvItem = lvDetails.ListItems.Add(6, , "Type")
lvItem.SubItems(1) = qFType(.Type).Name
Set lvItem = lvDetails.ListItems.Add(7, , "Size")
lvItem.SubItems(1) = .Size
Set lvItem = lvDetails.ListItems.Add(8, , "Default Value")
lvItem.SubItems(1) = .DefaultValue
Set lvItem = lvDetails.ListItems.Add(9, , "Indexed")
lvItem.SubItems(1) = .Index

End With

Case qDatabaseObjectEnum.qdRelation
With qlRelation(iRef)
lblDetails.Caption = "Relation: " & .Name
Set lvItem = lvDetails.ListItems.Add(1, , "Name")
lvItem.SubItems(1) = .Name
Set lvItem = lvDetails.ListItems.Add(2, , "Object")
lvItem.SubItems(1) = "Relation"
Set lvItem = lvDetails.ListItems.Add(3, , "Attributes")
lvItem.SubItems(1) = .Attributes
Set lvItem = lvDetails.ListItems.Add(4, , "Table")
lvItem.SubItems(1) = qlTable(.Table).Name
Set lvItem = lvDetails.ListItems.Add(5, , "Field")
lvItem.SubItems(1) = qlField(.Field).Name
Set lvItem = lvDetails.ListItems.Add(6, , "Foreign Table")
lvItem.SubItems(1) = qlTable(.ForeignTable).Name
Set lvItem = lvDetails.ListItems.Add(7, , "Foreign Field")
lvItem.SubItems(1) = qlField(.ForeignField).Name
End With

Case qDatabaseObjectEnum.qdQueries
lblDetails.Caption = "Queries"
Set lvItem = lvDetails.ListItems.Add(1, , "Count")
lvItem.SubItems(1) = qDB.Queries

Case qDatabaseObjectEnum.qdQuery
With qlQuery(iRef)
lblDetails.Caption = "Query: " & .Name
Set lvItem = lvDetails.ListItems.Add(1, , "Name")
lvItem.SubItems(1) = .Name
Set lvItem = lvDetails.ListItems.Add(2, , "Object")
lvItem.SubItems(1) = "Query"
Set lvItem = lvDetails.ListItems.Add(3, , "Fields")
lvItem.SubItems(1) = .Fields
Set lvItem = lvDetails.ListItems.Add(4, , "Type")
lvItem.SubItems(1) = .TypeText
End With

End Select




End Sub

Private Sub Information_FieldType()


qFType(DataTypeEnum.dbBigInt).Code = "dbBigInt"
qFType(DataTypeEnum.dbBigInt).Name = "Big Integer"
qFType(DataTypeEnum.dbBinary).Code = "dbBinary"
qFType(DataTypeEnum.dbBinary).Name = "Binary"
qFType(DataTypeEnum.dbBoolean).Code = "dbBoolean"
qFType(DataTypeEnum.dbBoolean).Name = "Boolean (True/False)"
qFType(DataTypeEnum.dbByte).Code = "dbByte"
qFType(DataTypeEnum.dbByte).Name = "Byte"
qFType(DataTypeEnum.dbChar).Code = "dbChar"
qFType(DataTypeEnum.dbChar).Name = "Fixed String"
qFType(DataTypeEnum.dbCurrency).Code = "dbCurrency"
qFType(DataTypeEnum.dbCurrency).Name = "Currency"
qFType(DataTypeEnum.dbDate).Code = "dbDate"
qFType(DataTypeEnum.dbDate).Name = "Date"
qFType(DataTypeEnum.dbDecimal).Code = "dbDecimal"
qFType(DataTypeEnum.dbDecimal).Name = "Decimal"
qFType(DataTypeEnum.dbDouble).Code = "dbDouble"
qFType(DataTypeEnum.dbDouble).Name = "Double"
qFType(DataTypeEnum.dbFloat).Code = "dbFloat"
qFType(DataTypeEnum.dbFloat).Name = "Float"
qFType(DataTypeEnum.dbGUID).Code = "dbGUID"
qFType(DataTypeEnum.dbGUID).Name = "GUID (Globally Unique Identifier)"
qFType(DataTypeEnum.dbInteger).Code = "dbInteger"
qFType(DataTypeEnum.dbInteger).Name = "Integer"
qFType(DataTypeEnum.dbLong).Code = "dbLong"
qFType(DataTypeEnum.dbLong).Name = "Long"
qFType(DataTypeEnum.dbLongBinary).Code = "dbLongBinary"
qFType(DataTypeEnum.dbLongBinary).Name = "Long Binary"
qFType(DataTypeEnum.dbMemo).Code = "dbMemo"
qFType(DataTypeEnum.dbMemo).Name = "Memo"
qFType(DataTypeEnum.dbNumeric).Code = "dbNumeric"
qFType(DataTypeEnum.dbNumeric).Name = "Numeric"
qFType(DataTypeEnum.dbSingle).Code = "dbSingle"
qFType(DataTypeEnum.dbSingle).Name = "Single"
qFType(DataTypeEnum.dbText).Code = "dbText"
qFType(DataTypeEnum.dbText).Name = "Text"
qFType(DataTypeEnum.dbTime).Code = "dbTime"
qFType(DataTypeEnum.dbTime).Name = "Time"
qFType(DataTypeEnum.dbTimeStamp).Code = "dbTimeStamp"
qFType(DataTypeEnum.dbTimeStamp).Name = "Time Stamp"
qFType(DataTypeEnum.dbVarBinary).Code = "dbVarBinary"
qFType(DataTypeEnum.dbVarBinary).Name = "Variable length Binary"




End Sub


Private Function Information_Index_Get(ByVal sName As String _
                                      , ByVal sType As qDatabaseObjectEnum _
                                      , ByVal iTable As Integer) As Integer

Dim iCount As Integer
Dim iHit As Integer


If sType = qdField Then
Do While iCount <= qDB.Fields - 1 Or iHit = 0
If qlField(iCount).Name = sName And qlField(iCount).Table = iTable Then
iHit = iCount + 1
End If
iCount = iCount + 1
Loop
Else
Do While iCount <= qDB.Tables - 1 Or iHit = 0
If qlTable(iCount).Name = sName Then
iHit = iCount + 1
End If
iCount = iCount + 1
Loop
End If
iHit = iHit - 1
If iHit < 0 Then
Stop
End If

Information_Index_Get = iHit

End Function

Private Function Database_Compile() As Boolean

Dim iTable As Integer
Dim iCount As Integer
Dim sBack As String
Dim sSubText As String
Dim iSubOption As Integer

On Error GoTo Database_CompileErr

' Create the code for the database
sText = "' ==============================================================" & vbCrLf
sText = sText & "' Module:       CreateDB" & vbCrLf
sText = sText & "' Purpose:      Create Database" & vbCrLf
sText = sText & "' ==============================================================" & vbCrLf
sText = sText & "' qbd DATABASE CODE CREATOR" & vbCrLf
sText = sText & "' ==============================================================" & vbCrLf
sText = sText & "' WHAT TO DO NEXT:" & vbCrLf
sText = sText & "' 1.  Add reference to Microsoft DA0 3.5x Library" & vbCrLf
sText = sText & "' 2.  Check the Database_Create() function for Optional Changes" & vbCrLf
sText = sText & "' 3.  To create a database use:" & vbCrLf
sText = sText & "'     bOkay = Database_Create sFilename" & vbCrLf
sText = sText & "'     Where sFilename is the Path and Name of the Database" & vbCrLf
sText = sText & "'     and bOkay is a boolean return value.  If return is false" & vbCrLf
sText = sText & "'     then the creation routine was unsuccessful." & vbCrLf
sText = sText & "' ==============================================================" & vbCrLf & vbCrLf
sText = sText & "Private dbData as Database" & vbCrLf
sText = sText & "Public Function Database_Create(byVal sFilename as String) As Boolean" & vbCrLf & vbCrLf
sText = sText & "' Code created by the qbd Database Code Creator" & vbCrLf
sText = sText & "' Use Find '#' to check optional settings" & vbCrLf & vbCrLf
sText = sText & "On Error Goto Database_Create_Error" & vbCrLf & vbCrLf
If qDB.Tables > 0 Then
sText = sText & "Dim dtTable as TableDef" & vbCrLf
End If
'If qDB.Relations > 0 Then
'sText = sText & "Dim drRelation as Relation" & vbCrLf
'End If
'If qDB.Indexes > 0 Then
'sText = sText & "Dim diIndex as Index" & vbCrLf
'End If
'If qDB.Fields > 0 Then
'sText = sText & "Dim dfField As Field" & vbCrLf
'End If
If qDB.Relations > 0 Then
iSubOption = iSubOption + 4
End If
If qDB.Indexes > 0 Then
iSubOption = iSubOption + 2
End If
If qDB.Fields > 0 Then
iSubOption = iSubOption + 1
End If


'If qDB.ItemCount Then
'sText = sText & "Dim iItems as Integer" & vbCrLf
'End If

sText = sText & vbCrLf
sText = sText & "' Create the Database" & vbCrLf
sText = sText & "' # Add password: insert '& """ & ";pwd=NewPassword" & """ after dbLangGeneral" & vbCrLf
sText = sText & "' # Encrypt: insert '+ dbEncrypt' after dbVersion30" & vbCrLf
sText = sText & "Set dbData = DBEngine.CreateDatabase(sFilename, dbLangGeneral, dbVersion30)" & vbCrLf & vbCrLf


iTable = 0
Do While iTable <= qDB.Tables - 1
If qlTable(iTable).Name = "#" Then
GoTo DC_Table_Complete
End If
sText = sText & "' Create table:'" & qlTable(iTable).Name & "'" & vbCrLf
sText = sText & "Set dtTable = dbData.CreateTableDef(""" & qlTable(iTable).Name & """"
If qlTable(iTable).Attributes = "" Then
sText = sText & ")" & vbCrLf
Else
sText = sText & ", " & qlTable(iTable).Attributes & ")" & vbCrLf
End If
sText = sText & vbCrLf
If qlTable(iTable).Indexes > 1 Then
sText = sText & vbCrLf & "' Create Indexes for table: " & qlTable(iTable).Name
ElseIf qlTable(iTable).Indexes = 1 Then
sText = sText & vbCrLf & "' Create Index for table: " & qlTable(iTable).Name
End If
sText = sText & vbCrLf
iCount = 0
Do While iCount <= qDB.Indexes - 1
With qlIndex(iCount)
If .Table = iTable Then
'sText = sText & "Set diIndex = dtTable.CreateIndex(""" & .Name & """)" & vbCrLf
'sText = sText & "Set dfField = diIndex.CreateField(""" & qlField(.FieldIndex).Name & """, " & qFType(qlField(.FieldIndex).Type).Code
'If qlField(.FieldIndex).Type = dbText Then
'sText = sText & ", " & qlField(.FieldIndex).Size & ")" & vbCrLf
'Else
'sText = sText & ")" & vbCrLf
'End If
'If .Sort Then
'sText = sText & "dfField.Attributes = dbDescending" & vbCrLf
'End If
'
'sText = sText & vbCrLf & "With diIndex" & vbCrLf
'sText = sText & "    .Fields.Append dfField" & vbCrLf
'sText = sText & "    .Primary = " & qlIndex(iCount).Primary & vbCrLf
'sText = sText & "    .Unique = " & qlIndex(iCount).Unique & vbCrLf
'sText = sText & "End With" & vbCrLf
'sText = sText & "dtTable.Indexes.Append diIndex" & vbCrLf & vbCrLf

sText = sText & "Index_Create dtTable, """ & .Name & """, """ & qlField(.FieldIndex).Name & """," _
              & qFType(qlField(.FieldIndex).Type).Code
sBack = ""
If qlIndex(iCount).Unique Then
sBack = ", True"
End If
If qlIndex(iCount).Primary Then
sBack = ", True" & sBack
ElseIf sBack > "" Then
sBack = ", " & sBack
End If
If qlIndex(iCount).Sort Then
sBack = ", True" & sBack
ElseIf sBack > "" Then
sBack = ", " & sBack
End If
If qlField(.FieldIndex).Type = dbText Then
sBack = ", " & qlField(.FieldIndex).Size & sBack
ElseIf sBack > "" Then
sBack = ", " & sBack
End If

sText = sText & sBack & vbCrLf


End If
End With

iCount = iCount + 1
Loop


sText = sText & "' Create field"
If qlTable(iTable).Fields > 1 Then
sText = sText & "s"
End If
sText = sText & vbCrLf
iCount = 0
Do While iCount <= qDB.Fields - 1
If qlField(iCount).Table <> iTable Then 'Or qlField(iCount).Index Then
GoTo DC_Field_Complete
End If

'sText = sText & "Set dfField = dtTable.CreateField(""" & qlField(iCount).Name & """, " & qFType(qlField(iCount).Type).Code
'If qlField(iCount).Type = dbText Then
'sText = sText & ", " & qlField(iCount).Size & ")" & vbCrLf
'Else
'sText = sText & ")" & vbCrLf
'End If
'
'sText = sText & "With dfField" & vbCrLf
'If qlField(iCount).Attributes > "" Then
'sText = sText & "    .Attributes = " & qlField(iCount).Attributes & vbCrLf
'End If
'sText = sText & "    .Required = " & qlField(iCount).Required & vbCrLf
'If qlField(iCount).DefaultValue > "" Then
'sText = sText & "    .DefaultValue = """ & qlField(iCount).DefaultValue & """" & vbCrLf
'End If
'sText = sText & "End With" & vbCrLf
'sText = sText & "dtTable.Fields.Append dfField" & vbCrLf & vbCrLf
sBack = ""
sText = sText & "Field_Create dtTable, """ & qlField(iCount).Name & """, " _
              & qFType(qlField(iCount).Type).Code


' v1.3.1 Improve default value setting
If IsNumeric(qlField(iCount).DefaultValue) Then
sBack = ", " & qlField(iCount).DefaultValue
ElseIf qlField(iCount).Type = dbBoolean And Len(qlField(iCount).DefaultValue) > 0 Then
sBack = ", " & qlField(iCount).DefaultValue
ElseIf VarType(qlField(iCount).DefaultValue) = vbString And qlField(iCount).DefaultValue > "" Then
sBack = ", " & Information_Default(qlField(iCount).DefaultValue)
End If
If qlField(iCount).Required = True Then
sBack = ", True" & sBack
ElseIf sBack > "" Then
sBack = ", " & sBack
End If
If qlField(iCount).Attributes > "" Then
sBack = ", " & qlField(iCount).Attributes & sBack
ElseIf sBack > "" Then
sBack = ", " & sBack
End If
If qlField(iCount).Type = dbText Then
sBack = ", " & qlField(iCount).Size & sBack
ElseIf sBack > "" Then
sBack = ", " & sBack
End If
sText = sText & sBack & vbCrLf

DC_Field_Complete:
iCount = iCount + 1
Loop

sText = sText & "dbData.TableDefs.Append dtTable" & vbCrLf & vbCrLf

DC_Table_Complete:
iTable = iTable + 1
Loop

If qDB.Relations > 1 Then
sText = sText & vbCrLf & "' Create Relations"
ElseIf qDB.Relations = 1 Then
sText = sText & vbCrLf & "' Create Relation"
End If
sText = sText & vbCrLf


iCount = 0
Do While iCount <= qDB.Relations - 1
With qlRelation(iCount)
sText = sText & "Relation_Create """ & .Name & """, """ & qlTable(.Table).Name _
              & """, """ & qlTable(.ForeignTable).Name & """, """ _
              & qlField(.Field).Name & """, """ & qlField(.ForeignField).Name & """"
If .Attributes > "" Then
sText = sText & ", " & .Attributes
End If
End With
sText = sText & vbCrLf
iCount = iCount + 1
Loop


If qDB.Tables > 0 Then
sText = sText & "Set dtTable = Nothing" & vbCrLf
End If
'If qDB.Relations > 0 Then
'sText = sText & "Set drRelation = Nothing" & vbCrLf
'End If
'If qDB.Indexes > 0 Then
'sText = sText & "Set diIndex = Nothing" & vbCrLf
'End If
'If qDB.Fields > 0 Then
'sText = sText & "Set dfField = Nothing" & vbCrLf
'End If
If qDB.Queries > 0 Then
sText = sText & "' Set up queries" & vbCrLf
sText = sText & "Query_Definition" & vbCrLf
End If

sText = sText & "Set dbData = Nothing" & vbCrLf & vbCrLf

sText = sText & "' Creation Successful" & vbCrLf
sText = sText & "Database_Create = True" & vbCrLf
sText = sText & "Exit Function" & vbCrLf & vbCrLf
sText = sText & "' Whoops an error occured" & vbCrLf
sText = sText & "Database_Create_Error:" & vbCrLf
sText = sText & "' #Add code to trap for errors" & vbCrLf
sText = sText & "Database_Create = False" & vbCrLf
sText = sText & "End Function" & vbCrLf & vbCrLf
sSubText = Add_Subroutines(iSubOption)

sText = sText & sSubText

' Set up Query Information
sQuery = "Private Sub Query_Definition()" & vbCrLf & vbCrLf
sQuery = sQuery & "Dim sSQLText As String" & vbCrLf
sQuery = sQuery & "Dim dqQuery As QueryDef" & vbCrLf & vbCrLf

iCount = 0
Do While iCount < qDB.Queries
sQuery = sQuery & "' QUERY: " & qlQuery(iCount).Name & vbCrLf
sQuery = sQuery & qlQuery(iCount).SQLText
sQuery = sQuery & "set dqQuery = dbData.CreateQueryDef(""" & qlQuery(iCount).Name & """, sSQLText)" & vbCrLf
iCount = iCount + 1
Loop
sQuery = sQuery & vbCrLf & "End Sub" & vbCrLf

Database_Compile = True
Exit Function

Database_CompileErr:
MsgBox "An error occured while analysing the Database." & vbCrLf & "Error: " & Err.Description
Database_Compile = False

End Function


Private Function Add_Subroutines(ByVal iOptions As Integer) As String

Dim sSub As String
If iOptions And 1 = 1 Then
sSub = sSub & "Private Sub Field_Create(dtTable as TableDef, _" & vbCrLf
sSub = sSub & "                         Name As String, _" & vbCrLf
sSub = sSub & "                         FieldType As Integer, _" & vbCrLf
sSub = sSub & "                         Optional Size As Integer = 0, _" & vbCrLf
sSub = sSub & "                         Optional Attributes As Long = 0, _" & vbCrLf
sSub = sSub & "                         Optional Required As Boolean = False, _" & vbCrLf
sSub = sSub & "                         Optional DefaultValue As String = """")" & vbCrLf
sSub = sSub & "Dim dfField As Field" & vbCrLf & vbCrLf
sSub = sSub & "On Error Goto Field_Create_Err" & vbCrLf & vbCrLf
sSub = sSub & "' Create Field in Table: dtTable" & vbCrLf & vbCrLf
sSub = sSub & "If FieldType = dbText Then" & vbCrLf
sSub = sSub & "  Set dfField = dtTable.CreateField(Name, FieldType, Size)" & vbCrLf
sSub = sSub & "Else" & vbCrLf
sSub = sSub & "  Set dfField = dtTable.CreateField(Name, FieldType)" & vbCrLf
sSub = sSub & "End If" & vbCrLf & vbCrLf
sSub = sSub & "dfField.Attributes = Attributes" & vbCrLf
sSub = sSub & "dfField.Required = Required" & vbCrLf
sSub = sSub & "dfField.DefaultValue = DefaultValue" & vbCrLf & vbCrLf
sSub = sSub & "dtTable.Fields.Append dfField" & vbCrLf & vbCrLf
sSub = sSub & "Set dfField = Nothing" & vbCrLf
sSub = sSub & "Exit Sub" & vbCrLf
sSub = sSub & "Field_Create_Err:" & vbCrLf
sSub = sSub & "' Whoops an error occured" & vbCrLf
sSub = sSub & "' #Add code to trap for errors" & vbCrLf
sSub = sSub & "Set dfField = Nothing" & vbCrLf
sSub = sSub & "End Sub" & vbCrLf
End If

If iOptions And 2 = 2 Then
sSub = sSub & "Private Sub Index_Create(dtTable As TableDef, _" & vbCrLf
sSub = sSub & "                         Name As String, _" & vbCrLf
sSub = sSub & "                         FieldName As String, _" & vbCrLf
sSub = sSub & "                         FieldType As DataTypeEnum, _" & vbCrLf
sSub = sSub & "                         Optional Size As Integer = 0, _" & vbCrLf
sSub = sSub & "                         Optional Sort As Boolean = False, _" & vbCrLf
sSub = sSub & "                         Optional Primary As Boolean = False, _" & vbCrLf
sSub = sSub & "                         Optional Unique As Boolean = False)" & vbCrLf & vbCrLf
sSub = sSub & "On Error GoTo Index_Create_Err" & vbCrLf & vbCrLf
sSub = sSub & "Dim diIndex As Index" & vbCrLf
sSub = sSub & "Dim dfField As Field" & vbCrLf & vbCrLf
sSub = sSub & "Set diIndex = dtTable.CreateIndex(Name)" & vbCrLf
sSub = sSub & "Set dfField = diIndex.CreateField(FieldName, FieldType)" & vbCrLf & vbCrLf
sSub = sSub & "If FieldType = dbText Then" & vbCrLf
sSub = sSub & "dfField.Size = Size" & vbCrLf
sSub = sSub & "End If" & vbCrLf & vbCrLf
sSub = sSub & "If Sort Then" & vbCrLf
sSub = sSub & "dfField.Attributes = dbDescending" & vbCrLf
sSub = sSub & "End If" & vbCrLf & vbCrLf
sSub = sSub & "With diIndex" & vbCrLf
sSub = sSub & "  .Fields.Append dfField" & vbCrLf
sSub = sSub & "  .Primary = Primary" & vbCrLf
sSub = sSub & "  .Unique = Unique" & vbCrLf
sSub = sSub & "End With" & vbCrLf & vbCrLf
sSub = sSub & "dtTable.Indexes.Append diIndex" & vbCrLf & vbCrLf
sSub = sSub & "Set diIndex = Nothing" & vbCrLf
sSub = sSub & "Set dfField = Nothing" & vbCrLf
sSub = sSub & "Exit Sub" & vbCrLf & vbCrLf
sSub = sSub & "Index_Create_Err:" & vbCrLf
sSub = sSub & "' Whoops an error occured" & vbCrLf
sSub = sSub & "' #Add code to trap for errors" & vbCrLf
sSub = sSub & "Set diIndex = Nothing" & vbCrLf
sSub = sSub & "Set dfField = Nothing" & vbCrLf & vbCrLf
sSub = sSub & "End Sub" & vbCrLf
End If

If iOptions And 4 = 4 Then
sSub = sSub & "Private Sub Relation_Create(Name As String, _" & vbCrLf
sSub = sSub & "                            Table As String, _" & vbCrLf
sSub = sSub & "                            ForeignTable As String, _" & vbCrLf
sSub = sSub & "                            Field As String, _" & vbCrLf
sSub = sSub & "                            ForeignField As String, _" & vbCrLf
sSub = sSub & "                            Optional Attributes As Long = 0)" & vbCrLf & vbCrLf
sSub = sSub & "On Error GoTo Relation_Create_Err" & vbCrLf & vbCrLf
sSub = sSub & "Dim drRelation As Relation" & vbCrLf
sSub = sSub & "Dim dfField As Field" & vbCrLf
sSub = sSub & "Set drRelation = dbdata.CreateRelation(Name, Table, ForeignTable, Attributes)" & vbCrLf
sSub = sSub & "drRelation.Fields.Append drRelation.CreateField(Field)" & vbCrLf
sSub = sSub & "drRelation.Fields(Field).ForeignName = ForeignField" & vbCrLf
sSub = sSub & "dbdata.Relations.Append drRelation" & vbCrLf & vbCrLf
sSub = sSub & "Set dfField = Nothing" & vbCrLf
sSub = sSub & "Set drRelation = Nothing" & vbCrLf & vbCrLf
sSub = sSub & "Exit Sub" & vbCrLf
sSub = sSub & "Relation_Create_Err:" & vbCrLf
sSub = sSub & "' Whoops an error occured" & vbCrLf
sSub = sSub & "' #Add code to trap for errors" & vbCrLf
sSub = sSub & "Set dfField = Nothing" & vbCrLf
sSub = sSub & "Set drRelation = Nothing" & vbCrLf & vbCrLf
sSub = sSub & "End Sub" & vbCrLf
End If

Add_Subroutines = sSub



End Function



Private Function Information_SQL(ByVal SQLText As String) As String

Dim iCount As Integer
Dim sChar As String
Dim sLine As String
Dim bQuote As Boolean
Dim bEnd As Boolean
Dim sReturn As String
Dim iLineItems As Integer

' Replace quotes
sReturn = ""
sLine = "sSQLText = " & Chr$(34)
iLineItems = 0
bQuote = True
iCount = 1
' v1.3.1 Correct last character omitted
Do While iCount <= Len(SQLText)
sChar = Mid$(SQLText, iCount, 1)
Select Case sChar
Case vbCr
bEnd = True
sChar = " & vbCrLf"
If bQuote Then
sChar = Chr$(34) & sChar
End If
bQuote = False
Case vbLf
bEnd = True
sChar = ""
Case Chr$(34)
sChar = " & Chr$(34)"
If bQuote Then
sChar = Chr$(34) & sChar
End If
bQuote = False
Case Else
If UCase(sChar) Like "[A-Z]" Then
bEnd = False
Else
bEnd = True
End If
If Not bQuote Then
sChar = " & " & Chr$(34) & sChar
End If
bQuote = True
End Select

sLine = sLine & sChar
iLineItems = iLineItems + Len(sChar)
If (Len(sLine) > 90 And bEnd) Or Len(sLine) > 110 Then
'Debug.Print sLine
If bQuote Then
sLine = sLine & Chr$(34)
End If
sReturn = sReturn & sLine & vbCrLf
sLine = "sSQLText = sSQLText & " & Chr$(34)
iLineItems = 0
bQuote = True
End If
iCount = iCount + 1
Loop
If iLineItems > 0 Then
If bQuote Then
sLine = sLine & Chr$(34)
End If
sReturn = sReturn & sLine & vbCrLf
End If

Information_SQL = sReturn

End Function
                          
Public Function Information_Default(ByVal sText As String) As String

Dim iCount As Integer
Dim sChar As String
Dim bQuote As Boolean
Dim bEnd As Boolean
Dim sReturn As String
Dim iLineItems As Integer

If Left$(sText, 1) <> Chr$(34) Then
sText = Chr$(34) & sText
End If
If Right$(sText, 1) <> Chr$(34) Then
sText = sText & Chr$(34)
End If

' Replace quotes
sReturn = ""
bQuote = True
iCount = 1

Do While iCount <= Len(sText)
sChar = Mid$(sText, iCount, 1)
If sChar = Chr$(34) Then
sChar = "Chr$(34)"
If Not bQuote Then
sChar = Chr$(34) & " & " & sChar
End If
bQuote = True
Else
If bQuote Then
sChar = " & " & Chr$(34) & sChar
bQuote = False
End If
End If

sReturn = sReturn & sChar
iCount = iCount + 1
Loop
If Not bQuote Then
sReturn = sReturn & Chr$(34)
End If

Information_Default = sReturn

End Function


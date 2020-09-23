Attribute VB_Name = "DeclarationsEXE"
' ==============================================================
' Module:       Declarations
' Purpose:      Global Type and Variable Declarations
' ==============================================================

Option Explicit

Public Type qTable
    Name As String
    Attributes As String
    Fields As Integer
    Indexes As Integer
End Type

Public Type qField
    Name As String
    Type As DataTypeEnum
    Size As Integer
    Attributes As String
    DefaultValue As String
    Required As Boolean
    Table As Integer
    Index As Boolean
End Type

Public Type qIndex
    Name As String
    FieldIndex As Integer
    Primary As Boolean
    Unique As Boolean
    Required As Boolean
    Sort As Boolean
    Table As Integer
End Type

Public Type qRelate
    Name As String
    Field As Integer
    ForeignField As Integer
    Table As Integer
    ForeignTable As Integer
    Attributes As String
End Type

Public Type qDatabase
    Name As String
    Tables As Integer
    Fields As Integer
    Indexes As Integer
    Relations As Integer
    Queries As Integer
    ItemCount As Boolean
End Type

Public Type qQuery
    Name As String
    SQLText As String
    Fields As Integer
    Type As QueryDefTypeEnum
    TypeText As String
End Type

Public Type qFieldDataType
    Name As String
    Code As String
End Type


Public Type qListView
    Name As String
    Type As qDatabaseObjectEnum
    Reference As Integer
End Type

Public qData As Database

Public qlTable() As qTable
Public qlField() As qField
Public qlIndex() As qIndex
Public qlRelation() As qRelate
Public qlQuery() As qQuery
Public qDB As qDatabase
Public qlNode() As qListView

Public qFType(0 To 23) As qFieldDataType

Public Enum qDatabaseObjectEnum
    qdNone = 0
    qdDatabase = 1
    qdTable = 2
    qdIndex = 3
    qdRelation = 4
    qdField = 5
    qdQueries = 6
    qdQuery = 7
End Enum




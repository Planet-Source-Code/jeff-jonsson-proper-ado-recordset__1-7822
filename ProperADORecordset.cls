VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProperADORecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private SoRidiculous        As ADODB.Recordset
    Private DaNoMatch           As Boolean
    Private DaFields()          As Field

Public Sub AddNew()
    SoRidiculous.AddNew
End Sub

Public Property Get BOF() As Boolean
    BOF = SoRidiculous.BOF
End Property

Public Sub CloseIt()
    SoRidiculous.Close
    Set SoRidiculous = Nothing
End Sub
Public Sub Edit()

End Sub


Public Property Get EOF() As Boolean
    EOF = SoRidiculous.EOF
End Property

Public Property Let Fields(index As Variant, NewValue As Variant)
    DaFields(index) = NewValue
End Property

Public Property Get Fields(index As Variant) As Field
    Set Fields = DaFields(index)
End Property
Sub MoveFirst()
    SoRidiculous.MoveFirst
End Sub

Sub MoveLast()
    SoRidiculous.MoveLast
End Sub


Sub MoveNext()
    SoRidiculous.MoveNext
End Sub


Sub MovePrevious()
    SoRidiculous.MovePrevious
End Sub

Public Sub OpenIt(Optional o1 As Variant, Optional o2 As Variant, Optional o3 As Variant, Optional o4 As Variant, Optional o5 As Variant)
    On Error Resume Next
    Dim Cnt As Integer
    
    Set SoRidiculous = New ADODB.Recordset

    SoRidiculous.Open o1, o2, o3, o4
    
    Erase DaFields
    ReDim DaFields(0 To SoRidiculous.Fields.Count - 1)
    For Cnt = LBound(DaFields) To UBound(DaFields)
        Set DaFields(Cnt) = SoRidiculous.Fields(Cnt)
    Next Cnt
End Sub

Public Property Let ProperADORecordset(index As Variant, NewValue As Variant)
    SoRidiculous(index) = NewValue
End Property


Public Property Get ProperADORecordset(index As Variant) As Variant
Attribute ProperADORecordset.VB_UserMemId = 0
    ProperADORecordset = SoRidiculous(index)
End Property



Public Sub Update()
    SoRidiculous.Update
End Sub

Private Sub Class_Initialize()
    Set SoRidiculous = New ADODB.Recordset
End Sub


Private Sub Class_Terminate()
    On Error Resume Next

    SoRidiculous.Close
    Set SoRidiculous = Nothing
End Sub

Public Function RecordCount()
RecordCount = SoRidiculous.RecordCount
End Function

Public Sub FindFirst(Filter As String)
Dim ST2 As ADODB.Recordset

DaNoMatch = True
'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
    Set ST2 = New ADODB.Recordset
    Set ST2 = SoRidiculous.Clone
    ST2.Filter = Filter
    If (ST2.RecordCount > 0) Then
        ST2.MoveFirst
        SoRidiculous.BookMark = ST2.BookMark
        DaNoMatch = False
    Else
        If (SoRidiculous.RecordCount > 0) Then
            SoRidiculous.MoveLast: SoRidiculous.MoveNext
            DaNoMatch = True
        End If
    End If
'Else
'    SoRidiculous.Find Filter
'    DaNoMatch = SoRidiculous.EOF
'End If
End Sub
Public Sub FindLast(Filter As String)
Dim ST2 As ADODB.Recordset

DaNoMatch = True
'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
    Set ST2 = New ADODB.Recordset
    Set ST2 = SoRidiculous.Clone
    ST2.Filter = Filter
    If (ST2.RecordCount > 0) Then
        ST2.MoveLast
        SoRidiculous.BookMark = ST2.BookMark
        DaNoMatch = False
    Else
        If (SoRidiculous.RecordCount > 0) Then
            SoRidiculous.MoveLast: SoRidiculous.MoveNext
            DaNoMatch = True
        End If
    End If
'Else
'    SoRidiculous.Find Filter
'    SoRidiculous.MoveLast
'    DaNoMatch = SoRidiculous.BOF
'End If
End Sub
Public Sub FindNext(Filter As String)
Dim ST2 As ADODB.Recordset

DaNoMatch = True
'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
    Set ST2 = New ADODB.Recordset
    Set ST2 = SoRidiculous.Clone
    ST2.Filter = Filter
    ST2.Sort = SoRidiculous.Sort
    If (ST2.RecordCount > 0) Then
        ST2.BookMark = SoRidiculous.BookMark
        ST2.MoveNext
        If (Not ST2.EOF) Then
            SoRidiculous.BookMark = ST2.BookMark
            DaNoMatch = False
        Else
            DaNoMatch = True
        End If
    Else
        If (SoRidiculous.RecordCount > 0) Then
            SoRidiculous.MoveLast: SoRidiculous.MoveNext
            DaNoMatch = True
        End If
    End If
'Else
'    SoRidiculous.Find Filter
'    DaNoMatch = SoRidiculous.EOF
'End If
End Sub
Public Sub FindPrevious(Filter As String)
Dim ST2 As ADODB.Recordset

DaNoMatch = True
'If (InStr(Filter, "(") > 0) Or (InStr(Filter, ")") > 0) Or (InStr(UCase(Filter), " AND ") > 0) Or (InStr(UCase(Filter), " OR ") > 0) Then
    Set ST2 = New ADODB.Recordset
    Set ST2 = SoRidiculous.Clone
    ST2.Filter = Filter
    ST2.Sort = SoRidiculous.Sort
    If (ST2.RecordCount > 0) Then
        ST2.BookMark = SoRidiculous.BookMark
        ST2.MovePrevious
        If (Not ST2.BOF) Then
            SoRidiculous.BookMark = ST2.BookMark
            DaNoMatch = False
        Else
            DaNoMatch = True
        End If
    Else
        If (SoRidiculous.RecordCount > 0) Then
            SoRidiculous.MoveFirst: SoRidiculous.MovePrevious
            DaNoMatch = True
        End If
    End If
'Else
'    SoRidiculous.Find Filter
'    DaNoMatch = SoRidiculous.BOF
'End If
End Sub
Public Property Get NoMatch() As Variant

NoMatch = DaNoMatch
End Property
Public Property Get BookMark() As Variant
BookMark = SoRidiculous.BookMark
End Property

Public Property Let BookMark(ByVal vNewValue As Variant)
SoRidiculous.BookMark = vNewValue
End Property

Public Sub Find(Criteria As String, Optional SkipRows As Long = 0, Optional SearchDirection As SearchDirectionEnum = adSearchForward, Optional Start As Variant)
Dim ST2 As ADODB.Recordset
Dim Cnt As Integer

DaNoMatch = True
If (InStr(Criteria, "(") > 0) Or (InStr(Criteria, ")") > 0) Or (InStr(UCase(Criteria), " AND ") > 0) Or (InStr(UCase(Criteria), " OR ") > 0) Then
    Set ST2 = New ADODB.Recordset
    Set ST2 = SoRidiculous.Clone
    ST2.Filter = Criteria
    If (ST2.RecordCount > 0) Then
        If (Not IsMissing(Start)) Then
            ST2.BookMark = Start
            If (SearchDirection = adSearchForward) Then
                For Cnt = 0 To SkipRows
                    If (Not ST2.EOF) Then ST2.MoveNext Else Exit For
                Next Cnt
                If (Not ST2.EOF) Then
                    SoRidiculous.BookMark = ST2.BookMark
                    DaNoMatch = False
                End If
            Else
                For Cnt = 0 To SkipRows
                    If (Not ST2.BOF) Then ST2.MovePrevious Else Exit For
                Next Cnt
                If (Not ST2.BOF) Then
                    SoRidiculous.BookMark = ST2.BookMark
                    DaNoMatch = False
                End If
            End If
        Else
            ST2.MoveFirst
            SoRidiculous.BookMark = ST2.BookMark
            DaNoMatch = False
        End If
    Else
        If (SoRidiculous.RecordCount > 0) Then
            SoRidiculous.MoveLast: SoRidiculous.MoveNext
            DaNoMatch = True
        End If
    End If
Else
    SoRidiculous.Find Criteria
End If
End Sub

Public Function AbsolutePosition()
    AbsolutePosition = SoRidiculous.AbsolutePosition
End Function

Public Sub Delete(Optional AffectRecords As AffectEnum = adAffectCurrent)
    SoRidiculous.Delete AffectRecords
End Sub

Public Function Clone(Optional LockType As LockTypeEnum = adLockUnspecified) As ADODB.Recordset
    Set Clone = SoRidiculous.Clone(LockType)
End Function


Public Function index()
    index = SoRidiculous.index
End Function

Public Property Get Filter() As Variant
    Filter = SoRidiculous.Filter
End Property

Public Property Let Filter(ByVal vNewValue As Variant)
SoRidiculous.Filter = vNewValue
End Property


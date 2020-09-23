VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Proper ADO Recordset Class"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click This"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim db      As ADODB.Connection
    Dim st      As ProperADORecordset

    Set db = New ADODB.Connection
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=c:\my documents\temp.mdb;"
    db.Open
    
    Set st = New ProperADORecordset
    ' So we couldn't duplicate the Open/Close methods 'cause vb won't allow
    ' use of those words as any form of procedure.
    ' So: OpenIt, CloseIt
    ' Hopefully vb7 (true oop?) will remedy this.
    ' ----------------------------------------------------------------------
    st.OpenIt "aTableHere", db, adOpenKeyset
    
    Do While Not mi.EOF
        Debug.Print mi("aField")
        
        mi.MoveNext
    Loop
End Sub


VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5520
      List            =   "Form1.frx":0019
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Report"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim MyDb As New ADODB.Connection
Dim MyRs As New ADODB.Recordset

Dim strPath As String

strPath = App.Path & "\test1.mdb"
 
'Set connection to database ( strpath )

MyDb.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & strPath

MyDb.Open


'Open The The Recordset
MyRs.ActiveConnection = MyDb

'Open  Keyset,  LockOptimistic
MyRs.Open Combo1.Text, MyDb, adOpenKeyset, adLockOptimistic, adCmdTableDirect

With DataReport1.Sections("Section1").Controls 'section1 mean that section you create in datareport
   .Item("text1").DataField = MyRs("Name").Name
   .Item("text2").DataField = MyRs("Address").Name
   .Item("text3").DataField = MyRs("Age").Name
End With

With DataReport1.Sections("Section2").Controls
   .Item("Label2").Caption = "Name"
   .Item("Label3").Caption = "Address"
   .Item("Label4").Caption = "Age"
End With

With DataReport1.Sections("Section4").Controls
   .Item("Label1").Caption = "My Address Book"
End With

Set DataReport1.DataSource = MyRs

'show datareport
DataReport1.Show

End Sub

Private Sub Command2_Click()
Unload Form1
End

End Sub


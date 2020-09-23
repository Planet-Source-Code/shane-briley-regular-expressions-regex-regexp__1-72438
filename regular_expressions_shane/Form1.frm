VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Regex by Shane"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Get String"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "http://www.the--source.org/"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'make sure when using Regex, you add the reference to Microsoft VBScript Regular Expressions 5.5
'Project -> References - > Microsoft VBScript Regular Expressions 5.5
'
'Nothing really to much to this example, its fairly simple once you get the hang of it.
'If you like it or use this example all we ask is that you give some credit to The--Source.
'
'By Shane
'http://www.the--source.org/

Private Sub Command1_Click()
On Error Resume Next
Dim T As String, RE As New RegExp, M As Match
'this is the data that the pattern will be searching through.
T = Text1.Text
'this is your pattern string, i have it set to search for proxies(ips).
'it will only return strings matching say 192.168.1.1:80
RE.Pattern = "\d\d\d?.\d\d\d?.\d\d\d?.\d\d\d?[:]\d+"
RE.Global = True
'this is searching through the whole pattern data and returning only matching strings.
For Each M In RE.Execute(T)
    'displays your returned data matching your pattern.
    MsgBox M.Value
Next
End Sub

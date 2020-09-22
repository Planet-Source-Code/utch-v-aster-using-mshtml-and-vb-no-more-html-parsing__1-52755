VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "HTML Element Parser"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1890
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1658
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3225
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   5689
      _Version        =   393217
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parse HTML Elements"
      Height          =   450
      Left            =   600
      TabIndex        =   0
      Top             =   30
      Width           =   2475
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Static URL As String
    
    If URL = "" Then
        URL = InputBox("Please enter a URL to parse elements:", "Enter URL", "http://www.alphamedia.net")
    Else
        URL = InputBox("Please enter a URL to parse elements:", "Enter URL", URL)
    End If
    
    If StrPtr(URL) = 0 Then GoTo Ending
    
    Command1.Enabled = False
    GetLinks (URL)
    
Ending:
    Command1.Enabled = True

End Sub

Sub GetLinks(URL As String)

   On Error GoTo ErrPoint

   Dim Web As New SHDocVw.InternetExplorer
   Dim Doc As New MSHTML.HTMLDocument
   Dim e As MSHTML.HTMLGenericElement
   Dim a As MSHTML.HTMLAnchorElement
   Dim i As MSHTML.HTMLImg
   Dim t As MSHTML.HTMLTitleElement
   Dim S As MSHTML.HTMLInputElement
      
   Call ResetTreeView
   Web.navigate URL
   
   Do While Web.Busy
    DoEvents
   Loop
   
   Set Doc = Web.document
   
   For Each e In Doc.All
      If e.tagName = "A" Then
         Set a = e
         If a.href <> "" Then Call AddToTreeView(a.href, "A", 2)
      ElseIf e.tagName = "IMG" Then
         Set i = e
         If i.src <> "" Then Call AddToTreeView(i.src, "IMG", 3)
      ElseIf e.tagName = "TITLE" Then
         Set t = e
         If t.Text <> "" Then Call AddToTreeView("<TITLE>: " & t.Text, "Doc", 4)
      ElseIf e.tagName = "INPUT" Then
         Set S = e
         If S.Name <> "" Then Call AddToTreeView("Name (" & S.Name & ")   Size (" & S.Size & ")   Value(" & S.Value & ")", "INPUT", 5)
      End If
   Next

ErrPoint:
   Call CountThem
   Set Web = Nothing

End Sub

Sub CountThem()
    Dim X As Integer
    For X = 1 To 4
        tv.Nodes(X).Text = tv.Nodes(X).Text & " (" & tv.Nodes(X).children & ")"
    Next
End Sub

Private Sub Form_Load()
    Height = 5000
    Width = 5000

    Call ResetTreeView
    
End Sub

Sub AddToTreeView(mText As String, mParent As String, Optional mImage As Integer)
    
    On Error GoTo ErrPoint
    Dim tvNode As Node
    If mImage = 0 Then
        Set tvNode = tv.Nodes.Add(mParent, tvwChild, Right(mText, 20), mText)
    Else
        Set tvNode = tv.Nodes.Add(mParent, tvwChild, Right(mText, 20), mText, mImage)
    End If

ErrPoint:

End Sub

Sub ResetTreeView()
    
    tv.Nodes.Clear
    
    Dim tvNode As Node
    
    Set tvNode = tv.Nodes.Add(, tvwparent, "Doc", "Document Elements", 1)
    Set tvNode = tv.Nodes.Add(, tvwparent, "A", "A's", 1)
    Set tvNode = tv.Nodes.Add(, tvwparent, "IMG", "IMG's", 1)
    Set tvNode = tv.Nodes.Add(, tvwparent, "INPUT", "INPUT's", 1)
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrPoint
    
    List1.Move 0, 0, ScaleWidth, ScaleHeight - (Command1.Height + 150)
    tv.Move 0, 0, ScaleWidth, ScaleHeight - (Command1.Height + 150)
    Command1.Move ScaleWidth / 2 - Command1.Width / 2, ScaleHeight - (Command1.Height + 120)
    
ErrPoint:

End Sub

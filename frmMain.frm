VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Duke Nukem 3D GRP Manager"
   ClientHeight    =   4155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3900
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1153
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F2D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "test"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnufilesaveas 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnufiledash 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuitem 
      Caption         =   "&Item"
      Begin VB.Menu mnuitemadd 
         Caption         =   "&Add Item"
      End
      Begin VB.Menu mnuitemextract 
         Caption         =   "&Extract Item"
      End
      Begin VB.Menu mnuitemdelete 
         Caption         =   "&Delete Item"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dukegrp As New clsDN3DGRP
Private hasName As Boolean

'Resize the Listview to the same size of the form
Private Sub Form_Resize()
    If Me.WindowState <> 1 Then
       ListView1.Width = Me.ScaleWidth: ListView1.Height = Me.ScaleHeight - StatusBar1.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload_Program
End Sub

'If the person right clicked, pop up the mnuitem button
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuitem
End Sub

'This function allows users to drag and drop their files
'Onto the form for easy archiving
Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim file As Variant, i As Integer
    If Not (Data.GetFormat(1)) Then
        For Each file In Data.Files
            If Dir(file) <> "" And Not vbDirectory Then
            AddItem dukegrp.addFile(CStr(file))
            End If
        Next
    End If
End Sub

'Add the item to the listview, with the right icon
Private Sub AddItem(name As String)
            Select Case LCase(Mid(name, InStrRev(name, ".") + 1, 3))
            Case "voc", "mid"
                ListView1.ListItems.Add , , name, 3, 3
            Case "anm", "dmo"
                ListView1.ListItems.Add , , name, 2, 2
            Case "art"
                ListView1.ListItems.Add , , name, 4, 4
            Case "map"
                ListView1.ListItems.Add , , name, 5, 5
            Case Else
                ListView1.ListItems.Add , , name, 1, 1
            End Select
End Sub
'When the person exits, clean up before closing
Private Sub mnufileexit_Click()
    Unload_Program
End Sub
'Start a new project
Private Sub mnuFileNew_Click()
    If dukegrp.Changed = True Then
        If MsgBox("File has changed." & vbCrLf & "Save now?", vbYesNo + vbQuestion, "Save Changes?") = vbYes Then mnufilesave_Click
    End If
    hasName = False
    Set dukegrp = Nothing
    Set dukegrp = New clsDN3DGRP
End Sub
'Open a grp file
Private Sub mnuFileOpen_Click()
    Dim fd As New clsFileDialog
    
    If dukegrp.Changed = True Then
        If MsgBox("File has changed." & vbCrLf & "Save now?", vbYesNo + vbQuestion, "Save Changes?") = vbYes Then mnufilesave_Click
    End If
    
    fd.Filter = "Duke Nukem 3d GRP Files|*.grp|All Files|*.*"
    fd.InitDir = App.Path
    fd.ShowOpen
    If Len(fd.filename) > 0 Then
        dukegrp.LoadFile (fd.filename)
        updateListview
        hasName = True
    End If
    Set fd = Nothing
    
End Sub

'Clean up the program nicely
Private Sub Unload_Program()
    If dukegrp.Changed = True Then
        If MsgBox("File has changed." & vbCrLf & "Save now?", vbYesNo + vbQuestion, "Save Changes?") = vbYes Then mnufilesave_Click
    End If
    Set dukegrp = Nothing
    Unload Me
    End
End Sub

'Save the GRP file
Private Sub mnufilesave_Click()
    If Not (hasName) Then mnufilesaveas_Click: Exit Sub
    dukegrp.saveFile
End Sub

'Save the GRP file with a new filename
Private Sub mnufilesaveas_Click()
    Dim fd As New clsFileDialog
    fd.Filter = "Duke Nukem 3d GRP Files|*.grp|All Files|*.*"
    fd.ShowSave
    If Len(fd.filename) > 0 Then dukegrp.saveFile fd.filename: hasName = True
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

'Add a file to the GRP file
Private Sub mnuitemadd_Click()
    Dim fd As New clsFileDialog
    fd.Filter = "All Files|*.*"
    fd.ShowOpen
    If Len(fd.filename) > 0 Then AddItem dukegrp.addFile(fd.filename)
    Set fd = Nothing
End Sub

'Delete a file from the GRP file
Private Sub mnuitemdelete_Click()
    dukegrp.deleteFile ListView1.SelectedItem.index
    ListView1.ListItems.Remove (ListView1.SelectedItem.index)
End Sub

'Extract a file from the GRP file
Private Sub mnuitemextract_Click()
    Dim fd As New clsFileDialog
    fd.Filter = "All Files|*.*"
    fd.filename = ListView1.SelectedItem
    fd.ShowSave
    If Len(fd.filename) > 0 Then
        If ListView1.SelectedItem.index <= dukegrp.NumberOfFiles Then dukegrp.extractFile ListView1.SelectedItem.index, fd.filename
    End If
    Set fd = Nothing
End Sub

'Update the listview
Private Sub updateListview()
    Dim i As Integer
    ListView1.ListItems.Clear
    For i = 1 To dukegrp.NumberOfFiles
        AddItem dukegrp.filename(i)
    Next i
End Sub

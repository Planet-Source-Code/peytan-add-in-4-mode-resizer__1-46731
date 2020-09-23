VERSION 5.00
Begin VB.Form frmAddIn 
   Caption         =   "Resizer Add In"
   ClientHeight    =   6330
   ClientLeft      =   2190
   ClientTop       =   1950
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.Frame fraFormProperties 
      Caption         =   "Form Properties"
      Height          =   855
      Left            =   2400
      TabIndex        =   30
      Tag             =   "0"
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdAddCode 
         Caption         =   "Add Necessary code to form"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.ComboBox cmbContainers 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Tag             =   "0000"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Resize and Move Properties"
      Height          =   5295
      Left            =   2400
      TabIndex        =   2
      Tag             =   "0011"
      Top             =   960
      Width           =   4695
      Begin VB.CommandButton cmdResizeProp 
         Caption         =   "Only Resize(prop)"
         Height          =   375
         Left            =   1680
         TabIndex        =   33
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdMoveProp 
         Caption         =   "Only Move(prop)"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton cmdOnlyResize 
         Caption         =   "Only Resize(static)"
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Tag             =   "1111"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton cmdOnlyMove 
         Caption         =   "Only Move(static)"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Tag             =   "1111"
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Tag             =   "0000"
         Top             =   600
         Width           =   3615
         Begin VB.OptionButton optTop 
            Caption         =   "Limited Static"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optTop 
            Caption         =   "Not"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optTop 
            Caption         =   "Proportional"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optTop 
            Caption         =   "Static"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   23
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame fraLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Tag             =   "0000"
         Top             =   1560
         Width           =   3495
         Begin VB.OptionButton optLeft 
            Caption         =   "Limited Static"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optLeft 
            Caption         =   "Static"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   20
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optLeft 
            Caption         =   "Proportional"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optLeft 
            Caption         =   "Not"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame fraWidth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Tag             =   "0000"
         Top             =   2520
         Width           =   3615
         Begin VB.OptionButton optWidth 
            Caption         =   "Limited Static"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optWidth 
            Caption         =   "Static"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   15
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optWidth 
            Caption         =   "Proportional"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optWidth 
            Caption         =   "Not"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame fraHeight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Tag             =   "0000"
         Top             =   3480
         Width           =   3615
         Begin VB.OptionButton optHeight 
            Caption         =   "Limited Static"
            Height          =   255
            Index           =   3
            Left            =   1800
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optHeight 
            Caption         =   "Static"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   10
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton optHeight 
            Caption         =   "Proportional"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optHeight 
            Caption         =   "Not"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Label lblHeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Height"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Tag             =   "0000"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblWidth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Width"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Tag             =   "0000"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Left"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Tag             =   "0000"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblTop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Top"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Tag             =   "0000"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ListBox lstControls 
      Height          =   5520
      Left            =   0
      TabIndex        =   1
      Tag             =   "0001"
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox cmbForms 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "0000"
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Resizer As New ControlResizer
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub GetAllForms()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
cmbForms.Clear
For Each comp In VBInstance.ActiveVBProject.VBComponents
    If (comp.Type = vbext_ct_MSForm Or _
        comp.Type = vbext_ct_UserControl _
          Or comp.Type = vbext_ct_VBForm Or _
         comp.Type = vbext_ct_VBMDIForm) Then
        cmbForms.AddItem comp.Name
    End If
Next
If cmbForms.ListCount <> 0 Then
cmbForms.ListIndex = 0
End If
End Sub
Private Sub GetContainers()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
If cmbForms.ListIndex = -1 Then
Exit Sub
End If
cmbContainers.Clear
cmbContainers.AddItem cmbForms.Text
For Each comp In VBInstance.ActiveVBProject.VBComponents
    If comp.Name = cmbForms.Text Then
        Set vbf = comp.Designer
        For Each ctrl In vbf.VBControls
            If ctrl.ClassName = "Frame" Or ctrl.ClassName = "PictureBox" Then
                cmbContainers.AddItem ctrl.Properties("Name")
            End If
        Next
    End If
Next
cmbContainers.ListIndex = 0
End Sub
Private Sub GetFormControls()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim ctrl2 As VBControl
Dim vbf As VBForm
If cmbForms.ListIndex = -1 Then
Exit Sub
End If
lstControls.Clear
lstControls.AddItem cmbContainers.Text
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbContainers.Text Then
            Set vbf = comp.Designer
            For Each ctrl In vbf.ContainedVBControls
                    lstControls.AddItem ctrl.Properties("Name")
            Next
        End If
    Next
End Sub
Private Sub GetControlControls()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim ctrl2 As VBControl
Dim ctrl3 As VBControl
Dim vbf As VBForm
If cmbForms.ListIndex = -1 Then
Exit Sub
End If
lstControls.Clear
lstControls.AddItem cmbContainers.Text
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbForms.Text Then
            Set vbf = comp.Designer
            For Each ctrl2 In vbf.VBControls
                If ctrl2.Properties("Name") = cmbContainers Then
                    For Each ctrl3 In ctrl2.ContainedVBControls
                        lstControls.AddItem ctrl3.Properties("Name")
                    Next
                    Exit Sub
                End If
            Next
        End If
    Next
End Sub


Private Sub cmbContainers_Click()
If cmbContainers.ListIndex = 0 Then
GetFormControls
Else:
GetControlControls
End If
End Sub

Private Sub cmbForms_Click()
VBInstance.ActiveVBProject.VBComponents(cmbForms.Text).Activate
GetContainers
GetFormControls
End Sub

Private Sub cmdApply_Click()
Call ApplyResizeTag
End Sub

Private Sub cmdMoveProp_Click()
optTop(1).Value = True
optLeft(1).Value = True
optWidth(0).Value = True
optHeight(0).Value = True
cmdApply_Click
End Sub

Private Sub cmdOnlyMove_Click()
optTop(2).Value = True
optLeft(2).Value = True
optWidth(0).Value = True
optHeight(0).Value = True
cmdApply_Click
End Sub

Private Sub cmdOnlyResize_Click()
optTop(0).Value = True
optLeft(0).Value = True
optWidth(2).Value = True
optHeight(2).Value = True
cmdApply_Click
End Sub

Private Sub cmdAddCode_Click()
AddCode
End Sub

Private Sub cmdResizeProp_Click()
optTop(0).Value = True
optLeft(0).Value = True
optWidth(1).Value = True
optHeight(1).Value = True
cmdApply_Click
End Sub

Private Sub Form_Load()
Resizer.InitResizer Me, Me.Width, Me.Height
AddClass
GetAllForms
GetFormControls
GetContainers
End Sub
Private Function ControlTag() As String
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
If cmbForms.ListIndex = -1 Then
Exit Function
End If
If cmbContainers.ListIndex = 0 Then
    If lstControls.ListIndex = 0 Then
        For Each comp In VBInstance.ActiveVBProject.VBComponents
            If comp.Name = cmbForms.Text Then
            ControlTag = comp.Properties("Tag")
            Exit Function
            End If
        Next
    Else:
    GoTo NormalControl
    End If
Else:
NormalControl:
For Each comp In VBInstance.ActiveVBProject.VBComponents
    If comp.Name = cmbForms.Text Then
        Set vbf = comp.Designer
        For Each ctrl In vbf.VBControls
            If ctrl.Properties("Name") = lstControls.Text Then
                ControlTag = ctrl.Properties("Tag")
                Exit Function
            End If
        Next
    End If
Next
End If
End Function

Private Sub lstControls_Click()
If lstControls.ListIndex = 0 And cmbContainers.ListIndex = 0 Then
fraFormProperties.Enabled = True
Else
fraFormProperties.Enabled = False
End If
ShowCurrentState (ControlTag)
End Sub
Private Sub ShowCurrentState(ResizeString As String)
Dim i As Integer
If ResizeString = "" Or Len(ResizeString) <> 4 Then
For i = 0 To 3
optTop(i).Value = False
Next i
For i = 0 To 3
optLeft(i).Value = False
Next i
For i = 0 To 3
optWidth(i).Value = False
Next i
For i = 0 To 3
optHeight(i).Value = False
Next i
Exit Sub
End If
optTop(CInt(Mid(ResizeString, 2, 1))).Value = True
optLeft(CInt(Mid(ResizeString, 1, 1))).Value = True
optWidth(CInt(Mid(ResizeString, 3, 1))).Value = True
optHeight(CInt(Mid(ResizeString, 4, 1))).Value = True
End Sub
Private Sub ApplyResizeTag()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim vbf As VBForm
If cmbForms.ListIndex = -1 Then
Exit Sub
End If
        If cmbContainers.ListIndex = 0 Then
            If lstControls.ListIndex = 0 Then
                For Each comp In VBInstance.ActiveVBProject.VBComponents
                    If comp.Name = cmbForms.Text Then
                        Set vbf = comp.Designer
                        comp.Properties("Tag") = ResizeTag
                        Exit Sub
                    End If
                Next
                Exit Sub
            End If
        End If
        For Each comp In VBInstance.ActiveVBProject.VBComponents
            If comp.Name = cmbForms.Text Then
                Set vbf = comp.Designer
                For Each ctrl In vbf.VBControls
                    If ctrl.Properties("Name") = lstControls.Text Then
                        ctrl.Properties("Tag") = ResizeTag
                        Exit Sub
                    End If
                Next
            End If
        Next
End Sub
Private Function ResizeTag() As String
Dim i As Integer
Dim ResizeString As String
ResizeString = ""
For i = 0 To 3
    If optLeft(i).Value = True Then
        ResizeString = ResizeString & CStr(i)
    End If
Next i
If Len(ResizeString) = 0 Then
ResizeString = ResizeString & "0"
End If
For i = 0 To 3
    If optTop(i).Value = True Then
        ResizeString = ResizeString & CStr(i)
    End If
Next i
If Len(ResizeString) = 1 Then
ResizeString = ResizeString & "0"
End If
For i = 0 To 3
    If optWidth(i).Value = True Then
        ResizeString = ResizeString & CStr(i)
    End If
Next i
If Len(ResizeString) = 2 Then
ResizeString = ResizeString & "0"
End If
For i = 0 To 3
    If optHeight(i).Value = True Then
        ResizeString = ResizeString & CStr(i)
    End If
Next i
If Len(ResizeString) = 3 Then
ResizeString = ResizeString & "0"
End If
ResizeTag = ResizeString
End Function

Private Sub optHeight_Click(Index As Integer)
cmdApply_Click
End Sub

Private Sub optLeft_Click(Index As Integer)
cmdApply_Click
End Sub

Private Sub optTop_Click(Index As Integer)
cmdApply_Click
End Sub

Private Sub optWidth_Click(Index As Integer)
cmdApply_Click
End Sub
Private Sub AddCode()
Dim comp As VBComponent
Dim ctrl As VBControl
Dim ctrl2 As VBControl
Dim vbf As VBForm
If cmbForms.ListIndex = -1 Then
Exit Sub
End If
lstControls.Clear
lstControls.AddItem cmbContainers.Text
    For Each comp In VBInstance.ActiveVBProject.VBComponents
        If comp.Name = cmbContainers.Text Then
            Set vbf = comp.Designer
            If comp.CodeModule.Find("Dim Resizer As New ControlResizer", 1, 1, comp.CodeModule.CountOfDeclarationLines, -1, True, True) = False Then
                comp.CodeModule.InsertLines 1, "Private Resizer As New ControlResizer"
            End If
            If comp.CodeModule.Find("Sub Form_Load", 1, 1, -1, -1) = False Then
                comp.CodeModule.CreateEventProc "Load", "Form"
            End If
            If comp.CodeModule.Find("Resizer.InitResizer Me,Me.Width,Me.Height", comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc), 1, comp.CodeModule.ProcCountLines("Form_Load", vbext_pk_Proc) + comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc), -1) = False Then
                comp.CodeModule.InsertLines comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc) + 1, "Resizer.InitResizer Me,Me.Width,Me.Height"
                If comp.Properties("MDIChild") = True Then
                    comp.CodeModule.InsertLines comp.CodeModule.ProcBodyLine("Form_Load", vbext_pk_Proc) + 1, "Me.Height=" & comp.Properties("Height") & vbCrLf & "Me.Width=" & comp.Properties("Width")
                End If
            End If
            If comp.CodeModule.Find("Sub Form_Resize", 1, 1, -1, -1) = False Then
                comp.CodeModule.CreateEventProc "Resize", "Form"
            End If
            If comp.CodeModule.Find("Resizer.InitResizer Me,Me.Width,Me.Height", comp.CodeModule.ProcBodyLine("Form_Resize", vbext_pk_Proc), 1, comp.CodeModule.ProcCountLines("Form_Resize", vbext_pk_Proc) + comp.CodeModule.ProcBodyLine("Form_Resize", vbext_pk_Proc), -1) = False Then
                comp.CodeModule.InsertLines comp.CodeModule.ProcBodyLine("Form_Resize", vbext_pk_Proc) + 1, "Resizer.FormResized Me"
            End If
        End If
    Next
cmbContainers_Click
Me.Show
End Sub
Private Sub AddClass()
Dim comp As VBComponent
Dim AlreadyExists As Boolean
FileCopy App.Path & "\ControlResizer.cls", ParsePath(VBInstance.ActiveVBProject.FileName, vbDirectory) & "\ControlResizer.cls"
For Each comp In VBInstance.ActiveVBProject.VBComponents
If comp.Name = "ControlResizer" Then
AlreadyExists = True
End If
Next
If AlreadyExists = False Then
VBInstance.ActiveVBProject.VBComponents.AddFile (App.Path & "\ControlResizer.cls")
End If
End Sub
Public Function ParsePath(strFullPathName As String, ReturnType As Integer, Optional StripLastBackslash) As String
    Dim strTemp As String, intX As Integer, strPathName As String, strFileName As String

    If IsMissing(StripLastBackslash) Then StripLastBackslash = False
    If Len(strFullPathName) > 0 Then
        strTemp = ""
        intX = Len(strFullPathName)
        Do While strTemp <> "\"
            strTemp = Mid(strFullPathName, intX, 1)
            If strTemp = "\" Then
                strPathName = Left(strFullPathName, intX + StripLastBackslash)
                strFileName = Right(strFullPathName, Len(strFullPathName) - intX)
            End If
            intX = intX - 1
        Loop

        Select Case ReturnType
        Case vbDirectory
            ParsePath = strPathName
        Case vbNormal
            ParsePath = strFileName
        Case Else
            ParsePath = strFullPathName
        End Select
    Else
        ParsePath = ""
    End If

End Function


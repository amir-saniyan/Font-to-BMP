VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mainForm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font to BMP"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame saveFrame 
      Caption         =   "Save :"
      Height          =   2175
      Left            =   120
      TabIndex        =   29
      Top             =   4920
      Width           =   4215
      Begin VB.CommandButton aboutCommand 
         Caption         =   "&About..."
         Height          =   495
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "About this program."
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox endAsciiRangeList 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "End ASCII range for generate picture(s)."
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox startAsciiRangeList 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Start ASCII range for generate picture(s)."
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton generateCommand 
         Caption         =   "&Generate"
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         ToolTipText     =   "Generate picture(s)."
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton savePathCommand 
         Caption         =   "..."
         Height          =   285
         Left            =   3840
         TabIndex        =   16
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox savePathText 
         Height          =   285
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "File name."
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label endRangeLabel 
         AutoSize        =   -1  'True
         Caption         =   "End ASCII Range :"
         Height          =   195
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label startRangrLabel 
         AutoSize        =   -1  'True
         Caption         =   "Start ASCII Range :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label savePathLabel 
         AutoSize        =   -1  'True
         Caption         =   "File Name :"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   795
      End
   End
   Begin VB.Frame fontPropertiesFrame 
      Caption         =   "Font Properties :"
      Height          =   4695
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   4215
      Begin VB.TextBox previewText 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   11
         Text            =   "A"
         ToolTipText     =   "Enter a letter."
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton previewCommand 
         Caption         =   "&Preview"
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         ToolTipText     =   "Preview font picture."
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Frame colorFrame 
         Caption         =   "Color :"
         Height          =   1095
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   3975
         Begin VB.PictureBox backGroundColorPicture 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            MouseIcon       =   "main.frx":08CA
            MousePointer    =   99  'Custom
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   9
            ToolTipText     =   "Click here for change back ground color."
            Top             =   480
            Width           =   495
         End
         Begin VB.PictureBox fontColorPicture 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2160
            MouseIcon       =   "main.frx":1194
            MousePointer    =   99  'Custom
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   8
            ToolTipText     =   "Click here for change font color."
            Top             =   480
            Width           =   495
         End
         Begin VB.Label fontColorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Font Color :"
            Height          =   195
            Left            =   2160
            TabIndex        =   28
            Top             =   240
            Width           =   810
         End
         Begin VB.Label backGroundColorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Back Ground Color :"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.Frame backGorundSizeFrame 
         Caption         =   "Back Ground Size :"
         Height          =   1695
         Left            =   2160
         TabIndex        =   23
         Top             =   960
         Width           =   1935
         Begin VB.TextBox backGroundHeightText 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "50"
            ToolTipText     =   "Enter a number for back ground height in pixel."
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox backGroundWidthText 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "50"
            ToolTipText     =   "Enter a number for back ground width in pixel."
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label backGroundHeightLabel 
            AutoSize        =   -1  'True
            Caption         =   "Back Ground Height :"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label backGroundWidthLabel 
            AutoSize        =   -1  'True
            Caption         =   "Back Ground Width :"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1500
         End
      End
      Begin VB.Frame effectsFrame 
         Caption         =   "Effects :"
         Height          =   1695
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1935
         Begin VB.CheckBox strikeoutCheck 
            Caption         =   "strikeout"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   -1  'True
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Strikeout."
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox underLineCheck 
            Caption         =   "Under Line"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Under Line."
            Top             =   960
            Width           =   1095
         End
         Begin VB.CheckBox italicCheck 
            Caption         =   "Italic"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Italic."
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox boldCheck 
            Caption         =   "Bold"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Bold."
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox fontSizeList 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Font size in pixel."
         Top             =   4200
         Width           =   1455
      End
      Begin VB.ComboBox fontNameList 
         Height          =   315
         ItemData        =   "main.frx":1A5E
         Left            =   120
         List            =   "main.frx":1A60
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select a font from list."
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label fontSizeLabel 
         AutoSize        =   -1  'True
         Caption         =   "Font Size (Pixel) :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label fontNameLabel 
         AutoSize        =   -1  'True
         Caption         =   "Font Name :"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.PictureBox fontPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6120
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6720
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' In the name of God.
' This program generates font pictures.
' Programmer : Amir Saniyan.
' Programmer E-mail : AmirSn@mail.com
' Date : Sunday and Monday, 23, 24 March 2003.
' Date : Sunday and Monday, 3, 4 Farvardin 1382.
' All rights reserved.

Private Sub aboutCommand_Click()
    ' Show about form.
    aboutForm.Show 1
End Sub

Private Sub backGroundColorPicture_Click()
    ' If cancel button pressed.
    On Error GoTo PASS
    CommonDialog.ShowColor
    backGroundColorPicture.BackColor = CommonDialog.Color
PASS:
End Sub

Private Sub exitMenu_Click()
    ' Exit.
    End
End Sub

Private Sub backGroundHeightText_Change()
    ' Numbers only.
    If Not IsNumeric(backGroundHeightText.Text) Then
        MsgBox "Please Enter a Number.", vbCritical, "Error"
        backGroundHeightText.Text = 50
        backGroundHeightText.SetFocus
    End If
End Sub

Private Sub backGroundWidthText_Change()
    ' Numbers only.
    If Not IsNumeric(backGroundWidthText.Text) Then
        MsgBox "Please Enter a Number.", vbCritical, "Error"
        backGroundWidthText.Text = 50
        backGroundWidthText.SetFocus
    End If
End Sub

Private Sub fontColorPicture_Click()
    ' If cancel button pressed.
    On Error GoTo PASS
    CommonDialog.ShowColor
    fontColorPicture.BackColor = CommonDialog.Color
PASS:
End Sub

Private Sub fontNameList_Click()
    ' Update preview text.
    previewText.FontName = fontNameList.Text
End Sub

Private Sub fontSizeList_Change()
    ' Numbers only.
    If Not IsNumeric(fontSizeList.Text) Then
        MsgBox "Please Enter a Number.", vbCritical, "Error"
        fontSizeList.ListIndex = 24
        fontSizeList.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim counter As Integer
    ' Adjusting font name menu.
    For counter = 1 To Screen.FontCount
        fontNameList.AddItem Screen.Fonts(counter - 1)
    Next
    fontNameList.ListIndex = 0
    
    ' Adjusting font size menu and default = 32 pixel.
    For counter = 8 To 72
       fontSizeList.AddItem counter
    Next
    fontSizeList.ListIndex = 24
    
    ' Adjusting start and end ASCII range menus and defaults = 32 ... 255.
    For counter = 32 To 255
        startAsciiRangeList.AddItem counter
        endAsciiRangeList.AddItem counter
    Next
    startAsciiRangeList.ListIndex = 0
    endAsciiRangeList.ListIndex = 223
    
    ' Defult save path.
    CommonDialog.Flags = cdlCCFullOpen Or cdlOFNHideReadOnly
    If Right(App.Path, 1) = "\" Then
        savePathText.Text = App.Path & fontNameList.Text & ".bmp"
    Else
        savePathText.Text = App.Path & "\" & fontNameList.Text & ".bmp"
    End If
End Sub

Private Sub generateCommand_Click()
    On Error GoTo PASS
    ' True range?
    If Int(startAsciiRangeList.Text) > Int(endAsciiRangeList.Text) Then
        MsgBox "Start ASCII Range is bigger than End ASCII Range.", vbCritical, "Error"
        startAsciiRangeList.SetFocus
        Exit Sub
    End If
    
    ' True path?
    If savePathText.Text = "" Then
        MsgBox "Please Enter File Name.", vbCritical, "Error"
        savePathText.SetFocus
        Exit Sub
    End If
    
    ' Ready hidden picturebox.
    With fontPicture
        .Width = backGroundWidthText.Text + 4
        .Height = backGroundHeightText.Text + 4
        .BackColor = backGroundColorPicture.BackColor
        .FontName = fontNameList.Text
        .FontBold = boldCheck.Value
        .FontItalic = italicCheck.Value
        .FontUnderline = underLineCheck.Value
        .FontStrikethru = strikeoutCheck.Value
        .FontSize = fontSizeList.Text
        .ForeColor = fontColorPicture.BackColor
    End With
    
    ' Generate "fontName_XXX.bmp" file names where XXX are number.
    Dim withoutExtension As String
    withoutExtension = ""
    If UCase(Right(savePathText.Text, 4)) = UCase(".bmp") Then
        withoutExtension = Left(savePathText.Text, Len(savePathText.Text) - 4)
    Else
        withoutExtension = savePathText.Text
    End If
    
    ' Generate pictures.
    Dim counter As Integer
    For counter = startAsciiRangeList.Text To endAsciiRangeList
        fontPicture.Cls
        fontPicture.Print Chr(counter)
        SavePicture fontPicture.Image, withoutExtension & "_" & counter & ".bmp"
    Next
    Exit Sub
PASS:
    MsgBox Err.Description, vbCritical, "Error"
    End
End Sub

Private Sub previewCommand_Click()
    ' Ready preview picturebox.
    With previewForm.previewPicture
        .Width = backGroundWidthText.Text + 4
        .Height = backGroundHeightText.Text + 4
        .BackColor = backGroundColorPicture.BackColor
        .FontName = fontNameList.Text
        .FontBold = boldCheck.Value
        .FontItalic = italicCheck.Value
        .FontUnderline = underLineCheck.Value
        .FontStrikethru = strikeoutCheck.Value
        .FontSize = fontSizeList.Text
        .ForeColor = fontColorPicture.BackColor
    End With
    
    ' Ready preview form.
    With previewForm
        .ScaleMode = vbTwips
        .previewPicture.ScaleMode = vbTwips
        .previewPicture.Top = 0
        .Height = .previewPicture.Height + 390
        .previewPicture.Left = 0
        .Width = .previewPicture.Width + 90
        .ScaleMode = vbPixels
        .previewPicture.ScaleMode = vbPixels
    End With
    
    ' Write preview.
    previewForm.previewPicture.Cls
    previewForm.previewPicture.Print previewText.Text
    previewForm.Show 1
End Sub

Private Sub previewText_KeyPress(KeyAscii As Integer)
    ' Clear last text.
    previewText.Text = ""
End Sub

Private Sub savePathCommand_Click()
    ' To show save dialog.
    On Error GoTo PASS
    CommonDialog.FileName = fontNameList.Text
    CommonDialog.Filter = "Picture (*.bmp)|*.bmp"
    CommonDialog.ShowSave
    savePathText.Text = CommonDialog.FileName
PASS:
End Sub

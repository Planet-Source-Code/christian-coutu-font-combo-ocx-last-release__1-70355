VERSION 5.00
Object = "*\AFont_Combo.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin Font_Combo.FontCombo FontCombo1 
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   210
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   556
      PreviewText     =   "Sample Text"
      ComboFontSize   =   12
      ComboWidth      =   300
      RecentMax       =   5
      ButtonOverColor =   0
      UseMouseWheel   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check7 
      Alignment       =   1  'Right Justify
      Caption         =   "XP Style"
      Height          =   285
      Left            =   3630
      TabIndex        =   11
      Top             =   1170
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtFont 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   300
      TabIndex        =   10
      Text            =   "Sample Text"
      Top             =   3480
      Width           =   6135
   End
   Begin VB.CheckBox Check6 
      Alignment       =   1  'Right Justify
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3630
      TabIndex        =   9
      Top             =   2970
      Width           =   2175
   End
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3630
      TabIndex        =   8
      Top             =   2670
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load Recents"
      Height          =   405
      Left            =   4980
      TabIndex        =   7
      Top             =   630
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Recents"
      Height          =   405
      Left            =   3660
      TabIndex        =   6
      Top             =   630
      Width           =   1245
   End
   Begin VB.CheckBox Check4 
      Alignment       =   1  'Right Justify
      Caption         =   "Sorted"
      Height          =   285
      Left            =   3630
      TabIndex        =   5
      Top             =   2370
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Font Name"
      Height          =   285
      Left            =   3630
      TabIndex        =   4
      Top             =   2070
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Preview"
      Height          =   285
      Left            =   3630
      TabIndex        =   3
      Top             =   1770
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Font in Combo"
      Height          =   285
      Left            =   3630
      TabIndex        =   2
      Top             =   1470
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4590
      TabIndex        =   1
      Text            =   "Arial"
      Top             =   180
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set font"
      Height          =   435
      Left            =   3630
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click()
FontCombo1.ShowFontInCombo = Check1.Value = 1
End Sub

Private Sub Check2_Click()
FontCombo1.ShowPreview = Check2.Value = 1
Check3.Enabled = Check2.Value = 1
End Sub


Private Sub Check3_Click()
FontCombo1.ShowFontName = Check3.Value = 1
End Sub


Private Sub Check4_Click()
FontCombo1.Sorted = Check4.Value = 1
End Sub


Private Sub Check5_Click()
FontCombo1.ComboFontBold = Check5.Value = 1
End Sub

Private Sub Check6_Click()
FontCombo1.ComboFontItalic = Check6.Value = 1
End Sub

Private Sub Check7_Click()
FontCombo1.XPStyle = Check7.Value = 1
End Sub

Private Sub Command1_Click()
FontCombo1.SelectedFont = Text1.Text
End Sub

Private Sub Command2_Click()
FontCombo1.SaveRecentFonts HKEY_CURRENT_USER, "Software", "FontCombo", "Settings"
End Sub

Private Sub Command3_Click()
FontCombo1.LoadRecentFonts HKEY_CURRENT_USER, "Software", "FontCombo", "Settings"
End Sub

Private Sub FontCombo1_FontNotFound(FontName As String)
MsgBox "Cant find this font: " & FontName
End Sub

Private Sub FontCombo1_SelectedFontChanged(NewFontName As String)
txtFont.FontName = NewFontName
FontCombo1.ClearUsedList
FontCombo1.AddToUsedList NewFontName
End Sub

Private Sub Form_Load()
FontCombo1.SelectedFont = txtFont.FontName
End Sub

Private Sub txtFont_Change()
FontCombo1.PreviewText = txtFont.Text
End Sub



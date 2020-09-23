VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RotFonts 
   Caption         =   "Circular Text"
   ClientHeight    =   7815
   ClientLeft      =   2715
   ClientTop       =   2235
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Rotfonts.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7815
   ScaleWidth      =   8685
   Begin MSComDlg.CommonDialog CD 
      Left            =   15
      Top             =   2205
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   5535
      Left            =   450
      ScaleHeight     =   5475
      ScaleWidth      =   7740
      TabIndex        =   12
      Top             =   2250
      Width           =   7800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   8535
      Begin VB.CommandButton Command2 
         Caption         =   "R 360° TEST ==>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7635
         TabIndex        =   22
         Top             =   1335
         Width           =   795
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H00C00000&
         Height          =   330
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1725
         Width           =   630
      End
      Begin VB.CommandButton txtSimulation 
         Caption         =   "L 360° TEST <=="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6795
         TabIndex        =   19
         Top             =   1335
         Width           =   810
      End
      Begin VB.TextBox txtFont 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2205
         TabIndex        =   17
         Text            =   "Arial"
         Top             =   1410
         Width           =   1470
      End
      Begin VB.TextBox TxtCY 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   15
         Text            =   "360"
         Top             =   1740
         Width           =   765
      End
      Begin VB.TextBox TxtCX 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Text            =   "360"
         Top             =   1425
         Width           =   765
      End
      Begin VB.TextBox txtSettore 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   10
         Text            =   "360"
         Top             =   1095
         Width           =   765
      End
      Begin VB.TextBox txtRaggio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Text            =   "2000"
         Top             =   750
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Draw text with this parameters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6810
         TabIndex        =   4
         Top             =   645
         Width           =   1620
      End
      Begin VB.TextBox txtDegree 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Text            =   "90"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtSize 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Text            =   "20"
         Top             =   1050
         Width           =   645
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Text            =   "TEXT EXAMPLE - printed circular - "
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   21
         Top             =   1770
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Font Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   825
         TabIndex        =   18
         Top             =   1425
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Center Y coord"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3855
         TabIndex        =   16
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Center X coord"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3855
         TabIndex        =   14
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Angular Sector"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3855
         TabIndex        =   11
         Top             =   1110
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Radius"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3855
         TabIndex        =   9
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   870
         TabIndex        =   7
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Starting Angle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   6
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Font size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2115
         TabIndex        =   5
         Top             =   1065
         Width           =   780
      End
   End
End
Attribute VB_Name = "RotFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Nome del Progetto: Project1
' Nome del Modulo: RotFonts
' Scopo: Circular Text
' Data: 12/05/2001
' Ora: 14.29
' Revisione:
' Autore: NDV Software
'------------------------------------------------------------
' ****************************************************************************************************
' Copyright © 1990 - 2001 NDV Software,
' Tutti i diritti riservati, ndv@interfree.it
' ****************************************************************************************************

Private Sub cmdColor_Click()
  CD.ShowColor
  Me.cmdColor.BackColor = CD.Color
  Call Command1_Click
End Sub
Private Sub Command1_Click()
  Me.Picture1.Cls
  Me.Picture1.FontSize = Me.txtSize
  Me.Picture1.FontName = Me.txtFont
  Me.Picture1.ForeColor = Me.cmdColor.BackColor
  Call drawCircularText(Me.Picture1, Me.Text1, Me.txtDegree, Me.txtRaggio, Me.TxtCX, Me.TxtCY, Me.txtSettore)
End Sub

Private Sub Command2_Click()
Dim I As Integer
  Me.Picture1.FontSize = 25
  Me.Picture1.FontName = "Arial black"
  Me.Picture1.ForeColor = vbRed
  For I = 90 To -270 Step -5
    Me.Picture1.Cls
    Call drawCircularText(Me.Picture1, "CIRCULAR TEXT - by NDV Software - ", CSng(I), 2500, Me.Picture1.Width / 2, Me.Picture1.Height / 2, 360)
    waitFor (0.01)
  Next I
End Sub

Private Sub Form_Activate()
  Call Command1_Click
End Sub

Private Sub Form_Load()
  Me.TxtCX = (Me.Picture1.Width / 2)
  Me.TxtCY = (Me.Picture1.Height / 2)
End Sub
Public Sub waitFor(Secondi As Single)
  Dim Start, finish As Single
  Start = Timer ' Set start time.
  Do While Timer < Start + Secondi
    'DoEvents  ' Yield to other processes.
  Loop
  finish = Timer  ' Set end time.
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
 Call Command1_Click
End Sub

Private Sub TxtCX_Validate(Cancel As Boolean)
 Call Command1_Click
End Sub

Private Sub TxtCY_Validate(Cancel As Boolean)
  Call Command1_Click
End Sub

Private Sub txtDegree_Validate(Cancel As Boolean)
  Call Command1_Click
End Sub

Private Sub txtRaggio_Validate(Cancel As Boolean)
 Call Command1_Click
End Sub

Private Sub txtSettore_Validate(Cancel As Boolean)
 Call Command1_Click
End Sub

Private Sub txtSimulation_Click()
Dim I As Integer
  Me.Picture1.FontSize = 25
  Me.Picture1.FontName = "Arial black"
  Me.Picture1.ForeColor = vbRed
  For I = 90 To 450 Step 5
    Me.Picture1.Cls
    Call drawCircularText(Me.Picture1, "CIRCULAR TEXT - by NDV Software - ", CSng(I), 2500, Me.Picture1.Width / 2, Me.Picture1.Height / 2, 360)
    waitFor (0.01)
  Next I
End Sub

Private Sub txtSize_Validate(Cancel As Boolean)
 Call Command1_Click
End Sub

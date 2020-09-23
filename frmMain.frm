VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XP Visual Styles - Example"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "More Info ..."
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Radio Buttons"
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   10
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Radio Buttons"
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   9
      Top             =   1920
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Radio Buttons"
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdMsg 
      Caption         =   "Message"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame fraInfo 
      Caption         =   " Information "
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      Begin SHDocVwCtl.WebBrowser wb 
         Height          =   135
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
         ExtentX         =   450
         ExtentY         =   238
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "- Have the manifest file in the same directory as the EXE"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "- Be running the compiled EXE"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "- Be running windows XP"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "To see the visual themes you must:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "This is an example of how to use XP visual styles in Visual Basic 6"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraExample 
      Caption         =   "Example Controls"
      Height          =   3495
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "Check Box"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Text            =   "Text Box"
         Top             =   480
         Width           =   2055
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   3120
         Width           =   2295
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long




Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdInfo_Click()
wb.Navigate "http://msdn.microsoft.com/library/en-us/dnwxp/html/xptheming.asp", , "_new"
End Sub

Private Sub cmdMsg_Click()
    MsgBox "This is an example Message Box!", vbInformation, "Message"
End Sub

Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls
End Sub


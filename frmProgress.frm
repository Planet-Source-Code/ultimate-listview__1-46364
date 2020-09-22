VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   Caption         =   "Progress"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   Icon            =   "frmProgress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   50000
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton cmdGO 
         Caption         =   "Go"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "C&ancel"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         Caption         =   "label1"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdGO_Click()

For x = 1 To ProgressBar1.Max
ProgressBar1.Value = x
Next x


End Sub

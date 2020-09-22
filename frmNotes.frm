VERSION 5.00
Begin VB.Form frmNotes 
   Caption         =   "Note"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Co&py"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtNote 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6615
   End
   Begin VB.Frame frmNote 
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdCopy_Click()

Clipboard.Clear
Clipboard.SetText txtNote.Text
End Sub

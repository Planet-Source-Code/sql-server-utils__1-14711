VERSION 5.00
Begin VB.Form frmBackup 
   Caption         =   "Backup Database"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBckPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdBck 
      Caption         =   "..."
      Height          =   300
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   330
   End
   Begin VB.Label Label5 
      Caption         =   "Backup &File"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

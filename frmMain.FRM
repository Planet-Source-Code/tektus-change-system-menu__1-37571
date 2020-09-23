VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "try to resize the form.."
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0000
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Usage:
    'RemoveSysMenu [FormName from which the menu will
    '               be removed], [Menu Item Number]
    
    RemoveSysMenu Me, 2
End Sub

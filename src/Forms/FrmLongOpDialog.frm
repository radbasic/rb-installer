VERSION 5.00
Begin VB.Form FrmLongOpDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Title"
   ClientHeight    =   948
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   948
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LblMessage 
      Caption         =   "Message"
      Height          =   252
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4452
   End
End
Attribute VB_Name = "FrmLongOpDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' RAD Basic Installer
' Copyright (c) 2019-2025 by RAD Basic Team. All rights reserved.
' Licensed under the MIT License. See License.txt in the project root for license information.
Option Explicit

Public Sub SetMessage(WindowTitle As String, Message As String)
    Me.Caption = WindowTitle
    LblMessage.Caption = Message
End Sub

Private Sub Form_Load()

End Sub

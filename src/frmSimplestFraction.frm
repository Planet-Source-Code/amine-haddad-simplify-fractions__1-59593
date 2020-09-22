VERSION 5.00
Begin VB.Form frmSimplestFraction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demo"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtNewDenum 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtNewNum 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtDenum 
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   615
   End
   Begin VB.Line lneLineNew 
      X1              =   1080
      X2              =   1680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line lneLineOld 
      X1              =   120
      X2              =   720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblEqual 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   135
   End
End
Attribute VB_Name = "frmSimplestFraction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()

    Dim cSF             As New cSimplestFraction
    Dim dblNewNum       As Double
    Dim dblNewDenum     As Double
    Dim strDesc         As String
    
    If cSF.FindFraction(txtNum.Text, txtDenum.Text, dblNewNum, dblNewDenum, strDesc) = False Then
        'The function failed.
        MsgBox strDesc, vbCritical + vbOKOnly, "Error"
        Exit Sub
    Else
        txtNewNum.Text = dblNewNum
        txtNewDenum.Text = dblNewDenum
    End If
    
End Sub

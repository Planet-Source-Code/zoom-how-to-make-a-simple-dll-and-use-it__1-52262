VERSION 5.00
Begin VB.Form Sample 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sample"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "test"
      Height          =   375
      Left            =   1170
      TabIndex        =   2
      Top             =   675
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "test 2"
      Height          =   420
      Left            =   1170
      TabIndex        =   1
      Top             =   1080
      Width           =   1230
   End
   Begin VB.CommandButton cmdAButton 
      Caption         =   "A Button"
      Height          =   375
      Left            =   1170
      TabIndex        =   0
      Top             =   270
      Width           =   1200
   End
End
Attribute VB_Name = "Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyObject As Object  'Declare your Object
Private Sub Form_Load()
    Set MyObject = CreateObject("MyDLL.MYCLASS") 'SET YOUR Object
End Sub
Private Sub cmdAButton_Click()
    MyObject.MyFunction "calc", vbNormalFocus
End Sub
Private Sub Command1_Click()
    MyObject.MyFunction2 ("hello")
End Sub
Private Sub Command2_Click()
    MyObject.MyFunction3 ("hello 2")
End Sub

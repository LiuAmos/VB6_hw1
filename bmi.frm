VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6765
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   3480
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BMI"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "�魫(����)"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "����:(����)"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wweight As Variant
Dim hheight As Variant
Dim bmi As Variant
Private Sub Command1_Click()

bmi = (wweight / (hheight ^ 2))
'List1.AddItem hheight
'List1.AddItem wweight
List1.AddItem bmi

If bmi >= 35 Then
    List1.AddItem "���תέD"
ElseIf bmi >= 30 Then
    List1.AddItem "���תέD"
ElseIf bmi >= 27 Then
    List1.AddItem "���תέD"
ElseIf bmi >= 24 Then
    List1.AddItem "�L��"
ElseIf bmi >= 18.5 Then
    List1.AddItem "���`�d��"
Else
    List1.AddItem "�魫�L��"
End If

'Debug.Print bmi

End Sub

Private Sub Text1_Change()
wweight = Text1.Text
End Sub

Private Sub Text2_Change()
hheight = Text2.Text
End Sub

VERSION 5.00
Begin VB.Form frmgame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392.347
   ScaleMode       =   0  'User
   ScaleWidth      =   342.671
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   0
      Top             =   4200
   End
   Begin VB.Image bird 
      Height          =   527
      Left            =   404
      Picture         =   "AngryBird.frx":0000
      Stretch         =   -1  'True
      Top             =   1022
      Width           =   690
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1185
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   1185
      TabIndex        =   0
      Top             =   1868
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   4860
      Left            =   0
      Picture         =   "AngryBird.frx":36D4B
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   4590
   End
   Begin VB.Menu level 
      Caption         =   "LEVEL"
      Begin VB.Menu level1 
         Caption         =   "LEVEL-1"
      End
      Begin VB.Menu level2 
         Caption         =   "LEVEL-2"
      End
      Begin VB.Menu level3 
         Caption         =   "LEVEL-3"
      End
      Begin VB.Menu level4 
         Caption         =   "LEVEL-4"
      End
      Begin VB.Menu level5 
         Caption         =   "LEVEL-5"
      End
   End
   Begin VB.Menu new 
      Caption         =   "NEW GAME"
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "frmgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yspeed, gravity, gap, score As Integer
Dim n As Integer
Dim pipe(2), toppipe(2) As PictureBox
Dim s, p As String
Dim pipespeed As Single
'exit application
Private Sub exit_Click()
Me.Cls
End
End Sub
'get character [space] from keyboard to jump the bird
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
Beep
yspeed = -12
End If
End Sub
'start timer initialize variables create pipes
Private Sub Form_Load()
Randomize
n = 4
score = 0
gap = 450
pipespeed = 2.5
s = "picturebox1"
p = "picturebox2"
gravity = 2
yspeed = 0
createpipes (1)
createtoppipes (1)
End Sub



'pipespeed for various levels
Private Sub Label1_Click()
Timer2.Enabled = True
End Sub

Private Sub level1_Click()
pipespeed = 2.5
End Sub

Private Sub level2_Click()
pipespeed = 3.5
End Sub

Private Sub level3_Click()
pipespeed = 4.5
End Sub

Private Sub level4_Click()
pipespeed = 5.5
End Sub

Private Sub level5_Click()
pipespeed = 6.5
End Sub
'new game
Private Sub new_Click()
Me.Cls
Unload Me
Me.Show
End Sub

Private Sub Timer1_Timer()
Label3.Caption = score
'falling of bird
yspeed = yspeed + gravity
bird.Top = bird.Top + yspeed
'moving pipes
For i = 0 To 1
pipe(i).Left = pipe(i).Left - pipespeed
toppipe(i).Left = toppipe(i).Left - pipespeed
If bird.Left > pipe(i).Left Then
score = score + 1
End If
If (collision(pipe(i), bird) = True) Or (collision(toppipe(i), bird) = True) Then
Call terminate
End If
'new position for ending pipe to be printed on
If pipe(i).Left < 0 Then
pipe(i).Top = 70 + 270 * Rnd()
toppipe(i).Left = pipe(i).Left + 400
pipe(i).Left = pipe(i).Left + 400
toppipe(i).Top = pipe(i).Top - gap
End If
Next
End Sub
'create bottom pipes
Private Sub createpipes(number As Integer)
Dim i As Integer
For i = 0 To number
Dim temp   As PictureBox
Set temp = Me.Controls.Add("VB.PictureBox", "s" & i)
temp.Height = 350
temp.Width = 50
temp.BackColor = &H800000
temp.BorderStyle = 0
temp.Top = 70 + 270 * Rnd()
temp.Left = (i * 200) + 300
Set pipe(i) = temp
pipe(i).Visible = True
Next
End Sub
'create top pipes
Private Sub createtoppipes(number As Integer)
Dim i As Integer
For i = 0 To number
Dim temp   As PictureBox
Set temp = Me.Controls.Add("VB.PictureBox", "p" & i)
temp.Height = 350
temp.Width = 50
temp.BackColor = &H800000
temp.BorderStyle = 0
temp.Top = pipe(i).Top - gap
temp.Left = (i * 200) + 300
Set toppipe(i) = temp
toppipe(i).Visible = True
Next
End Sub
'start menu
Private Sub Timer2_Timer()
n = n - 1
Label1.Caption = n
If Label1.Caption = "0" Then
Label1.Visible = False
Timer2.Enabled = False
Timer1.Enabled = True
End If
End Sub
'stop application
Private Sub terminate()
Timer1.Enabled = False
Label2.Visible = False
Label3.Visible = False
Label1.FontSize = 14
Label1.Caption = "SCORE"
Label1.ForeColor = &HFF&
Label4.Caption = score
Label1.Visible = True
Label4.Visible = True
End Sub
'collision detection
Private Function collision(ByVal object1 As Object, ByVal object2 As Object) As Boolean
If object1.Top + object1.Height >= object2.Top And _
   object2.Top + object2.Height >= object1.Top And _
   object1.Left + object1.Width >= object2.Left And _
   object2.Left + object2.Width >= object1.Left And object1.Visible = True And object2.Visible = True Then
   collision = True
   Else
   collision = False
    End If
End Function



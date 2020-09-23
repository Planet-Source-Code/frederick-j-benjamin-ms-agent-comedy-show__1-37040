VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form1 
   Caption         =   "Tangent MS Agents"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAgent 
      Caption         =   "Angel"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgent 
      Caption         =   "Genie"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgent 
      Caption         =   "Robby"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgent 
      Caption         =   "Peedy"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Design"
      Height          =   3975
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      Begin VB.ComboBox cboSelect 
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdSpeak 
         Caption         =   "Speak"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtSpeak 
         Height          =   1215
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboAnimations 
         Height          =   315
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdAnimate 
         Caption         =   "Animate"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   2295
      End
      Begin AgentObjectsCtl.Agent msA 
         Left            =   25
         Top             =   720
         _cx             =   847
         _cy             =   847
      End
   End
   Begin VB.CommandButton cmdAgent 
      Caption         =   "Merlin"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CI As String ' Character String (index)
'=====================================================
'You need to download and install the MS Agent Chars
'and a voice engine which takes about an hour
'Here's the URL (posted below) It's worth the download
'http://www.microsoft.com/msagent/downloads.htm#cpl
'or try just->  http://www.microsoft.com/msagent/
'======================================================
Private Sub cboSelect_Click()
Dim AnimationName As Variant
Select Case cboSelect.ListIndex
   Case Is = 0
      CI = "Merlin"
   Case Is = 1
      CI = "Peedy"
   Case Is = 2
      CI = "Robby"
   Case Is = 3
      CI = "Charlie"
End Select
msA.Characters(CI).Show
     '=============== load animations combobox
     cboAnimations.Clear
     For Each AnimationName In msA.Characters(CI).AnimationNames
      cboAnimations.AddItem AnimationName
     Next
'=========== now play the correct show for the selected agent
End Sub

Private Sub cmdAgent_Click(Index As Integer)
Dim AnimationName As Variant
'msA.Characters.Unload CI
Select Case Index
   Case Is = 0
      CI = "Merlin"
      cmdAgent(0).Enabled = False
   Case Is = 1
      CI = "Peedy"
      cmdAgent(1).Enabled = False
   Case Is = 2
      CI = "Robby"
      cmdAgent(2).Enabled = False
   Case Is = 3
      CI = "Genie"
      cmdAgent(3).Enabled = False
   Case Is = 4
      CI = "Angel"
      cmdAgent(4).Enabled = False
End Select
cboSelect.AddItem CI
msA.Characters.Load CI, CI & ".acs"
msA.Characters(CI).Width = 250
msA.Characters(CI).Height = 300
msA.Characters(CI).Show
     '=============== load animations combobox
     cboAnimations.Clear
     For Each AnimationName In msA.Characters(CI).AnimationNames
      cboAnimations.AddItem AnimationName
     Next
'=========== now play the correct show for the selected agent
Select Case Index
    Case Is = 0
       Merlin_Show
    Case Is = 1
       Peedy_Show
    Case Is = 2
       Robby_Show
    Case Is = 3
       Genie_Show
End Select
End Sub

Private Sub cmdAnimate_Click()
If cboAnimations.Text = "" Then Exit Sub
   msA.Characters(CI).Play cboAnimations.Text
End Sub

Private Sub cmdSpeak_Click()
msA.Characters(CI).Speak txtSpeak
End Sub

Private Sub Merlin_Show()
   With msA.Characters(CI)
    .MoveTo 457, 320, 10
    .Play "greet"
    .Play "gestureup"
    .Speak "Why didst thou summon Merlin? Does thou wish to have thy head shrunken to the size of a peanut?"
    .Play "GestureLeft"
    .Speak "Or would thou favor being smothered to death by yonder fat maiden."
    .Play "GetAttentionContinued"
    .Play "GetAttentionContinued"
    .Play "GetAttentionContinued"
    .Play "Decline"
    .Play "Alert"
    .Speak "Hello ?! I'm talking to you, you insignificant crusty little scab on the ass of society. What do you want."
    .Play "DoMagic1"
    .Speak "OK, fine...I will just shrink your penis for you then shall I?"
    .Play "DoMagic2"
    .Speak "Goodbye."
    .Play "Hide"
  End With
End Sub
Private Sub Peedy_Show()
   With msA.Characters(CI)
    .MoveTo 257, 120, 10
    .Play "greet"
    .Play "Announce"
    .Play "alert"
    .Speak "Yo, what's up human? I just walked over from the Birds Nest Titty Club... boy are my feet tired."
    .Play "GestureLeft"
    .Speak "Man that chick in there had some great breasts. Get it?... Get it?... Chicken Breasts! Hah hah ha ha ha!"
    .Play "GetAttentionContinued"
    .Play "Decline"
    .Play "Congratulate"
    .Speak "Hey, I won first prize for having the biggest pecker!... Get it? Ha hah ha ha haw haw!"
    .Play "Confused"
    .Speak "Whoa, I think I am stoned. What was in that bird seed anyway?"
    .Play "DontRecognize"
    .Speak "What was that you said?."
    .Play "LookDownRightBlink"
    .Play "LookDownRightBlink"
    .Play "LookDownRightBlink"
    .Speak "You are really starting to freak me out now... Later dude."
    .Play "Hide"
  End With
End Sub
Private Sub Robby_Show()
   With msA.Characters(CI)
    .MoveTo 557, 220, 10
    .Play "Greet"
    .Play "Confused"
    .Speak "You know you humanoids really are strange."
    .Play "Explain"
    .Speak "To propetuate the species you must engage in extreemly bizarre sexual sub-routines."
    .Play "Suggest"
    .Speak "We Robots can just pull out our schematics and pop a nut and we don't even have to take the bitch out for dinner and a movie."
    .Speak "Please stand-by... this just in..."
    .Play "Read"
    .Play "ReadReturn"
    .Speak "Your planet is to be destroyed at exactly 2:30pm mountain time, in order to make way for a new Hyperspace freeway."
    .Play "Sad"
    .Speak "Damn, and I was looking forward to pulling your heads off one by one."
    .Play "Think"
    .Speak "I wonder if I have enough time to kick at least half of you in the testicles?"
    .Speak "Fuck, oh dear, Oh Well... Ta!"
    .Play "Hide"
  End With
End Sub
Private Sub Genie_Show()
   With msA.Characters(CI)
    .MoveTo 357, 50, 10
    .Play "Process"
    .Play "Surprised"
    .Speak "Whoa, Just a sec. I'm feeling a bit queasy!"
    .Play "Uncertain"
    .Speak "I'm like 5000 years old... I need to stop spinning like that."
    .Play "Think"
    .Speak "OK, let's see now... Oh yeah! I am the Djinni of the lamp. I shall grant you three wishes."
    .Play "Pleased"
    .Speak "Wish number 1..."
    .Play "DoMagic1"
    .Play "DoMagic2"
    .Speak "Done... all your Ex-wives are now in extream pain and will die a horrible death within the hour!"
    .Play "Congratulate_2"
    .Play "Announce"
    .Speak "Wish Number 2..."
    .Play "DoMagic1"
    .Play "DoMagic2"
    .Speak "Done... Your penis is as large as a loaf of french bread and you now own several very expensive sports cars!"
    .Play "Congratulate_2"
    .Play "GetAttention"
    .Speak "No! Nooooo! I don't think that's a very good idea."
    .Play "Greet"
    .Speak "Still, your wish is my command."
    .Play "DoMagic1"
    .Play "DoMagic2"
    .Play "Sad"
    .Speak "I'm sorry I assumed that's what you wanted when you asked for a little head."
    .Play "Hide"
   End With
  End Sub


Private Sub Form_Load()
Me.Top = 150
Me.Left = 150
End Sub

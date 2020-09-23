VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Do Loop Picture FX"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDemo 
      Caption         =   "Demo"
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   6120
      Width           =   1455
   End
   Begin VB.OptionButton OptCandelasDown 
      Caption         =   "Decrease Candelas"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
   Begin VB.OptionButton OptB 
      Caption         =   "Split to Blue"
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   5880
      Width           =   2055
   End
   Begin VB.OptionButton OptG 
      Caption         =   "Split to Green"
      Height          =   315
      Left            =   4200
      TabIndex        =   10
      Top             =   5520
      Width           =   2055
   End
   Begin VB.OptionButton OptR 
      Caption         =   "Split to Red"
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   5160
      Width           =   2055
   End
   Begin VB.OptionButton OptCandelasUp 
      Caption         =   "Increase Candelas"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
   End
   Begin VB.OptionButton OptNeg 
      Caption         =   "Negative Image"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   5160
      Width           =   1455
   End
   Begin VB.OptionButton OptR90 
      Caption         =   "Rotate Right 90*"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.OptionButton OptL90 
      Caption         =   "Rotate Left 90*"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   1455
   End
   Begin VB.OptionButton OptMirrorH 
      Caption         =   "Mirror Horizontal"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.OptionButton OptMirrorV 
      Caption         =   "Mirror Vertical"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.PictureBox After 
      Height          =   4575
      Left            =   4560
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.PictureBox Before 
      Height          =   4575
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewXpixel, NewYpixel, Xpixel, Ypixel, R1, G1, B1 As Integer
Dim i ' i'll need these later
'X is across and Y is high! Remember that one?
Private Sub MirrorV()
'Tell the computer that Y is Zero so that after it is done you can do it again
'without it thinking Y is still 300...
Ypixel = 0
Do
If Xpixel <= 300 Then
NewXpixel = (300 - Xpixel)
After.PSet (NewXpixel, Ypixel), Before.Point(Xpixel, Ypixel)
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
'...I suppose i could have put it here though
'By the way, I know for a fact that this is a 300 by 300 bitmap, that is why
'I have put it in, but you could use scaleheight and scalewidth if you are
'intending to load pictures in and then do the effects
End Sub
Private Sub MirrorH()
'I'll explain how this works:
Ypixel = 0 'You know this already
Do 'Tell it what its about to do
If Xpixel <= 300 Then 'If the current pixel at x is less than or equal to 300
NewYpixel = (300 - Ypixel) 'then the new Y pixel is 300 (scaleheight) minus current pixel at y
After.PSet (Xpixel, NewYpixel), Before.Point(Xpixel, Ypixel) 'set the new pixels on After.PictureBox
Xpixel = (Xpixel + 1) 'add 1 to current x pixel
Else ' but if current x pixel exceeds 300...
Ypixel = (Ypixel + 1) 'add one to the current y pixel
Xpixel = 0 'and put current x pixel back to zero
End If 'end the if statement
Loop Until Ypixel = 300 'do it again and again until ypixel gets to 300
End Sub 'Well Duh!
Private Sub RR90()
'This was a bit of an illegitimate child!
'it works by getting the pixel at 0,0
'and putting it at 300,0
'then 1,0 at 300,1
'2,0 at 300,2
'etc
'and when we get to 300,0
'it goes
'1,1 at 299,1
'2,1 at 299,2
'etc, I hope I haven't lost you.
Ypixel = 0
NewXpixel = 300
NewYpixel = 0
Do
If Xpixel <= 300 Then
After.PSet (NewXpixel, NewYpixel), Before.Point(Xpixel, Ypixel)
Xpixel = (Xpixel + 1)
NewYpixel = (NewYpixel + 1)
Else
Ypixel = (Ypixel + 1)
NewYpixel = 0
NewXpixel = (NewXpixel - 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
'IMPORTANT!
'THIS IS A SQUARE IMAGE, IF YOU INTEND TO LOAD A DIFFERENT
'SHAPE, YOU WILL HAVE TO CHANGE AFTER.WIDTH TO
'BEFORE.HEIGHT AND AFTER.HEIGHT TO BEFORE.WIDTH
'SO THAT IT CAN SUCCESSFULLY ROTATE IN THIS AND THE NEXT
'SUB.
'E.G FROM THIS:

'********************
'*                  *
'*                  *
'********************
'TO THIS

'**********
'*        *
'*        *
'*        *
'*        *
'*        *
'**********

'GET IT?
End Sub
Private Sub RL90()
'This was easy after doing RR90
'See if you can spot the subtle differences
Ypixel = 0
NewXpixel = 0                      '} These two are swapped
NewYpixel = 300                    '}
Do
If Xpixel <= 300 Then
After.PSet (NewXpixel, NewYpixel), Before.Point(Xpixel, Ypixel)
Xpixel = (Xpixel + 1)
NewYpixel = (NewYpixel - 1) 'NewYpixel - 1 instead of + 1
Else
Ypixel = (Ypixel + 1)
NewYpixel = 300                  '300 instead of 0
NewXpixel = (NewXpixel + 1) 'NewXpixel + 1 instead of - 1
Xpixel = 0
End If
Loop Until Ypixel = 300
End Sub
Private Sub Negative()
'I sat for ages trying to work out how to do a negative image.
'So I opened Paint Shop Pro 5 and did negative on a simple image
'I was looking at the rgb value differences for ages and it was staring at me
'I'll show you what I mean:
'Right, say you've got the rgb value 230,123,165
'Now, you do negative on and get   25,132,90
'I was trying to figure out little formulas and stuff and then it just stuck out at me
'I felt like a right dip!
'the maximum is 255,255,255 right?
'230 + 25 = 255
'123 + 132 = 255
'165 + 90 = 255
'Its the differences
'Duh!
'If you are wondering about the bitmap, I was making an example logo for
'a playstion emulator I was hoping to write in the future
'I got the name from the CPU of the playstation
'I thought it would make a nice example for this project
'Oh, playstation is copyright of sony blah blah
'Just so I don't get in any S**T
'Paint Shop Pro is copyright of Jasc
'God
Ypixel = 0
Do
If Xpixel <= 300 Then
R1 = (Before.Point(Xpixel, Ypixel) And &HFF&) 'Get Red Value
G1 = ((Before.Point(Xpixel, Ypixel) And &HFF00&) / &H100&) 'Get Green Value
B1 = ((Before.Point(Xpixel, Ypixel) And &HFF0000) / &H10000) 'Get Blue Value
R1 = (255 - R1) 'Get the differences
G1 = (255 - G1)
B1 = (255 - B1)
After.PSet (Xpixel, Ypixel), RGB(R1, G1, B1)
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
End Sub
Private Sub Candelas()
'Candelas is Luminous Intensity by the way
'Or Brightness in other words
'Bet you didn't know that!
'Who am i kidding
Ypixel = 0
Do
If Xpixel <= 300 Then
R1 = (Before.Point(Xpixel, Ypixel) And &HFF&) 'Get values again
G1 = ((Before.Point(Xpixel, Ypixel) And &HFF00&) / &H100&)
B1 = ((Before.Point(Xpixel, Ypixel) And &HFF0000) / &H10000)
R1 = (R1 * i) 'Times them by i
If R1 > 255 Then R1 = 255 'if it ends up higher then put it to 255
If R1 < 0 Then R1 = 0 'lower than zero then set it to zero
G1 = (G1 * i)              'etc
If G1 > 255 Then G1 = 255
If G1 < 0 Then G1 = 0
B1 = (B1 * i)
If B1 > 255 Then B1 = 255
If B1 < 0 Then B1 = 0
After.PSet (Xpixel, Ypixel), RGB(R1, G1, B1)
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
End Sub                         'Done
Private Sub Split2Red()
Ypixel = 0
Do
If Xpixel <= 300 Then
R1 = (Before.Point(Xpixel, Ypixel) And &HFF&)
G1 = ((Before.Point(Xpixel, Ypixel) And &HFF00&) / &H100&)
B1 = ((Before.Point(Xpixel, Ypixel) And &HFF0000) / &H10000)
After.PSet (Xpixel, Ypixel), RGB(R1, R1, R1) '<<<<<<<<<Get it?
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
End Sub
Private Sub Split2Green()
Ypixel = 0
Do
If Xpixel <= 300 Then
R1 = (Before.Point(Xpixel, Ypixel) And &HFF&)
G1 = ((Before.Point(Xpixel, Ypixel) And &HFF00&) / &H100&)
B1 = ((Before.Point(Xpixel, Ypixel) And &HFF0000) / &H10000)
After.PSet (Xpixel, Ypixel), RGB(G1, G1, G1)  '<<<<<<<<<WELL?
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
End Sub
Private Sub Split2Blue()
Ypixel = 0
Do
If Xpixel <= 300 Then
R1 = (Before.Point(Xpixel, Ypixel) And &HFF&)
G1 = ((Before.Point(Xpixel, Ypixel) And &HFF00&) / &H100&)
B1 = ((Before.Point(Xpixel, Ypixel) And &HFF0000) / &H10000)
After.PSet (Xpixel, Ypixel), RGB(B1, B1, B1)  '<<<<<<<<<GOOD!
Xpixel = (Xpixel + 1)
Else
Ypixel = (Ypixel + 1)
Xpixel = 0
End If
Loop Until Ypixel = 300
End Sub

Private Sub CmdDemo_Click()
'Just do one after another for a little demo
'It doesn't try to do them all at once by the way
MirrorV
MirrorH
RR90
RL90
Negative
i = 1.5
Candelas
i = 0.5
Candelas
Split2Red
Split2Green
Split2Blue
End Sub

Private Sub CmdGo_Click()
'You know the drill...
If OptMirrorV.Value = True Then MirrorV
If OptMirrorH.Value = True Then MirrorH
If OptR90.Value = True Then RR90
If OptL90.Value = True Then RL90
If OptNeg.Value = True Then Negative
If OptCandelasUp.Value = True Then
'set i
i = 1.7 'increase this by about 0.1 each time too see the difference
Candelas
End If
If OptCandelasDown.Value = True Then
i = 0.5 'decrease by about 0.1 each time
Candelas
End If
If OptR.Value = True Then Split2Red
If OptG.Value = True Then Split2Green
If OptB.Value = True Then Split2Blue
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Here's me begging for votes:
MsgBox "Here is an example of using Do Loops for effects in pictures instead of calling from a library (gdi32.lib etc) to do it for you. I'm not really bothered about votes but i'd appreciate it if you took the time to give me one, that is, if you think that I deserve one. Serenity_2k@hotmail.com", vbInformation, "Bye!"
End 'The End!, I hope you've learned something!
End Sub
'Serenity_2k@ hotmail.com

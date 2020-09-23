VERSION 5.00
Begin VB.Form Thumb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thumbnails without OCX or DLL! Please vote if you like it!"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   2160
      TabIndex        =   13
      Top             =   4200
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Make thumbs"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   5640
      Width           =   3015
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   3480
      Pattern         =   "*.jpg;*.bmp;*.gif;*.ico;*.wmf;*.Emf;*.dib;*.cur"
      TabIndex        =   10
      Top             =   4200
      Width           =   1695
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DOWN"
      Height          =   1815
      Left            =   7680
      TabIndex        =   9
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "U P"
      Height          =   1815
      Left            =   7680
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame8 
      Height          =   1935
      Left            =   5760
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
      Begin VB.Image Image8 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1935
      Left            =   3840
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
      Begin VB.Image Image7 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1935
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
      Begin VB.Image Image6 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
      Begin VB.Image Image5 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      Begin VB.Image Image4 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   3840
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      Begin VB.Image Image3 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   1935
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label nisse 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3960
      Width           =   7935
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Thumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JB As Integer

Sub uppd()

Frame1.Visible = True
Frame2.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame5.Visible = True
Frame6.Visible = True
Frame7.Visible = True
Frame8.Visible = True


On Error GoTo balle1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame1.Caption = File1.FileName
Image1.Picture = LoadPicture(nisse.Caption)

On Error GoTo balle2
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame2.Caption = File1.FileName
Image2.Picture = LoadPicture(nisse.Caption)

On Error GoTo balle3
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame3.Caption = File1.FileName
Image3.Picture = LoadPicture(nisse.Caption)
On Error GoTo balle4
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame4.Caption = File1.FileName
Image4.Picture = LoadPicture(nisse.Caption)
On Error GoTo balle5
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame5.Caption = File1.FileName
Image5.Picture = LoadPicture(nisse.Caption)
On Error GoTo balle6
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame6.Caption = File1.FileName
Image6.Picture = LoadPicture(nisse.Caption)
On Error GoTo balle7
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame7.Caption = File1.FileName
Image7.Picture = LoadPicture(nisse.Caption)
On Error GoTo balle8
JB = JB + 1
File1.ListIndex = JB
If Right(File1.Path, 1) <> "\" Then
        nisse.Caption = File1.Path & "\" & File1.FileName
      Else
        nisse.Caption = File1.Path & File1.FileName
        End If
Frame8.Caption = File1.FileName
Image8.Picture = LoadPicture(nisse.Caption)

Exit Sub
balle1:
Frame1.Visible = False

Resume Next
Exit Sub
balle2:
Frame2.Visible = False

Resume Next
Exit Sub
balle3:
Frame3.Visible = False

Resume Next
Exit Sub
balle4:
Frame4.Visible = False

Resume Next
Exit Sub
balle5:
Frame5.Visible = False

Resume Next
Exit Sub
balle6:
Frame6.Visible = False

Resume Next
Exit Sub
balle7:
Frame7.Visible = False

Resume Next
Exit Sub
balle8:
Frame8.Visible = False

Resume Next

Exit Sub
End Sub


Private Sub Command1_Click()
If JB > 14 Then
JB = JB - 15
Call uppd
End If
End Sub

Private Sub Command2_Click()
If Frame1.Visible = False And Frame2.Visible = False And Frame3.Visible = False And Frame4.Visible = False And Frame5.Visible = False And Frame6.Visible = False And Frame7.Visible = False And Frame8.Visible = False Then
Command1.Value = True
Exit Sub
End If
JB = JB + 1
Call uppd
End Sub

Private Sub Command3_Click()
JB = 0
Call uppd
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo ju
Dir1.Path = Drive1.Drive
Exit Sub
ju:
MsgBox "Drive not ready!"
Exit Sub
End Sub



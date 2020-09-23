VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sharp Wizard - LCD Display Emulator"
   ClientHeight    =   6045
   ClientLeft      =   1050
   ClientTop       =   1065
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7440
   Begin VB.CommandButton Command6 
      Caption         =   "Click ME for an EXAMPLE"
      Height          =   405
      Left            =   270
      TabIndex        =   7
      Top             =   135
      Width           =   2460
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLS"
      Height          =   330
      Left            =   5460
      TabIndex        =   6
      Top             =   2955
      Width           =   555
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PSET Command"
      Height          =   315
      Left            =   315
      TabIndex        =   5
      Top             =   2970
      Width           =   1560
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LOCATE Command"
      Height          =   315
      Left            =   1995
      TabIndex        =   4
      Top             =   2970
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LINE Command"
      Height          =   330
      Left            =   3675
      TabIndex        =   3
      Top             =   2955
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox listing 
      Height          =   2520
      Left            =   105
      TabIndex        =   2
      Top             =   3390
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   4445
      _Version        =   393217
      TextRTF         =   $"WizLCD.frx":0000
   End
   Begin VB.PictureBox lcd 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   105
      ScaleHeight     =   2115
      ScaleWidth      =   7170
      TabIndex        =   1
      Top             =   705
      Width           =   7230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   360
      Left            =   6300
      TabIndex        =   0
      Top             =   2940
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "By:  Max Seim - mlseim@mmm.com    01/26/2001"
      Height          =   330
      Left            =   3390
      TabIndex        =   8
      Top             =   150
      Width           =   3720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim foundpos As Integer
Dim foundpos2 As Integer
Dim build As String
Dim scratch As String
Dim scratch2 As String
Dim index As Integer
Private Sub Command1_Click()
Set m_obj = Nothing
For index = Forms.Count - 1 To 0 Step -1
Unload Forms(index)
Next index
End Sub

Private Sub Command2_Click()
'
On Error Resume Next
foundpos = listing.Find("LINE", foundpos2, , rtfWholeWord)
If foundpos = -1 Then
foundpos2 = 0
Exit Sub
End If
' build LINE command
listing.SelStart = foundpos
listing.SelLength = 22
listing.SelColor = vbRed
build = Right$(listing.SelRTF, foundpos - foundpos + 30)
foundpos2 = foundpos + 1
c = 0
scratch = ""
scratch2 = ""
For x = 1 To 10 ' Parse the LINE Command for X1 and Y1 Coordinates
If Mid$(build, x, 4) = "LINE" Then
c = x + 6
10 '
    If Mid$(build, c, 1) <> "," Then
    scratch = scratch + Mid$(build, c, 1)
    X1 = Val(scratch)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 10
    End If
    c = c + 1
    
20 '
    If (Asc(Mid$(build, c, 1)) > 47) And (Asc(Mid$(build, c, 1)) < 58) Then
    scratch2 = scratch2 + Mid$(build, c, 1)
    Y1 = Val(scratch2)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 20
    End If
End If
Next x
k = c
scratch = ""
scratch2 = ""
For x = k To (k + 10) ' Parse the LINE Command for X1 and Y1 Coordinates
If Mid$(build, x, 1) = "(" Then
c = x + 1
30 '
    If Mid$(build, c, 1) <> "," Then
    scratch = scratch + Mid$(build, c, 1)
    X2 = Val(scratch)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 30
    End If
    c = c + 1
40 '
    If (Asc(Mid$(build, c, 1)) > 47) And (Asc(Mid$(build, c, 1)) < 58) Then
    scratch2 = scratch2 + Mid$(build, c, 1)
    Y2 = Val(scratch2)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 40
    End If
End If
Next x

If X1 > 238 Then
X1 = 238
End If
If X1 < 0 Then
X1 = 0
End If
If X2 > 238 Then
X2 = 238
End If
If X2 < 0 Then
X2 = 0
End If
If Y1 > 69 Then
Y1 = 69
End If
If Y1 < 0 Then
Y1 = 0
End If
If Y2 > 69 Then
Y2 = 69
End If
If Y2 < 0 Then
Y2 = 0
End If
If (X2 - X1) < 0 Then
X3 = X1
X1 = X2
X2 = X3
Y3 = Y1
Y1 = Y2
Y2 = Y3
End If
If (Y2 - Y1) < 0 Then
X3 = X1
X1 = X2
X2 = X3
Y3 = Y1
Y1 = Y2
Y2 = Y3
End If
If (X2 - X1) >= (Y2 - Y1) Then
m = 0
b = Y1
If (Y2 - Y1) <> 0 Then
m = (X1 - X2) / (Y1 - Y2)
b = 0 - (m * Y1)
End If
For x = X1 To X2
   y = Y1
   If m <> 0 Then
   y = (x - b) / m
   End If
      lcd.Line (x * 30, y * 30)-Step(15, 15), &H0, BF
Next x
End If

If (Y2 - Y1) > (X2 - X1) Then
m = 0
b = Y1
If (Y2 - Y1) <> 0 Then
m = (X1 - X2) / (Y1 - Y2)
b = 0 - (m * Y1)
End If
For y = Y1 To Y2
   x = X1
   If m <> 0 Then
   x = (y - b) / m
   End If
      lcd.Line (x * 30, y * 30)-Step(15, 15), &H0, BF
Next y
End If
End Sub

Private Sub Command3_Click()
'
On Error Resume Next
foundpos = listing.Find("LOCATE", foundpos2, , rtfWholeWord)
If foundpos = -1 Then
foundpos2 = 0
Exit Sub
End If
' build LOCATE command
listing.SelStart = foundpos
listing.SelLength = 15
listing.SelColor = vbBlue
build = Right$(listing.SelRTF, foundpos - foundpos + 30)
foundpos2 = foundpos + 1
c = 0
scratch = ""
scratch2 = ""
For x = 1 To 20 ' Parse the LOCATE Command for the coordinates.
If Mid$(build, x, 4) = "CATE" Then
c = x + 5
10 '
    If Mid$(build, c, 1) <> "," Then
    scratch = scratch + Mid$(build, c, 1)
    X1 = Val(scratch)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 10
    End If
    c = c + 1
       
20 '
    If (Asc(Mid$(build, c, 1)) > 47) And (Asc(Mid$(build, c, 1)) < 58) Then
    scratch2 = scratch2 + Mid$(build, c, 1)
    Y1 = Val(scratch2)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 20
    End If
End If
Next x
Y1 = (Y1 * 10) + 2
For t = X1 To X1 + 4
   For j = Y1 To Y1 + 7
lcd.Line (t * 30, j * 30)-Step(15, 15), vbBlue, BF
   Next j
Next t
End Sub

Private Sub Command4_Click()
'
On Error Resume Next
foundpos = listing.Find("PSET", foundpos2, , rtfWholeWord)
If foundpos = -1 Then
foundpos2 = 0
Exit Sub
End If
' build PSET command
listing.SelStart = foundpos
listing.SelLength = 13
listing.SelColor = &H8000&
build = Right$(listing.SelRTF, foundpos - foundpos + 30)
foundpos2 = foundpos + 1
c = 0
scratch = ""
scratch2 = ""
For x = 1 To 13 ' Parse the PSET Command for X and Y Coordinates
If Mid$(build, x, 4) = "PSET" Then
c = x + 5
10 '
    If Mid$(build, c, 1) <> "," Then
    scratch = scratch + Mid$(build, c, 1)
    X1 = Val(scratch)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 10
    End If
    c = c + 1
    
20 '
    If (Asc(Mid$(build, c, 1)) > 47) And (Asc(Mid$(build, c, 1)) < 58) Then
    scratch2 = scratch2 + Mid$(build, c, 1)
    Y1 = Val(scratch2)
    c = c + 1
       If c > 40 Then
       Exit Sub
       End If
    GoTo 20
    End If
End If
Next x
lcd.Line (X1 * 30, Y1 * 30)-Step(15, 15), &H8000&, BF
End Sub

Private Sub Command5_Click()
' Clear the LCD Screen
Call matrix
foundpos2 = 0
listing.LoadFile "blank.txt", 1
End Sub

Private Sub Command6_Click()
listing.LoadFile "example.txt", 1
End Sub

Private Sub Form_Load()
Call matrix ' Paint the LCD
End Sub
Private Sub matrix()
'Create LCD Matrix
For x = 0 To 238 ' Number of Pixels Across LCD
   For y = 0 To 69 ' Number of Pixels Down LCD
      lcd.Line (x * 30, y * 30)-Step(15, 15), &HE0E0E0, BF
   Next y
Next x
End Sub

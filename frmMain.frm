VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alphanumeric LED display"
   ClientHeight    =   3144
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5904
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3144
   ScaleWidth      =   5904
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkScroll 
      BackColor       =   &H00404040&
      Caption         =   "Scroll text"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optColour 
      BackColor       =   &H00404040&
      Caption         =   "Blue"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton optColour 
      BackColor       =   &H00404040&
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton optColour 
      BackColor       =   &H00404040&
      Caption         =   "Green"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtString 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label lblColour 
      BackColor       =   &H00404040&
      Caption         =   "Choose your display LEDs colour:"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblDigits 
      BackColor       =   &H00404040&
      Caption         =   "Use these buttons to set the number of digits to display:"
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label lblNumber 
      BackColor       =   &H00404040&
      Caption         =   "Enter the string to display:"
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'6 sprites HDCs: Green on/off, Red on/off, Blue on/off
Dim SpritesHDC(6) As Long
'Number of digits allowed (user selected)
'Byte because it varies from 1 to 15, but you can put whatever you want...
'as long as your screen is large enough :) Use smaller sprites to display plenty of them
Dim Digits As Byte
'Width and height of a LED
'Have to declare as Integer (and not byte) because of sprite functions reference type
Dim LedWidth, LedHeight As Integer
'Display starting position, i.e. upper-left display corner position
'Have to use Integer as in my case StartPosX is negative
Dim StartPosX, StartPosY As Integer

Private Sub Form_Load()
   'Initialise LEDs width and height, i.e. bitmap image size in pixels
   LedWidth = 8
   LedHeight = 8
   'Set upper-left corner display coordinates
   StartPosX = -50
   StartPosY = 20
   'Initialise digits for ASCII table (in modAsciiTable module)
   Call DeclareAsciiTable
   'Load sprites in memory
   'I HAD to use 40 and 10 instead of LedWidth and LedHeight
   'in the following HDCs and sprites declarations.
   'I've got no idea why, I'd appreciate any help on that :)
   SpritesHDC(0) = CreateMemHdc(Me.hdc, 8, 8)
   Call LoadBmpToHdc(SpritesHDC(0), "GreenOff.bmp") 'Green, off
   SpritesHDC(1) = CreateMemHdc(Me.hdc, 8, 8)
   Call LoadBmpToHdc(SpritesHDC(1), "GreenOn.bmp")  'Green, on
   SpritesHDC(2) = CreateMemHdc(Me.hdc, 8, 8)
   Call LoadBmpToHdc(SpritesHDC(2), "RedOff.bmp")   'Red, off
   SpritesHDC(3) = CreateMemHdc(Me.hdc, 8, 8)
   Call LoadBmpToHdc(SpritesHDC(3), "RedOn.bmp")    'Red, on
   SpritesHDC(4) = CreateMemHdc(Me.hdc, 8, 8)
   Call LoadBmpToHdc(SpritesHDC(4), "BlueOff.bmp")  'Blue, off
   SpritesHDC(5) = CreateMemHdc(Me.hdc, 8, 8)
   Call LoadBmpToHdc(SpritesHDC(5), "BlueOn.bmp")   'Blue, on
   'Set default digits number to 3. Change it to whatever you need.
   Digits = 3
   'Show display
   Call cmdDisplay_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'IMPORTANT! Release all created hDCs to prevent memory leaks!
   Call DestroyHdcs
   'And exit... properly please :)
   Unload Me
   Set frmMain = Nothing
End Sub

Private Sub cmdPlus_Click()
   'Allow maximum of 15 digits
   If Digits < 15 Then
      Digits = Digits + 1
   Else
      Exit Sub
   End If
   'Update form
   Call UpdateWithNewValue
End Sub

Private Sub cmdMinus_Click()
   'Allow minimum of 1 digits
   If Digits > 1 Then
      Digits = Digits - 1
   Else
      Exit Sub
   End If
   'Clear form to remove extra digits
   Me.Cls
   'Update form
   Call UpdateWithNewValue
End Sub

Private Sub UpdateWithNewValue()
   'As it is called from two locations above, I prefered to create a separate Sub
   'Redimension form (if 6 digits or more, else we would lose the controls)
   If Digits > 6 Then Me.Width = 1000 * Digits
   'Update display
   Call cmdDisplay_Click
End Sub

Private Sub cmdDisplay_Click()
   'Loops, byte is enough
   Dim i, j, k As Byte
   'LEDs colour: 0=green, 1=red, 2=blue
   'Thus, byte is enough
   Dim LedColour As Byte
   'ASCII value for current digit to process, goes from 0 to 255 => byte is enough
   Dim DigitValue As Byte
   'Power of 2 for "And" test. Store it to avoid redundant computations
   'From 1 to 64, thus byte is enough
   Dim Weight As Byte
   'Temporary string for scrolling
   Dim ScrollingText As String
   'Scrolling status/position. Byte, coz 255 characters should be enough
   Dim ScrollPosition As Byte
   'Boolean set to true when scrolling is completed
   Dim Scrolled As Boolean
   'GetTickCount values for timing
   Dim CurrTick, PrevTick As Long
   
   'Convert to upper case
   'Sorry, no lower-case for now, I'm lazy! Do it, go ahead! And send me the updated code :)
   txtString.Text = UCase(txtString.Text)
   'Set LEDs colour according to option selected: 0=green, 1=red, 2=blue
   'You should use a "Case Select" if you want to use more sprites sets
   If optColour(0).Value Then
      LedColour = 0
   ElseIf optColour(1).Value Then
      LedColour = 1
   Else
      LedColour = 2
   End If
   'Initialise scrolling to False if we want the text to scroll
   If chkScroll.Value = 1 Then
      Scrolled = False
   Else
      'No scrolling? Set to True to allow one single loop
      Scrolled = True
   End If
   'Initialise scrolling string position
   ScrollPosition = 0
   'Initialise timer
   PrevTick = GetTickCount
   Do
      CurrTick = GetTickCount
      'Replace the 150 in next line by whichever value you wish
      'The "Not Scrolled" allow to display the string once is "Scroll" is not checked
      If (CurrTick - PrevTick) < 150 And Not Scrolled Then
         DoEvents
      Else
         'Increment scrolling position
         ScrollPosition = ScrollPosition + 1
         'Trunk text to display from text box
         ScrollingText = Mid(txtString.Text, ScrollPosition, Digits)
         'If text is empty: done
         If Len(ScrollingText) = 0 Then Scrolled = True
         'Add spaces at the end of the string to display
         'The program would crash if we didn't do that :)
         If Len(ScrollingText) < Digits Then
            For i = Len(ScrollingText) To Digits
               ScrollingText = ScrollingText + " "
            Next
         End If
         'Process digits one by one, starting from left hand side character
         For i = 1 To Digits
            'Capture current digit to proceed
            DigitValue = Asc(Mid(ScrollingText, i, 1))
            For j = 0 To 6 'j is row, reduces number of 2^j computations
               Weight = 2 ^ j
               For k = 0 To 6
                  If (AlphaCodes(DigitValue).LedLine(k) And Weight) = Weight Then
                     'Display on LED
                     Call DisplayDigit(StartPosX + i * 8 * LedWidth + j * LedWidth, 20 + k * LedHeight, LedHeight, LedWidth, LedColour * 2 + 1)
                  Else
                     'Display off LED
                     Call DisplayDigit(StartPosX + i * 8 * LedWidth + j * LedWidth, 20 + k * LedHeight, LedHeight, LedWidth, LedColour * 2)
                  End If
               Next
            Next
         Next
         'Refresh to show updated display: the lame and easy way :) + AutoRedraw! Blech!
         Me.Refresh
         'Update timer
         PrevTick = CurrTick
      End If
   Loop Until Scrolled
   'Finished
End Sub

Private Sub DisplayDigit(PosX, PosY, LedWidth, LedHeight, SpriteNumber As Byte)
   'Display sprite SpriteNumber at coordinates PosX,PosY on form (Me.hdc)
   'Also use LedWidth and LedHeight here, so it's 100% flexible
   'SpriteNumber specifies which LED to display
   BitBlt Me.hdc, PosX, PosY, LedWidth, LedHeight, SpritesHDC(SpriteNumber), 0, 0, vbSrcCopy
End Sub

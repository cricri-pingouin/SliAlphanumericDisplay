Attribute VB_Name = "modAsciiTable"
Option Explicit

'Declare GetTickCount API for timing (replaces timer controls)
'I had to stick that somewhere and I didn't want to create a module just for that
Declare Function GetTickCount Lib "kernel32.dll" () As Long

'Characters consist in 7 lines of 7 rows
'As row contains only 7 LEDs, a byte per line is enough
'Array containing coding for one ASCII characters
Public Type AlphaCode
   LedLine(7) As Byte 'Each digit consists in 7 lines
End Type
'Array containing all ASCII characters coding, 255 of them
Public AlphaCodes(255) As AlphaCode

Public Sub DeclareAsciiTable()

'The following numbering for the 7 LEDs in a line is used (powers of 2)
' 1 2 4 8 16 32 64

   'Determine LEDs to switch on for the different ASCII characters
   'HAVE TO DO ALL OF THEM! BOOOOOORING!
   'Space
   AlphaCodes(48).LedLine(0) = 0 '0 everywhere to keep all LEDs off
   AlphaCodes(48).LedLine(1) = 0
   AlphaCodes(48).LedLine(2) = 0
   AlphaCodes(48).LedLine(3) = 0
   AlphaCodes(48).LedLine(4) = 0
   AlphaCodes(48).LedLine(5) = 0
   AlphaCodes(48).LedLine(6) = 0
   '0
   AlphaCodes(48).LedLine(0) = 28 'Example, here = 4 + 8 + 16 for the 3 LEDs in the center of the first line
   AlphaCodes(48).LedLine(1) = 34
   AlphaCodes(48).LedLine(2) = 81
   AlphaCodes(48).LedLine(3) = 73
   AlphaCodes(48).LedLine(4) = 69
   AlphaCodes(48).LedLine(5) = 34
   AlphaCodes(48).LedLine(6) = 28 'Lower line is the same then the first one. I hope you get it now :)
   '1
   AlphaCodes(49).LedLine(0) = 8
   AlphaCodes(49).LedLine(1) = 12
   AlphaCodes(49).LedLine(2) = 8
   AlphaCodes(49).LedLine(3) = 8
   AlphaCodes(49).LedLine(4) = 8
   AlphaCodes(49).LedLine(5) = 8
   AlphaCodes(49).LedLine(6) = 28
   '2
   AlphaCodes(50).LedLine(0) = 62
   AlphaCodes(50).LedLine(1) = 65
   AlphaCodes(50).LedLine(2) = 64
   AlphaCodes(50).LedLine(3) = 48
   AlphaCodes(50).LedLine(4) = 12
   AlphaCodes(50).LedLine(5) = 2
   AlphaCodes(50).LedLine(6) = 127
   '3
   AlphaCodes(51).LedLine(0) = 62
   AlphaCodes(51).LedLine(1) = 65
   AlphaCodes(51).LedLine(2) = 64
   AlphaCodes(51).LedLine(3) = 60
   AlphaCodes(51).LedLine(4) = 64
   AlphaCodes(51).LedLine(5) = 65
   AlphaCodes(51).LedLine(6) = 62
   '4
   AlphaCodes(52).LedLine(0) = 16
   AlphaCodes(52).LedLine(1) = 8
   AlphaCodes(52).LedLine(2) = 4
   AlphaCodes(52).LedLine(3) = 2
   AlphaCodes(52).LedLine(4) = 17
   AlphaCodes(52).LedLine(5) = 127
   AlphaCodes(52).LedLine(6) = 16
   '5
   AlphaCodes(53).LedLine(0) = 127
   AlphaCodes(53).LedLine(1) = 1
   AlphaCodes(53).LedLine(2) = 1
   AlphaCodes(53).LedLine(3) = 62
   AlphaCodes(53).LedLine(4) = 64
   AlphaCodes(53).LedLine(5) = 65
   AlphaCodes(53).LedLine(6) = 62
   '6
   AlphaCodes(54).LedLine(0) = 62
   AlphaCodes(54).LedLine(1) = 65
   AlphaCodes(54).LedLine(2) = 1
   AlphaCodes(54).LedLine(3) = 63
   AlphaCodes(54).LedLine(4) = 65
   AlphaCodes(54).LedLine(5) = 65
   AlphaCodes(54).LedLine(6) = 62
   '7
   AlphaCodes(55).LedLine(0) = 62
   AlphaCodes(55).LedLine(1) = 65
   AlphaCodes(55).LedLine(2) = 64
   AlphaCodes(55).LedLine(3) = 32
   AlphaCodes(55).LedLine(4) = 16
   AlphaCodes(55).LedLine(5) = 8
   AlphaCodes(55).LedLine(6) = 4
   '8
   AlphaCodes(56).LedLine(0) = 62
   AlphaCodes(56).LedLine(1) = 65
   AlphaCodes(56).LedLine(2) = 65
   AlphaCodes(56).LedLine(3) = 62
   AlphaCodes(56).LedLine(4) = 65
   AlphaCodes(56).LedLine(5) = 65
   AlphaCodes(56).LedLine(6) = 62
   '9
   AlphaCodes(57).LedLine(0) = 62
   AlphaCodes(57).LedLine(1) = 65
   AlphaCodes(57).LedLine(2) = 65
   AlphaCodes(57).LedLine(3) = 126
   AlphaCodes(57).LedLine(4) = 64
   AlphaCodes(57).LedLine(5) = 65
   AlphaCodes(57).LedLine(6) = 62
   'A
   AlphaCodes(65).LedLine(0) = 62
   AlphaCodes(65).LedLine(1) = 65
   AlphaCodes(65).LedLine(2) = 65
   AlphaCodes(65).LedLine(3) = 127
   AlphaCodes(65).LedLine(4) = 65
   AlphaCodes(65).LedLine(5) = 65
   AlphaCodes(65).LedLine(6) = 65
   'B
   AlphaCodes(66).LedLine(0) = 63
   AlphaCodes(66).LedLine(1) = 65
   AlphaCodes(66).LedLine(2) = 65
   AlphaCodes(66).LedLine(3) = 63
   AlphaCodes(66).LedLine(4) = 65
   AlphaCodes(66).LedLine(5) = 65
   AlphaCodes(66).LedLine(6) = 63
   'C
   AlphaCodes(67).LedLine(0) = 62
   AlphaCodes(67).LedLine(1) = 65
   AlphaCodes(67).LedLine(2) = 1
   AlphaCodes(67).LedLine(3) = 1
   AlphaCodes(67).LedLine(4) = 1
   AlphaCodes(67).LedLine(5) = 65
   AlphaCodes(67).LedLine(6) = 62
   'D
   AlphaCodes(68).LedLine(0) = 63
   AlphaCodes(68).LedLine(1) = 65
   AlphaCodes(68).LedLine(2) = 65
   AlphaCodes(68).LedLine(3) = 65
   AlphaCodes(68).LedLine(4) = 65
   AlphaCodes(68).LedLine(5) = 65
   AlphaCodes(68).LedLine(6) = 63
   'E
   AlphaCodes(69).LedLine(0) = 127
   AlphaCodes(69).LedLine(1) = 1
   AlphaCodes(69).LedLine(2) = 1
   AlphaCodes(69).LedLine(3) = 63
   AlphaCodes(69).LedLine(4) = 1
   AlphaCodes(69).LedLine(5) = 1
   AlphaCodes(69).LedLine(6) = 127
   'F
   AlphaCodes(70).LedLine(0) = 127
   AlphaCodes(70).LedLine(1) = 1
   AlphaCodes(70).LedLine(2) = 1
   AlphaCodes(70).LedLine(3) = 63
   AlphaCodes(70).LedLine(4) = 1
   AlphaCodes(70).LedLine(5) = 1
   AlphaCodes(70).LedLine(6) = 1
   'G
   AlphaCodes(71).LedLine(0) = 62
   AlphaCodes(71).LedLine(1) = 65
   AlphaCodes(71).LedLine(2) = 1
   AlphaCodes(71).LedLine(3) = 121
   AlphaCodes(71).LedLine(4) = 65
   AlphaCodes(71).LedLine(5) = 65
   AlphaCodes(71).LedLine(6) = 62
   'H
   AlphaCodes(72).LedLine(0) = 65
   AlphaCodes(72).LedLine(1) = 65
   AlphaCodes(72).LedLine(2) = 65
   AlphaCodes(72).LedLine(3) = 127
   AlphaCodes(72).LedLine(4) = 65
   AlphaCodes(72).LedLine(5) = 65
   AlphaCodes(72).LedLine(6) = 65
   'I
   AlphaCodes(73).LedLine(0) = 28
   AlphaCodes(73).LedLine(1) = 8
   AlphaCodes(73).LedLine(2) = 8
   AlphaCodes(73).LedLine(3) = 8
   AlphaCodes(73).LedLine(4) = 8
   AlphaCodes(73).LedLine(5) = 8
   AlphaCodes(73).LedLine(6) = 28
   'J
   AlphaCodes(74).LedLine(0) = 28
   AlphaCodes(74).LedLine(1) = 8
   AlphaCodes(74).LedLine(2) = 8
   AlphaCodes(74).LedLine(3) = 8
   AlphaCodes(74).LedLine(4) = 8
   AlphaCodes(74).LedLine(5) = 9
   AlphaCodes(74).LedLine(6) = 6
   'K
   AlphaCodes(75).LedLine(0) = 33
   AlphaCodes(75).LedLine(1) = 17
   AlphaCodes(75).LedLine(2) = 9
   AlphaCodes(75).LedLine(3) = 7
   AlphaCodes(75).LedLine(4) = 9
   AlphaCodes(75).LedLine(5) = 17
   AlphaCodes(75).LedLine(6) = 33
   'L
   AlphaCodes(76).LedLine(0) = 1
   AlphaCodes(76).LedLine(1) = 1
   AlphaCodes(76).LedLine(2) = 1
   AlphaCodes(76).LedLine(3) = 1
   AlphaCodes(76).LedLine(4) = 1
   AlphaCodes(76).LedLine(5) = 1
   AlphaCodes(76).LedLine(6) = 127
   'M
   AlphaCodes(77).LedLine(0) = 65
   AlphaCodes(77).LedLine(1) = 99
   AlphaCodes(77).LedLine(2) = 85
   AlphaCodes(77).LedLine(3) = 73
   AlphaCodes(77).LedLine(4) = 65
   AlphaCodes(77).LedLine(5) = 65
   AlphaCodes(77).LedLine(6) = 65
   'N
   AlphaCodes(78).LedLine(0) = 65
   AlphaCodes(78).LedLine(1) = 67
   AlphaCodes(78).LedLine(2) = 69
   AlphaCodes(78).LedLine(3) = 73
   AlphaCodes(78).LedLine(4) = 81
   AlphaCodes(78).LedLine(5) = 97
   AlphaCodes(78).LedLine(6) = 65
   'O
   AlphaCodes(79).LedLine(0) = 62
   AlphaCodes(79).LedLine(1) = 65
   AlphaCodes(79).LedLine(2) = 65
   AlphaCodes(79).LedLine(3) = 65
   AlphaCodes(79).LedLine(4) = 65
   AlphaCodes(79).LedLine(5) = 65
   AlphaCodes(79).LedLine(6) = 62
   'P
   AlphaCodes(80).LedLine(0) = 63
   AlphaCodes(80).LedLine(1) = 65
   AlphaCodes(80).LedLine(2) = 65
   AlphaCodes(80).LedLine(3) = 63
   AlphaCodes(80).LedLine(4) = 1
   AlphaCodes(80).LedLine(5) = 1
   AlphaCodes(80).LedLine(6) = 1
   'Q
   AlphaCodes(81).LedLine(0) = 62
   AlphaCodes(81).LedLine(1) = 65
   AlphaCodes(81).LedLine(2) = 65
   AlphaCodes(81).LedLine(3) = 65
   AlphaCodes(81).LedLine(4) = 81
   AlphaCodes(81).LedLine(5) = 33
   AlphaCodes(81).LedLine(6) = 94
   'R
   AlphaCodes(82).LedLine(0) = 63
   AlphaCodes(82).LedLine(1) = 65
   AlphaCodes(82).LedLine(2) = 65
   AlphaCodes(82).LedLine(3) = 63
   AlphaCodes(82).LedLine(4) = 17
   AlphaCodes(82).LedLine(5) = 33
   AlphaCodes(82).LedLine(6) = 65
   'S
   AlphaCodes(83).LedLine(0) = 62
   AlphaCodes(83).LedLine(1) = 65
   AlphaCodes(83).LedLine(2) = 1
   AlphaCodes(83).LedLine(3) = 62
   AlphaCodes(83).LedLine(4) = 64
   AlphaCodes(83).LedLine(5) = 65
   AlphaCodes(83).LedLine(6) = 62
   'T
   AlphaCodes(84).LedLine(0) = 127
   AlphaCodes(84).LedLine(1) = 8
   AlphaCodes(84).LedLine(2) = 8
   AlphaCodes(84).LedLine(3) = 8
   AlphaCodes(84).LedLine(4) = 8
   AlphaCodes(84).LedLine(5) = 8
   AlphaCodes(84).LedLine(6) = 8
   'U
   AlphaCodes(85).LedLine(0) = 65
   AlphaCodes(85).LedLine(1) = 65
   AlphaCodes(85).LedLine(2) = 65
   AlphaCodes(85).LedLine(3) = 65
   AlphaCodes(85).LedLine(4) = 65
   AlphaCodes(85).LedLine(5) = 65
   AlphaCodes(85).LedLine(6) = 62
   'V
   AlphaCodes(86).LedLine(0) = 65
   AlphaCodes(86).LedLine(1) = 65
   AlphaCodes(86).LedLine(2) = 65
   AlphaCodes(86).LedLine(3) = 65
   AlphaCodes(86).LedLine(4) = 34
   AlphaCodes(86).LedLine(5) = 20
   AlphaCodes(86).LedLine(6) = 8
   'W
   AlphaCodes(87).LedLine(0) = 65
   AlphaCodes(87).LedLine(1) = 65
   AlphaCodes(87).LedLine(2) = 65
   AlphaCodes(87).LedLine(3) = 73
   AlphaCodes(87).LedLine(4) = 85
   AlphaCodes(87).LedLine(5) = 99
   AlphaCodes(87).LedLine(6) = 65
   'X
   AlphaCodes(88).LedLine(0) = 65
   AlphaCodes(88).LedLine(1) = 34
   AlphaCodes(88).LedLine(2) = 20
   AlphaCodes(88).LedLine(3) = 8
   AlphaCodes(88).LedLine(4) = 20
   AlphaCodes(88).LedLine(5) = 34
   AlphaCodes(88).LedLine(6) = 65
   'Y
   AlphaCodes(89).LedLine(0) = 65
   AlphaCodes(89).LedLine(1) = 65
   AlphaCodes(89).LedLine(2) = 34
   AlphaCodes(89).LedLine(3) = 20
   AlphaCodes(89).LedLine(4) = 8
   AlphaCodes(89).LedLine(5) = 8
   AlphaCodes(89).LedLine(6) = 8
   'Z
   AlphaCodes(90).LedLine(0) = 127
   AlphaCodes(90).LedLine(1) = 32
   AlphaCodes(90).LedLine(2) = 16
   AlphaCodes(90).LedLine(3) = 8
   AlphaCodes(90).LedLine(4) = 4
   AlphaCodes(90).LedLine(5) = 2
   AlphaCodes(90).LedLine(6) = 127
End Sub


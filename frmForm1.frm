VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "A Particle's Tale"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmForm1.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   4560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Made by Cory Ului(mouak@crosswinds.net.
'I'd sure love it if this once went in a book. notify me if
'it ever might, so I know.. plzzz thanx! Luv ya.. hehe

'if a Particle ever found out... ouch!!! it'll be my butt!!
Private Declare Function PaINtSPloT Lib "gdi32" Alias "SetPixel" (ByVal SPLotLand As Long, ByVal encome As Long, ByVal wiSDom As Long, ByVal smELL As Long) As Long
Private Declare Function SnIFFSPLot Lib "gdi32" Alias "GetPixel" (ByVal SmELLySPLotLand As Long, ByVal encome As Long, ByVal wiSDom As Long) As Long
Private Declare Function SPloTLaNDTiME Lib "kernel32" Alias "GetTickCount" () As Long

Private Type Part
  x As Integer
  y As Integer
  L As Integer
  T As Byte
  D As Byte
  Dy As Byte
End Type

Dim Particle(0 To 599) As Part, ic As Long
Dim ix As Integer, iy As Integer, iend As Byte, idc As Long
'the above code is the matrix to it all. the GEN, the DNA, the BRAIN.
'okay okay okaYyyyyyy!!!!! it's just code.. geee! (NEVER TELL A
'PARTICLE........ NOT ONE)

Private Sub DoParticle()
Dim i As Integer
  iend = 1
  Do
    'Particle's need time to be bound to, just like us! (other wise were
    'out of control. no???
    If SPloTLaNDTiME - itc >= 15 Then
    
      For i = 0 To 599
        If Particle(i).T > 0 Then
          PaINtSPloT idc, Particle(i).x, Particle(i).y, 0
          
          'If Particle's Direction is Left, then Left it is! (same with right)
          If Particle(i).D = 1 Then
            Particle(i).x = Particle(i).x - 1
          ElseIf Particle(i).D = 2 Then
            Particle(i).x = Particle(i).x + 1
          End If
          
          'Make the Particle Fall
          Particle(i).y = Particle(i).y + 1
          
          '"A Particle shall never harm another Particle, only join and -
          ' -Shine better and brigher then before!" - Particle(104)
          
          'Make the Particle's life minus (the less life a Particle has
          'the more depressed, more dull, and week it gets.
          Particle(i).L = Particle(i).L - 4
          
          'If Particle's lucky number is choosen then
          'Particle might be able to apply for a bypass!
          If Int(10 * Rnd) = 1 Then
            'If Particle's life could handle a bypass operation then
            'give it one, lets the Particle live a little longer!
            'Else the Particle is beyond known medicene to cure! :(
            If Particle(i).L >= 128 Then Particle(i).L = Particle(i).L + 32
            
            'If Particle is actualy a Particle and not a human!!!
            '(u guys! GEE!) then Particle goes right, and HUMANS GO
            'LEFT!!!!!!
            If Int(2 * Rnd) = 1 Then
              Particle(i).x = Particle(i).x + 1
            Else
              Particle(i).x = Particle(i).x - 1
            End If
            
            'AGAIN!! HUMANS! ARGGHHHH.......
            'If Particle is ACTUALY A BLIMIN Particle and not a
            'FRIGGIN HU-MON!! then continue
            If Int(2 * Rnd) = 1 Then
              'ARGGGHHHHH.
              'If Paticle is ACTUALY a Particle and not a human
              'and not pretenting to be a Particle by pretending to
              'be a human, by pretenting to be a Particle then
              'FOR FRIGGIN SAKE, go UP, and you humans go down!!
              If Int(2 * Rnd) = 1 Then
                Particle(i).y = Particle(i).y + 1
              Else
                Particle(i).y = Particle(i).y - 1
              End If
            End If
          End If
          '"Without humans Particle don't exist! say humans! BUT  -
          ' -without Particles humans are nothing but ...ummmm errrr-
          ' - mare...... ummm... NOTHING" - Particle(462)
          
          
          'If Particle is alittle dizzy or disorientated from the fall
          'and/or movement of the "GREAT GOD BIG BARE MOUSE (no extras)"
          'that makes Particle go past known pysisics! then smack
          'Particle's face and bring Particle back to the ground and
          'not past it! and take 4 of Particle's life away for the
          'heck of it!
          If Particle(i).y > 240 Then Particle(i).y = 240: Particle(i).L = Particle(i).L - 4
          
          'If Particle *SOB*......... umm sorry.. if Particle *SOB* WAHHH
          'I can't go on.. I must.. if Particle is dead.. *SOB* *SNIFF*
          'then Particle must be 0 and Particle's Existence (T) is set
          'to Nothing. R.I.P
          'what's that you say? hmmm? Yes yes. a decent furual, with no
          'hu-mons in site.....
          If Particle(i).L <= 0 Then Particle(i).L = 0: Particle(i).T = 0
          
          '"ruff woof woof" - Particle(6) ((( DON'T ASK AYEE!!! )))
          
          'Somethings are just so hard to understand, you figure this
          'one out!... argghhhhh.
          If SnIFFSPLot(idc, Particle(i).x, Particle(i).y) > 0 Then
            PaINtSPloT idc, Particle(i).x, Particle(i).y, &H80FFFF
          Else
            PaINtSPloT idc, Particle(i).x, Particle(i).y, RGB(Particle(i).L, Particle(i).L, 0)
          End If
        End If
      Next
      
      'Set Particle Daylight time to some metrics with the 0 and
      '80151/214*4%.021 I HAVE NO IDEA!!! I've only got a
      'Particle pass this evening!
      itc = SPloTLaNDTiME
      
    End If
    'UR KIDDING RIGHT, U HAVE NO IDEA YOU SAY??? OH GEEE!!
    DoEvents
    
    'loop until the end of the Particle world, and guess what
    'you guest it. ONLY HUMANS KILL Particle world....
    'POP QUIZ HOT SHOT, WHADA DO....... YOUR A HUMAN WHAT ELSE.
  Loop Until iend = 0
End Sub

Private Sub Form_Load()
  'Show the stupid form!
  Me.Show
  
  'Get a pysical state where Particle's can be seen, and not
  'touched!
  idc = Me.hdc
  
  'In the begining..........
  DoParticle
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer, j As Byte

  '"GREAT GOD BIG BARE MOUSE (no extras)" creats 4 Particles in
  '"GREAT GOD BIG BARE MOUSE (no extras)" image, but he means
  'in the way he thinks and processes! 4 Particle are created
  'with every movement from clouds the "GREAT GOD BIG... yeh yeh!
  For j = 0 To 3
    'Due to population limits, that makes Particle's get crowed
    'and they bicker and fight, every Particle is counted and
    'checked for absence and stuff..
    For i = 0 To 599
      'if Particle is died R.I.P, dead. :( then
      If Particle(i).T = 0 Then
        'one out of 4 Particles are born a leader! If the Particle
        'that is getting investergated is that leader then
        If j = 0 Then
          'Particle's position in life, is the position as a
          'leader!
          Particle(i).x = x
          Particle(i).y = y
          'Particle's Life is fresh, healthy as a leader
          Particle(i).L = 255
          'Particle's Existance is vaild! Particle belongs!
          Particle(i).T = 1
        Else
          'Other wise you might as well be road kill
          'or here or there!
          Particle(i).x = x + 6 - Int(12 * Rnd)
          Particle(i).y = y + 6 - Int(12 * Rnd)
          'And your life pretty much sucks, you have stupid neighbours
          'and no leadership skill!
          Particle(i).L = 196 + Int(64 * Rnd)
          'But your given a chance just in case!
          Particle(i).T = 1
        End If
        
        '"GREAT GOD BIG BARE MOUSE (no extras)" either goes
        'left or right. Particle's life chances depending!
        If ix > x Then
          Particle(i).D = 1
        ElseIf ix < x Then
          Particle(i).D = 2
        ElseIf ix = x Then
          Particle(i).D = 0
        End If
        
        'Oh and or up or down. Particle's life is effected by
        'ever mood chance, swing and movement!
        If iy > y Then
          Particle(i).Dy = 1
        ElseIf iy < y Then
          Particle(i).Dy = 2
        ElseIf iy = y Then
          Particle(i).Dy = 0
        End If
        
        'so spawed a Particle
        PaINtSPloT idc, Particle(i).x, Particle(i).y, RGB(Particle(i).L, Particle(i).L, 0)
        
        Exit For
      End If
    Next i
    'ok ok.. I'll tell you, this is what it does. Do[es OTHER] Events!
    DoEvents
  Next j
  'And old acient god of "GREAT GOD BIG BARE... yeh yeh hears about
  'new "GREAT GOD BIG.. yeh ..... progress and moods, and position..
  'on the clouds!
  ix = x
  iy = y
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '*sob*.... how... how could you, why.....*YELLS* WHY!!!!!!... *sob*
  '*waaa*.... hu........mons!!
  iend = 0
End Sub

'PS (No offence is taken to humans, or ittented in anyway...)

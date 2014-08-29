VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "MSR605 Utility - Xyl2k!"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Top             =   4080
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Disco 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   600
      Top             =   0
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Get firmware version"
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Read raw data"
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Ram test"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Sensor test"
      Height          =   375
      Left            =   4800
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Get device model"
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Communication test"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Reset MSR"
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   4800
      TabIndex        =   10
      Top             =   360
      Width           =   3495
      Begin VB.OptionButton Option9 
         Caption         =   "Track 1 + Track 2 + Track 3"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Track 2 + Track 3"
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Track 1 + Track 3"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Track 1 + Track 2"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Track 3"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Track 2"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Track 1"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Eraze"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&WRITE"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&READ"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Low-Co"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hi-Co"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&DISCO !"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   4455
   End
   Begin VB.TextBox Track3 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Track2 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Track1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MSR Response:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Track 3:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Track 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Track 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   615
   End
   Begin VB.Label stat 
      Alignment       =   2  'Center
      Caption         =   "Ready"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public stuff As String
Dim disc0 As Integer

Private Sub Command1_Click() ' lecture de la carte
stuff = 1
Text1.Text = ""
Track1.Text = ""
Track2.Text = ""
Track3.Text = ""
MSComm1.Output = Chr(&H1B) + Chr(&H72)
stat.Caption = "SWIPE THE CARD !!!"
End Sub

Private Sub Command10_Click()
Text1.Text = ""
MSComm1.Output = Chr(&H1B) + Chr(&H6D)
stat.Caption = "SWIPE THE CARD !!!"
End Sub

Private Sub Command11_Click()
Text1.Text = ""
MSComm1.Output = Chr(&H1B) + Chr(&H76)
End Sub

Private Sub Command12_Click()
Unload Me
End Sub

Private Sub Command2_Click() 'check des tracks + écriture de la carte
If Track1.Text = "" Then
stat.Caption = "TRACK1 EMPTY !!!"
Else
If Track2.Text = "" Then
stat.Caption = "TRACK2 EMPTY !!!"
Else
If Track3.Text = "" Then
stat.Caption = "TRACK3 EMPTY !!!"
Else
stuff = 2
Text1.Text = ""
Track1.Locked = True
Track2.Locked = True
Track3.Locked = True
MSComm1.Output = Chr(&H1B) + Chr(&H77) + Chr(&H1B) + Chr(&H73) + Chr(&H1B) + Chr(&H1) + Track1.Text + Chr(&H1B) + Chr(&H2) + Track2.Text + Chr(&H1B) + Chr(&H3) + Track3.Text + Chr(&H3F) + Chr(&H1C)
stuff = 3
stat.Caption = "SWIPE THE CARD !!!"
End If
End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "&DISCO !" Then
Aleatoire
Disco.Enabled = True
Command3.Caption = "STOP !"
Else
If Command3.Caption = "STOP !" Then
Disco.Enabled = False
Command3.Caption = "&DISCO !"
MSComm1.Output = Chr(&H1B) + Chr(&H83)
End If
End If
End Sub

Private Sub Aleatoire()
Dim A, B As Integer 'vrai fausse fonction d'aléatoire !
A = 5
B = 1
disc0 = Int((A * Rnd) + B)
End Sub

Private Sub Command4_Click()
Text1.Text = ""
If Option3.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "000"
Else
End If
If Option4.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "010"
Else
End If
If Option5.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "100"
Else
End If
If Option6.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "011"
Else
End If
If Option7.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "101"
Else
End If
If Option8.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "110"
Else
End If
If Option9.Value = True Then
MSComm1.Output = Chr(&H1B) + Chr(&H63) + "111"
Else
End If
End Sub

Private Sub Command5_Click()
MSComm1.Output = Chr(&H1B) + Chr(&H61)
stat.Caption = "Reseted !"
End Sub

Private Sub Command6_Click()
Text1.Text = ""
MSComm1.Output = Chr(&H1B) + Chr(&H65)
End Sub

Private Sub Command7_Click()
Text1.Text = ""
MSComm1.Output = Chr(&H1B) + Chr(&H74)
End Sub

Private Sub Command8_Click()
Text1.Text = ""
stat.Caption = "Try to swipe a card or reset the MSR now !"
MSComm1.Output = Chr(&H1B) + Chr(&H86)
End Sub

Private Sub Command9_Click()
Text1.Text = ""
MSComm1.Output = Chr(&H1B) + Chr(&H87)
End Sub

Private Sub Disco_Timer()
   Select Case disc0 'feature a la con qui allume et éteint de façon aléatoire les ptites leds de la MSR
      Case 1
         MSComm1.Output = Chr(&H1B) + Chr(&H81) 'All LED off
         Aleatoire
      Case 2
         MSComm1.Output = Chr(&H1B) + Chr(&H82) 'All LED on
         Aleatoire
      Case 3
         MSComm1.Output = Chr(&H1B) + Chr(&H83) 'GREEN LED on
         Aleatoire
      Case 4
         MSComm1.Output = Chr(&H1B) + Chr(&H84) 'YELLOW LED on
         Aleatoire
      Case 5
         MSComm1.Output = Chr(&H1B) + Chr(&H85) 'RED LED on
         Aleatoire
   End Select
End Sub

Private Sub Option1_Click()
MSComm1.Output = Chr(&H1B) + Chr(&H78)
Text1.Text = ""
End Sub

Private Sub Option2_Click()
MSComm1.Output = Chr(&H1B) + Chr(&H79)
Text1.Text = ""
End Sub

Private Sub Text1_Change() 'Gestion des réponses de la MSR
If Text1.Text = Chr(&H1B) + "0" Then
stat.Caption = "OK !!!"
Else
End If
If Text1.Text = Chr(&H1B) + "1" Then
stat.Caption = "Read error !!!"
Else
End If
If Text1.Text = Chr(&H1B) + "2" Then
stat.Caption = "Command format error !!!"
Else
End If
If Text1.Text = Chr(&H1B) + "4" Then
stat.Caption = "Invalid command !!!"
Else
End If
If Text1.Text = Chr(&H1B) + "9" Then
stat.Caption = "Invalid card swipe when in write mode !!!"
Else
End If
If Text1.Text = Chr(&H1B) + "h" Then 'Set Hi-Co
Option1.Value = True
Else
End If
If Text1.Text = Chr(&H1B) + "l" Then 'Set Low-Co
Option2.Value = True
Else
End If
If Text1.Text = Chr(&H1B) + "y" Then 'Communication test
stat.Caption = "MSR OK !"
Else
End If
If Text1.Text = Chr(&H1B) + "A" Then
stat.Caption = "RAM error !!!"
Else
End If
If Text1.Text = Chr(&H1B) + "1S" Then
stat.Caption = "MSR206-1" 'Track 2
Else
End If
If Text1.Text = Chr(&H1B) + "2S" Then
stat.Caption = "MSR206-2" 'Track 2 & 3
Else
End If
If Text1.Text = Chr(&H1B) + "3S" Then
stat.Caption = "MSR206-3" 'Track 1,2 & 3
Else
End If
If Text1.Text = Chr(&H1B) + "5S" Then
stat.Caption = "MSR206-5" 'Track 1 & 2
Else
End If
Select Case stuff
Case Is = 1 'on parse
Track1.Text = funcParseStringFromString2String(Text1, "%", "?")
Track2.Text = funcParseStringFromString2String(Text1, ";", "?")
Track3.Text = funcParseStringFromString2String(Text1, Chr(&H3) + ";", "??")
Case Is = 2
End Select
Track1.Locked = False 'déverrouille
Track2.Locked = False
Track3.Locked = False
End Sub

Private Sub Track1_KeyPress(KeyAscii As Integer) 'Majuscule seulement sur le track1
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Track2_KeyPress(KeyAscii As Integer) 'chiffres seulement sur le track2
If (KeyAscii <> vbKeyBack) And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub Track3_KeyPress(KeyAscii As Integer) 'chiffres seulement sur le track3
If (KeyAscii <> vbKeyBack) And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub


Private Sub Form_Load()
On Error GoTo err
With MSComm1
.CommPort = 3 'on utilise le port COM3
.Handshaking = 0
.RThreshold = 1
.RTSEnable = True
.Settings = "9600,n,8,1"
.SThreshold = 1
.PortOpen = True
End With
MSComm1.Output = Chr(&H1B) + Chr(&H64)
Exit Sub
err:
stat.Caption = "MSR Not found on COM3 !"
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Command6.Enabled = True Then
MSComm1.PortOpen = False 'on ferme le port quand l'appli quitte
Else
End If
End Sub

Private Sub MSComm1_OnComm()
Dim Tampon As String

Select Case MSComm1.CommEvent
' On effectue la gestion des erreurs (cf. le modèle ci-dessus)
' Ici, on gère en fait pas grand-chose, mais c'est pour illustrer la démarche ;)

'liste des erreurs possibles
Case comEventBreak 'On a reçu un signal d’interruption (Break)
Case comEventCDTO ' Timeout de la porteuse
Case comEventCTSTO ' Timeout du signal CTS (Clear To Send)
Case comEventDSRTO ' Timeout du signal de réception
Case comEventFrame ' Erreur de trame
Case comEventOverrun ' Des données ont été perdues
Case comEventRxOver ' Tampon de réception saturé
Case comEventRxParity ' Erreur de parité
Case comEventTxFull ' Tampon d’envoi saturé
Case comEventDCB ' Erreur de réception DCB (jamais vu)

'liste des événements possibles qui sont, eux, normaux
Case comEvCD 'Changement dans la broche CD (porteuse)
Case comEvCTS 'Changement dans broche CTS
Case comEvDSR 'Changement dans broche DSR (réception)
Case comEvRing 'Changement dans broche RING (sonnerie)

'Chouette! on a reçu des données :)
Case comEvReceive
      Tampon = MSComm1.Input
      Call Traitement(Tampon) 'traitement données

Case comEvSend ' il y a des caractères à envoyer

Case comEvEOF 'on a reçu le caractère EOF
End Select
End Sub

Sub Traitement(Chaine As String)
'cette procédure sert à traiter l’information reçue dans le tampon
     Text1.SelStart = Len(Text1.Text)
     Text1.SelText = Chaine 'ici, on affiche le résultat dans un champ de texte
End Sub

Function funcParseStringFromString2String(sSourceString, sString1 As String, sString2 As String, Optional fCaseCaseInsensitive As Boolean = False) As String
 Dim sOutput As String
 Dim iLocationOfString1 As Long
 Dim iLocationOfString2 As Long
 Dim iCompareStyle As Long
 If fCaseCaseInsensitive Then
iCompareStyle = vbTextCompare
 Else
iCompareStyle = vbBinaryCompare
 End If
 sOutput = sSourceString
 iLocationOfString1 = InStr(1, sOutput, sString1, iCompareStyle)
 iLocationOfString2 = InStr(1, sOutput, sString2, iCompareStyle)
 If iLocationOfString1 = 0 And iLocationOfString2 = 0 Then
'non trouvé
sOutput = ""
 Else
If Len(sString1) = 0 And Len(sString2) = 0 Then
 'fais rien
ElseIf Len(sString1) = 0 Then
 If iLocationOfString2 <> 0 Then
sOutput = Mid(sOutput, 1, iLocationOfString2 - 1)
 End If
ElseIf Len(sString2) = 0 Then
 sOutput = Mid(sOutput, iLocationOfString1 + Len(sString1))
Else
 'coupe la premiere part
 If iLocationOfString1 <> 0 Then
sOutput = Mid(sOutput, iLocationOfString1 + Len(sString1))
 End If
 'coupe la derniere part
 iLocationOfString2 = InStr(1, sOutput, sString2, iCompareStyle)
 If iLocationOfString2 <> 0 Then
sOutput = Mid(sOutput, 1, iLocationOfString2 - 1)
 End If
End If
 End If
 funcParseStringFromString2String = sOutput
End Function

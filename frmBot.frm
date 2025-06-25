VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AI Bot"
   ClientHeight    =   4335
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5535
   Icon            =   "frmBot.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmBot.frx":0442
   ScaleHeight     =   4335
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Online:"
      Height          =   1095
      Left            =   4200
      TabIndex        =   23
      Top             =   120
      Width           =   1215
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Day(s): "
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Hour(s): "
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Minute(s): "
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblMin 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblHour 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblDay 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox Human 
      Height          =   405
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Computer 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Data datWords 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "nlp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Words"
      Top             =   4440
      Width           =   2700
   End
   Begin VB.PictureBox picprog 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleMode       =   0  'User
      ScaleWidth      =   44.473
      TabIndex        =   20
      Top             =   5760
      Width           =   2655
      Begin VB.Shape pbsp 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.CommandButton cmdturnl 
      Caption         =   "Turn >"
      Height          =   495
      Left            =   1800
      TabIndex        =   18
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdturnr 
      Caption         =   "< Turn"
      Height          =   495
      Left            =   1800
      TabIndex        =   17
      Top             =   3720
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Caption         =   "System"
      Height          =   1575
      Left            =   4200
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
      Begin VB.CheckBox chkTalk 
         Caption         =   "Respond"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox chkServtxt 
         Caption         =   "SText"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkWhisp 
         Caption         =   "Whispers"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkFollow 
         Caption         =   "Follow"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox chkServCode 
         Caption         =   "SCode"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmduse 
      Caption         =   "&Use"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdWho 
      Caption         =   "&Who"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdLay 
      Caption         =   "&Lay"
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdSE 
      Caption         =   "SE"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdSW 
      Caption         =   "SW"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdNE 
      Caption         =   "NE"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdNW 
      Caption         =   "NW"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Timer StayOnline 
      Interval        =   60000
      Left            =   120
      Top             =   600
   End
   Begin MSWinsockLib.Winsock sckFurc 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox txtFromFurc 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sign As String
Dim lastwalk As String
Dim whatwalk As String
Dim hit
Public Day, Hour, Minute As Integer
Public onet
Public twot
Public Desc As String
Public Connected As Boolean
'Bot Settings
Const BotName = "Cidaok"
Const BotPass = "0519aa"
Const descrip = "#SP"
Const ColorCode = "577)+<===< #!!#!"

Private Const UnwantedCharacters = "~`!@#$%^&*()-_=+\|[{]};:'"",<.>/?"

Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL = 1

Private Sub chkServCode_Click()
If chkServCode = 1 Then
chkServtxt = 2
chkServtxt.Enabled = False
End If
If chkServCode = 0 Then
chkServtxt = 1
chkServtxt.Enabled = True
End If
End Sub
Private Sub cmdGet_Click()
If Connected = True Then sckFurc.SendData "get" & vbLf
End Sub
Private Sub cmdlie_Click()
If Connected = True Then sckFurc.SendData "lie" & vbLf
End Sub
Private Sub cmdNE_Click()
If Connected = True Then sckFurc.SendData "m 9" & vbLf
End Sub
Private Sub cmdNW_Click()
If Connected = True Then sckFurc.SendData "m 7" & vbLf
End Sub
Private Sub cmdSE_Click()
If Connected = True Then sckFurc.SendData "m 3" & vbLf
End Sub
Private Sub cmdSW_Click()
If Connected = True Then sckFurc.SendData "m 1" & vbLf
End Sub
Private Sub cmdturnl_Click()
If Connected = True Then sckFurc.SendData ">" & vbLf
End Sub
Private Sub cmdturnr_Click()
If Connected = True Then sckFurc.SendData "<" & vbLf
End Sub
Private Sub cmduse_Click()
If Connected = True Then sckFurc.SendData "use" & vbLf
End Sub
Private Sub cmdWho_Click()
If Connected = True Then sckFurc.SendData "who" & vbLf
End Sub
Sub Form_Load()
Minute = 0
Hour = 0
Day = 0
Desc = descrip & " [Online: 0 Day(s), 0 Hour(s), 0 Minute(s)]"
Connected = False
End Sub
Private Sub cmdConnect_Click()
If Connected = False Then
sckFurc.RemoteHost = "64.191.51.88"
sckFurc.RemotePort = "6000"
sckFurc.Connect
Connected = True
lastwalk = "none"
End If
End Sub
Private Sub cmdDisconnect_Click()
If Connected = True Then
sckFurc.Close
Connected = False
End If
End Sub

Private Sub sckFurc_DataArrival(ByVal bytesTotal As Long)
Dim s As String
sckFurc.GetData s
X = Split(s, vbLf)
For r = 0 To UBound(X) - 1
RealText X(r)
Next
End Sub
Sub RealText(Txt)
On Error Resume Next
Dim dirtySentence As String

If chkServtxt.Value = Checked Or chkServtxt.Enabled = False Then
If chkServCode.Value = Checked Then txtFromFurc = txtFromFurc & Txt & vbCrLf
If chkServCode.Value = Unchecked Then
If Left(Txt, 1) = "(" Then txtFromFurc = txtFromFurc & Right(Txt, Len(Txt) - 1) & vbCrLf
End If
End If
If Txt = "END" Then
Connected = True
sckFurc.SendData "connect " & BotName & " " & BotPass & vbLf & "color " & ColorCode & vbLf & "desc " & Desc & vbLf
End If
If Txt = "]ccmarbled.pcx" Then
sckFurc.SendData "vascodagama" & vbLf
End If

'(keny: hello
If Left(Txt, 1) = "(" Then
If Right(Txt, 1) <> ")" Then
cmsg = Split(Txt, ": ", 2)
dirtySentence = cmsg(1)
cleanSentence = RemoveUnwantedCharacters(dirtySentence, UnwantedCharacters)
dirtySentence = cleanSentence
cleanSentence = RemoveExtraSpaces(dirtySentence)

'sckFurc.SendData Chr(34) & cleanSentence & vbLf
chat cleanSentence
End If
End If

If chkWhisp.Value = Checked Then
If Left(Txt, 3) = "([ " And Right(Txt, 10) = " to you. ]" Then
    tmsg = Split(Txt, " whispers, " & Chr(34), 2)
    Furre = Right(tmsg(0), Len(tmsg(0)) - 3)
    NMsg = Left(tmsg(1), Len(tmsg(1)) - 11)
    Msg = LCase(NMsg)
    DoWhisper Furre, Msg
End If
End If 'chkWhisp

'make the bot follow it owner
If chkFollow.Value = Checked Then
If Left(Txt, 11) = Chr(34) & "87.-6*<48" Then
        frl = Mid(Txt, 17, Len(Txt) - 0)
        whatwalk = Mid(frl, 1, Len(frl) - 4)
        'whatwalk = LCase(wwalk)
    dowalk whatwalk, lastwalk
End If
End If 'chkFollow



End Sub


Sub chat(Human)
Dim words(20) As String
On Error Resume Next
Computer = ""
Human = LCase(Trim(Human))

Human = "^" & Human & "^"
Human = " " & Human & " "
Start = 1
cword = 1
For a = 1 To Len(Human)
    If Mid(Human, a, 1) = " " And a > 1 Then
        words(cword) = Mid(Human, Start + 1, a - Start)
        If Len(words(cword)) > Len(maxword) Then maxword = Trim(words(cword))

        Start = a
        cword = cword + 1
        
    End If
Next a
rf = "'adkevriy'"
For a = 1 To cword - 1

If Trim(words(a)) <> "" Then
datWords.Recordset.AddNew

middle = Trim(words(a))
rf = rf & ",'" & middle & "'"
'For b = 0 To 4
'    If optrate(b).Value = True Then Exit For
'Next b

If a > 1 Then Previous = Trim(words(a - 1))
If a < cword Then nextT = Trim(words(a + 1))
datWords.Recordset.FindFirst ("middle = '" & Trim(words(a)) & "' and previous = '" & Trim(words(a - 1)) & "' and next = '" & Trim(words(a + 1)) & "'")
If datWords.Recordset.NoMatch = True Then
    datWords.Recordset.AddNew
    datWords.Recordset.Fields(0) = middle
    datWords.Recordset.Fields(1) = Previous
    datWords.Recordset.Fields(2) = nextT
    datWords.Recordset.Update
    
End If
End If
Next a
Human = ""

Set dbs = OpenDatabase(App.Path & "\nlp.mdb")
Set tdf = dbs.OpenRecordset("select middle,count(middle) from words where middle in (" & rf & ") group by middle order by count(middle)")
maxword = tdf.Fields(0)

'
datWords.Recordset.MoveFirst
'
datWords.Recordset.FindFirst "middle = '" & maxword & "'"

Computer = maxword
wrd = choose(maxword, False)

Do While wrd <> ""


'
nextword = datWords.Recordset.Fields(1)
Computer = wrd & " " & Computer

wrd = choose(wrd, False)

'
num = (Int((datWords.Recordset.RecordCount * Rnd) + 1))
'
datWords.Recordset.MoveFirst
'
datWords.Recordset.Move num

'
datWords.Recordset.FindNext "middle = '" & nextword & "'"
'
If datWords.Recordset.NoMatch = True Then datWords.Recordset.FindLast "middle = '" & nextword & "'"



Loop
'
datWords.Recordset.MoveFirst
'
datWords.Recordset.FindFirst "middle = '" & maxword & "'"
wrd = choose(maxword, True)

Do While wrd <> ""

Computer = Computer & " " & wrd

wrd = choose(wrd, True)

'Do While datWords.Recordset.Fields(2) <> ""
'
nextword = datWords.Recordset.Fields(2)
'Computer = Computer & " " & nextword



'
num = (Int((datWords.Recordset.RecordCount * Rnd) + 1))
'
datWords.Recordset.MoveFirst
'
datWords.Recordset.Move num

'
datWords.Recordset.FindNext "middle = '" & nextword & "'"
'
If datWords.Recordset.NoMatch = True Then datWords.Recordset.FindLast "middle = '" & nextword & "'"


Loop
Computer = Mid(Computer, 2, Len(Computer) - 2)
comp = Computer
Computer = ""
picprog.ScaleWidth = Len(comp)
For c = 1 To Len(comp)
    Computer = Computer & Mid(comp, c, 1)
    For a = 1 To Int(50000 * Rnd) + 5000
        stuff = 5 * 5 * 5 * 5 * 5 * 5 * 2
    Next a
    Computer.Refresh
    pbsp.Width = c
    picprog.Refresh
Next c
Human.SetFocus
pbsp.Width = 0

If chkTalk.Value = Checked Then
sckFurc.SendData Chr(34) & Computer & vbLf
End If

End Sub

Private Function RemoveUnwantedCharacters(From As String, What As String) As String
  
    From = Replace(From, "#SA", " ")
    From = Replace(From, "#SB", " ")
    From = Replace(From, "#SC", " ")
    From = Replace(From, "#SD", " ")
    From = Replace(From, "#SE", " ")
    From = Replace(From, "#SF", " ")
    From = Replace(From, "#SM", " ")
    From = Replace(From, "#SN", " ")
    From = Replace(From, "#SG", " ")
    From = Replace(From, "#SH", " ")
    From = Replace(From, "#SI", " ")
    From = Replace(From, "#SJ", " ")
    From = Replace(From, "#SK", " ")
    From = Replace(From, "#SL", " ")
    From = Replace(From, "#SO", " ")
    From = Replace(From, "#SP", " ")
    
    RemoveUnwantedCharacters = From
    
    For jVar = 1 To Len(What)
    
        RemoveUnwantedCharacters = Replace(RemoveUnwantedCharacters, Mid$(What, jVar, 1), "")
                
    Next jVar
   
End Function

Public Function RemoveExtraSpaces(TheString As String) As String

    Dim LastChar As String
    Dim NextChar As String
    LastChar = Left(TheString, 1)
    RemoveExtraSpaces = LastChar

    For i = 2 To Len(TheString)
    NextChar = Mid(TheString, i, 1)


    If NextChar = " " And LastChar = " " Then
    Else
        RemoveExtraSpaces = RemoveExtraSpaces & NextChar

End If

LastChar = NextChar
Next i

End Function

Public Function choose(word, forward As Boolean)
Set dbs = OpenDatabase(App.Path & "\nlp.mdb")
Set tdf = dbs.OpenRecordset("select * from words where middle = '" & word & "'")
rf = "'vkuseyvgwzelkbzwle'"
Do Until tdf.EOF
    chkstr = IIf(forward = True, tdf!Next, tdf!Previous)
    If InStr(1, Computer, chkstr) = 0 Or Len(rf) < 22 Then
        rf = rf & ",'" & IIf(forward = True, tdf!Next, tdf!Previous) & "'"
    
    End If
    tdf.MoveNext
Loop

    'New search routine - unstable and commented out for 1.1 bugfix - fixed 1.2 01/08/00
    Set rare = dbs.OpenRecordset("select middle,count(middle) from words where middle in (" & rf & ") and " & _
    IIf(forward = True, "previous = '" & word & "'", "next = '" & word & "'") & " group by middle order by count(middle)")
    'Set rare = dbs.OpenRecordset("select middle,count(middle) from words where middle in (" & rf & ") group by middle order by count(middle)")

If rare.EOF = True Then
    choose = ""
Else
    Do Until rare.EOF
        If Int((2 * Rnd) + 1) = 1 Then
            choose = rare!middle
            Exit Function
        Else
            rare.MoveNext
        End If
    Loop
    rare.MoveFirst
    choose = rare!middle
End If


End Function

Sub DoWhisper(Furre, Msg)

If Msg Like "help" Then
whspnum = 1
Else
whspnum = 0
End If
If whspnum = 0 Then sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
If whspnum = 1 Then sckFurc.SendData "wh " & Furre & " I dont understand. Try Whispering me help." & vbLf
End Sub

Sub dowalk(whatwalk, lastwalk)
If (lastwalk = "j") Or (lastwalk = "k") Or (lastwalk = "l") Then ' moveing up
    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
        sckFurc.SendData "m 9" & vbLf
        Else ' move left
            If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 7" & vbLf
        Else 'move right
            If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 3" & vbLf
        Else ' move down
            If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
            sckFurc.SendData "m 7" & vbLf
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 9" & vbLf
            sckFurc.SendData "m 3" & vbLf
        End If
        End If
        End If
    End If
lastwalk = whatwalk
End If
If (lastwalk = "f") Or (lastwalk = "g") Or (lastwalk = "h") Then 'moveing left
        If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
        sckFurc.SendData "m 7" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 9" & vbLf
                Else 'move down
                    If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                Else 'move right
                If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
                    sckFurc.SendData "m 9" & vbLf
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "b") Or (lastwalk = "c") Or (lastwalk = "d") Then 'moveing right
        If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
        sckFurc.SendData "m 3" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 9" & vbLf
                Else 'move down
                    If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                Else 'move left
                    If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
                    sckFurc.SendData "m 9" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "`") Or (lastwalk = "_") Or (lastwalk = "^") Then 'moveing down
        If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
        sckFurc.SendData "m 1" & vbLf
                Else 'move left
                    If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 7" & vbLf
                Else 'move right
                    If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                Else 'move up
                    If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
                    sckFurc.SendData "m 7" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 1" & vbLf
                    sckFurc.SendData "m 3" & vbLf
                End If
                End If
                End If
        End If
lastwalk = whatwalk
End If
If (lastwalk = "none") Then
        If (whatwalk = "`") Or (whatwalk = "_") Or (whatwalk = "^") Then
            sckFurc.SendData "m 1" & vbLf
        Else 'move left
            If (whatwalk = "f") Or (whatwalk = "g") Or (whatwalk = "h") Then
            sckFurc.SendData "m 7" & vbLf
        Else 'move right
            If (whatwalk = "b") Or (whatwalk = "c") Or (whatwalk = "d") Then
            sckFurc.SendData "m 3" & vbLf
        Else 'move left
            If (whatwalk = "j") Or (whatwalk = "k") Or (whatwalk = "l") Then
            sckFurc.SendData "m 9" & vbLf
        End If
        End If
        End If
        End If
lastwalk = whatwalk
End If
End Sub


Private Sub cmdExit_Click()
End
End Sub

Private Sub StayOnline_Timer()
If Connected = True Then
Minute = Minute + 1

If Minute = 60 Then
Hour = Hour + 1
Minute = 0
End If

If Hour = 24 Then
Hour = 0
Day = Day + 1
End If

lblDay.Caption = Day
lblHour.Caption = Hour
lblMin.Caption = Minute

sckFurc.SendData "desc " & descrip & " [Online: " & Day & " Day(s), " & Hour & " Hour(s), " & Minute & " Minute(s)]" & vbLf
End If
End Sub


Private Sub txtFromFurc_Change()
txtFromFurc.SelStart = Len(txtFromFurc)
If Len(txtFromFurc) > 10000 Then txtFromFurc = Right(txtFromFurc, 9000)
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sckFurc.SendData txtSend & vbLf
    txtSend = ""
    KeyAscii = 0
End If
End Sub

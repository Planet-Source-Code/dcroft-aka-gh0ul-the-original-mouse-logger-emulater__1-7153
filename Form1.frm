VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cursor Logger v1.75"
   ClientHeight    =   3060
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   4416
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4416
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   252
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Width           =   972
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2412
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   4255
      _Version        =   393216
      Tab             =   1
      TabHeight       =   420
      TabCaption(0)   =   "About"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtAbout"
      Tab(0).Control(1)=   "Shape3"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "New"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Timer2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Timer1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Play Previous"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdBrowse"
      Tab(2).Control(1)=   "txtOpenFile"
      Tab(2).Control(2)=   "cmdPlay"
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(4)=   "Shape1"
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1692
         Left            =   -74640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   480
         Width           =   3492
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   252
         Left            =   -71400
         TabIndex        =   12
         Top             =   720
         Width           =   372
      End
      Begin VB.TextBox txtOpenFile 
         Height          =   288
         Left            =   -74760
         TabIndex        =   10
         Top             =   720
         Width           =   3252
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play >"
         Height          =   252
         Left            =   -74640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1440
         Width           =   852
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Log"
         Height          =   252
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   972
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop Log"
         Height          =   252
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   972
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Replay"
         Height          =   252
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   2052
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4440
         Top             =   720
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4320
         Top             =   480
      End
      Begin VB.Shape Shape3 
         Height          =   1932
         Left            =   -74760
         Top             =   360
         Width           =   3732
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cursor file name:"
         Height          =   192
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   1176
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         Height          =   492
         Left            =   -74760
         Top             =   1320
         Width           =   3732
      End
      Begin VB.Label Label1 
         Caption         =   "Log the mouse (X,Y) coordinates, as well as the click events sent to windows then replay the movements and events.  "
         Height          =   732
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3492
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(X,Y)"
         Height          =   252
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         Height          =   252
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label Label5 
         Caption         =   "Cursor click..."
         Height          =   252
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   2412
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000C&
         BackStyle       =   1  'Opaque
         Height          =   852
         Left            =   240
         Top             =   1080
         Width           =   3732
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' '///////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' '
' ' ' Thanks to ~$[]D[][]V[][]D$~ for solving the major
' ' ' and most frustrating bug in this code. No more slow dragging
' ' ' on the scrollbars..... no you can record and playback
' ' ' every click, doubleclick, right click, drag, and even select
' ' ' text, Cut, Copy, and paste faster than ever. And replay perfectly.
' ' '
' '   gh0ul 2000
' '
' '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\//////////////////////////////

Private Sub cmdPlay_Click()
    
    cmdPlay.Enabled = False
    
    ' play this file
    If LCase(Right(txtOpenFile, 3)) <> "dat" Then
       MsgBox "Invalid File Type!!", vbCritical
       Exit Sub
    End If
    
    ' get the size to make the array
    Open txtOpenFile For Binary As #1
       Get #1, , cp
    Close #1
    
    ' now resize the array, preserving the last index
    ReDim Preserve cp(cp(0).last_idx)
    
    ' sized up the array, now fill in the data
    Open txtOpenFile For Binary As #1
       Get #1, , cp
    Close #1
    
    ' to ensure start at begining
    idx& = 0
    
    ' show the replay
    Timer2.Enabled = True
    
End Sub



Private Sub Command1_Click()
    ' Start Log
    idx& = 0
    Timer1.Enabled = True
    Command2.Enabled = True
    
End Sub

Private Sub Command2_Click()
    ' Stop Log
    Dim rv As String
    Dim SaveFN As String
    
    Timer1.Enabled = False
    
    ' save the Coordinates
    ' // call the open Procedure
    SaveFN$ = Save_File(hWnd)
    
    ' // Check the return value
    If SaveFN$ <> "" Then
       Fname$ = SaveFN$
    Else
       Fname$ = Fname$
       Command3.Enabled = True
       Exit Sub
    End If
    
    rv$ = Dir$(Fname$)
    If rv$ <> "" Then Kill Fname$
    
    ' store the last index reached in the first array position.
    ' since the data is stored and read from the 1 position
    ' the 0 position is left untouched, fill it with the last index.
    cp(0).last_idx& = idx&
        
    Open Fname$ For Binary As #1
        Put #1, , cp
    Close #1
    
    Command3.Enabled = True
End Sub

Private Sub Command3_Click()
    ' get the size to make the array
    Open Fname$ For Binary As #1
       Get #1, , cp
    Close #1
    
    ' // the array may have previously been destroyed so
    ' // reinitialize it as if it were empty.
    
    ' resize the array, preserving the last index
    ReDim Preserve cp(cp(0).last_idx)
    
    ' sized up the array, now fill in the data
    Open Fname$ For Binary As #1
       Get #1, , cp
    Close #1
    
    idx = 0
    
    ' show the replay
    Timer2.Enabled = True
End Sub


Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    Dim rvFileName As String
    ' // call the open Procedure
    rvFileName$ = Open_File(hWnd)
    
    ' // Check the return value
    If rvFileName$ <> "" Then
       txtOpenFile = rvFileName$
    Else
       txtOpenFile = ""
    End If
End Sub

Private Sub Form_Load()
    ''
    Command3.Enabled = False
    Command2.Enabled = False
    
    ' re-dimension the Mouse info Array
    ' the 0 position will be reserved to save the
    ' size of the array or the last index in the array
    ReDim cp(0)
        
    If txtOpenFile = "" Then
       cmdPlay.Enabled = False
    Else
       cmdPlay.Enabled = True
    End If
    
    DOAbout
End Sub



Private Sub Form_Unload(Cancel As Integer)
    '
    ReDim cp(0)
    End
End Sub

Private Sub Timer1_Timer()
     Dim pt As POINTAPI  ' receives coordinate points of the cursor
    Dim rv As Long  ' return value
    Dim WinText As String
    
    Static DontCount As Long
    DontCount& = DontCount& + 1
    
    idx& = idx& + 1
    Label3 = Format(idx&, "#00:00:#")
    If Len(Label3) > 2 Then Label3 = Format(idx&, "#00:0:#")
    
    
    ' read cursor location
    rv& = GetCursorPos(pt)
    
    ' Create array members
    ReDim Preserve cp(idx)
    ' save the coordinates
    cp(idx).x = pt.x
    cp(idx).y = pt.y
            
    ' report the coordinates
    Label5 = "(" & pt.x & "," & pt.y & ")"
    
    
    WinText$ = Space$(255)
    ' get the text of the window the mouse is over
    WinText$ = Left$(WinText$, GetWindowText(GetForegroundWindow, ByVal WinText$, 255))
                
    ' if it's the first time through then don't count the Left
    ' button down
    If DontCount < 50 Then
       Exit Sub
    Else
       'DontCount = 75
    End If
    ' check to see if a click event was sent
    If GetAsyncKeyState(VK_LBUTTON) Then
        ' Left Mouse button down
        With cp(idx)
           ' keep the button down until it is released
          .bDragging = True
          .click_x = pt.x
          .click_y = pt.y
          .Button = BT_LEFT
          .dblClicked = False
        End With
    ElseIf GetAsyncKeyState(VK_RBUTTON) Then
        ' Right mouse button down
        With cp(idx)
          .click_x = pt.x
          .click_y = pt.y
          .Button = BT_RIGHT
          .dblClicked = False
        End With
        
    Else
         ' released the mouse button
         
        ' Holding& = 0
         cp(idx).bDragging = False
         
    End If
    
    DoEvents
    
    
End Sub





Private Sub Timer2_Timer()
    
    
    ' move through the array repeating the cursor positions
    idx& = idx& + 1
    
    
    ' if the last member of the array has been
    ' met stop and return the cursor
    If idx& > cp(0).last_idx& Then
       Timer2.Enabled = False
       idx& = 0
       Form_Load
       Exit Sub
    End If
    
    Label3 = Format(idx&, "#00:00:#")
    If Len(Label3) > 2 Then Label3 = Format(idx&, "#00:0:#")
    
    ' replay the coordinates
    Call SetCursorPos(cp(idx&).x, cp(idx&).y)
    
    Label5 = "(" & cp(idx&).x & "," & cp(idx&).y & ")"
             
    ' was there a click Event here
    If cp(idx).click_x <> 0 And cp(idx).click_y <> 0 Then
       ' must've logged a click
       ' attempt to re-click it
       ' click the button previously clicked
       If cp(idx).Button = BT_LEFT Then
          Static l As Long  ' counter for the left button
          
          ' only allow this code to be executed once
          ' if the user did not hold down the button
          l& = l& + 1
          
          ' if the first time through, and dragging
          ' then let it pass.
          If l& <= 2 And (cp(idx&).bDragging) Then
             ' continue allow to drag
          ElseIf l& > 1 And (Not cp(idx&).bDragging) Then
            Exit Sub
          End If
          
          ' if the user is draggin a scrollbar don't
          ' release the button
          If cp(idx&).bDragging Then
            
            ' if the button is already down, don't do it again.
            If GetAsyncKeyState(VK_LBUTTON) Then
              ' do nothing the button is pressed
              Debug.Print "Mouse Down; DRAGGING"
            Else
              ' ok to Press left button
              Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            End If
          
          ElseIf (Not cp(idx&).bDragging) Then
            ' Press and then release the left mouse button only
            ' a single click was recieved.
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            Debug.Print "Mouse Down; CLICKING"
          End If
       
       ElseIf cp(idx).Button = BT_RIGHT Then
          Static r As Long  ' counter for the right button
          ' only allow this code to be executed once
          r& = r& + 1
          If r& > 1 Then Exit Sub
          ' Press and then release the right mouse button.
           Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
           Call mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
         
       End If
             
    End If
       
    ' were we previously dragging??
    If idx& > 2 Then
      ' if the last index was dragging and this one is not
      ' make sure the button releases
      If cp(idx& - 1).bDragging And (Not cp(idx&).bDragging) Then
         ' make sure the button releases
         Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
         Debug.Print "Mouse Released"
      End If
    End If
    
    ' no button pressed
    If cp(idx).Button = 0 Then
      ' reset button press counter
      l& = 0
      r& = 0
    End If
    DoEvents
End Sub

Private Sub txtOpenFile_Change()
    
    cmdPlay.Enabled = True
End Sub


Private Sub DOAbout()
   
   Dim msg As String
   
   msg = vbCrLf
   msg = msg & "Cursor Logger v.1.75" & vbCrLf
   msg = msg & "-----------------------------------" & vbCrLf
   msg = msg & vbCrLf
   msg = msg & "To log the movements of the cursor select " & _
               "the ""New"" tab and click the ""Start Log"" " & _
               "button. The cursor position and click events " & _
               "will be logged." & vbCrLf & vbCrLf
   
   msg = msg & "To Replay the cursor movements immediatley " & _
               "Click the ""RePlay"" button in the ""New Tab""" & vbCrLf & vbCrLf
               
   msg = msg & "To play any valid cursor log file, Select the " & vbCrLf & _
               """Play Previous"" tab and enter the file name " & _
               "Click the ""Play"" button to play a file." & vbCrLf & vbCrLf
   
   msg = msg & "Improvements: Since v.1.25" & vbCrLf
   msg = msg & "----------------------------------" & vbCrLf
   msg = msg & vbCrLf
   msg = msg & "1.) Added ability to log and replay double clicks" & vbCrLf
   msg = msg & "2.) Added drag and drop, select text, drag scrollbar "
   msg = msg & "and move window." & vbCrLf
   msg = msg & "3.) Increased single click response  ." & vbCrLf & vbCrLf
   
   msg = msg & "Usage Hint:" & vbCrLf
   msg = msg & "----------------------------------" & vbCrLf
   msg = msg & vbCrLf
   msg = msg & "When dragging a scrollbar. Be sure to establish a solid click before dragging."
   msg = msg & vbCrLf
   msg = msg & "Author      :  DCroft" & vbCrLf
   msg = msg & "E-Mail      :  gh0ul@homail.com" & vbCrLf
   msg = msg & "ICQ#        :  31047555" & vbCrLf
   msg = msg & "OS\Platform :  WINNT 4.0\VB 5.0" & vbCrLf & vbCrLf
   msg = msg & vbCrLf
   msg = msg & "Best results when used as an exe." & vbCrLf & vbCrLf
   txtAbout = msg
End Sub

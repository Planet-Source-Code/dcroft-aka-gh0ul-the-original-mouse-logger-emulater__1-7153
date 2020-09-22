Attribute VB_Name = "modDeclares"
Option Explicit


' GetCursorPos
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
'
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public LastPoint As POINTAPI

' SetCurPos
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const BT_LEFT = &H30
Public Const BT_RIGHT = &H40


Public Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10




   ' // returned file name
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
' // Type declarations
Public Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

' type for storing the cursor positions and click events

Public Type CUR_STATS
    x As Long                 ' Current Mouse position
    y As Long                 ' Current Mouse Position
    click_x As Long           ' X Coor where clicked
    click_y As Long           ' Y Coor Where clicked
    Button As Long            ' Right or left button
    dblClicked As Boolean     ' True if double clicked
    last_idx As Long          ' the last index reached in the array
    bDragging As Boolean      ' true if the user is holding down
End Type                      ' the button to say, move a scroll
                              ' bar

Public cp() As CUR_STATS     ' an array of the above type
Public idx As Long           ' array index
Public last_idx As Long      ' the last recorded index
Public Fname As String       ' the name of the file that stores
                             ' the coordinates
Public bForward As Boolean



' open FIle Function
Function Open_File(hWnd As Long) As String
   '
   Dim OpenFileDialog As OPENFILENAME
   Dim rv As Long
   
   ' // init dialog
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "Cursor Log Files (*.dat)" + Chr$(0) + "*.dat" + _
      Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = CurDir
     .lpstrTitle = "Open Cursor Log File..."
     .flags = 0
   End With
  
   ' // call API to show the dialog that was just initialized
   rv& = GetOpenFileName(OpenFileDialog)
   
   If (rv&) Then
      Open_File = Trim$(OpenFileDialog.lpstrFile)
   Else
      Open_File = ""
   End If
   
End Function


Function Save_File(hWnd As Long) As String
   '
   Dim SaveFileDialog As OPENFILENAME
   Dim rv As Long
   
   ' // init dialog
   With SaveFileDialog
     .lStructSize = Len(SaveFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "Cursor Log Files (*.dat)" + Chr$(0) + "*.dat" + _
      Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = CurDir
     .lpstrTitle = "Save Cursor Log File..."
     .flags = 0
   End With
  
   ' // call API to show the dialog that was just initialized
   rv& = GetSaveFileName(SaveFileDialog)
   
   If (rv&) Then
      Save_File = Trim$(SaveFileDialog.lpstrFile)
   Else
      Save_File = ""
   End If
End Function



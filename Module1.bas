Attribute VB_Name = "modGlobal"
Option Explicit

' Win32 Declares '
'----------------'
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
'Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
'----------------'

Const PROCESS_ALL_ACCESS = &H1F0FFF ' Process Access
 
' Help to minimize/maximize window so can refresh the view '
'------------------------------------------------------------------------'
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type WINDOWPLACEMENT
    length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const SW_SHOW As Long = 5
Private Const SW_HIDE As Long = 0
'------------------------------------------------------------------------'

' Variables for attaching to process '
'------------------------------------'
Public Window_hWnd As Long      ' Handle Of Window
Dim ProcessId As Long           ' Process Id
Dim Process As Long             ' Process
'------------------------------------'

Public intVert As Long
Public intHor As Long

Public refresh_view As Boolean
Dim lastValue As String

' Transparent form              '
'-------------------------------'
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Private Const WS_EX_TOOLWINDOW As Long = &H80&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'-------------------------------'
    
Public Function TranslucentForm(frm As Form, TranslucenceLevel As Byte) As Boolean
    Dim lngRetVal As Long
    On Error Resume Next
    
    If TranslucenceLevel <> 255 Then
        SetWindowLong frm.hWnd, GWL_EXSTYLE, WS_EX_LAYERED Or WS_EX_TOOLWINDOW
        SetLayeredWindowAttributes frm.hWnd, 0, TranslucenceLevel, LWA_ALPHA
    Else
        lngRetVal = GetWindowLong(frm.hWnd, GWL_EXSTYLE)
        lngRetVal = lngRetVal - WS_EX_LAYERED Or WS_EX_TOOLWINDOW
        SetWindowLong frm.hWnd, GWL_EXSTYLE, lngRetVal
    End If
    
    TranslucentForm = Err.LastDllError = 0
End Function

Public Function RefreshWindow()
Dim wp As WINDOWPLACEMENT, i&
wp.length = Len(wp)
    For i = 0 To 1
        GetWindowPlacement Window_hWnd, wp
        If i = 0 Then wp.showCmd = SW_HIDE
        If i = 1 Then wp.showCmd = SW_SHOW
        SetWindowPlacement Window_hWnd, wp
    Next i
End Function

'____________________________________________________________________________
'MEMORY STAFF \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'Write Memory -------------
Public Function Poke(PokeAddress As Variant, PokeValue As Variant, typeSize As Long) As Long
    If PokeAddress = "" Then Exit Function
    Poke = WriteProcessMemory(Process, CLng(PokeAddress), ConvertNumberToString(CDbl(PokeValue)), typeSize, 0&)
End Function

'Read Memory -------------
Public Function Peek(PeekAddress As Variant, Buffer As Long, typeSize As Long) As Integer
    If PeekAddress = "" Then Exit Function
    Peek = ReadProcessMemory(Process, CLng(PeekAddress), Buffer, typeSize, 0&)
End Function

Public Sub Attach()
    Call GetWindowThreadProcessId(Window_hWnd, ProcessId) ' Get the process identifier
    Process = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessId) ' Get open handle to the process
End Sub

Private Function ConvertNumberToString(number As Double) As String
'converts number to string will be searched in memory
Dim b4&, b3, b2, b1
If number < 256 Then ConvertNumberToString = Chr(number): Exit Function

If number < 65536 Then
    ConvertNumberToString = Chr(number And 255) & Chr((number And 65280) / 256)
    Exit Function
End If

b4 = number And 255: number = Int(number / 256)
b3 = number And 255: number = Int(number / 256)
b2 = number And 255: number = Int(number / 256)
b1 = number And 255: number = Int(number / 256)

ConvertNumberToString = Chr(b4) & Chr(b3) & Chr(b2) & Chr(b1)

End Function
'_____________________________________________________________________________
'END MEMORY STAFF ////////////////////////////////////////////////////////////

Public Sub LoadTable()
    'local variables..
    
    Dim i As Integer, PeekValue As Long
    
    ResetAll
    
    'load width,height
    If Peek(&H10056AC, PeekValue, 4) <> 0 Then
        If intHor <> PeekValue Then refresh_view = True
        intHor = PeekValue
    Else
        Exit Sub
    End If
    If Peek(&H10056A8, PeekValue, 4) <> 0 Then
        If intVert <> PeekValue Then refresh_view = True
        intVert = PeekValue
    Else
        Exit Sub
    End If
    
    'check if new game started
    Peek &H1005164, PeekValue, 1
    If (PeekValue = 0) And (lastValue <> CStr(PeekValue)) Then refresh_view = True
    lastValue = PeekValue
    
With frmMain
    'start by loading (intVert * intHor) index's.
    
    For i = 1 To (intVert * intHor)
        Load .lblCell(i)
    Next i
    
    'lets place up the first one.
    
    .lblCell(1).Left = 100
    .lblCell(1).Top = 100
    .lblCell(1).Visible = True
    
    'thats our first one, lets place the rest of that row
    
    For i = 2 To intHor
        .lblCell(i).Left = .lblCell(i - 1).Left + .lblCell(i - 1).width - 15
        .lblCell(i).Top = .lblCell(i - 1).Top
        .lblCell(i).Visible = True
    Next i
    
    'thats our first row, lets repeat it (intVer - 1) times
    
    For i = (intHor + 1) To (intHor * intVert)
        If i Mod intHor = 1 Then
            'first one in row, so its under the one over it.
            .lblCell(i).Left = .lblCell(i - intHor).Left
            .lblCell(i).Top = .lblCell(i - intHor).Top + .lblCell(i - intHor).Height - 15
            .lblCell(i).Visible = True
        Else
            'not first, so its next to the one before it.
            .lblCell(i).Left = .lblCell(i - 1).Left + .lblCell(i - 1).width - 15
            .lblCell(i).Top = .lblCell(i - 1).Top
            .lblCell(i).Visible = True
        End If
    Next i
    
    'load from memory and fill celles
    
    Dim cell As Integer, start_line As Boolean
    start_line = False
    
    For i = 1 To 1000
        If cell = .lblCell.Count - 1 Then Exit For 'check for celles count to exit
        Peek CLng(&H100535F) + i, PeekValue, 1 'peek byte value from memory
        
        'if Auto Check Mines is enabled change memory bytes from Mine to Flag
        If PeekValue = 143 And .mnuAutoCheckMines.Checked Then _
            Poke CLng(&H100535F) + i, 142, 1: _
            PeekValue = 142
        
        'use start_line var as start line flag
        If PeekValue = 16 And start_line = False Then
            start_line = True
        ElseIf PeekValue = 16 Then
            start_line = False
        End If
        
        If PeekValue <> 16 And start_line = True Then
            cell = cell + 1
            .lblCell(cell).Caption = ConvertToSymbol(PeekValue, cell)
        End If
    Next i
    
    'if AutoCheckMines and first run after process attached
    If refresh_view And .mnuAutoCheckMines.Checked Then _
                     RefreshWindow: refresh_view = False
    
    'min 01005361 - end of max row 0100537E \______some memory table info
    'max 0100565E                           /
    
    'resize form
    
    .Height = .lblCell(intHor * intVert).Top + .lblCell(intHor * intVert).Height + 750
    .width = .lblCell(intHor * intVert).Left + .lblCell(intHor * intVert).width + 200
End With
End Sub

Public Sub ResetAll()
Dim i As Integer
    
With frmMain
    
    'intHor = 0: intVert = 0
    For i = 1 To .lblCell.ubound
        Unload .lblCell(i)
    Next i
        
End With
End Sub

Private Function ConvertToSymbol(value As Long, cell As Integer) As String
With frmMain.lblCell(cell)
    Select Case value
        Case 13: ConvertToSymbol = "?" '? flag
        Case 14: ConvertToSymbol = "F": .ForeColor = vbRed  'mine flag
        Case 15: ConvertToSymbol = "!": .BackColor = &HFFE1CE 'know
        Case 16: ConvertToSymbol = "#"  'row start,end flags
        Case 64: ConvertToSymbol = ""   'empty cell
        Case 65: ConvertToSymbol = "1": .ForeColor = vbBlue
        Case 66: ConvertToSymbol = "2": .ForeColor = &H8000&
        Case 67: ConvertToSymbol = "3": .ForeColor = vbRed
        Case 68: ConvertToSymbol = "4": .ForeColor = &H800000
        Case 69: ConvertToSymbol = "5": .ForeColor = &H80&
        Case 70: ConvertToSymbol = "6": .ForeColor = &H808000
        Case 71: ConvertToSymbol = "7"
        Case 72: ConvertToSymbol = "8": .ForeColor = &H808080
        Case 138: ConvertToSymbol = "X": .BackColor = vbRed 'visible mine
        Case 141: ConvertToSymbol = "?" '? flag
        Case 142: ConvertToSymbol = "F": .ForeColor = vbRed 'mine flag
        Case 143: ConvertToSymbol = "X": .BackColor = vbRed 'hiden mine
        Case 204: ConvertToSymbol = "X": .BackColor = vbRed 'visible hited mine
        Case Else: Debug.Print value: ConvertToSymbol = "@"
    End Select
End With
End Function

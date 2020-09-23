Attribute VB_Name = "Module1"
''For Tibia 7.4
''
Option Explicit
''
Public Const BattleList_First = &H49A07C
Public Const Player_ID = &H49A018
''
Public Const Distance_Characters = 156
Public Const Distance_Light = 112
Public Const Distance_ID = -4
''
Global BattleList_Array(BattleList_First To (BattleList_First + 147 * Distance_Characters)) As Long
''
Global BattleList_Address As Long
Global Player_Light As Long
''
''// API Declarations
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
''
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const PROCESS_VM_READ = (&H10)
Public Const PROCESS_VM_WRITE = (&H20)
Public Const PROCESS_VM_OPERATION = (&H8)
Public Const PROCESS_QUERY_INFORMATION = (&H400)
Public Const PROCESS_READ_WRITE_QUERY = PROCESS_VM_READ + PROCESS_VM_WRITE + PROCESS_VM_OPERATION + PROCESS_QUERY_INFORMATION
''
Public Function Tibia_Hwnd() As Long

  'Return the value of the Tibia Window
  'Use to find the window alot
    
  'Find Tibia's hwnd or Window
  Dim tibiaclient As Long
  tibiaclient = FindWindow("tibiaclient", vbNullString)
  
  'Return hwnd to function
  Tibia_Hwnd = tibiaclient

End Function
Public Function Memory_ReadByte(Address As Long) As Byte
  
   ' Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Byte   ' Byte
    
   ' First get a handle to the "game" window
   If (Tibia_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId Tibia_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   If (phandle = 0) Then Exit Function
   
   ' Read Long
   ReadProcessMemory phandle, Address, valbuffer, 1, 0&
       
   ' Return
   Memory_ReadByte = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Function Memory_ReadLong(Address As Long) As Long
  
   ' Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   Dim valbuffer As Long   ' Long
    
   ' First get a handle to the "game" window
   If (Tibia_Hwnd = 0) Then Exit Function
   
   ' We can now get the pid
   GetWindowThreadProcessId Tibia_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_VM_READ, False, pid)
   If (phandle = 0) Then Exit Function
   
   ' Read Long
   ReadProcessMemory phandle, Address, valbuffer, 4, 0&
       
   ' Return
   Memory_ReadLong = valbuffer
   
   ' Close the Process Handle
   CloseHandle phandle
  
End Function
Public Sub Memory_WriteByte(Address As Long, valbuffer As Byte)

   'Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   If (Tibia_Hwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId Tibia_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Sub
   
   ' Write Long
   WriteProcessMemory phandle, Address, valbuffer, 1, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub
Public Sub Memory_WriteLong(Address As Long, valbuffer As Long)

   'Declare some variables we need
   Dim pid As Long         ' Used to hold the Process Id
   Dim phandle As Long     ' Holds the Process Handle
   
   ' First get a handle to the "game" window
   If (Tibia_Hwnd = 0) Then Exit Sub
   
   ' We can now get the pid
   GetWindowThreadProcessId Tibia_Hwnd, pid
   
   ' Use the pid to get a Process Handle
   phandle = OpenProcess(PROCESS_READ_WRITE_QUERY, False, pid)
   If (phandle = 0) Then Exit Sub
   
   ' Write Long
   WriteProcessMemory phandle, Address, valbuffer, 4, 0&
   
   ' Close the Process Handle
   CloseHandle phandle

End Sub
Public Function Player_Address() As Long

  '// Return the Player_Address used
  '// for all the distance calculations

If Tibia_Hwnd = 0 Then Exit Function

 Dim Temp1 As Long, Temp2 As Long
 
 'For loop through the addresses
 For BattleList_Address = LBound(BattleList_Array) To UBound(BattleList_Array) Step Distance_Characters
 
 'put both IDs into memory!
   Temp1 = Memory_ReadLong(BattleList_Address + Distance_ID)
   Temp2 = Memory_ReadLong(Player_ID)
   
   'check if they match!
   If Temp1 = Temp2 Then
   
     'Player has been found!
     Player_Address = CLng("&H" & Hex(BattleList_Address))
     Exit Function
   
    End If

  Next BattleList_Address
 
 End Function
Public Sub Hack_Light(lightAmount As Byte)
  
  '//Do Light Hack
  
  'Find Tibia
  If Tibia_Hwnd = 0 Then Exit Sub
  
  'Light Hack Address
  Player_Light = CLng(Player_Address + Distance_Light)
    
  'Check if player has already light to prevent a unecessary memory write
   If Memory_ReadByte(Player_Light) <> lightAmount Then
   
   'Write Memory
    Call Memory_WriteByte(Player_Light, lightAmount)
    
  End If
    
End Sub
Public Sub Window_Ontop(thewindow As Form)

  '// Make a window stay on top
  
  SetWindowPos thewindow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOMOVE

End Sub

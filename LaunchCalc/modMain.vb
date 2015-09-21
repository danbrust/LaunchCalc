'<div>Icon made by <a href="http://www.freepik.com" title="Freepik">Freepik</a> from <a href="http://www.flaticon.com" title="Flaticon">www.flaticon.com</a> is licensed under <a href="http://creativecommons.org/licenses/by/3.0/" title="Creative Commons BY 3.0">CC BY 3.0</a></div>

Option Explicit On

Imports System.IO
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Module modMain
  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
  Private Function AttachThreadInput(ByVal idAttach As UInt32, ByVal idAttachTo As UInt32, ByVal fAttach As Boolean) As Boolean
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
  Private Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
  End Function

  <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
  Private Function GetCurrentThreadId() As Integer
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
  Private Function GetForegroundWindow() As IntPtr
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
  Private Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
  Public Function IsIconic(hWnd As IntPtr) As Boolean
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
  Public Function IsZoomed(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
  Private Function SetFocus(ByVal hWnd As IntPtr) As IntPtr
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
  Private Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
  End Function

  <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
  Private Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Int32) As Boolean
  End Function

  Public Enum ShowWindowCommands As Integer
    'for info on these commands see, http://msdn.microsoft.com/en-us/library/windows/desktop/ms633548%28v=vs.85%29.aspx
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8 'similar to SW_SHOW, except the window is not activated.
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
  End Enum

  Public Sub Main()
    Try
      Const filePath As String = "calc.exe"
      Const processName As String = "Calculator"

      Dim startedApp As Boolean = False

      ' Detect OS.
      ' Win7 = 6.1
      ' Win8 = 6.2
      ' Win8.1 = 6.3
      ' Win10 = 10.0
      Dim isWin10 As Boolean = True
      If Environment.OSVersion.Version.Major < 10 Then
        isWin10 = False
      End If

      ' See if the Windows Calculator is running.
      Dim hWnd As IntPtr = 0
      hWnd = FindWindow("ApplicationFrameWindow", processName)

      ' If it wasn't found, then try checking using Win7-style CalcFrame.
      If hWnd = IntPtr.Zero Then
        hWnd = FindWindow("CalcFrame", processName)
      End If

      If hWnd = IntPtr.Zero Then
        ' Still haven't found a running instance? Then launch a new one.
        Dim process As Process = process.Start(filePath)
        startedApp = True

        If isWin10 Then
          ' Win10 launches calc.exe, which launches calculator.exe, then closes off. Need to wait for it to finish.
          process.WaitForExit(3000) ' Wait no longer than 3 seconds.
        End If

        ' We can't just grab the handle from the started process, as calc.exe launches calculator.exe. Instead, we now have to look for the new process.
        hWnd = FindWindow("ApplicationFrameWindow", processName)
        If hWnd = IntPtr.Zero Then
          hWnd = FindWindow("CalcFrame", processName)
        End If
        ForceForegroundWindow(hWnd, isWin10, startedApp)
      Else
        ForceForegroundWindow(hWnd, isWin10, startedApp)
      End If '' If hWnd = IntPtr.Zero Then

    Catch ex As Exception
      MsgBox(ex.Message, vbCritical, "Error!")

    End Try
  End Sub

  Private Function GetWindowHandle(processName As String) As IntPtr
    processName = System.IO.Path.GetFileNameWithoutExtension(processName)
    Dim processArray() As System.Diagnostics.Process = Process.GetProcessesByName(processName)
    If processArray.Length > 0 Then
      Return processArray(0).MainWindowHandle
    Else
      Return IntPtr.Zero
    End If
  End Function

  Private Function ForceForegroundWindow(ByVal hWnd As IntPtr, ByVal isWin10 As Boolean, ByVal startedApp As Boolean) As Boolean
    ' Windows works hard to stop us from focusing on external application, so I'm throwing everything I can at it. Something is bound to stick...

    Dim Result As Integer = 0

    ' Check to see if the window is already in the foreground; if so, do nothing.
    If hWnd = GetForegroundWindow() Then
      Return True
    Else
      ' Get thread responsible for the foreground window.
      Dim ForeThread As IntPtr = GetWindowThreadProcessId(GetForegroundWindow(), 0)

      ' Get thread running the passed (hWnd) window.
      Dim AppThread As IntPtr = GetWindowThreadProcessId(hWnd, 0)

      If ForeThread <> AppThread Then
        ' Need to attach new thread to foreground thread.
        AttachThreadInput(ForeThread, AppThread, True)
        Result = SetForegroundWindow(hWnd)
        SetFocus(hWnd)
        AttachThreadInput(ForeThread, AppThread, False)
      Else
        ' Just need to focus on window.
        Result = SetForegroundWindow(hWnd)
        SetFocus(hWnd)
      End If
      'End If

      ' Restore the window.
      If IsIconic(hWnd) Then
        ShowWindow(hWnd, ShowWindowCommands.SW_RESTORE)
      ElseIf IsZoomed(hWnd) Then
        If isWin10 Then
          ' Minimizing first increases reliability of showing the window on Win10 systems.
          ShowWindow(hWnd, ShowWindowCommands.SW_MINIMIZE)
          ShowWindow(hWnd, ShowWindowCommands.SW_SHOWMAXIMIZED)
        Else
          ShowWindow(hWnd, ShowWindowCommands.SW_SHOWMAXIMIZED)
        End If
      Else
        If isWin10 Then
          ' Minimizing first increases reliability of showing the window on Win10 systems.
          ShowWindow(hWnd, ShowWindowCommands.SW_MINIMIZE)
          ShowWindow(hWnd, ShowWindowCommands.SW_RESTORE)
        Else
          ShowWindow(hWnd, ShowWindowCommands.SW_SHOW)
        End If
      End If

      ForceForegroundWindow = CBool(Result)
    End If
  End Function

End Module

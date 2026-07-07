''' ============================================================================
''' Faxtur
''' Copyright © 2026 Frédéric Brouard
'''
''' This Source Code Form is subject to the terms of the
''' Mozilla Public License, v. 2.0.
''' If a copy of the MPL was not distributed with this file,
''' You can obtain one at https://mozilla.org/MPL/2.0/
''' ============================================================================
Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

AppDir = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = AppDir

Pythonw = AppDir & "\.venv\Scripts\pythonw.exe"
Python = AppDir & "\.venv\Scripts\python.exe"
MainPy = AppDir & "\main.py"

If FSO.FileExists(Pythonw) Then
  Cmd = Chr(34) & Pythonw & Chr(34) & " " & Chr(34) & MainPy & Chr(34)
ElseIf FSO.FileExists(Python) Then
  Cmd = Chr(34) & Python & Chr(34) & " " & Chr(34) & MainPy & Chr(34)
Else
  Cmd = "pythonw " & Chr(34) & MainPy & Chr(34)
End If

WshShell.Run Cmd, 0, False

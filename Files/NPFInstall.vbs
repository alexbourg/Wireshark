'## Application Name: NPF Installer
'## Created Date: 16.05.2021
'## Created By: Alex BOURG

'## Description: This script will install NPF silently

'========================================================================
'## Global Object and Variable Settings
'======================================================================== 
'Option Explicit

Dim WshShell:	Set WshShell = CreateObject("WScript.Shell")
Dim objFSO:	Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim SourceDir:	SourceDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
Dim value
WshShell.CurrentDirectory = SourceDir

DIM Subfolder
Dim command
'========================================================================
'## Global Object and Variable Settings
'======================================================================== 
On error resume next
wshshell.run "cmd /C ""C:\Program Files\Npcap\bin\NPFInstall.exe"" -i",0,true
Set WshShell = Nothing
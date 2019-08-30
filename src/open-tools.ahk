#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

!t::
Run "C:\Program Files\PowerShell\7-preview\pwsh.exe"
return

#c::
Path := Explorer_GetSelection()
Path := (Path = "" ? Clipboard : Path)
if(FileExist(Clipboard)) {
   Run, "code" %Path%
} else {
   Run, "code"
}
return

#e::
if(FileExist(Clipboard)) {
   run, Explore %Clipboard%
} else {
   run, Explore C:\Repos
}
return

#s::
try {
   run, www.bing.com/search?q=%Clipboard%
} catch e {
   run, www.bing.com
}

return

!g::
run, www.github.com
return

Explorer_GetSelection() {
   WinGetClass, winClass, % "ahk_id" . hWnd := WinExist("A")
   if !(winClass ~= "(Cabinet|Explore)WClass")
      Return
   for window in ComObjCreate("Shell.Application").Windows
      if (hWnd = window.HWND) && (oShellFolderView := window.document)
         break
   for item in oShellFolderView.SelectedItems
      result .= (result = "" ? "" : "`n") . item.path
   if !result
      result := oShellFolderView.Folder.Self.Path
   Return result
}
Set objHTML = CreateObject("htmlfile")
textFromClipboard = objHTML.ParentWindow.ClipboardData.GetData("text")

' Extract the filename from the URL (similar to your batch script)
url = textFromClipboard

' Construct the new URL
newUrl = "https://hub.gitmirror.com/raw.githubusercontent.com/sDreamForZzQ/img/refs/heads/master/wallpaper/" & url

' Copy to clipboard
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c echo " & newUrl & "| clip", 0, True
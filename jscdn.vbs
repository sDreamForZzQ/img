Set objHTML = CreateObject("htmlfile")
textFromClipboard = objHTML.ParentWindow.ClipboardData.GetData("text")

' Extract the filename from the URL (similar to your batch script)
url = textFromClipboard

' Construct the new URL
newUrl = "https://jsd.cdn.zzko.cn/gh/sDreamForZzQ/img@latest/wallpaper/" & url

' Copy to clipboard
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c echo " & newUrl & "| clip", 0, True
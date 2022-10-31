@echo off
for /f "delims=" %%i in ('mshta "JavaScript:new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(clipboardData.getData('text'));close()"') do (
    set "url=%%i"
)

set "s1=%url%"

echo https://cdn.jsdelivr.net/gh/sDreamForZzQ/img@latest/wallpaper/%s1% | clip
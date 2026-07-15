@echo off
:: PS Tracker Launcher
:: Finds uv automatically — works for any user on any Amazon machine

set "UVX=%USERPROFILE%\.aki\bin\uv.exe"
if not exist "%UVX%" set "UVX=uv"

"%UVX%" run --with requests --with openpyxl --with pycryptodome --with requests-negotiate-sspi --with requests-kerberos "%~dp0Launch PS Tracker.pyw"

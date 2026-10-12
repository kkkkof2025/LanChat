@echo off
cd /d %~dp0
chcp 65001
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { [Console]::OutputEncoding=[System.Text.Encoding]::UTF8; [Console]::InputEncoding=[System.Text.Encoding]::UTF8; & '.\LanChat.ps1' }"
pause

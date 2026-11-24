@echo off
chcp 65001 >nul
title 開啟自主檢查表簽章轉換系統
cd /d "%~dp0"

echo 正在為您開啟前端網頁介面...
start "" "index.html"
exit

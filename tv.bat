@echo off
title TradingView Market Movers

REM ===========================
REM Change this path
REM ===========================

cd /d "C:\Users\ATHARVA\pension"

REM ===========================
REM Run Python Script
REM ===========================

python tv.py

REM If python isn't recognized, uncomment the next line
REM py tradingview.py

echo.
echo ======================================
echo Script has stopped.
echo ======================================
pause
@echo off
title %cd%
call conda activate
python run-mp3.py
pause
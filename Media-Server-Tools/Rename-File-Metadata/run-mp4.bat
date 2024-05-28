@echo off
title %cd%
call conda activate
python run-mp4.py
pause
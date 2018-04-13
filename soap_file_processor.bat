@echo off
set src=%1
set dest=%2
echo ===========================================================================
echo Source file = %src% and Destination file = %dest%
echo ===========================================================================
python soap_file_processor.py %src% %dest%
Pause





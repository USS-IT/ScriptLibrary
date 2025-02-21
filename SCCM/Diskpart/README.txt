Formats disk0 cleanly for UEFI boot

1) Place file on imaging stick
2) After booting into PXE, open Command Prompt before selecting Task Sequence
3) Enter command:
cd /d D:\
diskpart /s diskpart_reformat_disk0.txt

...where D:\ is the drive letter for the imaging stick (may be different)
4) Remove imaging stick and start imaging

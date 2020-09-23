set Path=%Path%;c:\winnt\system32
for %%a in (*.ocx) do regsvr32 /s /c %%a

set Path=%Path%;c:\winnt\system32;
for %%a in (*.dll) do regsvr32 /s /c %%a

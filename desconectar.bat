@echo off
for /f "skip=1 tokens=3" %%s in ('query user %USERNAME%') do (
    tscon %%s /dest:console
)
exit
pause
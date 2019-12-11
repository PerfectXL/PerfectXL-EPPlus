@echo off
cd /d %~dp0
del /q PerfectXL.EPPlus\bin\Release\PerfectXL.EPPlus.*.nupkg
rd /s/q PerfectXL.EPPlus\bin
rd /s/q PerfectXL.EPPlus\obj
dotnet pack -c Release
dotnet nuget push PerfectXL.EPPlus\bin\Release\PerfectXL.EPPlus.*.nupkg --api-key q5ZDX89QvzM01aG4F8ZJ --source https://nuget.perfectxl.com/nuget/default
pause
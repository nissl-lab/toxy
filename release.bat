@echo off

set fdir=%WINDIR%\Microsoft.NET\Framework64

if not exist %fdir% (
	set fdir=%WINDIR%\Microsoft.NET\Framework
)

set msbuild=%fdir%\v4.0.30319\msbuild.exe

%msbuild% ToxyFramework\ToxyFramework_dotnet45.csproj /p:Configuration=Release /t:Rebuild /p:OutputPath=..\Release\Net45

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework_dotnet4.csproj /p:Configuration=Release /t:Rebuild /p:OutputPath=..\Release\Net40

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework.csproj  /p:Configuration=Release /t:Rebuild /p:OutputPath=..\Release\Net35

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

pause
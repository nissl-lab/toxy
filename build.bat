@echo off

set fdir=%WINDIR%\Microsoft.NET\Framework64

if not exist %fdir% (
	set fdir=%WINDIR%\Microsoft.NET\Framework
)

set msbuild=%fdir%\v4.0.30319\msbuild.exe

%msbuild% ToxyFramework\ToxyFramework_dotnet45.csproj /p:Configuration=Release /t:Rebuild /p:OutputPath=..\Build\Net45\Release

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework_dotnet4.csproj /p:Configuration=Release /t:Rebuild /p:OutputPath=..\Build\Net40\Release

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework.csproj  /p:Configuration=Release /t:Rebuild /p:OutputPath=..\Build\Net20\Release

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework_dotnet45.csproj /p:Configuration=Debug /t:Rebuild /p:OutputPath=..\Build\Net45\Debug

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework_dotnet4.csproj /p:Configuration=Debug /t:Rebuild /p:OutputPath=..\Build\Net40\Debug

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

%msbuild% ToxyFramework\ToxyFramework.csproj  /p:Configuration=Debug /t:Rebuild /p:OutputPath=..\Build\Net20\Debug

FOR /F "tokens=*" %%G IN ('DIR /B /AD /S obj') DO RMDIR /S /Q "%%G"

pause
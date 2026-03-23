@echo off
REM FILE: csolver/build.bat
REM PURPOSE: Compiles solver.c into solver.dll using GCC (MinGW-w64).
REM
REM HOW TO RUN:
REM     cd csolver
REM     build.bat
REM
REM OUTPUT:
REM     solver.dll in the project root directory (one level up)
REM
REM PREREQUISITES:
REM     GCC (MinGW-w64) must be installed and on PATH.
REM     Install via: winget install BrechtSanders.WinLibs.POSIX.UCRT
REM
REM CHANGE LOG:
REM | Date       | Change                                    | Why                              |
REM |------------|-------------------------------------------|----------------------------------|
REM | 23-03-2026 | Created — DLL build script                | Phase 6 C solver                 |

echo Building solver.dll...

gcc -shared -O2 -o ..\solver.dll solver.c -lm

if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: solver.dll created in project root.
) else (
    echo FAILED: Compilation error. Check GCC is installed.
    echo Install GCC: winget install BrechtSanders.WinLibs.POSIX.UCRT
    exit /b 1
)

@echo off
REM Quick Test Commands for Excel Comparison Tool
REM This batch file provides shortcuts for running common test commands

setlocal enabledelayedexpansion

REM Check if pytest is installed
python -m pytest --version >nul 2>&1
if errorlevel 1 (
    echo Installing pytest...
    pip install pytest pytest-cov
)

REM Get the command argument
set "command=%1"

if "%command%"=="" (
    echo.
    echo Excel Comparison Tool - Test Commands
    echo ====================================
    echo.
    echo Usage: test.bat [command]
    echo.
    echo Commands:
    echo   all          Run all tests
    echo   fast         Run tests without coverage (fastest)
    echo   coverage     Run tests with coverage report
    echo   html         Generate HTML coverage report
    echo   engine       Run comparison engine tests only
    echo   report       Run report generator tests only
    echo   integration  Run integration tests only
    echo   verbose      Run all tests with verbose output
    echo   watch        Run tests and re-run on file changes
    echo   help         Show this help message
    echo.
    goto end
)

if "%command%"=="all" (
    echo Running all tests...
    python -m pytest tests/ -v
    goto end
)

if "%command%"=="fast" (
    echo Running all tests (fast mode, no coverage)...
    python -m pytest tests/ -v --tb=short
    goto end
)

if "%command%"=="coverage" (
    echo Running tests with coverage...
    python -m pytest tests/ --cov=src --cov-report=term-missing
    goto end
)

if "%command%"=="html" (
    echo Generating HTML coverage report...
    python -m pytest tests/ --cov=src --cov-report=html
    echo.
    echo Coverage report generated. Opening in browser...
    start htmlcov\index.html
    goto end
)

if "%command%"=="engine" (
    echo Running comparison engine tests...
    python -m pytest tests/test_comparison_engine.py -v
    goto end
)

if "%command%"=="report" (
    echo Running report generator tests...
    python -m pytest tests/test_report_generator.py -v
    goto end
)

if "%command%"=="integration" (
    echo Running integration tests...
    python -m pytest tests/test_integration.py -v
    goto end
)

if "%command%"=="verbose" (
    echo Running all tests with verbose output...
    python -m pytest tests/ -vv
    goto end
)

if "%command%"=="watch" (
    echo Installing pytest-watch...
    pip install pytest-watch
    echo Watching for file changes...
    ptw tests/
    goto end
)

if "%command%"=="help" (
    echo.
    echo Excel Comparison Tool - Test Commands
    echo ====================================
    echo.
    echo Usage: test.bat [command]
    echo.
    echo Commands:
    echo   all          Run all tests
    echo   fast         Run tests without coverage (fastest)
    echo   coverage     Run tests with coverage report
    echo   html         Generate HTML coverage report
    echo   engine       Run comparison engine tests only
    echo   report       Run report generator tests only
    echo   integration  Run integration tests only
    echo   verbose      Run all tests with verbose output
    echo   watch        Run tests and re-run on file changes
    echo   help         Show this help message
    echo.
    goto end
)

echo Unknown command: %command%
echo Run "test.bat help" for available commands
exit /b 1

:end
endlocal

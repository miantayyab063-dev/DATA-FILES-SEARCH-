/* GUI entry point for the bundled Windows Python runtime. */
#include <windows.h>
#include <wchar.h>
#include <stdio.h>

#define CAP 4096

int main(void) {
    wchar_t directory[CAP];
    wchar_t python[CAP];
    wchar_t boot[CAP];
    wchar_t command[CAP * 2 + 16];
    DWORD length = GetModuleFileNameW(NULL, directory, CAP);
    wchar_t *last;
    STARTUPINFOW start = {0};
    PROCESS_INFORMATION process = {0};

    if (!length || length >= CAP - 1) {
        MessageBoxW(NULL, L"Cannot find the application folder.",
                    L"HACKER M-T181", MB_OK | MB_ICONERROR);
        return 1;
    }
    last = wcsrchr(directory, L'\\');
    if (!last) return 1;
    *last = L'\0';

    if (wcslen(directory) + 20 >= CAP) return 1;
    swprintf(python, CAP, L"%ls\\pythonw.exe", directory);
    swprintf(boot, CAP, L"%ls\\boot.py", directory);
    swprintf(command, CAP * 2 + 16, L"\"%ls\" \"%ls\"", python, boot);
    start.cb = sizeof(start);

    if (!CreateProcessW(python, command, NULL, NULL, FALSE, 0,
                        NULL, directory, &start, &process)) {
        MessageBoxW(NULL,
                    L"Cannot start the program. Extract the whole ZIP first, "
                    L"then keep HACKER_M_T181.exe, pythonw.exe, boot.py and "
                    L"the data folder together.",
                    L"HACKER M-T181", MB_OK | MB_ICONERROR);
        return 1;
    }
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return 0;
}

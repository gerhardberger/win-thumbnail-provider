# win-thumbnail-provider

A DLL shell extension that when registered can generate custom thumbnails for
files with a specified extension.

## Usage

1. Clone the repository, or copy the `thumbnail_provider.cpp` and\
`thumbnail_provider.def` file into your project.
2. Compile the `thumbnail_provider.dll` using `npm run build`. Note: Make sure
to compile your DLL for the correct environment, e.g. x64 as that can be a
problem when registering.
3. Ship `thumbnail_provider.dll` with your app or installer.
4. Register it with the following command (important to use the absolute path to the DLL):

``` bash
$ regsvr32 /absolute/path/to/thumbnail_provider.dll`
```

To uninstall the shell extension use run the following:

``` bash
$ regsvr32 /u /absolute/path/to/thumbnail_provider.dll`
```

## Debugging

### Logging

To debug you can add the following method to the code and use it add logging:

``` c++
void LogEvent(LPCWSTR message) {
    HANDLE hEventLog = RegisterEventSourceW(NULL, L"ThumbnailProvider");
    if (hEventLog) {
        LPCWSTR strings[1] = { message };
        ReportEventW(hEventLog, EVENTLOG_INFORMATION_TYPE, 0, 0, NULL, 1, 0, strings, NULL);
        DeregisterEventSource(hEventLog);
    }
}
```

You can check these logs in the `Event Viewer` app under 'Windows Logs' -> 'Application'

### Restarting explorer

Also useful to restart the explorer easily with the following command:

``` bash
$ taskkill /f /im explorer.exe && start explorer.exe
```

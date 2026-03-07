# Classic Repair Toolbox


## Compiler instructions for OS
- [Windows](#windows)
- [Linux](#linux)
- [macOS](#macos) 


## Windows

Load `Classic-Repair-Toolbox.slnx` in Visual Studio and build it.


## Linux

### Common for all Linux builds

- Make sure .NET10 SDK is installed.
  - Install via your internal package management system
  - .. or download from here, https://dotnet.microsoft.com/en-us/download/dotnet/10.0
- Fork the _CRT_ GitHub repository
- Clone the fork to your local computer

### **Fedora**
- Compile RELEASE build
  - `dotnet publish -c Release -f net10.0 --self-contained`
- Run application:  
  - `./bin/Debug/net10.0/Classic-Repair-Toolbox`

### **Gentoo**
- Show all available .NET SDK versions
  - `eselect dotnet list`
- Choose .NET10 SDK, which is profile (1) in this example
  - `eselect dotnet set 1`
- Reload system environment variables
  - `. /etc/profile`
- Verify the active .NET SDK
  - `dotnet --list-sdks`
- Compile RELEASE build
  - `dotnet publish -c Release -f net10.0 --self-contained`
- Run application:
  - `bin/Release/net10.0/linux-x64/Classic-Repair-Toolbox`

Note that it is recommened you always create a `RELEASE` version, as it otherwise will not check for a new version online. If this is a `DEBUG` build, it will always show-case a dummy update to visualize the UI for it.


## macOS

No help on this yet - please help me in providing these steps.

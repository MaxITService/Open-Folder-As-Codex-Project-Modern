# Open Folder As Codex Project - Modern

Local-admin/dev helper for adding **Open project in Codex** to the Windows 11
modern context menu.

![Open project in Codex in the Windows 11 modern context menu](<Media-GitHub/how it works modern menu.png>)

It uses the Windows 11-supported shape:

- a native `IExplorerCommand` shell extension DLL,
- a small signed MSIX package that gives the command app identity,
- packaged COM registration through `com:ComServer`,
- `desktop4:FileExplorerContextMenus` entries for folders and folder background.

## Install

Run **Install-Win11ModernContextMenu.bat**.

It asks for administrator rights automatically when needed.

```cmd
Install-Win11ModernContextMenu.bat
```

If the package is already installed, the installer asks whether to uninstall,
reinstall/update, or cancel.

The installer keeps the console open and writes a log to `artifacts\install.log`.

## Useful options

```cmd
Install-Win11ModernContextMenu.bat -CodexExe "C:\Path\To\Codex.exe"
Install-Win11ModernContextMenu.bat -Configuration Debug
Install-Win11ModernContextMenu.bat -NoRestartExplorer
Install-Win11ModernContextMenu.bat -Uninstall
```

`x64` is the default platform. `-Platform ARM64` is available for Windows on
ARM devices. There is no 32-bit build.

## What It Does

The installer:

1. finds Codex Desktop, or uses `-CodexExe`,
2. generates and trusts a local self-signed certificate,
3. builds the shell extension DLL with MSBuild,
4. renders the package manifest,
5. creates and signs the MSIX with `MakeAppx.exe` and `SignTool.exe`,
6. installs the package with `Add-AppxPackage`,
7. restarts Explorer.

## Requirements

- Windows 11.
- Local administrator rights for trusting the dev certificate and installing the package.
- Visual Studio 2022 Build Tools or Visual Studio with C++ desktop workload.
- Windows 10/11 SDK with `MakeAppx.exe` and `SignTool.exe`.
- PowerShell 5.1 or newer.

## Notes

- This follows the same architectural idea as
  [microsoft/vscode-explorer-command](https://github.com/microsoft/vscode-explorer-command):
  a small shell-extension DLL plus a package installed by the owning app/installer.
- No Microsoft source is vendored here. The C++ implementation is deliberately small and
  uses WRL and Windows SDK headers only.
- The package name, publisher, CLSID, and AppId are development defaults. Change them before
  using this in a distributed installer.
- Windows 11 package manifests do not accept `Drive` as a modern context-menu item type,
  so drive support belongs in the classic registry version.
- If `MakeAppx.exe` reports visual asset validation issues, replace the generated placeholder
  logos in `artifacts\PackageRoot\Assets` with real PNG assets and rerun.

## TODO before production

- Replace the development certificate with a production code-signing certificate.
- Decide whether the command should remain per-user or move to a managed machine-wide install.
- Add CI builds for `x64` and `arm64`.
- Add a smoke-test helper that checks package registration and COM activation diagnostics.

## Related Version

This is the native Windows 11 modern-menu version.

Want the simpler registry-based classic menu version?
Use [Open-Folder-As-Codex-Project](https://github.com/MaxITService/Open-Folder-As-Codex-Project).

## License

MIT License.

---

## My Other Projects

- [AivoRelay: AI Voice Relay for Windows](https://github.com/MaxITService/AIVORelay)
- [OneClickPrompts: Your Quick Prompt Companion for Multiple AI Chats](https://github.com/MaxITService/OneClickPrompts)
- [Console2Ai: Send PowerShell buffer to AI](https://github.com/MaxITService/Console2Ai)
- [AI for Complete Beginners: Guide to LLMs](https://medium.com/@maxim.fomins/ai-for-complete-beginners-guide-llms-f19c4b8a8a79)
- [Ping-Plotter: PowerShell-only ping plotting script](https://github.com/MaxITService/Ping-Plotter-PS51)

#include <windows.h>
#include <shellapi.h>
#include <shlobj_core.h>
#include <shlwapi.h>
#include <wrl/client.h>
#include <wrl/implements.h>
#include <wrl/module.h>

#include <string>

using Microsoft::WRL::ClassicCom;
using Microsoft::WRL::ComPtr;
using Microsoft::WRL::InhibitRoOriginateError;
using Microsoft::WRL::Module;
using Microsoft::WRL::ModuleType;
using Microsoft::WRL::RuntimeClass;
using Microsoft::WRL::RuntimeClassFlags;

namespace {

constexpr wchar_t kSettingsSubKey[] = L"Software\\OpenFolderAsCodexProject\\Win11Modern";
constexpr wchar_t kFallbackTitle[] = L"Open project in Codex";
constexpr wchar_t kFallbackCodexExe[] = L"%LOCALAPPDATA%\\OpenAI\\Codex\\app\\Codex.exe";
HINSTANCE g_module = nullptr;

std::wstring ExpandEnvironmentStringsToWString(const std::wstring& value) {
  DWORD required = ExpandEnvironmentStringsW(value.c_str(), nullptr, 0);
  if (required == 0) {
    return value;
  }

  std::wstring expanded(required, L'\0');
  DWORD written = ExpandEnvironmentStringsW(value.c_str(), expanded.data(), required);
  if (written == 0 || written > required) {
    return value;
  }

  if (!expanded.empty() && expanded.back() == L'\0') {
    expanded.pop_back();
  }
  return expanded;
}

bool ReadRegistryString(HKEY root, const wchar_t* valueName, std::wstring* value) {
  HKEY key = nullptr;
  LONG result = RegOpenKeyExW(root, kSettingsSubKey, 0, KEY_QUERY_VALUE | KEY_WOW64_64KEY, &key);
  if (result != ERROR_SUCCESS) {
    return false;
  }

  DWORD type = 0;
  DWORD byteCount = 0;
  result = RegQueryValueExW(key, valueName, nullptr, &type, nullptr, &byteCount);
  if (result != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ) || byteCount == 0) {
    RegCloseKey(key);
    return false;
  }

  std::wstring buffer(byteCount / sizeof(wchar_t), L'\0');
  result = RegQueryValueExW(key, valueName, nullptr, &type, reinterpret_cast<LPBYTE>(buffer.data()), &byteCount);
  RegCloseKey(key);
  if (result != ERROR_SUCCESS) {
    return false;
  }

  while (!buffer.empty() && buffer.back() == L'\0') {
    buffer.pop_back();
  }

  *value = type == REG_EXPAND_SZ ? ExpandEnvironmentStringsToWString(buffer) : buffer;
  return true;
}

std::wstring GetConfiguredString(const wchar_t* valueName, const wchar_t* fallback) {
  std::wstring value;
  if (ReadRegistryString(HKEY_CURRENT_USER, valueName, &value) ||
      ReadRegistryString(HKEY_LOCAL_MACHINE, valueName, &value)) {
    return value;
  }

  return ExpandEnvironmentStringsToWString(fallback);
}

bool IsEnabled() {
  std::wstring value = GetConfiguredString(L"Enabled", L"1");
  return value != L"0";
}

std::wstring QuoteForCommandLineArg(const std::wstring& arg) {
  if (arg.find_first_of(L" \\\"") == std::wstring::npos) {
    return arg;
  }

  std::wstring output;
  output.push_back(L'"');
  for (size_t i = 0; i < arg.size(); ++i) {
    if (arg[i] == L'\\') {
      size_t end = i;
      while (end < arg.size() && arg[end] == L'\\') {
        ++end;
      }

      size_t backslashCount = end - i;
      if (end == arg.size() || arg[end] == L'"') {
        backslashCount *= 2;
      }

      output.append(backslashCount, L'\\');
      i = end - 1;
    } else if (arg[i] == L'"') {
      output.append(L"\\\"");
    } else {
      output.push_back(arg[i]);
    }
  }
  output.push_back(L'"');
  return output;
}

HRESULT GetItemPath(IShellItem* item, std::wstring* path) {
  PWSTR rawPath = nullptr;
  HRESULT hr = item->GetDisplayName(SIGDN_FILESYSPATH, &rawPath);
  if (FAILED(hr)) {
    return hr;
  }

  *path = rawPath;
  CoTaskMemFree(rawPath);
  return S_OK;
}

HRESULT LaunchCodexForPath(const std::wstring& path) {
  std::wstring codexExe = GetConfiguredString(L"CodexExe", kFallbackCodexExe);
  std::wstring parameters = L"--open-project " + QuoteForCommandLineArg(path);
  HINSTANCE result = ShellExecuteW(nullptr, L"open", codexExe.c_str(), parameters.c_str(), nullptr, SW_SHOWNORMAL);
  return reinterpret_cast<INT_PTR>(result) <= HINSTANCE_ERROR ? HRESULT_FROM_WIN32(GetLastError()) : S_OK;
}

}  // namespace

extern "C" BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID) {
  if (reason == DLL_PROCESS_ATTACH) {
    g_module = instance;
    DisableThreadLibraryCalls(instance);
  }
  return TRUE;
}

class __declspec(uuid("5f220f1e-376f-4f5b-9d5e-d42d924ff811"))
    OpenFolderAsCodexCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>, IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) override {
    return SHStrDupW(GetConfiguredString(L"Title", kFallbackTitle).c_str(), name);
  }

  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) override {
    std::wstring codexExe = GetConfiguredString(L"CodexExe", kFallbackCodexExe);
    return SHStrDupW(codexExe.c_str(), icon);
  }

  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* infoTip) override {
    *infoTip = nullptr;
    return E_NOTIMPL;
  }

  IFACEMETHODIMP GetCanonicalName(GUID* guidCommandName) override {
    *guidCommandName = __uuidof(OpenFolderAsCodexCommand);
    return S_OK;
  }

  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* cmdState) override {
    *cmdState = IsEnabled() ? ECS_ENABLED : ECS_HIDDEN;
    return S_OK;
  }

  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) override {
    *flags = ECF_DEFAULT;
    return S_OK;
  }

  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** enumCommands) override {
    *enumCommands = nullptr;
    return E_NOTIMPL;
  }

  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) override {
    if (items == nullptr) {
      return E_INVALIDARG;
    }

    DWORD count = 0;
    HRESULT hr = items->GetCount(&count);
    if (FAILED(hr)) {
      return hr;
    }

    for (DWORD index = 0; index < count; ++index) {
      ComPtr<IShellItem> item;
      hr = items->GetItemAt(index, &item);
      if (FAILED(hr)) {
        return hr;
      }

      std::wstring path;
      hr = GetItemPath(item.Get(), &path);
      if (FAILED(hr)) {
        return hr;
      }

      hr = LaunchCodexForPath(path);
      if (FAILED(hr)) {
        return hr;
      }
    }

    return S_OK;
  }
};

CoCreatableClass(OpenFolderAsCodexCommand)
CoCreatableClassWrlCreatorMapInclude(OpenFolderAsCodexCommand)

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
  if (ppv == nullptr) {
    return E_POINTER;
  }

  *ppv = nullptr;
  return Module<ModuleType::InProc>::GetModule().GetClassObject(rclsid, riid, ppv);
}

STDAPI DllCanUnloadNow() {
  return Module<ModuleType::InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}

STDAPI DllGetActivationFactory(HSTRING activatableClassId, IActivationFactory** factory) {
  return Module<ModuleType::InProc>::GetModule().GetActivationFactory(activatableClassId, factory);
}

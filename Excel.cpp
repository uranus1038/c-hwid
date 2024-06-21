#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>
#include <vector>
#pragma comment(lib, "wbemuuid.lib")

void PrintDeviceInformation(IWbemServices* pSvc, const wchar_t* query, const wchar_t* property, std::vector<std::wstring>& serialNumbers)
{
    IEnumWbemClassObject* pEnumerator = nullptr;
    HRESULT hr = pSvc->ExecQuery(
        _bstr_t("WQL"),
        _bstr_t(query),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator
    );
    if (FAILED(hr)) {
        std::cerr << "Query failed: " << hr << std::endl;
        return;
    }

    IWbemClassObject* pclsObj = nullptr;
    ULONG uReturn = 0;
    while (pEnumerator) {
        hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (uReturn == 0) {
            break;
        }

        VARIANT vtProp;
        hr = pclsObj->Get(property, 0, &vtProp, 0, 0);
        if (FAILED(hr)) {
            std::cerr << "Failed to get property " << property << ": " << hr << std::endl;
            pclsObj->Release();
            continue;
        }

        std::wcout << property << ": " << vtProp.bstrVal << std::endl;
        serialNumbers.push_back(vtProp.bstrVal);

        VariantClear(&vtProp);
        pclsObj->Release();
    }

    pEnumerator->Release();
}

void UpdateDeviceInformation(IWbemServices* pSvc, const wchar_t* query, const wchar_t* property, const std::wstring& newValue)
{
    IEnumWbemClassObject* pEnumerator = nullptr;
    HRESULT hr = pSvc->ExecQuery(
        _bstr_t("WQL"),
        _bstr_t(query),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator
    );
    if (FAILED(hr)) {
        std::cerr << "Query failed: " << hr << std::endl;
        return;
    }

    IWbemClassObject* pclsObj = nullptr;
    ULONG uReturn = 0;
    while (pEnumerator) {
        hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (uReturn == 0) {
            break;
        }

        VARIANT vtProp;
        hr = pclsObj->Get(property, 0, &vtProp, 0, 0);
        if (FAILED(hr)) {
            std::cerr << "Failed to get property " << property << ": " << hr << std::endl;
            pclsObj->Release();
            continue;
        }

        // Update the property with the new value
        vtProp.vt = VT_BSTR;
        vtProp.bstrVal = SysAllocString(newValue.c_str());
        hr = pclsObj->Put(property, 0, &vtProp, 0);
        if (FAILED(hr)) {
            std::cerr << "Failed to put property " << property << ": " << hr << std::endl;
            VariantClear(&vtProp);
            pclsObj->Release();
            continue;
        }

        // Save the changes in WMI
        hr = pSvc->PutInstance(pclsObj, WBEM_FLAG_UPDATE_ONLY, NULL, NULL);
        if (FAILED(hr)) {
            if (hr == WBEM_E_NOT_FOUND) {
                std::cerr << "Instance not found." << std::endl;
            }
            else if (hr == WBEM_E_ACCESS_DENIED) {
                std::cerr << "Access denied." << std::endl;
            }
            else {
                std::cerr << "Failed to update instance. Error code: 0x" << std::hex << hr << std::endl;
            }
            VariantClear(&vtProp); // Clear the variant before releasing
            pclsObj->Release();
            continue;
        }

        std::wcout << "Updated " << property << " to: " << newValue << std::endl;

        VariantClear(&vtProp);
        pclsObj->Release();
    }

    pEnumerator->Release();
}

int main() {
    std::vector<std::wstring> biosSerialNumbers;

    HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        std::cerr << "Failed to initialize COM library: " << hr << std::endl;
        return 1;
    }

    hr = CoInitializeSecurity(
        NULL,                     // pVoid
        -1,                       // cAuthSvc
        NULL,                     // asAuthSvc
        NULL,                     // pReserved1
        RPC_C_AUTHN_LEVEL_DEFAULT,// dwAuthnLevel
        RPC_C_IMP_LEVEL_IMPERSONATE, // dwImpLevel
        NULL,                     // pAuthList
        EOAC_NONE,                // dwCapabilities
        NULL                      // pReserved3
    );
    if (FAILED(hr)) {
        std::cerr << "Failed to initialize security: " << hr << std::endl;
        CoUninitialize();
        return 1;
    }

    IWbemLocator* pLoc = nullptr;
    hr = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hr)) {
        std::cerr << "Failed to create IWbemLocator object: " << hr << std::endl;
        CoUninitialize();
        return 1;
    }

    IWbemServices* pSvc = nullptr;
    hr = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL,
        NULL,
        0,
        NULL,
        0,
        0,
        &pSvc
    );
    if (FAILED(hr)) {
        std::cerr << "Failed to connect to ROOT\\CIMV2 namespace: " << hr << std::endl;
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    hr = CoSetProxyBlanket(
        pSvc,                        // Indicates the proxy to set
        RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
        RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
        NULL,                        // Server principal name 
        RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
        RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
        NULL,                        // client identity
        EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hr)) {
        std::cerr << "Failed to set proxy blanket: " << hr << std::endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return 1;
    }

    // Update a specific property (example: Win32_DiskDrive SerialNumber)
    std::wstring newSerialNumber = L"1234567890ZA";
    UpdateDeviceInformation(pSvc, L"SELECT * FROM Win32_DiskDrive", L"SerialNumber", newSerialNumber);

    // Clean up 
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return 0;
}

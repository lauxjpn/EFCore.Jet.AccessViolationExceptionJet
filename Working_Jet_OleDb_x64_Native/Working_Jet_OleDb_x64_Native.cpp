#include <windows.h>
#include <strsafe.h>
#include <msdasc.h>
#include <stdio.h>

#define MAX_CONNECTION_STRING_LEN MAX_PATH + 128

#define RELEASE(pUnk) if (pUnk) { pUnk->Release(); pUnk = NULL; }

VOID RunNorthwindTest();

int main()
{
    //
    // Working_Jet_OleDb_x64_Native will always work and never result in a memory violation issue:
    //

    BOOL is64Bit = sizeof(VOID*) == 8;
    wprintf(L"Running as %s process.\r\n\r\n", (is64Bit ? L"x64" : L"x86"));
    
    if (is64Bit)
    {
        Sleep(1 * 1000);

        HRESULT hr;
        if (SUCCEEDED(hr = CoInitialize(NULL)))
        {
            RunNorthwindTest();
            CoUninitialize();
        }
    }
}

VOID RunNorthwindTest()
{
    HRESULT hr;

    IDataInitialize* pDataInitialize = NULL;
    if (SUCCEEDED(hr = CoCreateInstance(CLSID_MSDAINITIALIZE, NULL, CLSCTX_INPROC_SERVER, IID_IDataInitialize, (VOID**)&pDataInitialize)))
    {
        WCHAR wszPath[MAX_PATH + 1];
        if (GetFullPathNameW(L"..\\..\\Northwind.accdb", sizeof(wszPath) / sizeof(wszPath[0]), wszPath, NULL) > 0)
        {
            //
            // Use either `Microsoft.ACE.OLEDB.16.0` or `Microsoft.ACE.OLEDB.12.0`:
            //   - Microsoft.ACE.OLEDB.16.0: https://www.microsoft.com/en-us/download/details.aspx?id=54920
            //   - Microsoft.ACE.OLEDB.12.0: https://www.microsoft.com/en-us/download/details.aspx?id=13255
            //

            WCHAR wszConnectionString[MAX_CONNECTION_STRING_LEN];
            if (SUCCEEDED(hr = StringCchPrintfW(wszConnectionString, sizeof(wszConnectionString) / sizeof(wszConnectionString[0]), L"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=%s", wszPath)))
            {
                IDBInitialize* pDbInitialize = NULL;
                if (SUCCEEDED(hr = pDataInitialize->GetDataSource(NULL, CLSCTX_INPROC_SERVER, wszConnectionString, IID_IDBInitialize, (IUnknown**)&pDbInitialize)))
                {
                    if (SUCCEEDED(hr = pDbInitialize->Initialize()))
                    {
                        IDBCreateSession* pDbCreateSession = NULL;

                        if (SUCCEEDED(hr = pDbInitialize->QueryInterface(IID_IDBCreateSession, (VOID**)&pDbCreateSession)))
                        {
                            IUnknown* pSession = NULL;
                            if (SUCCEEDED(hr = pDbCreateSession->CreateSession(NULL, IID_IOpenRowset, &pSession)))
                            {
                                IDBCreateCommand* pDbCreateCommand = NULL;

                                if (SUCCEEDED(hr = pSession->QueryInterface(IID_IDBCreateCommand, (VOID**)&pDbCreateCommand)))
                                {
                                    for (int i = 0; i < 1000; i++)
                                    {
                                        wprintf(L"%.3d\r\n", i);

                                        IUnknown* pCommand = NULL;
                                        if (SUCCEEDED(hr = pDbCreateCommand->CreateCommand(NULL, IID_ICommand, &pCommand)))
                                        {
                                            ICommandText* pCommandText = NULL;
                                            if (SUCCEEDED(hr = pCommand->QueryInterface(IID_ICommandText, (VOID**)&pCommandText)))
                                            {
                                                PCWSTR pwszCommandText = L"SELECT [c].[Address]\n"
                                                    L"FROM[Customers] AS[c]\n"
                                                    L"WHERE[c].[City] = 'Berlin'\n"
                                                    L"UNION\n"
                                                    L"SELECT[c0].[Address]\n"
                                                    L"FROM[Customers] AS[c0]\n"
                                                    L"WHERE[c0].[City] = 'London'";

                                                if (SUCCEEDED(hr = pCommandText->SetCommandText(DBGUID_DEFAULT, pwszCommandText)))
                                                {
                                                    IUnknown* pRowset = NULL;
                                                    if (SUCCEEDED(hr = pCommandText->Execute(NULL, IID_IRowset, NULL, NULL, &pRowset)))
                                                    {
                                                        RELEASE(pRowset);
                                                    }
                                                }

                                                RELEASE(pCommandText);
                                            }

                                            RELEASE(pCommand);
                                        }

                                        if (SUCCEEDED(hr = pDbCreateCommand->CreateCommand(NULL, IID_ICommand, &pCommand)))
                                        {
                                            ICommandText* pCommandText = NULL;
                                            if (SUCCEEDED(hr = pCommand->QueryInterface(IID_ICommandText, (VOID**)&pCommandText)))
                                            {
                                                PCWSTR pwszCommandText = L"SELECT 1\n"
                                                    L"FROM[Customers] AS[c]";

                                                if (SUCCEEDED(hr = pCommandText->SetCommandText(DBGUID_DEFAULT, pwszCommandText)))
                                                {
                                                    IUnknown* pRowset = NULL;
                                                    if (SUCCEEDED(hr = pCommandText->Execute(NULL, IID_IRowset, NULL, NULL, &pRowset)))
                                                    {
                                                        RELEASE(pRowset);
                                                    }
                                                }

                                                RELEASE(pCommandText);
                                            }

                                            RELEASE(pCommand);
                                        }
                                    }
                                    
                                    RELEASE(pDbCreateCommand);
                                }

                                RELEASE(pSession);
                            }

                            RELEASE(pDbCreateSession);
                        }
                    }

                    RELEASE(pDbInitialize);
                }
            }
        }

        RELEASE(pDataInitialize);
    }
}

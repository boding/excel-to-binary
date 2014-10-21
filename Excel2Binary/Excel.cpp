/************************************************
 *
 * file  : Excel.cpp
 * author: bobding
 * date  : 2014-10-20
 * detail:
 *
************************************************/

#if defined(_WIN32) || defined(_WIN64)

#include "Excel.h"
#include "Log.h"

ADODB::_ConnectionPtr Excel::Open(const char* filePath)
{
    if (NULL == filePath || 0 == strlen(filePath))
    {
        LogError("[Excel::Open] failed, invalid file path.\n");
        return NULL;
    }

    CoInitialize(NULL);

    ADODB::_ConnectionPtr connectionPtr = NULL;
    HRESULT hr = connectionPtr.CreateInstance(__uuidof(ADODB::Connection));
    if (S_OK != hr)
    {
        LogError("[Excel::Open] failed, when creating connection instance, file: %s.\n", filePath);
        return NULL;
    }

    _bstr_t connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+_bstr_t(filePath)+";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'";
    connectionPtr->CursorLocation = ADODB::adUseClient;

    try
    {
        hr = connectionPtr->Open(connectionString, _bstr_t(""), _bstr_t(""), ADODB::adConnectUnspecified);
    }
    catch (_com_error& e)
    {
        LogError("[Excel::Open] exception: '%s' when opennig file: %s.\n", (const char*)e.Description(), filePath);
    }

    if (S_OK != hr)
    {
        LogError("[Excel::Open] failed, when openning file: %s.\n", filePath);
        return NULL;
    }

    return connectionPtr.Detach();
}

void Excel::Close(ADODB::_ConnectionPtr connectionPtr)
{
    try
    {
        connectionPtr->Close();
    }
    catch (_com_error& e)
    {
        LogError("[Excel::Close] exception: '%s' when closing connection.\n", (const char*)e.Description());
    }

    CoUninitialize();
}

ADODB::_RecordsetPtr Excel::Query(const char* sql, ADODB::_ConnectionPtr connectionPtr)
{
    ADODB::_RecordsetPtr recordsetPtr;
    _variant_t affected;

    try
    {
        recordsetPtr = connectionPtr->Execute(_bstr_t(sql), &affected, ADODB::adCmdText);

        while (ADODB::adStateExecuting == connectionPtr->GetState())
        {
            LogPrompt("[Excel::Query] querying, please wait...\n");
        }
    }
    catch (_com_error& e)
    {
        LogError("[Excel::Query] exception: '%s' when querying.\n", (const char*)e.Description());
    }

    return recordsetPtr.Detach();
}

#endif // _WIN32 || _WIN64
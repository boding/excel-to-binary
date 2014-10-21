#include "Excel.h"
#include "Log.h"
#include <string>

#define IntFromCollect(result, name, row)                       \
try                                                             \
{                                                               \
    result = (int)recordsetPtr->GetCollect(name);               \
}                                                               \
catch (_com_error&)                                             \
{                                                               \
    LogWarn("[exception] column: %s,\trow: %d.\n", name, row);  \
}

#define ShortFromCollect(result, name, row)                     \
try                                                             \
{                                                               \
    result = (short)recordsetPtr->GetCollect(name);             \
}                                                               \
catch (_com_error&)                                             \
{                                                               \
    LogWarn("[exception] column: %s,\trow: %d.\n", name, row);  \
}

#define StringFromCollect(result, length, name, row)            \
try                                                             \
{                                                               \
    _bstr_t s = recordsetPtr->GetCollect(name);                 \
    _snprintf(result, length - 1, "%s", (const char*)s);        \
}                                                               \
catch (_com_error&)                                             \
{                                                               \
    LogWarn("[exception] column: %s,\trow: %d.\n", name, row);  \
}

std::string MakeSqlString(const char* collectNames[], int count, const char* sheetName)
{
    std::string sqlString = "SELECT";
    for (int i = 0; i < count; ++i)
    {
        sqlString += std::string(" [") + collectNames[i] + "],";
    }

    sqlString[sqlString.length() - 1] = ' ';
    sqlString += std::string("FROM [") + sheetName + "$]";

    return sqlString;
}

void ConvertMap()
{
    LogCritical("[ConvertMap] convert map begin!\n");

    static const char* columnNames[] = { "赛道1", "赛道2", "赛道3", "赛道4", "赛道5", "上边道", "下边道" };
    static const int numColumns = sizeof(columnNames) / sizeof(columnNames[0]);
    static const std::string sqlString = MakeSqlString(columnNames, numColumns, "Sheet1");

    ADODB::_ConnectionPtr connectionPtr = Excel::Open("map.xlsx");
    if (0 == connectionPtr)
    {
        Excel::Close(connectionPtr);
        LogError("[ConvertMap] convert map failed!\n");
        getchar();
    }

    ADODB::_RecordsetPtr recordsetPtr = Excel::Query(sqlString.c_str(), connectionPtr);

    unsigned int count = recordsetPtr->GetRecordCount(), index = 0;
    short* data = new short[count * numColumns];
    memset(data, 0, count * numColumns * sizeof(short));
    recordsetPtr->MoveFirst();

    while (VARIANT_TRUE != recordsetPtr->GetADOEOF() && index < count)
    {
        for (int i = 0; i < numColumns; ++i)
        {
            ShortFromCollect(data[index * numColumns + i], columnNames[i], index + 2);
        }

        recordsetPtr->MoveNext();
        ++index;
    }

    char buffer[128];
    memset(buffer, 0, sizeof(buffer));
    recordsetPtr = Excel::Query("SELECT [地图] FROM [Sheet1$]", connectionPtr);
    recordsetPtr->MoveFirst();
    StringFromCollect(buffer, sizeof(buffer), "地图", 0);

    FILE* fp = fopen("map.bin", "wb");
    fwrite(buffer, sizeof(buffer), 1, fp);
    fwrite(&count, sizeof(int), 1, fp);
    fwrite(data, count * numColumns * sizeof(short), 1, fp);
    fclose(fp);

    delete[] data;
    Excel::Close(connectionPtr);
    LogCritical("[ConvertMap] convert map success!\n");
}

void ConvertActor()
{
    #pragma pack(1)
    struct Entry 
    {
        short Id;
        char Model[128];
        short Speed;
    };
    #pragma pack()

    LogCritical("[ConvertActor] convert actor begin!\n");

    static const char* columnNames[] = { "ID", "模型", "速度" };
    static const int numColumns = sizeof(columnNames) / sizeof(columnNames[0]);
    static const std::string sqlString = MakeSqlString(columnNames, numColumns, "Sheet1");

    ADODB::_ConnectionPtr connectionPtr = Excel::Open("actor.xlsx");
    if (0 == connectionPtr)
    {
        Excel::Close(connectionPtr);
        LogError("[ConvertActor] convert actor failed!\n");
    }

    ADODB::_RecordsetPtr recordsetPtr = Excel::Query(sqlString.c_str(), connectionPtr);

    unsigned int count = recordsetPtr->GetRecordCount(), index = 0;
    Entry* data = new Entry[count];
    memset(data, 0, count * sizeof(Entry));
    recordsetPtr->MoveFirst();

    while (VARIANT_TRUE != recordsetPtr->GetADOEOF() && index < count)
    {
        ShortFromCollect(data[index].Id, columnNames[0], index + 2);
        StringFromCollect(data[index].Model, sizeof(data[index].Model), columnNames[1], index + 2);
        ShortFromCollect(data[index].Speed, columnNames[2], index + 2);

        recordsetPtr->MoveNext();
        ++index;
    }

    FILE* fp = fopen("actor.bin", "wb");
    fwrite(&count, sizeof(int), 1, fp);
    fwrite(data, count * sizeof(Entry), 1, fp);
    fclose(fp);

    delete[] data;
    Excel::Close(connectionPtr);
    LogCritical("[ConvertActor] convert actor success!\n");
}

int main()
{
    ConvertMap();
    ConvertActor();

    getchar();

    return 0;
}
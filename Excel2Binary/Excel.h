/************************************************
 *
 * file  : Excel.h
 * author: bobding
 * date  : 2014-10-20
 * detail:
 *
************************************************/

#if defined(_WIN32) || defined(_WIN64)

#ifndef _EXCEL_H_
#define _EXCEL_H_

#pragma warning(disable:4146)

#include "msado15_i.h"

#include <string>
using std::string;

class Excel
{
public:
    static ADODB::_ConnectionPtr Open(const char* filePath);
    static void Close(ADODB::_ConnectionPtr connectionPtr);
    static ADODB::_RecordsetPtr Query(const char* sql, ADODB::_ConnectionPtr connectionPtr);
};

#endif // _EXCEL_H_

#endif // _WIN32 || _WIN64
// TestCM++.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <Windows.h>
#define INITGUID
#include <devpkey.h>
#include <propkey.h>
#include <functiondiscoverykeys_devpkey.h>

#include <iostream>

std::wostream &operator<<(std::wostream &os, GUID const &rhs)
{
    // Convert rhs to a string and write it to os
    wchar_t guidString[39];
    (void)StringFromGUID2(rhs, guidString, 39);
    os << guidString;
    return os;
}

// Stream opeartor overload for PROPERTYKEY
std::wostream &operator<<(std::wostream &os, PROPERTYKEY const &rhs)
{
    // Write the fmtid and pid to os
    os << rhs.fmtid << "," << rhs.pid;
    return os;
}

// Stream opeartor overload for DEVPROPKEY
std::wostream &operator<<(std::wostream &os, DEVPROPKEY const &rhs)
{
    // Write the fmtid and pid to os
    os << rhs.fmtid << "," << rhs.pid;
    return os;
}

int main()
{
    std::wcout << DEVPKEY_Device_Children.fmtid << L"," << DEVPKEY_Device_Children.pid << std::endl;
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file

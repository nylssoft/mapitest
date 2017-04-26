#include "stdafx.h"

#include <mapi.h>

#include <iostream>

bool test_logoff(const std::wstring & profile)
{
	try
	{
		mapi::mailsystem system;
		// MAPIInitialize with NT_SERVICE flag
		system.init(mapi::init_option::service);
		// MAPILogonEx
		auto session = system.logon(profile);
		// IMAPISession->Logoff
		session.release();
		// MAPIUninitialize
		system.release();
		return true;
	}
	catch (mapi::mapi_error & mapierror)
	{
		std::wcout << L"MAPI error: " << mapierror.str() << std::endl;
		return false;
	}
}

int wmain(int argc, wchar_t * args[])
{
	if (argc < 3)
	{
		std::wcout << L"Usage: mapitest.exe [logoff] <profile>" << std::endl;
		return -1;
	}
	std::wstring test = args[1];
	std::wstring profile = args[2];
	mapi::trace::set_listener(mapi::trace_console);
	bool failed = true;
	if (test == L"logoff")
	{
		failed = test_logoff(profile);
	}
	if (failed)
	{
		std::wcout << L"MAPI test failed!" << std::endl;
	}
    return 0;
}


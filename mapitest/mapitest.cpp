#include "stdafx.h"

#include <mapi.h>

#include <iostream>

bool test_init(const std::wstring & profile, mapi::init_option init_option)
{
	try
	{
		mapi::mailsystem system;
		// MAPIInitialize
		system.init(init_option);
		// MAPILogonEx
		auto session = system.logon(profile);
		// IMAPISession->Logoff, IMAPISession->Release
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
		std::wcout << L"Usage: mapitest.exe <test> <profile>" << std::endl;
		std::wcout << std::endl;
		std::wcout << L"<test> := init-service " << std::endl;
		std::wcout << L"<test> := init-noservice " << std::endl;
		std::wcout << std::endl;
		return -1;
	}
	std::wstring test = args[1];
	std::wstring profile = args[2];
	mapi::trace::set_listener(mapi::trace_console);
	bool ok = false;
	if (test == L"init-service")
	{
		ok = test_init(profile, mapi::init_option::service);
	}
	else if (test == L"init-noservice")
	{
		ok = test_init(profile, mapi::init_option::noservice);
	}
	std::wcout << L"MAPI test '" << test << L"' " << (ok ?  L"succeeded" : L"failed") << L"." << std::endl;
	return ok ? 0 : -1;
}


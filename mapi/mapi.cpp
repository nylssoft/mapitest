#include "stdafx.h"
#include "mapi.h"

#include <sstream>
#include <iostream>
#include <iomanip>

#include <MAPIX.h>

namespace mapi
{
#pragma region macros
#define TRACE(classname,method) trace trace(classname, method, __FILE__)
#define TRACE_ENTER(call) trace.enter(__LINE__, call)
#define TRACE_RESULT(hr) trace.result(__LINE__, hr)
#define TRACE_LEAVE trace.leave(__LINE__)
#pragma endregion

#pragma region session
	struct session::_pimpl_struct
	{
		_pimpl_struct() : pMAPISession(nullptr) {}
		IMAPISession * pMAPISession;
	};

	session::session() noexcept
	{
		_pimpl = std::make_unique<_pimpltype>();
	}

	session::session(const session & s) noexcept : session()
	{
		_pimpl->pMAPISession = s._pimpl->pMAPISession;
		if (_pimpl->pMAPISession != nullptr)
		{
			_pimpl->pMAPISession->AddRef();
		}
	}

	session::session(session && s) noexcept : session()
	{
		_pimpl->pMAPISession = s._pimpl->pMAPISession;
		s._pimpl->pMAPISession = nullptr;
	}

	session::~session() noexcept
	{
		release();
	}

	session::operator bool() const noexcept
	{
		return _pimpl->pMAPISession != nullptr;
	}

	session & session::operator = (const session & s) noexcept
	{
		if (_pimpl->pMAPISession != s._pimpl->pMAPISession)
		{
			release();
			_pimpl->pMAPISession = s._pimpl->pMAPISession;
			if (_pimpl->pMAPISession != nullptr)
			{
				_pimpl->pMAPISession->AddRef();
			}
		}
		return *this;
	}

	session & session::operator = (session && s) noexcept
	{
		if (_pimpl->pMAPISession != s._pimpl->pMAPISession)
		{
			release();
			_pimpl->pMAPISession = s._pimpl->pMAPISession;
			s._pimpl->pMAPISession = nullptr;
		}
		return *this;
	}

	void session::release() noexcept
	{
		if (_pimpl->pMAPISession != nullptr)
		{
			TRACE(L"session", L"release");
			TRACE_ENTER(L"IMAPISession::Logoff");
			HRESULT hr = _pimpl->pMAPISession->Logoff(0L, 0L, 0L);
			if (FAILED(hr)) TRACE_RESULT(hr);
			TRACE_LEAVE;
			TRACE_ENTER(L"IMAPISession::Release");
			ULONG cnt = _pimpl->pMAPISession->Release();
			if (cnt != 0L) TRACE_RESULT(cnt);
			TRACE_LEAVE;
			_pimpl->pMAPISession = nullptr;
		}
	}
#pragma endregion

#pragma region mailsystem
	struct mailsystem::_pimpl_struct
	{
		_pimpl_struct() : init(false), option(init_option::noservice) {}
		bool init;
		init_option option;
	};

	mailsystem::mailsystem() noexcept
	{
		_pimpl = std::make_unique<_pimpltype>();
	}

	mailsystem::~mailsystem() noexcept
	{
		release();
	}

	mailsystem::operator bool() const noexcept
	{
		return _pimpl->init;
	}

	void mailsystem::init(init_option opt)
	{
		TRACE(L"mailsystem", L"init");
		if (*this) throw std::logic_error("invalid state");
		ULONG serviceFlag = opt == init_option::service ? MAPI_NT_SERVICE : 0L;
		ULONG initFlags = MAPI_MULTITHREAD_NOTIFICATIONS | serviceFlag;
		MAPIINIT_0 mapiInit = { MAPI_INIT_VERSION, initFlags };
		TRACE_ENTER(L"::MAPIInitialize");
		HRESULT hr = ::MAPIInitialize(&mapiInit);
		if (FAILED(hr)) TRACE_RESULT(hr);
		TRACE_LEAVE;
		if (FAILED(hr)) throw mapi_error(hr, trace.what(), trace.context());
		_pimpl->option = opt;
		_pimpl->init = true;
	}

	session mailsystem::logon(const std::wstring & profile) const
	{
		TRACE(L"mailsystem", L"logon");
		if (!*this) throw std::logic_error("invalid state");
		session mapiSession;
		IMAPISession * pMAPISession = nullptr;
		ULONG serviceFlag = _pimpl->option == init_option::service ? MAPI_NT_SERVICE : 0L;
		FLAGS logonFlags = MAPI_EXTENDED | MAPI_NEW_SESSION | MAPI_UNICODE | MAPI_NO_MAIL | serviceFlag;
		TRACE_ENTER(L"::MAPILogonEx");
		HRESULT hr = ::MAPILogonEx(0L, (LPTSTR)profile.c_str(), nullptr, logonFlags, &pMAPISession);
		if (FAILED(hr)) TRACE_RESULT(hr);
		TRACE_LEAVE;
		if (FAILED(hr)) throw mapi_error(hr, trace.what(), trace.context());
		mapiSession._pimpl->pMAPISession = pMAPISession;
		return mapiSession;
	}

	void mailsystem::release() noexcept
	{
		if (_pimpl->init)
		{
			TRACE(L"mailsystem", L"release");
			TRACE_ENTER(L"::MAPIUninitialize");
			::MAPIUninitialize();
			TRACE_LEAVE;
			_pimpl->init = false;
		}
	}
#pragma endregion

#pragma region mapi trace
	bool trace::_enabled = false;

	trace::listener_type trace::_listener = trace::do_nothing;

	trace::trace(const wchar_t * classname, const wchar_t * method, const char * filename)
		: _classname(classname), _method(method), _filename(filename), _call(nullptr) {}

	void trace::enter(unsigned int lineno, const wchar_t * call)
	{
		_call = call;
		if (_enabled) _listener(_classname, _method, _filename, lineno, level::enter, _call, S_OK);
	}

	void trace::result(unsigned int lineno, long result)
	{
		if (_enabled) _listener(_classname, _method, _filename, lineno, level::result, _call, result);
	}

	void trace::leave(unsigned int lineno)
	{
		if (_enabled) _listener(_classname, _method, _filename, lineno, level::leave, _call, S_OK);
	}

	std::string trace::what() const noexcept
	{
		std::string ret;
		if (_call != nullptr)
		{
			std::wstring w(_call);
			ret = std::string(w.begin(), w.end());
		}
		return ret;
	}

	const wchar_t * trace::call() const noexcept { return _call; }

	const wchar_t * trace::method() const noexcept { return _method; }

	const wchar_t * trace::classname() const noexcept { return _classname; }

	std::wstring trace::context() const noexcept
	{
		std::wstring ret;
		if (_classname != nullptr && _method != nullptr)
		{
			ret = _classname;
			ret.append(L"::");
			ret.append(_method);
		}
		return ret;
	}

	void trace::set_listener(listener_type listener)
	{
		_enabled = true;
		_listener = listener;
	}
#pragma endregion

#pragma region mapi_error
	std::wstring mapi_error::str() const
	{
		std::wostringstream wss;
		wss << L"MAPI call '" << method() << L"' failed in '" << context() << L"'.";
		wss << L" Error code is " << hex(code()) << L". ";
		wss << last_error();
		return wss.str();
	}
#pragma endregion

#pragma region utilities
std::wstring hex(unsigned long value)
{
	std::wostringstream wss;
	wss << L"0x" << std::uppercase << std::hex << std::setw(8) << std::setfill(L'0') << value;
	return wss.str();
}

void trace_console(
	const wchar_t *			classname,
	const wchar_t *			method,
	const char *			filename,
	unsigned int			lineno,
	mapi::trace::level		level,
	const wchar_t *			call,
	long					result)
{
	std::wstring prefix;
	switch (level)
	{
	case mapi::trace::level::enter:
		prefix = L"ENTER ";
		break;
	case mapi::trace::level::leave:
		prefix = L"LEAVE ";
		break;
	case mapi::trace::level::result:
		prefix = L"RESULT";
		break;
	default:
		break;
	}
	std::wcout << L"TRACE " << prefix << L" " << call;
	if (level == mapi::trace::level::result)
	{
		std::wcout << L" => " << mapi::hex(result);
	}
	std::wcout << " <";
	std::wcout << classname << L"::" << method << L",";
	std::wcout << filename << "," << lineno << ">" << std::endl;
}
#pragma endregion
};

#pragma once

#include <memory>
#include <functional>
#include <string>

namespace mapi
{
#pragma region session
	class session
	{
	public:
		session() noexcept;
		session(const session & s) noexcept;
		session(session && s) noexcept;

		~session() noexcept;

		session & operator = (const session & s) noexcept;
		session & operator = (session && s) noexcept;

		explicit operator bool() const noexcept;

		void release() noexcept;


	private:
		struct _pimpl_struct;
		using _pimpltype = struct _pimpl_struct;

		std::unique_ptr<_pimpltype> _pimpl;

		friend class mailsystem;
	};
#pragma endregion

#pragma region mailsystem
	enum class init_option { service, noservice };

	class mailsystem
	{
	public:
		mailsystem() noexcept;

		~mailsystem() noexcept;

		explicit operator bool() const noexcept;

		void release() noexcept;

		void init(init_option t);

		session logon(const std::wstring & profile) const;

	private:
		struct _pimpl_struct;
		using _pimpltype = struct _pimpl_struct;

		std::unique_ptr<_pimpltype> _pimpl;

		mailsystem(mailsystem & s) = delete;
	};
#pragma endregion

#pragma region mapi_error
	class mapi_error : public std::exception
	{
	public:
		mapi_error() noexcept : _code(0L) {}

		mapi_error(const mapi_error & mapierror) noexcept
			: exception(mapierror.what()), _code(mapierror._code), _context(mapierror._context), _last_error(mapierror._last_error) {}

		explicit mapi_error(long code, const std::string & what) noexcept
			: exception(what.c_str()), _code(code) {}

		explicit mapi_error(long code, const std::string & what, const std::wstring & context) noexcept
			: exception(what.c_str()), _code(code), _context(context) {}

		explicit mapi_error(long code, const std::string & what, const std::wstring & context, const std::wstring & last_error) noexcept
			: exception(what.c_str()), _code(code), _context(context), _last_error(last_error) {}

		long code() const { return _code; }

		std::wstring str() const;

		std::wstring method() const
		{
			if (what() != nullptr)
			{
				std::string w = what();
				return std::wstring(w.begin(), w.end());
			}
			return std::wstring();
		}

		const std::wstring & context() const { return _context; }

		const std::wstring & last_error() const { return _last_error; }

	private:
		long _code;
		std::wstring _context;
		std::wstring _last_error;
	};
#pragma endregion

#pragma region trace
	class trace
	{
	public:
		enum class level { enter, leave, result };

		using listener_type = std::function<void(
			const wchar_t * classname,
			const wchar_t * method,
			const char *	filename,
			unsigned int	lineno,
			level			lev,
			const wchar_t *	mapicall,
			long			result)>;

		trace() = delete;
		trace(trace &) = delete;
		explicit trace(const wchar_t * classname, const wchar_t * method, const char * filename);

		void enter(unsigned int lineno, const wchar_t * call);

		void result(unsigned int lineno, long result);

		void leave(unsigned int lineno);

		std::string what() const noexcept;

		const wchar_t * call() const noexcept;

		const wchar_t * method() const noexcept;

		const wchar_t * classname() const noexcept;

		std::wstring context() const noexcept;

		static void set_listener(listener_type listener);

		static void do_nothing(const wchar_t *, const wchar_t *, const char *, unsigned int, level, const wchar_t *, long) {};

	private:
		static listener_type _listener;
		static bool _enabled;
		const wchar_t * _classname;
		const wchar_t * _method;
		const char * _filename;
		const wchar_t * _call;
	};
#pragma endregion

#pragma region utilities
	std::wstring hex(unsigned long value);

	void trace_console(
		const wchar_t *			classname,
		const wchar_t *			method,
		const char *			filename,
		unsigned int			lineno,
		mapi::trace::level		level,
		const wchar_t *			call,
		long					result);
#pragma endregion
}

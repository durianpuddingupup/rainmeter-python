/*
  Copyright (C) 2013 Johannes Blume

  This program is free software; you can redistribute it and/or
  modify it under the terms of the GNU General Public License
  as published by the Free Software Foundation; either version 2
  of the License, or (at your option) any later version.

  This program is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
  GNU General Public License for more details.

  You should have received a copy of the GNU General Public License
  along with this program; if not, write to the Free Software
  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
*/
#define PY_SSIZE_T_CLEAN
#include <Windows.h>
#include <Python.h>
#include <string>
#include "RainmeterAPI.h"

PyObject* CreateRainmeterObject(void *rm);
static PyTypeObject rainmeterType;
struct Measure;
static int active = 0;

void RedirectStdErr()
{
	try
	{
		PyObject* iomodule = PyImport_AddModule("io"); // Borrowed reference.
		if (iomodule == NULL) throw L"Cannot import `io`";
		PyObject* stringIOClass = PyObject_GetAttrString(iomodule, "StringIO");
		if (stringIOClass == NULL) throw L"Cannot get `io.StringIO`";
		PyObject* stringIO = PyObject_CallNoArgs(stringIOClass);
		Py_DECREF(stringIOClass);
		if (stringIO == NULL) throw L"Cannot create a `io.StringIO` object";
		int result = PySys_SetObject("stderr", stringIO);
		Py_DECREF(stringIO);
		if (result == -1) throw L"Cannot set `sys.stderr`";
	}
	catch (wchar_t* error)
	{
		PyErr_Clear();
		RmLog(LOG_ERROR, (std::wstring(L"Error redirecting `sys.stderr`: ") + error).c_str());
	}
}

void LogPythonError()
{
	try
	{
		if (PyErr_Occurred() == NULL) return;
		PyErr_Print();
		PyObject* stdErr = PySys_GetObject("stderr"); // Borrowed reference.
		if (stdErr == NULL) throw L"Cannot get `sys.stderr`";
		PyObject* getvalue = PyObject_GetAttrString(stdErr, "getvalue");
		if (getvalue == NULL) throw L"Cannot get `getvalue`";
		PyObject* value = PyObject_CallNoArgs(getvalue);
		Py_DECREF(getvalue);
		if (value == NULL) throw L"Cannot get exception string";
		wchar_t* str = PyUnicode_AsWideCharString(value, NULL);
		Py_DECREF(value);
		RmLog(LOG_ERROR, str);
		PyMem_Free((void*)str);
		RedirectStdErr();
	}
	catch (wchar_t* error)
	{
		PyErr_Clear();
		RmLog(LOG_ERROR, (std::wstring(L"Error logging an error in python scripts: ") + error).c_str());
	}
}

void AddDirToPath(LPCWSTR dir)
{
	PyObject* pathObj = PySys_GetObject("path"); // Borrowed reference.
	PyObject* scriptDirObj = PyUnicode_FromWideChar(dir, -1);
	if (!PySequence_Contains(pathObj, scriptDirObj))
	{
		PyList_Append(pathObj, scriptDirObj);
	}
	Py_DECREF(scriptDirObj);
}

PyObject* LoadScript(LPCWSTR scriptPath, char* fileName, LPCWSTR className)
{
	try
	{
		FILE* f = _Py_wfopen(scriptPath, L"r");
		if (f == NULL) throw L"Cannot open Python script";

		PyObject* globals = PyModule_GetDict(PyImport_AddModule("__main__")); // Borrowed reference.

		PyObject* result = PyRun_FileEx(f, fileName, Py_file_input, globals, globals, 1);
		if (result == NULL) throw L"";
		Py_DECREF(result);

		PyObject* classNameObj = PyUnicode_FromWideChar(className, -1);
		if (classNameObj == NULL) throw L"";

		PyObject* classObj = PyDict_GetItemWithError(globals, classNameObj); // Borrowed reference.
		if (classObj == NULL && PyErr_Occurred() == NULL)
			_PyErr_SetKeyError(classNameObj);
		Py_DECREF(classNameObj);
		if (classObj == NULL) throw L"";

		PyObject* measureObj = PyObject_CallNoArgs(classObj);
		if (measureObj == NULL) throw L"";

		return measureObj;
	}
	catch (wchar_t* error)
	{
		if (wcscmp(error, L"")) RmLog(LOG_ERROR, error);
		LogPythonError();
		return nullptr;
	}
}

struct Measure
{
	PyObject* measureObject;
	LPCWSTR pythonHome;
	LPCWSTR scriptPath;
	LPCWSTR className;
	bool initialized;
	Measure(void* rm) : measureObject(NULL), pythonHome(L""), scriptPath(L""), className(L""), initialized(false)
	{
		Reload(rm);
	}
	void Reload(void* rm)
	{
		LPCWSTR _pythonHome = RmReadString(rm, L"PythonHome", NULL, FALSE);
		if (_pythonHome != NULL && wcscmp(_pythonHome, L"") && wcscmp(_pythonHome, pythonHome))
		{
			pythonHome = _pythonHome;
			if (pythonHome != NULL && !wcscmp(pythonHome, L""))
			{
				Py_SetPythonHome(pythonHome);
			}
		}
		LPCWSTR _scriptPath = RmReadPath(rm, L"ScriptPath", L"default.py");
		wchar_t scriptBaseName[_MAX_FNAME];
		wchar_t scriptExt[_MAX_EXT];
		wchar_t scriptDir[_MAX_DIR];
		_wsplitpath_s(_scriptPath, NULL, 0, scriptDir, _MAX_DIR, scriptBaseName, _MAX_FNAME, scriptExt, _MAX_EXT);
		AddDirToPath(scriptDir);

		wchar_t fileName[_MAX_FNAME + 1 + _MAX_EXT];
		lstrcpyW(fileName, scriptBaseName);
		lstrcatW(fileName, scriptExt);
		char fileNameMb[_MAX_FNAME + 1 + _MAX_EXT];
		wcstombs_s(NULL, fileNameMb, sizeof(fileNameMb), fileName, sizeof(fileName));
		LPCWSTR _className = RmReadString(rm, L"ClassName", L"Measure", FALSE);
		if (wcscmp(_scriptPath, scriptPath) || wcscmp(_className, className))
		{
			scriptPath = _scriptPath;
			className = _className;
			measureObject = LoadScript(scriptPath, fileNameMb, className);
		}
	}
	~Measure()
	{
		Py_XDECREF(measureObject);
	}
};

PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
	++active;
	if (!Py_IsInitialized())
	{
		Py_Initialize();
		RedirectStdErr();
	}
	PyGILState_STATE gstate = PyGILState_Ensure();
	Measure* measure = new Measure(rm);
	*data = measure;

	if (measure->measureObject != NULL)
	{
		if (PyObject_HasAttrString(measure->measureObject, "Initialize"))
		{
			PyObject* rainmeterObject = CreateRainmeterObject(rm);
			if (rainmeterObject == NULL)
			{
				LogPythonError();
			}
			else
			{
				PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "Initialize", "O", rainmeterObject);
				if (resultObj == NULL) LogPythonError(); else Py_DECREF(resultObj);
				Py_DECREF(rainmeterObject);
			}
		}
	}
	PyGILState_Release(gstate);
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
	Measure *measure = (Measure*) data;
	PyGILState_STATE gstate = PyGILState_Ensure();
	measure->Reload(rm);

	if (measure->measureObject != NULL)
	{
		if (PyObject_HasAttrString(measure->measureObject, "Reload"))
		{
			PyObject* rainmeterObject = CreateRainmeterObject(rm);
			if (rainmeterObject == NULL)
			{
				LogPythonError();
			}
			else
			{
				PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "Reload", "Od", rainmeterObject, maxValue);
				if (resultObj == NULL) LogPythonError(); else Py_DECREF(resultObj);
				Py_DECREF(rainmeterObject);
			}
		}
	}

	PyGILState_Release(gstate);
}

PLUGIN_EXPORT double Update(void* data)
{
	Measure *measure = (Measure*) data;
	if (measure->measureObject == NULL)
	{
		return 0.0;
	}
	PyGILState_STATE gstate = PyGILState_Ensure();
	double result = 0.0;
	if (PyObject_HasAttrString(measure->measureObject, "Update"))
	{
		PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "Update", NULL);
		if (resultObj != NULL)
		{
			result = PyFloat_Check(resultObj) ? PyFloat_AsDouble(resultObj) : 0.0;
			Py_DECREF(resultObj);
		}
		else
		{
			LogPythonError();
		}
	}
	PyGILState_Release(gstate);
	return result;
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
	Measure *measure = (Measure*) data;
	if (measure->measureObject == NULL)
	{
		return nullptr;
	}

	PyGILState_STATE gstate = PyGILState_Ensure();
	wchar_t* result = nullptr;
	if (PyObject_HasAttrString(measure->measureObject, "GetString"))
	{
		PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "GetString", NULL);
		if (resultObj != NULL)
		{
			if (resultObj != Py_None)
			{
				PyObject* strObj = PyObject_Str(resultObj);
				wchar_t* _result = PyUnicode_AsWideCharString(strObj, NULL);
				size_t len = wcslen(_result) + 1;
				result = (wchar_t*)malloc(sizeof(wchar_t) * len);
				if (result != NULL)
					wcscpy_s(result, len, _result);
				else
					RmLog(LOG_ERROR, L"Error allocating memory for result string.");
				PyMem_Free(_result);
				Py_DECREF(strObj);
			}
			Py_DECREF(resultObj);
		}
		else
		{
			LogPythonError();
		}
	}
	PyGILState_Release(gstate);
	return result;
}

PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
{
	Measure *measure = (Measure*) data;
	if (measure->measureObject == NULL)
	{
		return;
	}

	PyGILState_STATE gstate = PyGILState_Ensure();
	if (PyObject_HasAttrString(measure->measureObject, "ExecuteBang"))
	{
		PyObject* argsObj = PyUnicode_FromWideChar(args, -1);
		if (argsObj == NULL)
		{
			LogPythonError();
		}
		else
		{
			PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "ExecuteBang", "O", argsObj);
			Py_DECREF(argsObj);
			if (resultObj != NULL)
				Py_DECREF(resultObj);
			else
				LogPythonError();
		}
	}
	PyGILState_Release(gstate);
}

PLUGIN_EXPORT void Finalize(void* data)
{
	Measure *measure = (Measure*) data;
	PyGILState_STATE gstate = PyGILState_Ensure();
	if (measure->measureObject != NULL)
		if (PyObject_HasAttrString(measure->measureObject, "Finalize"))
		{
			PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "Finalize", NULL);
			if (resultObj != NULL)
			{
				Py_DECREF(resultObj);
			}
			else
			{
				LogPythonError();
			}
		}
	delete measure;
	if ((--active) <= 0)
	{
		if (active < 0)
		{
			RmLog(LOG_ERROR, L"active < 0");
			active = 0;
		}
		// if (Py_FinalizeEx()) RmLog(LOG_ERROR, L"Py_FinalizeEx failed."); // Python causes memory leaking issues... so...
	}
}

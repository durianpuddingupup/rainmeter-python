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

#include <Windows.h>
#include <Python.h>
#include <string>
#include "RainmeterAPI.h"

PyObject* CreateRainmeterObject(void *rm);
struct Measure;

PyThreadState *mainThreadState = NULL;

void RedirectStdErr()
{
	try
	{
		PyObject* iomodule = PyImport_ImportModule("io");
		if (iomodule == NULL) throw L"Cannot import `io`";
		PyObject* stringIOClass = PyObject_GetAttrString(iomodule, "StringIO");
		Py_DECREF(iomodule);
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
		PyObject* stdErr = PySys_GetObject("stderr"); // `PySys_GetObject` Return value: Borrowed reference.
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
	PyObject* pathObj = PySys_GetObject("path");
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
		if (f == NULL) throw L"Error opening Python script";

		PyObject* globals = PyModule_GetDict(PyImport_AddModule("__main__"));

		PyObject* result = PyRun_FileEx(f, fileName, Py_file_input, globals, globals, 1);
		if (result == NULL) throw;
		Py_DECREF(result);

		PyObject* classNameObj = PyUnicode_FromWideChar(className, -1);
		if (classNameObj == NULL) throw;

		PyObject* classObj = PyDict_GetItem(globals, classNameObj);
		Py_DECREF(classNameObj);
		if (classObj == NULL) throw;

		PyObject* measureObj = PyObject_CallNoArgs(classObj);
		if (measureObj == NULL) throw;

		return measureObj;
	}
	catch (wchar_t* error)
	{
		if (error != NULL) RmLog(LOG_ERROR, error);
		LogPythonError();
		return nullptr;
	}
}

struct Measure
{
	Measure(void* rm)
	{
		scriptPath = RmReadPath(rm, L"ScriptPath", L"default.py");
		wchar_t scriptBaseName[_MAX_FNAME];
		wchar_t scriptExt[_MAX_EXT];
		wchar_t scriptDir[_MAX_DIR];
		_wsplitpath_s(scriptPath, NULL, 0, scriptDir, _MAX_DIR, scriptBaseName, _MAX_FNAME, scriptExt, _MAX_EXT);
		AddDirToPath(scriptDir);

		wchar_t fileName[_MAX_FNAME + 1 + _MAX_EXT];
		lstrcpyW(fileName, scriptBaseName);
		lstrcatW(fileName, L".");
		lstrcatW(fileName, scriptExt);
		char fileNameMb[_MAX_FNAME + 1 + _MAX_EXT];
		wcstombs_s(NULL, fileNameMb, sizeof(fileNameMb), fileName, sizeof(fileName));
		className = RmReadString(rm, L"ClassName", L"Measure", FALSE);
		measureObject = LoadScript(scriptPath, fileNameMb, className);
	}

	PyObject *measureObject;
	LPCWSTR scriptPath;
	LPCWSTR className;
	~Measure()
	{
		Py_XDECREF(measureObject);
	}
};

PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
	LPCWSTR pythonHome = RmReadString(rm, L"PythonHome", NULL, FALSE);
	if (pythonHome != NULL && !wcscmp(pythonHome, L""))
	{
		Py_SetPythonHome(pythonHome);
	}
	Py_Initialize();
	RedirectStdErr();

	Measure* measure = new Measure(rm);
	*data = measure;

	if (measure->measureObject != NULL)
	{
		PyObject* rainmeterObject = CreateRainmeterObject(rm);
		if (PyObject_HasAttrString(measure->measureObject, "Initialize"))
		{
			PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "Initialize", "O", rainmeterObject);
			if (resultObj == NULL) LogPythonError(); else Py_DECREF(resultObj);
		}
		Py_DECREF(rainmeterObject);
	}
	mainThreadState = PyEval_SaveThread();
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
	Measure *measure = (Measure*) data;
	PyEval_RestoreThread(mainThreadState);

	if (RmReadInt(rm, L"DynamicVariables", 0))
	{
		// TODO
	}

	if (measure->measureObject != NULL)
	{
		PyObject *rainmeterObject = CreateRainmeterObject(rm);
		if (PyObject_HasAttrString(measure->measureObject, "Reload"))
		{
			PyObject* resultObj = PyObject_CallMethod(measure->measureObject, "Reload", "Od", rainmeterObject, maxValue);
			if (resultObj == NULL) LogPythonError(); else Py_DECREF(resultObj);
		}
		Py_DECREF(rainmeterObject);
	}

	PyEval_SaveThread();
}

PLUGIN_EXPORT double Update(void* data)
{
	Measure *measure = (Measure*) data;
	if (measure->measureObject == NULL)
	{
		return 0.0;
	}
	PyEval_RestoreThread(mainThreadState);
	PyObject *resultObj = PyObject_CallMethod(measure->measureObject, "Update", NULL);
	double result = 0.0;
	if (resultObj != NULL)
	{
		result = PyFloat_Check(resultObj) ? PyFloat_AsDouble(resultObj) : 0.0;
		Py_DECREF(resultObj);
	}
	else
	{
		LogPythonError();
	}
	PyEval_SaveThread();
	return result;
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
	Measure *measure = (Measure*) data;
	if (measure->measureObject == NULL)
	{
		return nullptr;
	}

	PyEval_RestoreThread(mainThreadState);
	PyObject *resultObj = PyObject_CallMethod(measure->measureObject, "GetString", NULL);
	wchar_t* result = nullptr;
	if (resultObj != NULL)
	{
		if (resultObj != Py_None)
		{
			PyObject *strObj = PyObject_Str(resultObj);
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
	PyEval_SaveThread();
	return result;
}

PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args)
{
	Measure *measure = (Measure*) data;
	if (measure->measureObject == NULL)
	{
		return;
	}

	PyEval_RestoreThread(mainThreadState);
	PyObject *argsObj = PyUnicode_FromWideChar(args, -1);
	PyObject *resultObj = PyObject_CallMethod(measure->measureObject, "ExecuteBang", "O", argsObj);
	if (resultObj != NULL)
	{
		Py_DECREF(resultObj);
	}
	else
	{
		LogPythonError();
	}
	Py_DECREF(argsObj);
	PyEval_SaveThread();
}

PLUGIN_EXPORT void Finalize(void* data)
{
	Measure *measure = (Measure*) data;
	PyEval_RestoreThread(mainThreadState);
	if (measure->measureObject != NULL)
	{
		PyObject *resultObj = PyObject_CallMethod(measure->measureObject, "Finalize", NULL);
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
	Py_Finalize();
}

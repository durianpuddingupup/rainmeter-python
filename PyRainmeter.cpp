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
#include "structmember.h"
#include "rainmeter-plugin-sdk/API/RainmeterAPI.h"

typedef struct RainmeterObject {
	PyObject_HEAD
	PyObject* rm;
	int logError;
	int logWarning;
	int logNotice;
	int logDebug;

	RainmeterObject() : logError(LOG_ERROR), logWarning(LOG_WARNING), logNotice(LOG_NOTICE), logDebug(LOG_DEBUG) {}
} RainmeterObject;

static PyObject *Rainmeter_RmReadString(RainmeterObject *self, PyObject *args)
{
	PyObject *option, *defValue;
	int replaceMeasures;
	PyArg_ParseTuple(args, "UUp", &option, &defValue, &replaceMeasures);
	wchar_t *optionStr = PyUnicode_AsWideCharString(option, NULL);
	wchar_t* defValueStr = PyUnicode_AsWideCharString(defValue, NULL);
	void* rm = PyCapsule_GetPointer(self->rm, NULL);
	if (rm == NULL) return NULL;
	LPCWSTR result = RmReadString(rm, optionStr, defValueStr, replaceMeasures);
	PyMem_Free(defValueStr);
	PyMem_Free(optionStr);
	if (result == NULL)
	{
		Py_INCREF(Py_None);
		return Py_None;
	}
	return PyUnicode_FromWideChar(result, -1);
}

static PyObject *Rainmeter_RmReadPath(RainmeterObject *self, PyObject *args)
{
	PyObject *option, *defValue;
	PyArg_ParseTuple(args, "UU", &option, &defValue);
	wchar_t *optionStr = PyUnicode_AsWideCharString(option, NULL);
	wchar_t *defValueStr = PyUnicode_AsWideCharString(defValue, NULL);
	void* rm = PyCapsule_GetPointer(self->rm, NULL);
	if (rm == NULL) return NULL;
	LPCWSTR result = RmReadPath(rm, optionStr, defValueStr);
	PyMem_Free(defValueStr);
	PyMem_Free(optionStr);
	if (result == NULL)
	{
		Py_INCREF(Py_None);
		return Py_None;
	}
	return PyUnicode_FromWideChar(result, -1);
}

static PyObject *Rainmeter_RmReadDouble(RainmeterObject *self, PyObject *args)
{
	PyObject *option;
	double defValue;
	PyArg_ParseTuple(args, "Ud", &option, &defValue);
	wchar_t *optionStr = PyUnicode_AsWideCharString(option, NULL);
	void* rm = PyCapsule_GetPointer(self->rm, NULL);
	if (rm == NULL) return NULL;
	double result = RmReadDouble(rm, optionStr, defValue);
	PyMem_Free(optionStr);
	return PyFloat_FromDouble(result);
}

static PyObject *Rainmeter_RmReadInt(RainmeterObject *self, PyObject *args)
{
	PyObject *option;
	int defValue;
	PyArg_ParseTuple(args, "Ui", &option, &defValue);
	wchar_t *optionStr = PyUnicode_AsWideCharString(option, NULL);
	void* rm = PyCapsule_GetPointer(self->rm, NULL);
	if (rm == NULL) return NULL;
	int result = RmReadInt(rm, optionStr, defValue);
	PyMem_Free(optionStr);
	return PyLong_FromLong(result);
}

static PyObject *Rainmeter_RmGetMeasureName(RainmeterObject *self)
{
	void* rm = PyCapsule_GetPointer(self->rm, NULL);
	if (rm == NULL) return NULL;
	LPCWSTR result = RmGetMeasureName(rm);
	if (result == NULL)
	{
		Py_INCREF(Py_None);
		return Py_None;
	}
	return PyUnicode_FromWideChar(result, -1);
}

static PyObject *Rainmeter_RmExecute(RainmeterObject *self, PyObject *args)
{
	PyObject *command;
	if (!PyArg_ParseTuple(args, "U", &command)) return NULL;
	wchar_t *commandStr = PyUnicode_AsWideCharString(command, NULL);
	if (commandStr == NULL) return NULL;
	void* rm = PyCapsule_GetPointer(self->rm, NULL);
	if (rm == NULL) return NULL;
	RmExecute(RmGetSkin(rm), commandStr);
	PyMem_Free(commandStr);
	Py_INCREF(Py_None);
	return Py_None;
}

static PyObject *Rainmeter_RmLog(RainmeterObject *self, PyObject *args)
{
	int level;
	PyObject *message;
	PyArg_ParseTuple(args, "iU", &level, &message);
	wchar_t *messageStr = PyUnicode_AsWideCharString(message, NULL);
	RmLog(level, messageStr);
	PyMem_Free(messageStr);
	Py_INCREF(Py_None);
	return Py_None;
}

static PyObject* Rainmeter_RmGetSkinWindow(RainmeterObject* self)
{
	try
	{
		PyObject* wintypesModule = PyImport_AddModule("ctypes.wintypes"); // Borrowed reference.
		if (wintypesModule == NULL) throw "";
		PyObject* HWNDType = PyObject_GetAttrString(wintypesModule, "HWND");
		void* rm = PyCapsule_GetPointer(self->rm, NULL);
		if (rm == NULL) return NULL;
		HWND _window = RmGetSkinWindow(rm);
		PyObject* window = PyLong_FromLongLong((long long)_window);
		PyObject* result;
		if (HWNDType != NULL && window != NULL)
			result = PyObject_CallOneArg(HWNDType, window);
		Py_XDECREF(HWNDType);
		Py_XDECREF(window);
		if (HWNDType == NULL || window == NULL) throw "";
		return result;
	}
	catch (char* error)
	{
		PyErr_SetString(PyExc_RuntimeError, error ? error : "Cannot get skin window.");
		return NULL;
	}
}

static PyObject* Rainmeter_new(PyTypeObject* type, PyObject* args, PyObject* kwds)
{
	RainmeterObject* self;
	PyObject* rm;
	if (!PyArg_ParseTuple(args, "O", &rm)) // Borrowed reference.
		return NULL;
	self = (RainmeterObject*)type->tp_alloc(type, 0);
	Py_INCREF(rm);
	self->rm = rm;
	return (PyObject*)self;
}

static void Rainmeter_dealloc(RainmeterObject* self)
{
	Py_XDECREF(self->rm);
	Py_TYPE(self)->tp_free((PyObject*)self);
}

static PyMethodDef Rainmeter_methods[] = {
	{"RmReadString", (PyCFunction)Rainmeter_RmReadString, METH_VARARGS, ""},
	{"RmReadPath", (PyCFunction)Rainmeter_RmReadPath, METH_VARARGS, ""},
	{"RmReadDouble", (PyCFunction)Rainmeter_RmReadDouble, METH_VARARGS, ""},
	{"RmReadInt", (PyCFunction)Rainmeter_RmReadInt, METH_VARARGS, ""},
	{"RmGetMeasureName", (PyCFunction)Rainmeter_RmGetMeasureName, METH_NOARGS, ""},
	{"RmExecute", (PyCFunction)Rainmeter_RmExecute, METH_VARARGS, ""},
	{"RmLog", (PyCFunction)Rainmeter_RmLog, METH_VARARGS, ""},
	{"RmGetSkinWindow", (PyCFunction)Rainmeter_RmGetSkinWindow, METH_NOARGS, ""},
	{NULL}
};

static PyMemberDef Rainmeter_members[] = {
	{"LOG_ERROR", T_INT, offsetof(RainmeterObject, logError), READONLY, ""},
	{"LOG_WARNING", T_INT, offsetof(RainmeterObject, logWarning), READONLY, ""},
	{"LOG_NOTICE", T_INT, offsetof(RainmeterObject, logNotice), READONLY, ""},
	{"LOG_DEBUG", T_INT, offsetof(RainmeterObject, logDebug), READONLY, ""},
	{NULL}
};

static PyTypeObject rainmeterType = {
	PyVarObject_HEAD_INIT(NULL, 0)
	"Rainmeter",               /* tp_name */
	sizeof(RainmeterObject),   /* tp_basicsize */
	0,                         /* tp_itemsize */
	(destructor) Rainmeter_dealloc, /* tp_dealloc */
	0,                         /* tp_print */
	0,                         /* tp_getattr */
	0,                         /* tp_setattr */
	0,                         /* tp_reserved */
	0,                         /* tp_repr */
	0,                         /* tp_as_number */
	0,                         /* tp_as_sequence */
	0,                         /* tp_as_mapping */
	0,                         /* tp_hash  */
	0,                         /* tp_call */
	0,                         /* tp_str */
	0,                         /* tp_getattro */
	0,                         /* tp_setattro */
	0,                         /* tp_as_buffer */
	Py_TPFLAGS_DEFAULT,        /* tp_flags */
	"Rainmeter objects",       /* tp_doc */
	0,                         /* tp_traverse */
	0,                         /* tp_clear */
	0,                         /* tp_richcompare */
	0,                         /* tp_weaklistoffset */
	0,                         /* tp_iter */
	0,                         /* tp_iternext */
	Rainmeter_methods,         /* tp_methods */
	Rainmeter_members,         /* tp_members */
	0,                         /* tp_getset */
	0,                         /* tp_base */
	0,                         /* tp_dict */
	0,                         /* tp_descr_get */
	0,                         /* tp_descr_set */
	0,                         /* tp_dictoffset */
	0,                         /* tp_init */
	0,                         /* tp_alloc */
	Rainmeter_new,             /* tp_new */
};

PyObject* CreateRainmeterObject(void* rm)
{
	if (PyType_Ready(&rainmeterType) < 0)
	{
		PyErr_SetString(PyExc_RuntimeError, "Cannot create rainmeter object.");
		return NULL;
	}
	PyObject* rmCap = PyCapsule_New(rm, NULL, NULL);
	if (rmCap == NULL) return NULL;
	RainmeterObject* rmObj = PyObject_New(RainmeterObject, &rainmeterType);
	if (rmObj != NULL) rmObj->rm = rmCap; else Py_DECREF(rmCap);
	return (PyObject*)rmObj;
}
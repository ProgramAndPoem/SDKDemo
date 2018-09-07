#ifndef __GS_OLE_HELP_H__
#define __GS_OLE_HELP_H__

#include "windows.h"
#include <stdio.h>
#include <TCHAR.h>

namespace gs
{
//double
VARIANT ToVariant(double value);
// int
VARIANT ToVariant(int value);


// long
VARIANT ToVariant(LONG value);


// unsigned long
VARIANT ToVariant(ULONG value);

// unsigned long long
VARIANT ToVariant(unsigned long long value);


// unsigned int
VARIANT ToVariant(unsigned int value);

// DWORD*
VARIANT ToVariant(DWORD* value);

// LONG*
VARIANT ToVariant(LONG* value);

// void*
VARIANT ToVariant(void* value);


// BSTR
VARIANT ToVariant(BSTR value);


// SAFEARRAY*
VARIANT ToVariant(SAFEARRAY* value);


// VARIANT*
VARIANT ToVariant(VARIANT* value);

// IDispatch
VARIANT ToVariant(IDispatch* value);

// double
void FromVariant(const VARIANT& v, double& value);

// char
void FromVariant(const VARIANT& v, char& value);

// int
void FromVariant(const VARIANT& v, int& value);

// long
void FromVariant(const VARIANT& v, LONG& value);


// UCHAR
void FromVariant(const VARIANT& v, UCHAR& value);


// unsigned int
void FromVariant(const VARIANT& v, unsigned int& value);

// ULONG
void FromVariant(const VARIANT& v, ULONG& value);

// BSTR
void FromVariant(const VARIANT& v, BSTR& value);


// VARIANT*
void FromVariant(const VARIANT& v, VARIANT*& value);

// IDispatch
void FromVariant(const VARIANT& v, IDispatch* &value);

template <class T>
void FromVariant(const VARIANT& v, T* &value)
{
	value = (T*) v.pdispVal;
}


//------------------------------------------------------------------------------


HRESULT OLEMethod(int nType, VARIANT *pvResult, IDispatch *pDisp, LPCOLESTR ptName, int cArgs...);



//CreateOleObject
HRESULT CreateOleObject(LPCOLESTR name, IDispatch* &disp);
HRESULT CreateOleObject(LPCOLESTR name, IDispatch* &disp, CLSCTX clsctx);


// OlePropertyGet
template <class T>
HRESULT OlePropertyGet(IDispatch* disp, LPCOLESTR name, T& retVal)
{
	VARIANT v;
	HRESULT hr = OLEMethod(DISPATCH_PROPERTYGET, &v, disp, name, 0);
	
	FromVariant(v, retVal);
	return hr;
}



// OlePropertyGet
template <class T1, class T2>
HRESULT OlePropertyGet(IDispatch* disp, LPCOLESTR name, T1 arg, T2& retVal)
{
	VARIANT v;

	HRESULT hr = OLEMethod(DISPATCH_PROPERTYGET, &v, disp, name, 1, ToVariant(arg) );

	FromVariant(v, retVal);

	return hr;
}



// OlePropertyPut
template <class T>
HRESULT OlePropertyPut(IDispatch* disp, LPCOLESTR name, T value)
{
	return OLEMethod(DISPATCH_PROPERTYPUT, NULL, disp, name, 1, ToVariant(value) );
}

template <class T>
HRESULT OleFunction(IDispatch* disp, LPCOLESTR name, T& retVal)
{
	VARIANT v;

	HRESULT hr =   OLEMethod(DISPATCH_METHOD, &v, disp, name, 0 );
	FromVariant(v, retVal);

	return hr;
}



template <class T1, class T2>
HRESULT OleFunction(IDispatch* disp, LPCOLESTR name, T1 arg1, T2& retVal)
{
	VARIANT v;

	HRESULT hr = OLEMethod(DISPATCH_METHOD, &v, disp, name, 1, ToVariant(arg1));
	FromVariant(v, retVal);

	return hr;
}

template <class T1, class T2>
HRESULT OleFunction(IDispatch* disp, LPCOLESTR name, T1 arg1, IDispatch* &retVal)
{
	VARIANT v;

	HRESULT hr = OLEMethod(DISPATCH_METHOD, &v, disp, name, 1, ToVariant(arg1));
	FromVariant(v, retVal);

	return hr;
}


template <class T1, class T2, class T3>
HRESULT OleFunction(IDispatch* disp, LPCOLESTR name, T1 arg1, T2 arg2, T3& retVal)
{
	VARIANT v;

	HRESULT hr =  OLEMethod(DISPATCH_METHOD, &v, disp, name, 2, ToVariant(arg1), ToVariant(arg2));
	FromVariant(v, retVal);

	return hr;
}

template <class T1, class T2, class T3, class T4>
HRESULT OleFunction(IDispatch* disp, LPCOLESTR name, T1 arg1, T2 arg2, T3 arg3, T4& retVal)
{
	VARIANT v;

	HRESULT hr = OLEMethod(DISPATCH_METHOD,  &v, disp, name, 3, ToVariant(arg1), ToVariant(arg2), ToVariant(arg3) );

	FromVariant(v, retVal);

	return hr;

}


template <class T1, class T2, class T3, class T4, class T5>
HRESULT OleFunction(IDispatch* disp, LPCOLESTR name, T1 arg1, T2 arg2, T3 arg3, T4 arg4, T5& retVal)
{
	VARIANT v;

	HRESULT hr = OLEMethod(DISPATCH_METHOD,  &v, disp, name, 4, ToVariant(arg1), ToVariant(arg2), ToVariant(arg3), ToVariant(arg4) );

	FromVariant(v, retVal);

	return hr;

}


HRESULT OleMethod(IDispatch* disp, LPCOLESTR name);


template <class T>
HRESULT OleMethod(IDispatch* disp, LPCOLESTR name, T arg)
{
	HRESULT hr = OLEMethod(DISPATCH_METHOD, NULL, disp, name, 1, ToVariant(arg));

	return hr;
}


template <class T1, class T2>
HRESULT OleMethod(IDispatch* disp, LPCOLESTR name, T1 arg1, T2 arg2)
{
	HRESULT hr = OLEMethod(DISPATCH_METHOD, NULL, disp, name, 2, ToVariant(arg1), ToVariant(arg2));

	return hr;
}


template <class T1, class T2, class T3>
HRESULT OleMethod(IDispatch* disp, LPOLESTR name, T1 arg1, T2 arg2, T3 arg3)
{
	HRESULT hr = OLEMethod(DISPATCH_METHOD, NULL, disp, name, 3, ToVariant(arg1), ToVariant(arg2), ToVariant(arg3));

	return hr;
}

}


#endif //__GS_OLE_HELP_H__
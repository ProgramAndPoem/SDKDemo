#include "olehelp.h"


namespace gs
{
//double
VARIANT ToVariant(double value)
{

	VARIANT v;

	v.vt = VT_I4;
	v.dblVal = value;

	return v;
}

// int
VARIANT ToVariant(int value)
{
	VARIANT v;

	v.vt = VT_I4;
	v.lVal = value;

	return v;
}



// long
VARIANT ToVariant(LONG value)
{
	VARIANT v;

	v.vt = VT_I4;
	v.lVal = value;

	return v;
}


// unsigned int
VARIANT ToVariant(unsigned int value)
{
	VARIANT v;

	v.vt = VT_UI4;
	v.ulVal = value;

	return v;
}

// unsigned long
VARIANT ToVariant(ULONG value)
{
	VARIANT v;
	
	v.vt = VT_UI4;
	v.ulVal = value;

	return v;
}

// unsigned long long
VARIANT ToVariant(unsigned long long value)
{
	VARIANT v;

	v.vt = VT_UI8;
	v.ulVal = value;

	return v;
}

// LONG*
VARIANT ToVariant(LONG* value)
{
	VARIANT v;

	v.vt = VT_I4 | VT_BYREF;
	v.plVal = value;

	return v;
}


// DWORD*
VARIANT ToVariant(DWORD* value)
{
	VARIANT v;
	
	v.vt = VT_UI4|VT_BYREF;
	v.pulVal = value;

	return v;
}


VARIANT ToVariant(void* value)
{
	VARIANT v;

	v.vt = VT_BYREF;
	v.byref = value;

	return v;
}

// BSTR
VARIANT ToVariant(BSTR value)
{
	VARIANT v;
	
	v.vt = VT_BSTR;
	v.bstrVal = value;

	return v;
}


// SAFEARRAY*
VARIANT ToVariant(SAFEARRAY* value)
{
	VARIANT v;

	v.vt = VT_ARRAY|VT_I1;
	v.parray = value;

	return v;
}


// VARIANT*
VARIANT ToVariant(VARIANT* value)
{
	VARIANT v;

	v.vt = VT_VARIANT|VT_BYREF;
	v.pvarVal = value;

	return v;

	//return *value;
}

// IDispatch
VARIANT ToVariant(IDispatch* value)
{
	VARIANT v;

	v.vt = VT_DISPATCH;
	v.pdispVal = value;

	return v;
}
//double
void FromVariant(const VARIANT& v, double& value)
{
	value = v.dblVal;
}

// char
void FromVariant(const VARIANT& v, char& value)
{
	value = v.cVal;
}


// int
void FromVariant(const VARIANT& v, int& value)
{
	value = v.lVal;
}

// long
void FromVariant(const VARIANT& v, LONG& value)
{
	value = v.lVal;
}


// UCHAR
void FromVariant(const VARIANT& v, UCHAR& value)
{
	value = v.cVal;
}


// unsigned int
void FromVariant(const VARIANT& v, unsigned int& value)
{
	value = v.lVal;
}

// ULONG
void FromVariant(const VARIANT& v, ULONG& value)
{
	value = v.lVal;
}


// BSTR
void FromVariant(const VARIANT& v, BSTR& value)
{
	value = v.bstrVal;
}

// VARIANT*
void FromVariant(const VARIANT& v, VARIANT*& value)
{
	value = v.pvarVal;
}


// IDispatch
void FromVariant(const VARIANT& v, IDispatch*& value)
{
	value = v.pdispVal;
}


#include <stdio.h>
HRESULT OLEMethod(int nType, VARIANT *pvResult, IDispatch *pDisp, LPCOLESTR ptName, int cArgs...)
{
	if(!pDisp) 
		return E_FAIL;

	LPOLESTR name = (LPOLESTR)ptName;

	va_list marker;
	va_start(marker, cArgs);

	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;

	// Get DISPID for name passed...
	HRESULT hr= pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispID);
	if(FAILED(hr)) 
	{
		return hr;
	}
	// Allocate memory for arguments...
	VARIANT *pArgs = new VARIANT[cArgs+1];
	// Extract arguments...
	for(int i = 0; i < cArgs; i++)
	{
		pArgs[cArgs - i - 1] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts!
	if(nType & DISPATCH_PROPERTYPUT) 
	{
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call!
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, nType, &dp, pvResult, NULL, NULL);
	if(FAILED(hr))
	{
		return hr;
	}
	// End variable-argument section...
	va_end(marker);

	delete [] pArgs;
	return hr;
}


//CreateOleObject
HRESULT CreateOleObject(LPCOLESTR name, IDispatch* &disp)
{
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(name, &clsid);
	if ( SUCCEEDED(hr) )
	{
		hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, 
			                  IID_IDispatch, (void **) &disp);
	}

	return hr;
}

//CreateOleObject
HRESULT CreateOleObject(LPCOLESTR name, IDispatch* &disp, CLSCTX clsctx)
{
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(name, &clsid);

	if ( SUCCEEDED(hr) )
	{
		hr = CoCreateInstance(clsid, NULL, clsctx,
			                  IID_IDispatch, (void **) &disp);
	}

	return hr;
}

HRESULT OleMethod(IDispatch* disp, LPCOLESTR name)
{
	HRESULT hr = OLEMethod(DISPATCH_METHOD, NULL, disp, name, 0);

	return hr;
}


}


// Test.cpp : Defines the entry point for the console application.
//

#include <Windows.h>
#include <vector>

#include "olehelp.h"

// Ole dispatch method
void SdkTest(void)
{

	HRESULT hr;
	IDispatch* disp;
	IDispatch* dispDevices;
	IDispatch* dispDevice;
	IDispatch* dispLights;
	IDispatch* dispLight;
	std::vector<IDispatch*> devices;
	std::vector<IDispatch*> lights;

	LONG count = 0;
	UCHAR blue = 0x00;
	UCHAR green = 0x00;
	UCHAR red = 0xFF;

	CoInitializeEx(NULL, COINIT_MULTITHREADED);
	hr = gs::CreateOleObject(L"aura.sdk", disp, CLSCTX_INPROC_SERVER);
	hr = gs::OleFunction(disp, L"Enumerate", (ULONG)0, dispDevices);
	hr = gs::OlePropertyGet(dispDevices, L"Count", count);

	for (LONG i = 0; i < count; ++i)
	{
		hr = gs::OlePropertyGet(dispDevices, L"Item", i, dispDevice);
		devices.push_back(dispDevice);
	}

	hr = gs::OleMethod(disp, L"SwitchMode");
	for (LONG i = 0; i < devices.size(); ++i)
	{
		ULONG type;
		BSTR name;
		hr = gs::OlePropertyGet(devices[i], L"Name", name);
		hr = gs::OlePropertyGet(devices[i], L"Lights", dispLights);
		hr = gs::OlePropertyGet(devices[i], L"Type", type);
		hr = gs::OlePropertyGet(dispLights, L"Count", count);


		for (LONG l = 0; l < count; ++l)
		{
			hr = gs::OlePropertyGet(dispLights, L"Item", l, dispLight);
			hr = gs::OlePropertyPut(dispLight, L"Blue", blue);
			hr = gs::OlePropertyPut(dispLight, L"Green", green);
			hr = gs::OlePropertyPut(dispLight, L"Red", red);

		}
		hr = gs::OleMethod(devices[i], L"SetMode", 0);
		hr = gs::OleMethod(devices[i], L"Apply");

	}

	printf("Press any key to ReleaseControl. \n");
	getchar();
	DWORD reserve = 0;
	hr = gs::OleMethod(disp, L"ReleaseControl", reserve);
	printf("Press any key to Exit. \n");
	getchar();

	for (auto item : lights)
		item->Release();

	for (auto item : devices)
		item->Release();

	
	dispDevices->Release();
	disp->Release();

	CoUninitialize();
}

int main()
{
	SdkTest();
    return 0;
}


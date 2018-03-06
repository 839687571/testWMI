// TestWMI.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#include <comdef.h>
/*#include <Wbemidl.h>*/
#include "WMI_Helper.h"
/*#pragma comment(lib, "wbemuuid.lib")*/
#include <iostream>
#include <iomanip>
#include <sstream>

using namespace  std;

int _tmain(int argc, _TCHAR* argv[])
{

	try
	{
		wcout.imbue(std::locale("chs"));
		WMI_Helper  wmi("ROOT\\CIMV2");
		wmi.getOsInfo();
		std::wcout << L"OS :" << wmi.osInfo.name << "  Build Number:" << wmi.osInfo.buildNumber 
			<< "   Archtecture:" << wmi.osInfo.architecture
			<< "   CSDVersion:" << wmi.osInfo.csdversion << endl;

		wmi.getSystemModel();
		std::wcout << L"BOIS Info:" << wmi.boisString << endl;


		wmi.getMonitor();
		std::wcout << L"Monitor Resolution:" << wmi.monitorString << endl;

		wmi.getCpuInfo();
		std::wcout << L"CPU info :" << wmi.cpuString << L"    CPU Cores:" << wmi.numberCores << endl;

		wmi.getRamInfo();
		std::wcout << L"Logical Ram Info:" << endl;
		int iSizeRam = wmi.ramInfos.size();
		for (int i = 0; i < iSizeRam; i++) {
			std::wcout << wmi.ramInfos.at(i) << endl;
		}


		wmi.getStorage();
		int iSize = wmi.diskInfo.size();

		std::wcout << L"Logical disk Info:"<< endl;
		std::wcout << L"Logical disk Number:" << iSize << endl;
		for (int i = 0; i < iSize; i++) {
			auto &v = wmi.diskInfo.at(i);
			double totalGBSize = v.totalSize * 1.0 / pow(1024, 3);
			double freeSize = v.totalSize * 1.0 / pow(1024, 3);
			std::wstringstream buf;
			buf << std::setiosflags(std::ios::fixed)  << std::setprecision(1) << totalGBSize;
			std::wstring sTotalGbSize = buf.str();

			buf.str(L"");
			buf << std::setiosflags(std::ios::fixed) << std::setprecision(1) << freeSize;
			std::wstring sFreeGbSize = buf.str();

			std::wcout << v.description << L"[" << v.name << L"]" << L"Total Size = " << sTotalGbSize << L" GB   Free Size = " << sFreeGbSize << L" GB" << endl;
		}
	}
	catch (std::exception& e)
	{
		std::cout << e.what() << std::endl;
		std::getwchar();
	}
	

	std::getwchar();

	return 0;
}


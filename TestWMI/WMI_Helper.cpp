#include "WMI_Helper.h"

#include <sstream>
#include <iostream>

using namespace  std;

static std::wstring RAMFormFactors[24] {
	L"Unknown form factor",
		L"",
		L"SIP",
		L"DIP",
		L"ZIP",
		L"SOJ",
		L"Proprietary",
		L"SIMM",
		L"DIMM",
		L"TSOP",
		L"PGA",
		L"RIMM",
		L"SODIMM",
		L"SRIMM",
		L"SMD",
		L"SSMP",
		L"QFP",
		L"TQFP",
		L"SOIC",
		L"LCC",
		L"PLCC",
		L"BGA",
		L"FPBGA",
		L"LGA"
};
static wstring RAMMemoryTypes[26] {
	L"DDR3",
		L"",
		L"SDRAM",
		L"Cache DRAM",
		L"EDO",
		L"EDRAM",
		L"VRAM",
		L"SRAM",
		L"RAM",
		L"ROM",
		L"Flash",
		L"EEPROM",
		L"FEPROM",
		L"CDRAM",
		L"3DRAM",
		L"SDRAM",
		L"SGRAM",
		L"RDRAM",
		L"DDR",
		L"DDR2",
		L"DDR2 FB-DIMM",
		L"DDR2",
		L"DDR3",
		L"FBD2"
};



WMI_Helper::WMI_Helper(std::string wmi_namespace, std::string wmi_class, bool autoconnect/*=true*/):
	m_wmi_namespace(wmi_namespace),
	m_wmi_class(wmi_class)
{
	//if the user wants to autoconnect
	if(autoconnect)
	{
		//start the connection
		connect();
	}
}

WMI_Helper::WMI_Helper(std::string wmi_namespace)
	:m_wmi_namespace(wmi_namespace)
{
	connect();
}

WMI_Helper::~WMI_Helper()
{
	m_pSvc->Release();
	m_pLoc->Release();
	//close down the WMI Service
	CoUninitialize();
}

void WMI_Helper::connect()
{
	HRESULT hres;
	std::stringstream error;

    // Step 1: --------------------------------------------------
    // Initialize COM. ------------------------------------------

    hres =  CoInitializeEx(0, COINIT_MULTITHREADED); 

	//if we failed to intialize COM library
    if (FAILED(hres))
    {
		//throw an exception, Program has failed.
		error << "Failed to initialize COM library. Error code = 0x" << std::hex << hres;
		throw std::exception(error.str().c_str());
    }

	// Step 2: --------------------------------------------------
    // Set general COM security levels --------------------------
    // Note: If you are using Windows 2000, you need to specify -
    // the default authentication credentials for a user by using
    // a SOLE_AUTHENTICATION_LIST structure in the pAuthList ----
    // parameter of CoInitializeSecurity ------------------------

    hres =  CoInitializeSecurity(
        NULL, 
        -1,                          // COM authentication
        NULL,                        // Authentication services
        NULL,                        // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
        RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
        NULL,                        // Authentication info
        EOAC_NONE,                   // Additional capabilities 
        NULL                         // Reserved
        );

              
	//if we failed to initialize security
    if (FAILED(hres))
    {
		CoUninitialize();

		//throw an exception, Program has failed.
        error << "Failed to initialize security. Error code = 0x" << std::hex << hres;
		throw std::exception(error.str().c_str());
    }

	// Step 3: ---------------------------------------------------
    // Obtain the initial locator to WMI -------------------------



    hres = CoCreateInstance(
        CLSID_WbemLocator,             
        0, 
        CLSCTX_INPROC_SERVER, 
		IID_IWbemLocator, (LPVOID *)&m_pLoc);
 
	//if we failed to create the initial locator to WMI
    if (FAILED(hres))
    {
		CoUninitialize();

		//throw an exception, Program has failed.
        error << "Failed to create IWbemLocator object. Error code = 0x" << std::hex << hres;
        throw std::exception(error.str().c_str());
    }

	// Step 4: -----------------------------------------------------
    // Connect to WMI through the IWbemLocator::ConnectServer method


 
    // Connect to a namespace with
    // the current user and obtain pointer pSvc
    // to make IWbemServices calls.
	hres = m_pLoc->ConnectServer(
		_bstr_t(m_wmi_namespace.c_str()),	// Object path of WMI namespace
         NULL,								// User name. NULL = current user
         NULL,								// User password. NULL = current
         0,									// Locale. NULL indicates current
         NULL,								// Security flags.
         0,									// Authority (for example, Kerberos)
         0,									// Context object 
		 &m_pSvc								// pointer to IWbemServices proxy
         );
    
    if (FAILED(hres))
    {
		m_pLoc->Release();
        CoUninitialize();

		//throw an exception, program has failed
        error << "Could not connect to namespace " << m_wmi_namespace << ". Error code = 0x" << std::hex << hres;
        throw std::exception(error.str().c_str());
    }

	// Step 5: --------------------------------------------------
    // Set security levels on the proxy -------------------------

    hres = CoSetProxyBlanket(
		m_pSvc,                        // Indicates the proxy to set
       RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
       RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
       NULL,                        // Server principal name 
       RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
       RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
       NULL,                        // client identity
       EOAC_NONE                    // proxy capabilities 
    );

    if (FAILED(hres))
    {
		m_pSvc->Release();
		m_pLoc->Release();
        CoUninitialize();

		//throw an exception, Program has failed.
        error << "Could not set proxy blanket. Error code = 0x" << std::hex << hres;
		throw std::exception(error.str().c_str());
    }

	// Step 6: --------------------------------------------------
    // Use the IWbemServices pointer to make requests of WMI ----

	//build the query to send
// 	std::stringstream query;
// 	query << "SELECT * FROM " << m_wmi_class;
// 
// 	m_enumerator = NULL;
// 
// 	hres = m_pSvc->ExecQuery(
//         bstr_t("WQL"), 
//         bstr_t(query.str().c_str()),
//         WBEM_FLAG_RETURN_IMMEDIATELY, 
//         NULL,
// 		&m_enumerator);
//     
//     if (FAILED(hres))
//     {
// 		m_pSvc->Release();
// 		m_pLoc->Release();
//         CoUninitialize();
// 
// 		//throw an exception, Program has failed
//         error << "Query: '" << query.str() << "' has failed. Error code = 0x" << std::hex << hres;
// 		throw std::exception(error.str().c_str());
//     }

	// Cleanup
    // ========
    
}

void  WMI_Helper::getOsInfo() 
{
	HRESULT hres;
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = m_pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);

	if (FAILED(hres)) {
		std::stringstream error;
		error << "Query for operating system name failed."
			<< " Error code = 0x"
			<< std::hex << hres << std::endl;

		throw std::exception(error.str().c_str());
	}
	IWbemClassObject *pclsObj = NULL;
	ULONG uReturn = 0;

	while (pEnumerator) {
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn) {
			break;
		}

		VARIANT vtProp;

		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to find value for " << "Name"<< " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		std::wstring OSArchitecture;
		std::wstring OSNameWide;
		std::wstring OSBuildNumber;

		if (vtProp.bstrVal != nullptr) {
			OSNameWide = vtProp.bstrVal;
		}
		try   //XP   can't find 
		{
			hr = pclsObj->Get(L"OSArchitecture", 0, &vtProp, 0, 0);
			if (FAILED(hr)) {
				std::stringstream error;
				error << "Failed to find value for " << "OSArchitecture" << " Error code = 0x"
					<< std::hex << hres << std::endl;
				throw std::exception(error.str().c_str());
			}
			if (vtProp.bstrVal != nullptr) {
				OSArchitecture = vtProp.bstrVal;
			} else {
				OSArchitecture = L"32-Bit";
			}
		}
		catch (std::exception &e)
		{
			std::cout << e.what() << std::endl;
			OSArchitecture = L"32-Bit";
		}

		osInfo.architecture = OSArchitecture;
		
		int garbageIndex = OSNameWide.find(L"|");
		if (garbageIndex != std::wstring::npos) {
			OSNameWide = OSNameWide.erase(garbageIndex, OSNameWide.length() - garbageIndex);
		}
		osInfo.name = OSNameWide;

		hr = pclsObj->Get(L"BuildNumber", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to find value for " << "BuildNumber" << " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		if (vtProp.bstrVal != nullptr) {
			OSBuildNumber = vtProp.bstrVal;
		}
		osInfo.buildNumber = OSBuildNumber;


		hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to find value for " << "Version" << " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		if (vtProp.bstrVal != nullptr) {
			osInfo.version = vtProp.bstrVal;
		}

		hr = pclsObj->Get(L"CSDVersion", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to find value for " << "CSDVersion" << " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		if (vtProp.bstrVal != nullptr) {
			osInfo.csdversion = vtProp.bstrVal;
		}

		VariantClear(&vtProp);
		pclsObj->Release();
	}
	pEnumerator->Release();

}
void WMI_Helper::getMonitor()
{
	int width = GetSystemMetrics(SM_CXSCREEN);
	int height = GetSystemMetrics(SM_CYSCREEN);

	std::wstringstream sizeStream;
	sizeStream << width << L"x" << height;
	setMonitorString(sizeStream.str());
}
void WMI_Helper::getCpuInfo() 
{
	HRESULT hres;
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = m_pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_Processor"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);

	if (FAILED(hres)) {
		std::stringstream error;
		error << "Failed to find value for " << "BuildNumber"
			<< " Error code = 0x"
			<< std::hex << hres << std::endl;
		throw std::exception(error.str().c_str());
	}
	IWbemClassObject *pclsObj = NULL;
	ULONG uReturn = 0;
	while (pEnumerator) {
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn) {
			break;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to get : " << "Name" << " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		std::wstring fullCPUString; //name + @clock
		if (vtProp.bstrVal) {
			fullCPUString = vtProp.bstrVal;
		}
		setCpuString(fullCPUString);


		hr = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to get : " << "NumberOfCores"
				<< " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		int numberCores = 0;
		numberCores = vtProp.iVal;
		setNumberCores(numberCores);


		VariantClear(&vtProp);
		pclsObj->Release();
	}
	pEnumerator->Release();
}

std::wstring getActualPhysicalMemory(HRESULT hres,
	IWbemServices *pSvc,
	IWbemLocator *pLoc)
{
	std::wstring ram;
	IEnumWbemClassObject *pEnumerator = NULL;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_PhysicalMemory"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);

	if (FAILED(hres)) {
		std::stringstream error;
		error << "Query for operating system name failed"
			<< " Error code = 0x"
			<< std::hex << hres << std::endl;
		throw std::exception(error.str().c_str());
	}
	IWbemClassObject *pclsObj = NULL;
	ULONG uReturn = 0;
	double accumulatedRAM = 0;
	while (pEnumerator) {
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn) {
			break;
		}

		VARIANT vtProp;

		hr = pclsObj->Get(L"Capacity", 0, &vtProp, 0, 0);
		double cap;
		double capacity;

		std::wstring temp;
		TCHAR tempChar[100];
		temp = vtProp.bstrVal;
		wcscpy(tempChar, temp.c_str());
		swscanf(tempChar, L"%lf", &cap);

		cap /= (pow(1024, 3));
		accumulatedRAM += cap;
		VariantClear(&vtProp);

		pclsObj->Release();
	}
	TCHAR capacityStrBuff[100];
	_swprintf(capacityStrBuff, L"%.2lf", accumulatedRAM);
	ram = std::wstring(capacityStrBuff);
	pEnumerator->Release();
	return ram;
}


void WMI_Helper::getRamInfo()
{
	HRESULT hres;
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = m_pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_PhysicalMemory"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);

	if (FAILED(hres)) {
		std::stringstream error;
		error << "Query for operating system name failed"
			<< " Error code = 0x"
			<< std::hex << hres << std::endl;
		throw std::exception(error.str().c_str());
	}
	IWbemClassObject *pclsObj = NULL;
	ULONG uReturn = 0;
	while (pEnumerator) {
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn) {
			break;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		//format
		//gb channel ddr @ mhz (no timings yet)

		std::wstring clockStr;
		UINT32 clock;
		UINT16 formFactor;
		std::wstring name;
		std::wstring formFactorStr;
		std::wstring memoryTypeStr;
		UINT16 memoryType;
		//TCHAR capacityStrBuff[10];
		TCHAR clockStrBuff[10];

		// Get the value of the Name property
		hr = pclsObj->Get(L"FormFactor", 0, &vtProp, 0, 0);
		formFactor = vtProp.uintVal;
		if (formFactor < 24 && formFactor >= 0) {
			formFactorStr = RAMFormFactors[formFactor];
		}

		hr = pclsObj->Get(L"MemoryType", 0, &vtProp, 0, 0);
		memoryType = vtProp.uintVal;
		if (memoryType < 26 && memoryType >= 0) {
			memoryTypeStr = RAMMemoryTypes[memoryType];
		}
		hr = pclsObj->Get(L"Speed", 0, &vtProp, 0, 0);
		clock = vtProp.uintVal;
		wsprintf(clockStrBuff, L"%d", clock);
		clockStr = wstring(clockStrBuff);

		hr = pclsObj->Get(L"Capacity", 0, &vtProp, 0, 0);
		wstring ram = L"0";
		if (vtProp.bstrVal) {
			double cap;
			std::wstring temp;
			TCHAR tempChar[100];
			temp = vtProp.bstrVal;
			wcscpy(tempChar, temp.c_str());
			swscanf(tempChar, L"%lf", &cap);

			cap /= (pow(1024, 3));
			TCHAR capacityStrBuff[100];
			_swprintf(capacityStrBuff, L"%.2lf", cap);
			ram = std::wstring(capacityStrBuff);
		}
		ramInfos.push_back(ram +
			L" GB " + formFactorStr + L" " + memoryTypeStr + L" " + clockStr + L"MHz");
		
		VariantClear(&vtProp);
		pclsObj->Release();
	}
	pEnumerator->Release();
}

void WMI_Helper::getStorage()
{
	HRESULT hres;
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = m_pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_LogicalDisk"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator);

	if (FAILED(hres)) {
		std::stringstream error;
		error << "Query for getStorage."
			<< " Error code = 0x"
			<< std::hex << hres << std::endl;
		throw std::exception(error.str().c_str());
	}
	IWbemClassObject *pclsObj = NULL;
	ULONG uReturn = 0;


	while (pEnumerator) {
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1,
			&pclsObj, &uReturn);

		if (0 == uReturn) {
			break;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		if (FAILED(hr)) {
			std::stringstream error;
			error << "Failed to get : " << "Name"<< " Error code = 0x"
				<< std::hex << hres << std::endl;
			throw std::exception(error.str().c_str());
		}
		std::wstring diskName; //name + @clock
		if (vtProp.bstrVal) {
			diskName = vtProp.bstrVal;

			storageInfo info;
			info.name = diskName;

			hr = pclsObj->Get(L"FreeSpace", 0, &vtProp, 0, 0);
			if (FAILED(hr)) {
				std::stringstream error;
				error << "Failed to get : " << "FreeSpace"
					<< " Error code = 0x"
					<< std::hex << hres << std::endl;
				throw std::exception(error.str().c_str());
			}

			if (vtProp.bstrVal) {
				std::wstring strFreeSize = vtProp.bstrVal;
				info.freeSize = _wtoll(strFreeSize.c_str());
			}

			hr = pclsObj->Get(L"Size", 0, &vtProp, 0, 0);
			if (FAILED(hr)) {
				std::stringstream error;
				error << "Failed to get : " << "Size"
					<< " Error code = 0x"
					<< std::hex << hres << std::endl;
				throw std::exception(error.str().c_str());
			}
			if (vtProp.bstrVal) {
				info.totalSize = _wtoll(vtProp.bstrVal);
			}

			hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
			if (FAILED(hr)) {
				std::stringstream error;
				error << "Failed to get : " << "Description"
					<< " Error code = 0x"
					<< std::hex << hres << std::endl;
				throw std::exception(error.str().c_str());
			}
			if (vtProp.bstrVal) {
				info.description = vtProp.bstrVal;
			}
			addDisk(info);
		}
		VariantClear(&vtProp);
		pclsObj->Release();
	}
	pEnumerator->Release();
}


/*
WMI_Helper::wmiValues WMI_Helper::request(std::vector<std::string> valuesToGet)
{
IWbemClassObject *pclsObj;
ULONG uReturn = 0;

wmiValues result;
unsigned int numValuesToGet = valuesToGet.size();

//reset the enumerator back to the start, to allow for multiple requests
m_enumerator->Reset();

//loop through every item that was found
while (m_enumerator)
{
//get the next item
HRESULT hr = m_enumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

if(uReturn == 0)
{
break;
}

wmiValue valueToAdd;

//for each value that the user requested
for(unsigned int i = 0; i < numValuesToGet; i++)
{
//get the next value from the requested list
std::string valueName = valuesToGet.at(i);

//convert to a wstring
std::wstring stemp = std::wstring(valueName.begin(), valueName.end());

VARIANT vtProp;

// Get the value of the requested property
hr = pclsObj->Get(stemp.c_str(), 0, &vtProp, 0, 0);

//if this value wasn't found
if(FAILED(hr))
{
std::stringstream error;
error << "Failed to find value for " << valueName;
throw std::exception(error.str().c_str());
}

//add to the map
valueToAdd[valueName] = vtProp;
}

//add the new map of values to the result vector
result.push_back(valueToAdd);

pclsObj->Release();
}

//return the generated vector of maps
return result;
}
 
*/


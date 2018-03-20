#pragma once

#define _WIN32_DCOM
#include <comdef.h>
#include <Wbemidl.h>

# pragma comment(lib, "wbemuuid.lib")

#include <string>
#include <memory>
#include <vector>
#include <map>
#include <list>

/*
  �ο�:
  https://github.com/rwstoneback/WMI_Helper
  https://github.com/CherryPill/system_info
*/

typedef struct  _storageInfo{
	std::wstring description;
	std::wstring name;
	int64_t  totalSize;   //bit
	int64_t  freeSize;    //bit

	_storageInfo()
	{
		name = description = L"";
		totalSize = 0;
		freeSize = 0;

	}
	_storageInfo(const _storageInfo &info)
	{
		description = info.description;
		name = info.name;
		totalSize = info.totalSize;
		freeSize = info.freeSize;
	}
}storageInfo;

// typedef struct  _ramInfo {
// 	std::wstring name;
// 	int64_t  size;
// 	_ramInfo()
// 	{
// 		name =  L"";
// 		size = 0;
// 	}
// 	_ramInfo(const _ramInfo &info)
// 	{
// 		name = info.name;
// 		size = info.size;
// 	}
// }ramInfo;

typedef struct _osInfo {
	std::wstring name;
	std::wstring buildNumber;
	std::wstring version;
	std::wstring csdversion; // ϵͳ�������汾
	std::wstring architecture;
	_osInfo()
	{

	}
}OsInfo;

//Class: WMI_Helper
//	WMI_Helper makes it simple to interface and request information from the WMI.
class WMI_Helper
{
public:
	//Typedef: wmiValue
	//	A map of string property names to their VARIANT values
	typedef std::map<std::string, VARIANT> wmiValue;

	//Typedef: wmiValues
	//	A vector of wmiValue maps
	typedef std::vector<wmiValue> wmiValues;

public:

	WMI_Helper(std::string wmi_namespace, std::string wmi_class);

	WMI_Helper(std::string wmi_namespace);
	
	// ϵͳ��Ϣ
	void  getOsInfo();
	// ��ʾ�� �ֱ���
	void getMonitor();
	// CUP �ͺ�, ���� ��Ƶ��
	void getCpuInfo();

	// �ڴ���Ϣ �ͺ�,��С
	void getRamInfo();

	// Ӳ�̴�С �ܴ�С ����
	void getStorage();

	// Ӳ����Ϣ
	void getSystemModel();

	// GPU info
	void getGPUInfo();


	void setMonitorString(std::wstring str)
	{
		monitorString = str;
	}

	void setCpuString(std::wstring str)
	{
		cpuString = str;
	}

	void setNumberCores(int v)
	{
		numberCores = v;
	}

	void setMemString(std::wstring str)
	{
		memString = str;
	}


	void addDisk(storageInfo &info)
	{
		diskInfo.push_back(info);
	}


	~WMI_Helper();
	//wmic MEMORYCHIP get speed
public:
	OsInfo       osInfo;
	std::wstring monitorString;

	std::wstring cpuString;
	int  numberCores;

	std::wstring memString;
	std::wstring memSize;

	// �̷�, �ܴ�С, ���ô�С
	std::vector<storageInfo> diskInfo;

	//  �ڴ�. ���ܲ����ڴ���.
	std::vector <std::wstring> ramInfos;

	// 
	std::wstring boisString;


	// gpu
	std::wstring gpuString;

public:

	void connect();

	///wmiValues request(std::vector<std::string> valuesToGet);

private:

	std::string m_wmi_namespace;
	std::string m_wmi_class;

	IWbemServices *m_pSvc = NULL;
	IWbemLocator  *m_pLoc = NULL;

	bool bInitCom = false;

private:
	IEnumWbemClassObject* m_enumerator;
};


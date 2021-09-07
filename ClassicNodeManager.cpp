#include "stdafx.h"
#include "IClassicBaseNodeManager.h"
#include "../IpcHelper/IpcHelper.h"
#include <ATLComTime.h>
#include <vector>
#include <string>
#include <map>
#include <algorithm>

//-----------------------------------------------------------------------------
// MACROS
//-----------------------------------------------------------------------------
#define CHECK_PTR(p) {if ((p)== NULL) throw E_OUTOFMEMORY;}
#define CHECK_ADD(f) {if (!(f)) throw E_OUTOFMEMORY;}
#define CHECK_BOOL(f) {if (!(f)) throw E_FAIL;}
#define CHECK_RESULT(f) {HRESULT hres = f; if (FAILED( hres )) throw hres;}

using DeviceTag = std::pair<std::wstring, VARIANT>;
std::vector<DeviceTag> GetDetectTags(const wchar_t* name)
{
	std::vector<DeviceTag> tags;

	if (name == nullptr)
		return tags;

	const SDeviceInfo* pDevInfo = IpcGetDeviceInfo(name);
	if (pDevInfo == nullptr)
		return tags;

	const STaskRecord& sTask = pDevInfo->lastTask;

	wchar_t szBuffer[64];
	VARIANT val;

	//V_DATE(&val) = COleDateTime(sTask.taskTime).m_dt;
	//V_VT(&val) = VT_DATE;

	swprintf_s(szBuffer, L"%d/%d/%d %d:%d:%d",
		sTask.taskTime.wYear, sTask.taskTime.wMonth, sTask.taskTime.wDay,
		sTask.taskTime.wHour, sTask.taskTime.wMinute, sTask.taskTime.wSecond);
	V_BSTR(&val) = SysAllocString(szBuffer);
	V_VT(&val) = VT_BSTR;

	tags.emplace_back(L"Time", val);

	if (pDevInfo->detectRangeNum > 0 && sTask.detectRepeat > 0)
	{
		int nNumberAll = 0;
		int arrNumberStat[MAX_DETECT_RANGE + 1] = { 0 };

		double dWeightAll = 0.0;
		double arrWeightStat[MAX_DETECT_RANGE + 1] = { 0.0 };

		for (int i = 0; i < sTask.detectRepeat; ++i)
		{
			const SDetectRecord& sDetect = sTask.detects[i];
			for (int j = 0; j <= pDevInfo->detectRangeNum; ++j)
			{
				arrNumberStat[j] += sDetect.numberStat[j];
				nNumberAll += sDetect.numberStat[j];

				arrWeightStat[j] += sDetect.weightStat[j];
				dWeightAll += sDetect.weightStat[j];
			}
		}

		V_I4(&val) = nNumberAll / sTask.detectRepeat;
		V_VT(&val) = VT_I4;
		tags.emplace_back(L"All", val);

		swprintf_s(szBuffer, L"N,<%g", pDevInfo->detectRanges[0]);
		std::replace(szBuffer, szBuffer + 64, L'.', L'_');
		V_I4(&val) = arrNumberStat[0] / sTask.detectRepeat;
		V_VT(&val) = VT_I4;
		tags.emplace_back(szBuffer, val);

		swprintf_s(szBuffer, L"%%,<%g", pDevInfo->detectRanges[0]);
		std::replace(szBuffer, szBuffer + 64, L'.', L'_');
		V_R8(&val) = arrWeightStat[0] * 100 / (dWeightAll + DBL_EPSILON);
		V_VT(&val) = VT_R8;
		tags.emplace_back(szBuffer, val);

		for (int i = 0; i < pDevInfo->detectRangeNum - 1; ++i)
		{
			swprintf_s(szBuffer, L"N,%g~%g", pDevInfo->detectRanges[i], pDevInfo->detectRanges[i + 1]);
			std::replace(szBuffer, szBuffer + 64, L'.', L'_');
			V_I4(&val) = arrNumberStat[i + 1] / sTask.detectRepeat;
			V_VT(&val) = VT_I4;
			tags.emplace_back(szBuffer, val);

			swprintf_s(szBuffer, L"%%,%g~%g", pDevInfo->detectRanges[i], pDevInfo->detectRanges[i + 1]);
			std::replace(szBuffer, szBuffer + 64, L'.', L'_');
			V_R8(&val) = arrWeightStat[i + 1] * 100 / (dWeightAll + DBL_EPSILON);
			V_VT(&val) = VT_R8;
			tags.emplace_back(szBuffer, val);
		}

		swprintf_s(szBuffer, L"N,>%g", pDevInfo->detectRanges[pDevInfo->detectRangeNum - 1]);
		std::replace(szBuffer, szBuffer + 64, L'.', L'_');
		V_I4(&val) = arrNumberStat[pDevInfo->detectRangeNum] / sTask.detectRepeat;
		V_VT(&val) = VT_I4;
		tags.emplace_back(szBuffer, val);

		swprintf_s(szBuffer, L"%%,>%g", pDevInfo->detectRanges[pDevInfo->detectRangeNum - 1]);
		std::replace(szBuffer, szBuffer + 64, L'.', L'_');
		V_R8(&val) = arrWeightStat[pDevInfo->detectRangeNum] * 100 / (dWeightAll + DBL_EPSILON);
		V_VT(&val) = VT_R8;
		tags.emplace_back(szBuffer, val);
	}

	return tags;
}

//-----------------------------------------------------------------------------
// IClassicBaseNodeManager FUNCTION IMPLEMENTATION
//-----------------------------------------------------------------------------

namespace IClassicBaseNodeManager{

static int		g_nClients = 0;
static void*	gDeviceItem_RequestShutdownCommand = NULL;
static std::map<std::wstring, void*>	g_mapDeviceItems;

HRESULT OnCreateServerItems()
{
	HRESULT		hr = S_OK;
	FILETIME	TimeStamp;
	VARIANT		varVal;

	try {
		CoFileTimeNow(&TimeStamp);

		SetServerState(ServerState::NoConfig);

		V_VT(&varVal) = VT_BSTR;						// canonical data type
		V_BSTR(&varVal) = SysAllocString(L" ");
		CHECK_RESULT(AddItem(
			L"Commands.RequestShutdown",				// ItemID
			ReadWritable,		// DaAccessRights
			&varVal, 									// Data Type and Initial Value
			&gDeviceItem_RequestShutdownCommand));		// It's an item with simulated data
		CHECK_RESULT(SetItemValue(
			gDeviceItem_RequestShutdownCommand,
			nullptr,
			(OPC_QUALITY_GOOD | OPC_LIMIT_OK),
			TimeStamp));

		SetServerState(ServerState::Running);
	}
	catch (HRESULT hresEx) {
		SetServerState(ServerState::Failed);
		hr = hresEx;
	}
	catch (_com_error& e) {
		SetServerState(ServerState::Failed);
		hr = e.Error();
	}
	catch (...) {
		SetServerState(ServerState::Failed);
		hr = E_FAIL;
	}

	return hr;
}

ClassicServerDefinition* OnGetDaServerDefinition(void)
{
	// Server Registry Definitions
	// ---------------------------
	//    Specifies all definitions required to register the server in
	//    the Registry.
#ifdef _WIN64
	static ClassicServerDefinition  DaServerDefinition = {
		// CLSID of current Server 
		L"{C3610A08-9E91-46DD-9544-FE7689B64FD5}",
		// CLSID of current Server AppId
		L"{C3610A08-9E91-46DD-9544-FE7689B64FD5}",
		// Version independent Prog.Id. 
		L"GrainDetector.DaServer.x64",
		// Prog.Id. of current Server
		L"GrainDetector.DaServer.x64.10",
		// Friendly name of server
		L"Grain Detector DA Server x64",
		// Friendly name of current server version
		L"Grain Detector DA Server x64 V1.0",
		// Companmy Name
		L"FNST"
	};
#else
	static ClassicServerDefinition  DaServerDefinition = {
		// CLSID of current Server 
		L"{57314941-7D89-477E-A40B-C205E2733CAC}",
		// CLSID of current Server AppId
		L"{57314941-7D89-477E-A40B-C205E2733CAC}",
		// Version independent Prog.Id. 
		L"GrainDetector.DaServer.x86",
		// Prog.Id. of current Server
		L"GrainDetector.DaServer.x86.10",
		// Friendly name of server
		L"Grain Detector DA Server x86",
		// Friendly name of current server version
		L"Grain Detector DA Server x86 V1.0",
		// Companmy Name
		L"FNST"
	};
#endif
	return &DaServerDefinition;
}

ClassicServerDefinition* OnGetAeServerDefinition(void)
{
	return NULL;
}

HRESULT OnGetDaServerParameters(int* updatePeriod, WCHAR* branchDelimiter, DaBrowseMode* browseMode)
{
	// Data Cache update rate in milliseconds
	*updatePeriod = 200;
	*branchDelimiter = '.';
	*browseMode = Generic;            // browse the generic server address space
	return S_OK;
}

HRESULT OnGetDaOptimizationParameters(
	bool * useOnRequestItems,
	bool * useOnRefreshItems,
	bool * useOnAddItem,
	bool * useOnRemoveItem)
{
	*useOnRequestItems = true;
	*useOnRefreshItems = true;
	*useOnAddItem = false;
	*useOnRemoveItem = false;

	return S_OK;
}

void OnStartupSignal(char* commandLine)
{
}

void OnShutdownSignal()
{
}

HRESULT OnQueryProperties(
	void* deviceItemHandle,
	int* noProp,
	int** ids)
{
	// item has no custom properties
	*noProp = 0;
	*ids = NULL;
	return S_FALSE;
}

HRESULT OnGetPropertyValue(
	void* deviceItemHandle,
	int propertyId,
	LPVARIANT propertyValue)
{
	// Item property is not available
	propertyValue = NULL;
	return S_FALSE; //E_INVALID_PID;
}

//----------------------------------------------------------------------------
// OPC Server SDK C++ API Dynamic address space Handling Methods 
// (Called by the generic server)
//----------------------------------------------------------------------------

HRESULT OnBrowseChangePosition(
	DaBrowseDirection browseDirection,
	LPCWSTR position,
	LPWSTR * actualPosition)
{
	// not supported in this default implementation
	return E_INVALIDARG;
}

 HRESULT OnBrowseItemIds(
	LPWSTR actualPosition,
	DaBrowseType browseFilterType,
	LPWSTR filterCriteria,
	VARTYPE dataTypeFilter,
	DaAccessRights accessRightsFilter,
	int * noItems,
	LPWSTR ** itemIds)
{
	// not supported in this default implementation
	*noItems = 0;
	*itemIds = NULL;
	return E_INVALIDARG;
}

HRESULT OnBrowseGetFullItemId(
	LPWSTR actualPosition,
	LPWSTR itemName,
    LPWSTR * fullItemId)
{
	// not supported in this default implementation
	*fullItemId = NULL;
	return E_INVALIDARG;
}

//----------------------------------------------------------------------------
// OPC Server SDK C++ API additional Methods
// (Called by the generic server)
//----------------------------------------------------------------------------

HRESULT OnClientConnect()
{
	if (g_nClients == 0)
		IpcInitialize(false);

	++g_nClients;

	return S_OK;
}

HRESULT OnClientDisconnect()
{
	--g_nClients;

	if (g_nClients == 0)
	{
		for (auto& pair : g_mapDeviceItems)
			RemoveItem(pair.second);
		g_mapDeviceItems.clear();

		RemoveItem(gDeviceItem_RequestShutdownCommand);
		gDeviceItem_RequestShutdownCommand = nullptr;

		IpcFinalize();
	}

	return S_OK;
}

HRESULT OnRefreshItems(
	/* in */       int        numItems,
	/* in */       void    ** itemHandles)
{
	HRESULT		hr = S_OK;
	FILETIME	TimeStamp;
	wchar_t		tagName[1024] = L"\0";

	if (!IpcCheckUpdate(true))
		return hr;

	try
	{
		CoFileTimeNow(&TimeStamp);

		int nDeviceCount = IpcGetDeviceCount();
		for (int i = 0; i < nDeviceCount; ++i)
		{
			const wchar_t* name = IpcGetDeviceName(i);
			const wchar_t* folderName = wcsrchr(name, L'.');
			if (folderName)
				folderName++;
			else
				folderName = name;

			auto tags = GetDetectTags(name);
			for (auto& tag : tags)
			{
				swprintf_s(tagName, L"%s.%s", folderName, tag.first.c_str());
				auto iter = g_mapDeviceItems.find(tagName);
				if (iter == g_mapDeviceItems.end())
				{
					void* item = nullptr;
					CHECK_RESULT(AddItem(
						tagName,
						Readable,
						&tag.second,
						&item));
					CHECK_RESULT(SetItemValue(
						item,
						nullptr,
						(OPC_QUALITY_GOOD | OPC_LIMIT_OK),
						TimeStamp));
					g_mapDeviceItems.emplace(tagName, item);
				}
				else
				{
					CHECK_RESULT(SetItemValue(
						iter->second,
						&tag.second,
						(OPC_QUALITY_GOOD | OPC_LIMIT_OK),
						TimeStamp));
				}
			}
		}
	}
	catch (HRESULT hresEx) {
		SetServerState(ServerState::Failed);
		hr = hresEx;
	}
	catch (_com_error& e) {
		SetServerState(ServerState::Failed);
		hr = e.Error();
	}
	catch (...) {
		SetServerState(ServerState::Failed);
		hr = E_FAIL;
	}

	return hr;
}

HRESULT OnAddItem(
	/* in */       void*	  deviceItem)
{
	return S_OK;
}

HRESULT OnRemoveItem(
	/* in */       void*	  deviceItem)
{
	DeleteItem(deviceItem);
	return S_OK;
}

HRESULT OnWriteItems(
	int          numItems,
	void      ** deviceItems,
	OPCITEMVQT*  itemVQTs,
	HRESULT   *  errors)
{
	for (int i = 0; i < numItems; ++i)              // handle all items
	{
		if (deviceItems[i] == gDeviceItem_RequestShutdownCommand)
		{
			FireShutdownRequest(V_BSTR(&itemVQTs[i].vDataValue));
		}
		errors[i] = S_OK;						// init to S_OK
	}

	return S_OK;
}

HRESULT OnTranslateToItemId(int conditionId, int subConditionId, int attributeId, LPWSTR* itemId, LPWSTR* nodeName, CLSID* clsid)
{
	return S_OK;
}

HRESULT OnAckNotification(int conditionId, int subConditionId)
{
	return S_OK;
}

/**
 * @fn  LogLevel OnGetLogLevel();
 *
 * @brief   Gets the logging level to be used.
 *
 * @return  A LogLevel.
 */
LogLevel OnGetLogLevel( )
{
#ifdef _DEBUG
	return Info;
#else
	return Disabled;
#endif
}

/**
 * @fn  void OnGetLogPath(char * logPath);
 *
 * @brief   Gets the logging pazh to be used.
 *
 * @param [in,out]  logPath    Path to be used for logging.
 */
void OnGetLogPath(const char * logPath)
{
	logPath = "";
}

HRESULT OnRequestItems(int numItems, LPWSTR *fullItemIds, VARTYPE *dataTypes)
{
	// no valid item in this default implementation
	return S_FALSE;
}

}
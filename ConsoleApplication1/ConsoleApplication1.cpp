#include "stdafx.h"
#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")

int main(int argc, char **argv)
{
	HRESULT hres;
	/*тип hresult является одним из средств контроля ошибок в com / dcom.этот тип представляет собой 32 - битное число,
		в котором кодируется результат операции.старший бит этого числа равен 1, если была ошибка,
		и 0, если всё прошло нормально.следующие 4 бита зарезервированы для дальнейшего использования.
		следующие 11 бит показывают, где возникла ошибка(это значение обычно называется facility code, что можно приблизительно перевести как код устройства,
		если подразумевать под устройством не только аппаратные, но и логические устройства).младшие 16 бит кодируют собственно ошибку.*/

	// Step 1: --------------------------------------------------
	// Initialize COM. ------------------------------------------

	hres = CoInitializeEx(0, COINIT_MULTITHREADED);//OLE объект может быть вызван в любом потоке
	/*Инициализирует COM - библиотеку для использования вызывающим потоком,
		устанавливает модель параллелизма потока и создает новую квартиру для потока, если требуется.*/
	if (FAILED(hres))
	{
		cout << "Failed to initialize COM library. Error code = 0x"
			<< hex << hres << endl;
		return 1;                  // Program has failed.
	}

	// Step 2: --------------------------------------------------
	// Set general COM security levels --------------------------

	/* Инициализация безопасности(хз вообще) и устанавливает значения безопасности по умолчанию для процесса.*/
	hres = CoInitializeSecurity(
		NULL, //Разрешения доступа, которые сервер будет использовать для приема вызовов.
		//Этот параметр используется COM только тогда, когда сервер вызывает CoInitializeSecurity. У нас не сервер
		-1,                          // COM authentication
		//Если этот параметр равен 0, службы аутентификации не будут зарегистрированы, и сервер не сможет получать защищенные вызовы.
		//Значение -1 сообщает COM, чтобы выбрать,
		//какие службы проверки подлинности необходимо зарегистрировать, и если это так, параметр asAuthSvc должен быть NULL
		NULL,                        // Authentication services
		NULL,                        // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication Массив служб аутентификации, которые сервер желает использовать для приема вызова.
		RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  Уровень аутентификации по умолчанию для процесса.
		//Оба сервера и клиенты используют этот параметр при вызове CoInitializeSecurity.
		NULL,                        // Authentication info Этот параметр используется COM только тогда, когда клиент вызывает CoInitializeSecurity.
		EOAC_NONE,                   // Additional capabilities  Дополнительные возможности клиента или сервера,
		//заданные установкой одного или нескольких значений EOLE_AUTHENTICATION_CAPABILITIES. EOAC_NONE Указывает, что флаги возможностей не установлены.
		NULL                         // Reserved
	);


	if (FAILED(hres))
	{
		cout << "Failed to initialize security. Error code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return 1;                    // Program has failed.
	}

	// Step 3: ---------------------------------------------------
	// Obtain the initial locator to WMI Получаем исходный локатор в WMI -------------------------

	IWbemLocator *pLoc = NULL;
	//Используйте интерфейс IWbemLocator, чтобы получить начальный указатель пространства имен для интерфейса IWbemServices для WMI на определенном хост-компьютере.
	//IWBEMSERVICE - Интерфейс IWbemServices используется клиентами и поставщиками для доступа к службам WMI.
	//Интерфейс реализован WMI и WMI-провайдерами и является основным интерфейсом WMI.

	hres = CoCreateInstance( //Создает один неинициализированный объект класса, связанного с указанным CLSID.CLSID — аббревиатура для идентификатора класса
		// Пример "Мой компьютер" - 	{20D04FE0-3AEA-1069-A2D8-08002B30309D} 
		CLSID_WbemLocator,
		0,// Позволяет клиентам получать указатели на другие интерфейсы на определенном объекте с помощью метода QueryInterface 
		//и управлять существованием объекта с помощью методов AddRef и Release. Все другие COM-интерфейсы наследуются, прямо или косвенно, от IUnknown.
		//Поэтому три метода в IUnknown - это первые записи в VTable для каждого интерфейса.
		CLSCTX_INPROC_SERVER,//Контекст, в котором будет выполняться код, управляющий вновь созданным объектом. Значения взяты из перечисления CLSCTX.
		//Код, который создает и управляет объектами этого класса,
		//представляет собой DLL, которая выполняется в том же процессе, что и вызывающая функция, определяющая контекст класса.
		IID_IWbemLocator,// Ссылка на идентификатор интерфейса, который будет использоваться для связи с объектом.
		(LPVOID *)&pLoc); //Адрес переменной указателя, который получает указатель интерфейса, запрошенный в riid.
	//При успешном возврате * ppv содержит запрошенный указатель интерфейса. После сбоя * ppv содержит NULL.

	if (FAILED(hres))
	{
		cout << "Failed to create IWbemLocator object."
			<< " Err code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return 1;                 // Program has failed.
	}

	// Step 4: -----------------------------------------------------
	// Connect to WMI through the IWbemLocator::ConnectServer method

	IWbemServices *pSvc = NULL;//Интерфейс IWbemServices используется клиентами и поставщиками для доступа к службам WMI.
	//Интерфейс реализован WMI и WMI-провайдерами и является основным интерфейсом WMI.

	// Connect to the root\cimv2 namespace with
	// the current user and obtain pointer pSvc
	// to make IWbemServices calls.
	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
		NULL,                    // User name. NULL = current user
		NULL,                    // User password. NULL = current
		0,                       // Locale. NULL indicates current
		NULL,                    // Security flags.
		0,                       // Authority (for example, Kerberos)
		0,                       // Context object 
		&pSvc                    // pointer to IWbemServices proxy
	);

	if (FAILED(hres))
	{
		cout << "Could not connect. Error code = 0x"
			<< hex << hres << endl;
		pLoc->Release();
		CoUninitialize();
		return 1;                // Program has failed.
	}

	cout << "Connected to ROOT\\CIMV2 WMI namespace" << endl;


	// Step 5: --------------------------------------------------
	// Set security levels on the proxy -------------------------

	hres = CoSetProxyBlanket(
		pSvc,                        // Indicates the proxy to set
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
		cout << "Could not set proxy blanket. Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 1;               // Program has failed.
	}

	// Step 6: --------------------------------------------------
	// Use the IWbemServices pointer to make requests of WMI ----

	// For example, get the name of the operating system
	IEnumWbemClassObject* pEnumerator = NULL; //Интерфейс IEnumWbemClassObject используется для перечисления объектов Common Information Model (CIM)
	//и похож на стандартный COM-счетчик.
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY,// 1.Этот флаг заставляет перечислитель пересылать только вперед. Перечислители только в прямом направлении, как правило, 
		//намного быстрее и используют меньше памяти, чем обычные счетчики, но не позволяют звонить на Clone или Reset.
		NULL,// Указатель на контекст
		&pEnumerator);//Если ошибка не возникает, это получает перечислитель, который позволяет вызывающему пользователю извлекать экземпляры
	//в результирующем наборе запроса. Это не ошибка для запроса иметь набор результатов с 0 экземплярами. Это определяется только попыткой итерации через экземпляры. 
	//Этот объект возвращается с положительным счетчиком ссылок. Вызывающий должен вызвать Release, когда объект больше не требуется.
	

	if (FAILED(hres))
	{
		cout << "Query for operating system name failed."
			<< " Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return 1;               // Program has failed.
	}

	// Step 7: -------------------------------------------------
	// Get the data from the query in step 6 -------------------

	IWbemClassObject *pclsObj = NULL;//Интерфейс IWbemClassObject содержит и обрабатывает определения классов и экземпляры объектов класса.
	ULONG uReturn = 0;

	while (pEnumerator)
	{
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE,// время 
			1,//Количество запрошенных объектов.
			&pclsObj,//Куда положить
			&uReturn);//Указатель на ULONG, который получает количество возвращенных объектов.
		//Это число может быть меньше количества, указанного в uCount. Этот указатель не может быть NULL.

		if (0 == uReturn)
		{
			break;
		}

		VARIANT vtProp;//Variant в языке C++ - универсальный тип, который может принимать значения разных типов данных.

		// Get the value of the Name property
		hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0); //Имя параметра, зарезервирован должен быть 0, Параметру присваевается тип, TBD,Может быть NULL.
		//Если не NULL, значение LONG указывает на получение информации о происхождении свойства.
		wcout << " OS Name : " << vtProp.bstrVal << endl;
		VariantClear(&vtProp);
		pclsObj->Release();
	}

	// Cleanup
	// ========

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();
	system("pause");

	return 0;   // Program successfully completed.

}

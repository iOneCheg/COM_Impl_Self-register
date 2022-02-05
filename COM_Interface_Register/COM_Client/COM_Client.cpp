
#include <iostream>
#include <string>
#include "../Dll_COM/Interfaces.h"

using namespace std;

int main()
{
	setlocale(LC_ALL, "ru");
	IEnumerator* pIEnumerator = NULL;
	ICollection* pICollection = NULL;

	CoInitialize(NULL);
	HRESULT hrICollection = CoCreateInstance(
		CLSID_CList, NULL, CLSCTX_INPROC_SERVER,
		IID_ICollection, (void**)&pICollection);
	
	unsigned int count = 0;
	string addElement = "Добавлен объект:", deleteElement = "Удалён элемент:";

	Object number, phi, e, pi, array;
	Object targetObj;

	number.Type = otInt;
	phi.Type = otDouble;
	e.Type = otDouble;
	pi.Type = otDouble;
	array.Type = otArray;

	number.Value.Int = 821543;
	phi.Value.Double = 1.6180339887;
	e.Value.Double = 2.718281828;
	pi.Value.Double = 3.1415926535;

	array.Value.Array = new ObjectArray;
	array.Value.Array->Data = new Object[4];
	array.Value.Array->Count = 4;
	array.Value.Array->Data[0] = number;
	array.Value.Array->Data[1] = phi;
	array.Value.Array->Data[2] = e;
	array.Value.Array->Data[3] = pi;

	if (SUCCEEDED(hrICollection))
	{
		cout << "Работа с указателем на ICollection" << endl;
		pICollection->Add(e);
		cout << addElement + "e" << endl;
		pICollection->Add(pi);
		cout << addElement + "pi" << endl;
		pICollection->GetCount(&count);
		cout << "В списке " << count << " элемента" << endl;
		pICollection->Add(number);
		cout << addElement + "number" << endl;
		pICollection->Remove(pi);
		cout << deleteElement + "pi" << endl;
		pICollection->Add(phi);
		cout << addElement + "phi" << endl;
		pICollection->Add(array);
		cout << addElement + "array типа ObjectArray" << endl;
		pICollection->GetCount(&count);
		cout << "В списке " << count << " элемента" << endl;
		ObjectArray* list;
		pICollection->ToArray(&list);
		cout << "Список преобразован в массив:" << endl;
		for (unsigned int i = 0; i < list->Count; i++)
		{
			cout << "Элемент " << i << ":";
			switch (list->Data[i].Type)
			{

			case otInt:
				cout << list->Data[i].Value.Int << endl;
				break;

			case otDouble:
				cout << list->Data[i].Value.Double << endl;
				break;

			case otArray:
				cout << "Список из " << list->Data[i].Value.Array->Count << " элементов:" << endl;
				for (unsigned int j = 0; j < list->Data[i].Value.Array->Count; j++)
				{
					cout <<"\t"<< "Элемент " << j << ":";
					

					switch (list->Data[i].Value.Array->Data[j].Type)
					{
					case otInt:
						cout << list->Data[i].Value.Array->Data[j].Value.Int << endl;
						break;

					case otDouble:
						cout << list->Data[i].Value.Array->Data[j].Value.Double << endl;
						break;

					default:
						break;
					}
				}
				break;

			default:
				break;

			}

		}
		cout << "Завершение работы  указателем на ICollection" << endl;

	}

	if (SUCCEEDED(pICollection->QueryInterface(IID_IEnumerator, (void**)&pIEnumerator)))
	{
		cout << "Работа с указателем на IEnumerator" << endl;

		pIEnumerator->Reset();
		cout << "Сброс итератора" << endl;

		if ((pIEnumerator->MoveNext(&count) == S_OK) && (count == TRUE))
		{
			cout << "Смещение итератора на 1" << endl;
		}
		else
		{
			cout << "Нет смещения итератора" << endl;
		}

		if (SUCCEEDED(pIEnumerator->GetCurrent(&targetObj)))
		{
			switch (targetObj.Type)
			{
			case otInt:
				cout << "Текущий объект:";
				cout << targetObj.Value.Int << endl;
				break;

			case otDouble:
				cout << "Текущий объект:";
				cout << targetObj.Value.Double << endl;
				break;

			case otArray:
				cout << "Текущий объект: список на "<<targetObj.Value.Array->Count<<"элементов";
			}
		}
		else
		{
			cout << "Объект не найден" << endl;
		}


		if ((pIEnumerator->MoveNext(&count) == S_OK) && (count == TRUE))
		{
			cout << "Смещение итератора на 1" << endl;
		}
		else
		{
			cout << "Нет смещения итератора" << endl;
		}

		if (SUCCEEDED(pIEnumerator->GetCurrent(&targetObj)))
		{
			switch (targetObj.Type)
			{
			case otInt:
				cout << "Текущий объект:";
				cout << targetObj.Value.Int << endl;
				break;

			case otDouble:
				cout << "Текущий объект:";
				cout << targetObj.Value.Double << endl;
				break;

			case otArray:
				cout << "Текущий объект: список на " << targetObj.Value.Array->Count << "элементов";
			}
		}
		
		cout << endl;
		pIEnumerator->Release();
		cout << "Завершение работы  указателем на IEnumerator" << endl;
		cout << endl;
	}

	pICollection->Release();
	CoUninitialize();
	return 0;
}


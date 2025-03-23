#pragma once
#include <vcclr.h>
#include <vector>
#include "../models/NotebookEntry.h"

using namespace System;
using namespace System::Collections::Generic;
using namespace System::Windows::Forms;
using namespace System::Text;
using namespace System::IO;

public ref class NotebookManager {
private:
    List<NotebookEntry<int>^>^ entries;
    String^ currentFilePath;
    static Encoding^ fileEncoding = Encoding::UTF8;

    // Вспомогательный класс для предиката удаления
    ref class RemoveEntryPredicate {
    private:
        int targetId;
    public:
        RemoveEntryPredicate(int id) : targetId(id) {}
        bool Match(NotebookEntry<int>^ entry) {
            return entry->GetId() == targetId;
        }
    };

    // Вспомогательный класс для сравнения при сортировке
    ref class LastNameComparer : IComparer<NotebookEntry<int>^> {
    private:
        bool ascending;
    public:
        LastNameComparer(bool asc) : ascending(asc) {}
        virtual int Compare(NotebookEntry<int>^ x, NotebookEntry<int>^ y) {
            if (ascending) {
                return String::Compare(x->GetLastName(), y->GetLastName());
            }
            return String::Compare(y->GetLastName(), x->GetLastName());
        }
    };

public:
    // Конструктор
    NotebookManager() {
        entries = gcnew List<NotebookEntry<int>^>();
        currentFilePath = "";
    }

    // Добавление новой записи
    void AddEntry(NotebookEntry<int>^ entry) {
        if (entry->IsValid()) {
            entries->Add(entry);
        }
        else {
            throw gcnew Exception("Недопустимая запись: обязательные поля должны быть заполнены");
        }
    }

    // Удаление записи по ID
    bool RemoveEntry(int id) {
        RemoveEntryPredicate^ predicate = gcnew RemoveEntryPredicate(id);
        return entries->RemoveAll(gcnew Predicate<NotebookEntry<int>^>(predicate, &RemoveEntryPredicate::Match)) > 0;
    }

    // Получение всех записей
    List<NotebookEntry<int>^>^ GetAllEntries() {
        return entries;
    }

    // Поиск по имени
    List<NotebookEntry<int>^>^ SearchByFirstName(String^ firstName) {
        List<NotebookEntry<int>^>^ results = gcnew List<NotebookEntry<int>^>();
        for each (NotebookEntry<int>^ entry in entries) {
            if (entry->GetFirstName()->ToLower()->Contains(firstName->ToLower())) {
                results->Add(entry);
            }
        }
        return results;
    }

    // Поиск по фамилии
    List<NotebookEntry<int>^>^ SearchByLastName(String^ lastName) {
        List<NotebookEntry<int>^>^ results = gcnew List<NotebookEntry<int>^>();
        for each (NotebookEntry<int>^ entry in entries) {
            if (entry->GetLastName()->ToLower()->Contains(lastName->ToLower())) {
                results->Add(entry);
            }
        }
        return results;
    }

    // Поиск по номеру телефона
    List<NotebookEntry<int>^>^ SearchByPhone(String^ phone) {
        List<NotebookEntry<int>^>^ results = gcnew List<NotebookEntry<int>^>();
        for each (NotebookEntry<int>^ entry in entries) {
            if (entry->GetPhoneNumber()->ToLower()->Contains(phone->ToLower())) {
                results->Add(entry);
            }
        }
        return results;
    }

    // Поиск по email
    List<NotebookEntry<int>^>^ SearchByEmail(String^ email) {
        List<NotebookEntry<int>^>^ results = gcnew List<NotebookEntry<int>^>();
        for each (NotebookEntry<int>^ entry in entries) {
            if (!String::IsNullOrEmpty(entry->GetEmail()) && 
                entry->GetEmail()->ToLower()->Contains(email->ToLower())) {
                results->Add(entry);
            }
        }
        return results;
    }

    // Поиск по адресу
    List<NotebookEntry<int>^>^ SearchByAddress(String^ address) {
        List<NotebookEntry<int>^>^ results = gcnew List<NotebookEntry<int>^>();
        for each (NotebookEntry<int>^ entry in entries) {
            if (!String::IsNullOrEmpty(entry->GetAddress()) && 
                entry->GetAddress()->ToLower()->Contains(address->ToLower())) {
                results->Add(entry);
            }
        }
        return results;
    }

    // Поиск по любому полю
    List<NotebookEntry<int>^>^ SearchByAnyField(String^ query, int searchType) {
        switch (searchType) {
            case 0: // By First Name
                return SearchByFirstName(query);
            case 1: // By Last Name
                return SearchByLastName(query);
            case 2: // By Phone
                return SearchByPhone(query);
            case 3: // By Email
                return SearchByEmail(query);
            case 4: // By Address
                return SearchByAddress(query);
            default:
                return GetAllEntries();
        }
    }

    // Сортировка по фамилии
    void SortByLastName(bool ascending) {
        LastNameComparer^ comparer = gcnew LastNameComparer(ascending);
        entries->Sort(comparer);
    }

    // Сохранение в файл
    void SaveToFile(String^ filePath) {
        try {
            StreamWriter^ writer = nullptr;
            try {
                // Создаем writer с явным указанием кодировки UTF-8 с BOM
                writer = gcnew StreamWriter(filePath, false, gcnew UTF8Encoding(true));
                
                for each (NotebookEntry<int>^ entry in entries) {
                    writer->WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}",
                        entry->GetId(),
                        entry->GetFirstName(),
                        entry->GetLastName(),
                        entry->GetPhoneNumber(),
                        entry->GetBirthDate(),
                        entry->GetEmail(),
                        entry->GetAddress(),
                        entry->GetNotes());
                }
                currentFilePath = filePath;
            }
            finally {
                if (writer != nullptr) {
                    writer->Close();
                    delete writer;
                }
            }
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Ошибка при сохранении файла: " + ex->Message);
        }
    }

    // Загрузка из файла
    void LoadFromFile(String^ filePath) {
        try {
            StreamReader^ reader = nullptr;
            try {
                entries->Clear();
                // Открываем reader с автоопределением кодировки
                reader = gcnew StreamReader(filePath, true);
                
                String^ line;
                while ((line = reader->ReadLine()) != nullptr) {
                    array<String^>^ parts = line->Split('\t');
                    if (parts->Length >= 8) {
                        NotebookEntry<int>^ entry = gcnew NotebookEntry<int>(
                            Int32::Parse(parts[0]),
                            parts[1],
                            parts[2],
                            parts[3],
                            parts[4],
                            parts[5],
                            parts[6],
                            parts[7]
                        );
                        entries->Add(entry);
                    }
                }
                currentFilePath = filePath;
            }
            finally {
                if (reader != nullptr) {
                    reader->Close();
                    delete reader;
                }
            }
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Ошибка при загрузке файла: " + ex->Message);
        }
    }

    // Экспорт в Excel
    void ExportToExcel(String^ filePath) {
        try {
            Type^ excelType = Type::GetTypeFromProgID("Excel.Application");
            if (excelType == nullptr) {
                throw gcnew Exception("Microsoft Excel не установлен на компьютере");
            }

            Object^ excelObj = Activator::CreateInstance(excelType);
            Object^ workbooks = excelType->InvokeMember("Workbooks", 
                System::Reflection::BindingFlags::GetProperty, nullptr, excelObj, nullptr);
            Object^ workbook = workbooks->GetType()->InvokeMember("Add",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, workbooks, nullptr);
            Object^ worksheet = workbook->GetType()->InvokeMember("ActiveSheet",
                System::Reflection::BindingFlags::GetProperty, nullptr, workbook, nullptr);

            // Заголовки
            array<String^>^ headers = {"ID", "Имя", "Фамилия", "Телефон", "Дата рождения", "Email", "Адрес", "Заметки"};
            for (int i = 0; i < headers->Length; i++) {
                Object^ cell = worksheet->GetType()->InvokeMember("Cells",
                    System::Reflection::BindingFlags::GetProperty, nullptr, worksheet,
                    gcnew array<Object^> { 1, i + 1 });
                cell->GetType()->InvokeMember("Value",
                    System::Reflection::BindingFlags::SetProperty, nullptr, cell,
                    gcnew array<Object^> { headers[i] });
            }

            // Данные
            int row = 2;
            for each (NotebookEntry<int>^ entry in entries) {
                array<Object^>^ rowData = {
                    entry->GetId(),
                    entry->GetFirstName(),
                    entry->GetLastName(),
                    entry->GetPhoneNumber(),
                    entry->GetBirthDate(),
                    entry->GetEmail(),
                    entry->GetAddress(),
                    entry->GetNotes()
                };

                for (int col = 0; col < rowData->Length; col++) {
                    Object^ cell = worksheet->GetType()->InvokeMember("Cells",
                        System::Reflection::BindingFlags::GetProperty, nullptr, worksheet,
                        gcnew array<Object^> { row, col + 1 });
                    cell->GetType()->InvokeMember("Value",
                        System::Reflection::BindingFlags::SetProperty, nullptr, cell,
                        gcnew array<Object^> { rowData[col] });
                }
                row++;
            }

            // Автоподбор ширины столбцов
            Object^ usedRange = worksheet->GetType()->InvokeMember("UsedRange",
                System::Reflection::BindingFlags::GetProperty, nullptr, worksheet, nullptr);
            Object^ columns = usedRange->GetType()->InvokeMember("Columns",
                System::Reflection::BindingFlags::GetProperty, nullptr, usedRange, nullptr);
            columns->GetType()->InvokeMember("AutoFit",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, columns, nullptr);

            // Сохранение
            workbook->GetType()->InvokeMember("SaveAs",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, workbook,
                gcnew array<Object^> { filePath });

            // Закрытие
            workbook->GetType()->InvokeMember("Close",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, workbook,
                gcnew array<Object^> { true });
            excelType->InvokeMember("Quit",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, excelObj, nullptr);

            System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheet);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(workbooks);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(excelObj);
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Ошибка при экспорте в Excel: " + ex->Message);
        }
    }
}; 
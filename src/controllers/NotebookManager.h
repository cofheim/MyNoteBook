#pragma once
#include <vcclr.h>
#include <vector>
#include "../models/NotebookEntry.h"

using namespace System;
using namespace System::Collections::Generic;
using namespace System::Windows::Forms;
using namespace System::Text;
using namespace System::IO;
using namespace Newtonsoft::Json;

public ref class NotebookManager {
private:
    List<NotebookEntry<int>^>^ entries;
    String^ currentFilePath;
    static Encoding^ fileEncoding = Encoding::UTF8;
    String^ defaultJsonPath = "contacts.json";

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
        currentFilePath = defaultJsonPath;
        
        // Автоматически загружаем контакты из JSON файла при запуске, если он существует
        if (File::Exists(defaultJsonPath)) {
            try {
                LoadFromJsonFile(defaultJsonPath);
            }
            catch (...) {
                // Если не удалось загрузить - создаем пустой список
                entries = gcnew List<NotebookEntry<int>^>();
            }
        }
    }

    // Добавление новой записи
    void AddEntry(NotebookEntry<int>^ entry) {
        if (entry->IsValid()) {
            entries->Add(entry);
            // Автоматически сохраняем в JSON после добавления
            SaveToJsonFile(defaultJsonPath);
        }
        else {
            throw gcnew Exception("Invalid entry: required fields must be filled");
        }
    }

    // Удаление записи по ID
    bool RemoveEntry(int id) {
        RemoveEntryPredicate^ predicate = gcnew RemoveEntryPredicate(id);
        bool result = entries->RemoveAll(gcnew Predicate<NotebookEntry<int>^>(predicate, &RemoveEntryPredicate::Match)) > 0;
        if (result) {
            // Автоматически сохраняем в JSON после удаления
            SaveToJsonFile(defaultJsonPath);
        }
        return result;
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
    
    // Сохранение в JSON файл
    void SaveToJsonFile(String^ filePath) {
        try {
            // Преобразуем список записей в JSON
            String^ json = JsonConvert::SerializeObject(entries, Formatting::Indented);
            
            // Записываем JSON в файл
            File::WriteAllText(filePath, json, gcnew UTF8Encoding(true));
            currentFilePath = filePath;
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Error saving to JSON file: " + ex->Message);
        }
    }
    
    // Загрузка из JSON файла
    void LoadFromJsonFile(String^ filePath) {
        try {
            if (File::Exists(filePath)) {
                // Читаем JSON из файла
                String^ json = File::ReadAllText(filePath, gcnew UTF8Encoding(true));
                
                // Десериализуем JSON в список записей
                entries = JsonConvert::DeserializeObject<List<NotebookEntry<int>^>^>(json);
                if (entries == nullptr) {
                    entries = gcnew List<NotebookEntry<int>^>();
                }
                currentFilePath = filePath;
            }
            else {
                throw gcnew Exception("File does not exist: " + filePath);
            }
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Error loading from JSON file: " + ex->Message);
        }
    }

    // Сохранение в файл (для совместимости)
    void SaveToFile(String^ filePath) {
        // Если файл имеет расширение .json, используем JSON формат
        if (filePath->EndsWith(".json", StringComparison::OrdinalIgnoreCase)) {
            SaveToJsonFile(filePath);
            return;
        }
        
        // Иначе используем старый текстовый формат
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
            throw gcnew Exception("Error saving file: " + ex->Message);
        }
    }

    // Загрузка из файла (для совместимости)
    void LoadFromFile(String^ filePath) {
        // Если файл имеет расширение .json, используем JSON формат
        if (filePath->EndsWith(".json", StringComparison::OrdinalIgnoreCase)) {
            LoadFromJsonFile(filePath);
            return;
        }
        
        // Иначе используем старый текстовый формат
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
            throw gcnew Exception("Error loading file: " + ex->Message);
        }
    }

    // Экспорт в Excel
    void ExportToExcel(String^ filePath, bool appendToExisting) {
        try {
            Type^ excelType = Type::GetTypeFromProgID("Excel.Application");
            if (excelType == nullptr) {
                throw gcnew Exception("Microsoft Excel is not installed on this computer");
            }

            Object^ excelObj = Activator::CreateInstance(excelType);
            Object^ workbooks = excelType->InvokeMember("Workbooks", 
                System::Reflection::BindingFlags::GetProperty, nullptr, excelObj, nullptr);
            
            Object^ workbook = nullptr;
            if (appendToExisting && File::Exists(filePath)) {
                // Открываем существующий файл
                workbook = workbooks->GetType()->InvokeMember("Open",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, workbooks, 
                    gcnew array<Object^> { filePath });
            } else {
                // Создаем новый файл
                workbook = workbooks->GetType()->InvokeMember("Add",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, workbooks, nullptr);
            }
            
            // Добавляем новый лист в конец (если открыли существующий файл)
            Object^ worksheets = workbook->GetType()->InvokeMember("Worksheets",
                System::Reflection::BindingFlags::GetProperty, nullptr, workbook, nullptr);
            
            Object^ worksheet = nullptr;
            if (appendToExisting && File::Exists(filePath)) {
                int count = Convert::ToInt32(worksheets->GetType()->InvokeMember("Count",
                    System::Reflection::BindingFlags::GetProperty, nullptr, worksheets, nullptr));
                
                // Получаем последний лист и добавляем после него новый
                Object^ lastSheet = worksheets->GetType()->InvokeMember("Item",
                    System::Reflection::BindingFlags::GetProperty, nullptr, worksheets, 
                    gcnew array<Object^> { count });
                
                worksheet = worksheets->GetType()->InvokeMember("Add",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, worksheets, 
                    gcnew array<Object^> { System::Reflection::Missing::Value, lastSheet });
                
                // Задаем имя нового листа с текущей датой/временем
                String^ sheetName = "Contacts " + DateTime::Now.ToString("yyyy-MM-dd HH:mm");
                worksheet->GetType()->InvokeMember("Name",
                    System::Reflection::BindingFlags::SetProperty, nullptr, worksheet,
                    gcnew array<Object^> { sheetName });
            } else {
                // Используем активный лист в новом файле
                worksheet = workbook->GetType()->InvokeMember("ActiveSheet",
                    System::Reflection::BindingFlags::GetProperty, nullptr, workbook, nullptr);
            }

            // Заголовки
            array<String^>^ headers = {"ID", "First Name", "Last Name", "Phone", "Birth Date", "Email", "Address", "Notes"};
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
            if (!appendToExisting || !File::Exists(filePath)) {
                workbook->GetType()->InvokeMember("SaveAs",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, workbook,
                    gcnew array<Object^> { filePath });
            } else {
                workbook->GetType()->InvokeMember("Save",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, workbook, nullptr);
            }

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
            throw gcnew Exception("Error exporting to Excel: " + ex->Message);
        }
    }
    
    // Перегрузка для обратной совместимости
    void ExportToExcel(String^ filePath) {
        ExportToExcel(filePath, false);
    }
    
    // Экспорт в Word
    void ExportToWord(String^ filePath, bool appendToExisting) {
        try {
            Type^ wordType = Type::GetTypeFromProgID("Word.Application");
            if (wordType == nullptr) {
                throw gcnew Exception("Microsoft Word is not installed on this computer");
            }

            Object^ wordObj = Activator::CreateInstance(wordType);
            Object^ documents = wordType->InvokeMember("Documents", 
                System::Reflection::BindingFlags::GetProperty, nullptr, wordObj, nullptr);
            
            Object^ document = nullptr;
            if (appendToExisting && File::Exists(filePath)) {
                // Открываем существующий документ
                document = documents->GetType()->InvokeMember("Open",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, documents, 
                    gcnew array<Object^> { filePath });
                
                // Перемещаем курсор в конец документа
                Object^ selection = wordType->InvokeMember("Selection",
                    System::Reflection::BindingFlags::GetProperty, nullptr, wordObj, nullptr);
                    
                selection->GetType()->InvokeMember("EndKey",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, selection,
                    gcnew array<Object^> { 6 });  // wdStory = 6
                
                // Добавляем разрыв страницы
                selection->GetType()->InvokeMember("InsertBreak",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, selection,
                    gcnew array<Object^> { 7 });  // wdPageBreak = 7
            } else {
                // Создаем новый документ
                document = documents->GetType()->InvokeMember("Add",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, documents, nullptr);
            }
            
            // Получаем объект selection для работы с документом
            Object^ selection = wordType->InvokeMember("Selection",
                System::Reflection::BindingFlags::GetProperty, nullptr, wordObj, nullptr);
            
            // Добавляем заголовок
            selection->GetType()->InvokeMember("Font",
                System::Reflection::BindingFlags::GetProperty, nullptr, selection, nullptr)
                ->GetType()->InvokeMember("Size",
                    System::Reflection::BindingFlags::SetProperty, nullptr, 
                    selection->GetType()->InvokeMember("Font",
                        System::Reflection::BindingFlags::GetProperty, nullptr, selection, nullptr),
                    gcnew array<Object^> { 16 });
                    
            selection->GetType()->InvokeMember("TypeText",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection,
                gcnew array<Object^> { "Contacts List" });
                
            selection->GetType()->InvokeMember("TypeParagraph",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection, nullptr);
                
            selection->GetType()->InvokeMember("TypeParagraph",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection, nullptr);
            
            // Добавляем текущую дату
            selection->GetType()->InvokeMember("Font",
                System::Reflection::BindingFlags::GetProperty, nullptr, selection, nullptr)
                ->GetType()->InvokeMember("Size",
                    System::Reflection::BindingFlags::SetProperty, nullptr, 
                    selection->GetType()->InvokeMember("Font",
                        System::Reflection::BindingFlags::GetProperty, nullptr, selection, nullptr),
                    gcnew array<Object^> { 10 });
                    
            selection->GetType()->InvokeMember("TypeText",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection,
                gcnew array<Object^> { "Generated: " + DateTime::Now.ToString() });
                
            selection->GetType()->InvokeMember("TypeParagraph",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection, nullptr);
                
            selection->GetType()->InvokeMember("TypeParagraph",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection, nullptr);
            
            // Получаем доступ к таблицам в документе
            Object^ tables = document->GetType()->InvokeMember("Tables",
                System::Reflection::BindingFlags::GetProperty, nullptr, document, nullptr);
            
            // Добавляем таблицу (8 колонок, строк по числу контактов + 1 для заголовка)
            Object^ table = tables->GetType()->InvokeMember("Add",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, tables,
                gcnew array<Object^> { 
                    selection->GetType()->InvokeMember("Range",
                        System::Reflection::BindingFlags::GetProperty, nullptr, selection, nullptr),
                    entries->Count + 1, 8, 1, 1 });
            
            // Заголовки
            array<String^>^ headers = {"ID", "First Name", "Last Name", "Phone", "Birth Date", "Email", "Address", "Notes"};
            Object^ headerRow = table->GetType()->InvokeMember("Rows",
                System::Reflection::BindingFlags::GetProperty, nullptr, table, nullptr)
                ->GetType()->InvokeMember("Item",
                    System::Reflection::BindingFlags::GetProperty, nullptr, 
                    table->GetType()->InvokeMember("Rows",
                        System::Reflection::BindingFlags::GetProperty, nullptr, table, nullptr),
                    gcnew array<Object^> { 1 });
            
            // Устанавливаем фон для заголовка
            headerRow->GetType()->InvokeMember("Shading",
                System::Reflection::BindingFlags::GetProperty, nullptr, headerRow, nullptr)
                ->GetType()->InvokeMember("BackgroundPatternColor",
                    System::Reflection::BindingFlags::SetProperty, nullptr,
                    headerRow->GetType()->InvokeMember("Shading",
                        System::Reflection::BindingFlags::GetProperty, nullptr, headerRow, nullptr),
                    gcnew array<Object^> { 16777215 });  // wdColorGray15 или светло-серый
            
            // Заполняем заголовки
            for (int i = 0; i < headers->Length; i++) {
                Object^ cell = table->GetType()->InvokeMember("Cell",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, table,
                    gcnew array<Object^> { 1, i + 1 });
                    
                cell->GetType()->InvokeMember("Range",
                    System::Reflection::BindingFlags::GetProperty, nullptr, cell, nullptr)
                    ->GetType()->InvokeMember("Text",
                        System::Reflection::BindingFlags::SetProperty, nullptr,
                        cell->GetType()->InvokeMember("Range",
                            System::Reflection::BindingFlags::GetProperty, nullptr, cell, nullptr),
                        gcnew array<Object^> { headers[i] });
            }
            
            // Заполняем данные
            int row = 2;
            for each (NotebookEntry<int>^ entry in entries) {
                array<Object^>^ rowData = {
                    entry->GetId().ToString(),
                    entry->GetFirstName(),
                    entry->GetLastName(),
                    entry->GetPhoneNumber(),
                    entry->GetBirthDate(),
                    entry->GetEmail(),
                    entry->GetAddress(),
                    entry->GetNotes()
                };

                for (int col = 0; col < rowData->Length; col++) {
                    Object^ cell = table->GetType()->InvokeMember("Cell",
                        System::Reflection::BindingFlags::InvokeMethod, nullptr, table,
                        gcnew array<Object^> { row, col + 1 });
                        
                    cell->GetType()->InvokeMember("Range",
                        System::Reflection::BindingFlags::GetProperty, nullptr, cell, nullptr)
                        ->GetType()->InvokeMember("Text",
                            System::Reflection::BindingFlags::SetProperty, nullptr,
                            cell->GetType()->InvokeMember("Range",
                                System::Reflection::BindingFlags::GetProperty, nullptr, cell, nullptr),
                            gcnew array<Object^> { rowData[col] });
                }
                row++;
            }
            
            // Автоподбор таблицы
            table->GetType()->InvokeMember("AutoFitBehavior",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, table,
                gcnew array<Object^> { 1 });  // wdAutoFitContent = 1
            
            // Перемещаем курсор после таблицы
            selection->GetType()->InvokeMember("MoveDown",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, selection, nullptr);
            
            // Сохранение
            if (!appendToExisting || !File::Exists(filePath)) {
                document->GetType()->InvokeMember("SaveAs",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, document,
                    gcnew array<Object^> { filePath });
            } else {
                document->GetType()->InvokeMember("Save",
                    System::Reflection::BindingFlags::InvokeMethod, nullptr, document, nullptr);
            }

            // Закрытие
            document->GetType()->InvokeMember("Close",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, document,
                gcnew array<Object^> { true });
            wordType->InvokeMember("Quit",
                System::Reflection::BindingFlags::InvokeMethod, nullptr, wordObj, nullptr);

            System::Runtime::InteropServices::Marshal::ReleaseComObject(document);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(documents);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(wordObj);
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Error exporting to Word: " + ex->Message);
        }
    }
    
    // Перегрузка для обратной совместимости
    void ExportToWord(String^ filePath) {
        ExportToWord(filePath, false);
    }
}; 
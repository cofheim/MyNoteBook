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

// Не используем общие пространства имен для Office, чтобы избежать конфликтов
// using namespace Microsoft::Office::Interop::Word;
// using namespace Microsoft::Office::Interop::Excel;

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

    // Вспомогательный класс для сортировки по ID
    ref class IdComparer : IComparer<NotebookEntry<int>^> {
    public:
        virtual int Compare(NotebookEntry<int>^ x, NotebookEntry<int>^ y) {
            return x->GetId() - y->GetId();
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
    
    // Сортировка по ID
    void SortById() {
        if (entries->Count > 0) {
            IdComparer^ comparer = gcnew IdComparer();
            entries->Sort(comparer);
            // Автоматически сохраняем в JSON после сортировки
            SaveToJsonFile(defaultJsonPath);
        }
    }
    
    // Обновление и сортировка контактов
    bool RefreshAndSortContacts() {
        if (entries->Count > 0) {
            // Сортировка по ID
            SortById();
            return true;
        }
        return false;
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
            Object^ missingValue = Type::Missing;
            
            // Создаем COM-объект Excel - используем полностью квалифицированные имена типов
            Microsoft::Office::Interop::Excel::Application^ excelApp = 
                gcnew Microsoft::Office::Interop::Excel::Application();
            excelApp->Visible = false;
            
            Microsoft::Office::Interop::Excel::Workbook^ workbook;
            Microsoft::Office::Interop::Excel::Worksheet^ worksheet;
            
            if (appendToExisting && File::Exists(filePath)) {
                // Открываем существующий файл - используем все параметры явно
                workbook = excelApp->Workbooks->Open(
                    filePath,                 // Filename
                    missingValue,             // UpdateLinks
                    false,                    // ReadOnly
                    missingValue,             // Format
                    missingValue,             // Password
                    missingValue,             // WriteResPassword
                    missingValue,             // IgnoreReadOnlyRecommended
                    missingValue,             // Origin
                    missingValue,             // Delimiter
                    missingValue,             // Editable
                    missingValue,             // Notify
                    missingValue,             // Converter
                    missingValue,             // AddToMru
                    missingValue,             // Local
                    missingValue              // CorruptLoad
                );
                
                // Добавляем новый лист с именем, включающим текущую дату и время
                String^ sheetName = "Contacts_" + DateTime::Now.ToString("yyyy-MM-dd_HH-mm-ss");
                worksheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(
                    workbook->Sheets->Add(
                        missingValue, 
                        workbook->Sheets[workbook->Sheets->Count], 
                        missingValue, 
                        missingValue
                    )
                );
                worksheet->Name = sheetName;
            } else {
                // Создаем новый файл
                workbook = excelApp->Workbooks->Add(missingValue);
                worksheet = safe_cast<Microsoft::Office::Interop::Excel::Worksheet^>(workbook->Sheets[1]);
                worksheet->Name = "Contacts";
            }
            
            // Заголовки колонок
            array<String^>^ headers = {"ID", "First Name", "Last Name", "Phone", "Birth Date", "Email", "Address", "Notes"};
            
            // Заголовок таблицы
            for (int i = 0; i < headers->Length; i++) {
                worksheet->Cells[1, i + 1] = headers[i];
                
                // Форматирование заголовка
                Microsoft::Office::Interop::Excel::Range^ headerCell = 
                    safe_cast<Microsoft::Office::Interop::Excel::Range^>(worksheet->Cells[1, i + 1]);
                    
                headerCell->Font->Bold = true;
                headerCell->Interior->Color = System::Drawing::ColorTranslator::ToOle(System::Drawing::Color::LightGray);
                headerCell->Borders->LineStyle = Microsoft::Office::Interop::Excel::XlLineStyle::xlContinuous;
                headerCell->HorizontalAlignment = Microsoft::Office::Interop::Excel::XlHAlign::xlHAlignCenter;
            }
            
            // Данные
            int row = 2;
            for each (NotebookEntry<int>^ entry in entries) {
                worksheet->Cells[row, 1] = entry->GetId();
                worksheet->Cells[row, 2] = entry->GetFirstName();
                worksheet->Cells[row, 3] = entry->GetLastName();
                
                // Установим формат ячейки с телефоном как текстовый перед вставкой данных
                Microsoft::Office::Interop::Excel::Range^ phoneCell = safe_cast<Microsoft::Office::Interop::Excel::Range^>(worksheet->Cells[row, 4]);
                phoneCell->NumberFormat = "@"; // Установка текстового формата
                phoneCell->Value2 = entry->GetPhoneNumber(); // Используем Value2 вместо простого присваивания
                
                worksheet->Cells[row, 5] = entry->GetBirthDate();
                worksheet->Cells[row, 6] = entry->GetEmail();
                worksheet->Cells[row, 7] = entry->GetAddress();
                worksheet->Cells[row, 8] = entry->GetNotes();
                
                // Добавляем границы ячеек с данными
                Microsoft::Office::Interop::Excel::Range^ dataRange = worksheet->Range[
                    worksheet->Cells[row, 1], 
                    worksheet->Cells[row, headers->Length]];
                dataRange->Borders->LineStyle = Microsoft::Office::Interop::Excel::XlLineStyle::xlContinuous;
                
                row++;
            }
            
            // Автоподбор ширины колонок
            Microsoft::Office::Interop::Excel::Range^ usedRange = worksheet->UsedRange;
            usedRange->Columns->AutoFit();
            
            // Сохранение и закрытие
            if (!appendToExisting || !File::Exists(filePath)) {
                workbook->SaveAs(
                    filePath,                                                // Filename
                    missingValue,                                            // FileFormat
                    missingValue,                                            // Password
                    missingValue,                                            // WriteResPassword
                    missingValue,                                            // ReadOnlyRecommended
                    missingValue,                                            // CreateBackup
                    Microsoft::Office::Interop::Excel::XlSaveAsAccessMode::xlNoChange,  // AccessMode
                    missingValue,                                            // ConflictResolution
                    missingValue,                                            // AddToMru
                    missingValue,                                            // TextCodepage
                    missingValue,                                            // TextVisualLayout
                    missingValue                                             // Local
                );
            } else {
                workbook->Save();
            }
            
            workbook->Close(true, missingValue, missingValue);
            excelApp->Quit();
            
            // Освобождаем ресурсы
            System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheet);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
            System::Runtime::InteropServices::Marshal::ReleaseComObject(excelApp);
        }
        catch (Exception^ ex) {
            throw gcnew Exception("Error exporting to Excel: " + ex->Message);
        }
    }
    
    // Перегрузка для обратной совместимости
    void ExportToExcel(String^ filePath) {
        ExportToExcel(filePath, false);
    }
    
    // Получение максимального ID
    int GetMaxId() {
        int maxId = 0;
        for each (NotebookEntry<int>^ entry in entries) {
            if (entry->GetId() > maxId) {
                maxId = entry->GetId();
            }
        }
        return maxId;
    }
}; 
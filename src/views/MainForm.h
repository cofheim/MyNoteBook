#pragma once
#include "../controllers/NotebookManager.h"
#include "../utils/ValidationUtils.h"

namespace NBapp {

using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Windows::Forms;
using namespace System::Data;
using namespace System::Drawing;
using namespace System::Text;
using namespace System::Globalization;
using namespace System::Threading;

public ref class MainForm : public System::Windows::Forms::Form
{
public:
    MainForm(void)
    {
        // Инициализация компонентов формы
        InitializeComponent();
        
        // Инициализация менеджера записей
        manager = gcnew NotebookManager();
        
        // Установим начальный ID
        // Если в списке уже есть записи (загруженные из JSON), используем максимальный ID + 1
        currentId = manager->GetMaxId() + 1;
        
        // Настройка обработчиков ввода
        SetupInputHandlers();
        
        // Обновляем dataGridView
        RefreshDataGrid();
    }

protected:
    ~MainForm()
    {
        if (components)
        {
            delete components;
        }
    }

private:
    NotebookManager^ manager;
    int currentId;
    System::ComponentModel::Container^ components;

    // Компоненты формы
    System::Windows::Forms::MenuStrip^ menuStrip;
    System::Windows::Forms::ToolStripMenuItem^ fileMenu;
    System::Windows::Forms::ToolStripMenuItem^ newFileMenuItem;
    System::Windows::Forms::ToolStripMenuItem^ openFileMenuItem;
    System::Windows::Forms::ToolStripMenuItem^ saveFileMenuItem;
    System::Windows::Forms::ToolStripMenuItem^ exportMenu;
    System::Windows::Forms::ToolStripMenuItem^ exportExcelMenuItem;
    System::Windows::Forms::ToolStripMenuItem^ exportExcelNewMenuItem;
    System::Windows::Forms::ToolStripMenuItem^ exportExcelExistingMenuItem;
    System::Windows::Forms::ToolStripSeparator^ toolStripSeparator;
    System::Windows::Forms::ToolStripMenuItem^ exitMenuItem;

    System::Windows::Forms::DataGridView^ dataGridView;
    System::Windows::Forms::GroupBox^ searchGroupBox;
    System::Windows::Forms::TextBox^ searchTextBox;
    System::Windows::Forms::ComboBox^ searchTypeComboBox;
    System::Windows::Forms::Button^ searchButton;
    System::Windows::Forms::Button^ showAllButton;
    System::Windows::Forms::Button^ refreshButton;

    System::Windows::Forms::GroupBox^ addEntryGroupBox;
    System::Windows::Forms::TextBox^ firstNameTextBox;
    System::Windows::Forms::TextBox^ lastNameTextBox;
    System::Windows::Forms::TextBox^ phoneTextBox;
    System::Windows::Forms::TextBox^ birthDateTextBox;
    System::Windows::Forms::TextBox^ emailTextBox;
    System::Windows::Forms::TextBox^ addressTextBox;
    System::Windows::Forms::TextBox^ notesTextBox;
    System::Windows::Forms::Button^ addButton;
    System::Windows::Forms::Button^ deleteButton;

    void InitializeComponent(void)
    {
        // Настраиваем параметры формы
        this->components = gcnew System::ComponentModel::Container();
        this->Size = System::Drawing::Size(1024, 768);
        this->Text = "My Notebook";
        this->StartPosition = FormStartPosition::CenterScreen;
        this->Padding = System::Windows::Forms::Padding(0);
        this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
        this->Font = gcnew System::Drawing::Font("Microsoft Sans Serif", 9);

        // Инициализация меню
        this->menuStrip = gcnew MenuStrip();
        this->fileMenu = gcnew ToolStripMenuItem("File");
        this->newFileMenuItem = gcnew ToolStripMenuItem("New");
        this->openFileMenuItem = gcnew ToolStripMenuItem("Open");
        this->saveFileMenuItem = gcnew ToolStripMenuItem("Save");
        this->exportMenu = gcnew ToolStripMenuItem("Export");
        this->exportExcelMenuItem = gcnew ToolStripMenuItem("Export to Excel");
        this->exportExcelNewMenuItem = gcnew ToolStripMenuItem("New File");
        this->exportExcelExistingMenuItem = gcnew ToolStripMenuItem("Existing File");
        this->toolStripSeparator = gcnew ToolStripSeparator();
        this->exitMenuItem = gcnew ToolStripMenuItem("Exit");

        // Настраиваем подменю экспорта
        this->exportExcelMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
            this->exportExcelNewMenuItem,
            this->exportExcelExistingMenuItem
        });

        this->exportMenu->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) {
            this->exportExcelMenuItem
        });

        this->fileMenu->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(6) {
            this->newFileMenuItem,
            this->openFileMenuItem,
            this->saveFileMenuItem,
            this->exportMenu,
            this->toolStripSeparator,
            this->exitMenuItem
        });

        this->menuStrip->Items->Add(this->fileMenu);
        this->Controls->Add(this->menuStrip);

        // Инициализация DataGridView
        this->dataGridView = gcnew DataGridView();
        this->dataGridView->Location = Point(10, 100);
        this->dataGridView->Size = System::Drawing::Size(800, 400);
        this->dataGridView->AllowUserToAddRows = false;
        this->dataGridView->AllowUserToDeleteRows = false;
        this->dataGridView->ReadOnly = true;
        this->dataGridView->MultiSelect = false;
        this->dataGridView->SelectionMode = DataGridViewSelectionMode::FullRowSelect;
        this->dataGridView->AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode::Fill;

        // Добавление столбцов
        this->dataGridView->Columns->Add("Id", "ID");
        this->dataGridView->Columns->Add("FirstName", "First Name");
        this->dataGridView->Columns->Add("LastName", "Last Name");
        this->dataGridView->Columns->Add("Phone", "Phone");
        this->dataGridView->Columns->Add("BirthDate", "Birth Date");
        this->dataGridView->Columns->Add("Email", "Email");
        this->dataGridView->Columns->Add("Address", "Address");
        this->dataGridView->Columns->Add("Notes", "Notes");

        // Инициализация группы поиска
        this->searchGroupBox = gcnew GroupBox();
        this->searchGroupBox->Text = "Search";
        this->searchGroupBox->Location = Point(10, 30);
        this->searchGroupBox->Size = System::Drawing::Size(800, 60);

        this->searchTypeComboBox = gcnew ComboBox();
        this->searchTypeComboBox->Location = Point(10, 20);
        this->searchTypeComboBox->Size = System::Drawing::Size(150, 25);
        this->searchTypeComboBox->Items->AddRange(gcnew cli::array<String^>(5) { 
            "By First Name", 
            "By Last Name", 
            "By Phone", 
            "By Email", 
            "By Address" 
        });
        this->searchTypeComboBox->SelectedIndex = 0;

        this->searchTextBox = gcnew TextBox();
        this->searchTextBox->Location = Point(170, 20);
        this->searchTextBox->Size = System::Drawing::Size(400, 25);

        this->searchButton = gcnew Button();
        this->searchButton->Text = "Search";
        this->searchButton->Location = Point(470, 20);
        this->searchButton->Size = System::Drawing::Size(100, 25);
        this->searchButton->Click += gcnew EventHandler(this, &MainForm::SearchButton_Click);
        
        this->showAllButton = gcnew Button();
        this->showAllButton->Text = "Show All";
        this->showAllButton->Location = Point(690, 20);
        this->showAllButton->Size = System::Drawing::Size(100, 25);
        this->showAllButton->Click += gcnew EventHandler(this, &MainForm::ShowAllButton_Click);
        
        this->refreshButton = gcnew Button();
        this->refreshButton->Text = "Refresh";
        this->refreshButton->Location = Point(580, 20);
        this->refreshButton->Size = System::Drawing::Size(100, 25);
        this->refreshButton->Click += gcnew EventHandler(this, &MainForm::RefreshButton_Click);

        // Инициализация группы добавления записи
        this->addEntryGroupBox = gcnew GroupBox();
        this->addEntryGroupBox->Text = "Add Entry";
        this->addEntryGroupBox->Location = Point(10, 510);
        this->addEntryGroupBox->Size = System::Drawing::Size(800, 200);

        // Создание меток и полей ввода
        Label^ firstNameLabel = gcnew Label();
        firstNameLabel->Text = "First Name:";
        firstNameLabel->Location = Point(10, 20);
        firstNameLabel->Size = System::Drawing::Size(200, 20);

        this->firstNameTextBox = gcnew TextBox();
        this->firstNameTextBox->Location = Point(10, 40);
        this->firstNameTextBox->Size = System::Drawing::Size(200, 25);

        Label^ lastNameLabel = gcnew Label();
        lastNameLabel->Text = "Last Name:";
        lastNameLabel->Location = Point(220, 20);
        lastNameLabel->Size = System::Drawing::Size(200, 20);

        this->lastNameTextBox = gcnew TextBox();
        this->lastNameTextBox->Location = Point(220, 40);
        this->lastNameTextBox->Size = System::Drawing::Size(200, 25);

        Label^ phoneLabel = gcnew Label();
        phoneLabel->Text = "Phone:";
        phoneLabel->Location = Point(430, 20);
        phoneLabel->Size = System::Drawing::Size(200, 20);

        this->phoneTextBox = gcnew TextBox();
        this->phoneTextBox->Location = Point(430, 40);
        this->phoneTextBox->Size = System::Drawing::Size(200, 25);

        Label^ birthDateLabel = gcnew Label();
        birthDateLabel->Text = "Birth Date:";
        birthDateLabel->Location = Point(640, 20);
        birthDateLabel->Size = System::Drawing::Size(150, 20);

        this->birthDateTextBox = gcnew TextBox();
        this->birthDateTextBox->Location = Point(640, 40);
        this->birthDateTextBox->Size = System::Drawing::Size(150, 25);

        Label^ emailLabel = gcnew Label();
        emailLabel->Text = "Email:";
        emailLabel->Location = Point(10, 60);
        emailLabel->Size = System::Drawing::Size(300, 20);

        this->emailTextBox = gcnew TextBox();
        this->emailTextBox->Location = Point(10, 80);
        this->emailTextBox->Size = System::Drawing::Size(300, 25);

        Label^ addressLabel = gcnew Label();
        addressLabel->Text = "Address:";
        addressLabel->Location = Point(320, 60);
        addressLabel->Size = System::Drawing::Size(470, 20);

        this->addressTextBox = gcnew TextBox();
        this->addressTextBox->Location = Point(320, 80);
        this->addressTextBox->Size = System::Drawing::Size(470, 25);

        Label^ notesLabel = gcnew Label();
        notesLabel->Text = "Notes:";
        notesLabel->Location = Point(10, 100);
        notesLabel->Size = System::Drawing::Size(780, 20);

        this->notesTextBox = gcnew TextBox();
        this->notesTextBox->Location = Point(10, 120);
        this->notesTextBox->Size = System::Drawing::Size(780, 25);

        // Кнопки
        this->addButton = gcnew Button();
        this->addButton->Text = "Add";
        this->addButton->Location = Point(10, 160);
        this->addButton->Size = System::Drawing::Size(100, 30);
        this->addButton->Click += gcnew EventHandler(this, &MainForm::AddButton_Click);

        this->deleteButton = gcnew Button();
        this->deleteButton->Text = "Delete";
        this->deleteButton->Location = Point(120, 160);
        this->deleteButton->Size = System::Drawing::Size(100, 30);
        this->deleteButton->Click += gcnew EventHandler(this, &MainForm::DeleteButton_Click);

        // Добавление элементов управления на форму
        this->Controls->Add(this->menuStrip);
        this->Controls->Add(this->dataGridView);
        this->Controls->Add(this->searchGroupBox);
        this->Controls->Add(this->addEntryGroupBox);

        // Добавление элементов в группу поиска
        this->searchGroupBox->Controls->Add(this->searchTypeComboBox);
        this->searchGroupBox->Controls->Add(this->searchTextBox);
        this->searchGroupBox->Controls->Add(this->searchButton);
        this->searchGroupBox->Controls->Add(this->showAllButton);
        this->searchGroupBox->Controls->Add(this->refreshButton);

        // Добавление элементов в группу добавления записи
        this->addEntryGroupBox->Controls->Add(firstNameLabel);
        this->addEntryGroupBox->Controls->Add(lastNameLabel);
        this->addEntryGroupBox->Controls->Add(phoneLabel);
        this->addEntryGroupBox->Controls->Add(birthDateLabel);
        this->addEntryGroupBox->Controls->Add(emailLabel);
        this->addEntryGroupBox->Controls->Add(addressLabel);
        this->addEntryGroupBox->Controls->Add(notesLabel);
        this->addEntryGroupBox->Controls->Add(this->firstNameTextBox);
        this->addEntryGroupBox->Controls->Add(this->lastNameTextBox);
        this->addEntryGroupBox->Controls->Add(this->phoneTextBox);
        this->addEntryGroupBox->Controls->Add(this->birthDateTextBox);
        this->addEntryGroupBox->Controls->Add(this->emailTextBox);
        this->addEntryGroupBox->Controls->Add(this->addressTextBox);
        this->addEntryGroupBox->Controls->Add(this->notesTextBox);
        this->addEntryGroupBox->Controls->Add(this->addButton);
        this->addEntryGroupBox->Controls->Add(this->deleteButton);

        // Привязка обработчиков событий меню
        this->newFileMenuItem->Click += gcnew EventHandler(this, &MainForm::NewFile_Click);
        this->openFileMenuItem->Click += gcnew EventHandler(this, &MainForm::OpenFile_Click);
        this->saveFileMenuItem->Click += gcnew EventHandler(this, &MainForm::SaveFile_Click);
        this->exportExcelNewMenuItem->Click += gcnew EventHandler(this, &MainForm::ExportExcel_Click);
        this->exportExcelExistingMenuItem->Click += gcnew EventHandler(this, &MainForm::ExportExcel_Click);
        this->exitMenuItem->Click += gcnew EventHandler(this, &MainForm::Exit_Click);
    }

    // Настройка обработчиков ввода
    void SetupInputHandlers()
    {
        // Обработчики для имени и фамилии (только буквы)
        this->firstNameTextBox->KeyPress += gcnew KeyPressEventHandler(this, &MainForm::NameField_KeyPress);
        this->lastNameTextBox->KeyPress += gcnew KeyPressEventHandler(this, &MainForm::NameField_KeyPress);
        
        // Обработчик для телефона (не буквы)
        this->phoneTextBox->KeyPress += gcnew KeyPressEventHandler(this, &MainForm::PhoneField_KeyPress);
    }

    // Обработчик ввода для полей имени и фамилии
    System::Void NameField_KeyPress(System::Object^ sender, KeyPressEventArgs^ e)
    {
        // Если вводимый символ не буква, пробел, дефис или управляющий символ - отменяем ввод
        if (!ValidationUtils::IsLetterInput(e->KeyChar)) {
            e->Handled = true; // Отменяем ввод символа
        }
    }

    // Обработчик ввода для поля телефона
    System::Void PhoneField_KeyPress(System::Object^ sender, KeyPressEventArgs^ e)
    {
        // Если вводимый символ - буква - отменяем ввод
        if (Char::IsLetter(e->KeyChar)) {
            e->Handled = true; // Отменяем ввод символа
        }
    }

    // Обработчики событий
    System::Void AddButton_Click(System::Object^ sender, System::EventArgs^ e)
    {
        // Упрощенная валидация
        if (String::IsNullOrEmpty(firstNameTextBox->Text) || String::IsNullOrEmpty(lastNameTextBox->Text)) {
            MessageBox::Show("First name and last name are required!", "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
            return;
        }
        
        if (String::IsNullOrEmpty(phoneTextBox->Text)) {
            MessageBox::Show("Phone number is required!", "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
            return;
        }

        // Создание новой записи
        NotebookEntry<int>^ entry = gcnew NotebookEntry<int>(
            currentId++,
            firstNameTextBox->Text,
            lastNameTextBox->Text,
            phoneTextBox->Text,
            birthDateTextBox->Text,
            emailTextBox->Text,
            addressTextBox->Text,
            notesTextBox->Text
        );

        // Добавление записи
        manager->AddEntry(entry);

        // Обновление таблицы
        RefreshDataGrid();

        // Очистка полей ввода
        ClearInputFields();
    }

    System::Void DeleteButton_Click(System::Object^ sender, System::EventArgs^ e)
    {
        if (dataGridView->SelectedRows->Count > 0) {
            if (MessageBox::Show("Are you sure you want to delete the selected entry?", "Confirmation",
                MessageBoxButtons::YesNo, MessageBoxIcon::Question) == System::Windows::Forms::DialogResult::Yes)
            {
                int id = Convert::ToInt32(dataGridView->SelectedRows[0]->Cells["Id"]->Value);
                if (manager->RemoveEntry(id)) {
                    RefreshDataGrid();
                }
            }
        }
    }

    System::Void SearchButton_Click(System::Object^ sender, System::EventArgs^ e)
    {
        // Поиск с использованием выбранного фильтра
        List<NotebookEntry<int>^>^ results = manager->SearchByAnyField(
            searchTextBox->Text, 
            searchTypeComboBox->SelectedIndex
        );

        // Отображение результатов
        dataGridView->Rows->Clear();
        for each (NotebookEntry<int>^ entry in results) {
            dataGridView->Rows->Add(
                entry->GetId(),
                entry->GetFirstName(),
                entry->GetLastName(),
                entry->GetPhoneNumber(),
                entry->GetBirthDate(),
                entry->GetEmail(),
                entry->GetAddress(),
                entry->GetNotes()
            );
        }
    }

    System::Void ShowAllButton_Click(System::Object^ sender, System::EventArgs^ e)
    {
        // Очищаем поле поиска
        searchTextBox->Clear();
        
        // Обновляем таблицу всеми записями
        RefreshDataGrid();
    }

    System::Void RefreshButton_Click(System::Object^ sender, System::EventArgs^ e)
    {
        // Обновляем и сортируем контакты, если они есть
        if (manager->RefreshAndSortContacts()) {
            // Обновляем таблицу отсортированными записями
            RefreshDataGrid();
            MessageBox::Show("Contact list has been updated and sorted by ID.", "Information", 
                MessageBoxButtons::OK, MessageBoxIcon::Information);
        } else {
            MessageBox::Show("Contact list is empty. Nothing to update.", "Information", 
                MessageBoxButtons::OK, MessageBoxIcon::Information);
        }
    }

    System::Void NewFile_Click(System::Object^ sender, System::EventArgs^ e)
    {
        if (MessageBox::Show("Create a new file? Unsaved data will be lost.", "Confirmation",
            MessageBoxButtons::YesNo, MessageBoxIcon::Question) == System::Windows::Forms::DialogResult::Yes)
        {
            manager = gcnew NotebookManager();
            currentId = 1;
            RefreshDataGrid();
        }
    }

    System::Void OpenFile_Click(System::Object^ sender, System::EventArgs^ e)
    {
        OpenFileDialog^ openFileDialog = gcnew OpenFileDialog();
        openFileDialog->Filter = "Text files (*.txt)|*.txt|JSON files (*.json)|*.json|All files (*.*)|*.*";
        openFileDialog->Title = "Open File";

        if (openFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
            try {
                manager->LoadFromFile(openFileDialog->FileName);
                RefreshDataGrid();
                
                // Обновляем currentId на максимальный ID + 1
                currentId = manager->GetMaxId() + 1;
            }
            catch (Exception^ ex) {
                MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
            }
        }
    }

    System::Void SaveFile_Click(System::Object^ sender, System::EventArgs^ e)
    {
        SaveFileDialog^ saveFileDialog = gcnew SaveFileDialog();
        saveFileDialog->Filter = "JSON files (*.json)|*.json|Text files (*.txt)|*.txt|All files (*.*)|*.*";
        saveFileDialog->Title = "Save File";
        saveFileDialog->DefaultExt = "json";

        if (saveFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
            try {
                manager->SaveToFile(saveFileDialog->FileName);
                MessageBox::Show("File saved successfully!", "Information", 
                    MessageBoxButtons::OK, MessageBoxIcon::Information);
            }
            catch (Exception^ ex) {
                MessageBox::Show(ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
            }
        }
    }

    System::Void ExportExcel_Click(System::Object^ sender, System::EventArgs^ e)
    {
        SaveFileDialog^ saveFileDialog = gcnew SaveFileDialog();
        saveFileDialog->Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
        saveFileDialog->Title = "Export to Excel";
        saveFileDialog->DefaultExt = "xlsx";

        // Определяем, нужно ли добавлять в существующий файл
        bool appendToExisting = false;
        if (sender == exportExcelExistingMenuItem) {
            appendToExisting = true;
            saveFileDialog->Title = "Select existing Excel file";
            saveFileDialog->CheckFileExists = true;
        }

        if (saveFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK) {
            try {
                // Проверяем на существование файла при опции добавления
                if (appendToExisting && !File::Exists(saveFileDialog->FileName)) {
                    MessageBox::Show("The selected file does not exist.", "Error", 
                        MessageBoxButtons::OK, MessageBoxIcon::Error);
                    return;
                }

                manager->ExportToExcel(saveFileDialog->FileName, appendToExisting);
                MessageBox::Show("Export completed successfully!", "Information", 
                    MessageBoxButtons::OK, MessageBoxIcon::Information);
            }
            catch (Exception^ ex) {
                MessageBox::Show("Export error: " + ex->Message, "Error", 
                    MessageBoxButtons::OK, MessageBoxIcon::Error);
            }
        }
    }

    System::Void Exit_Click(System::Object^ sender, System::EventArgs^ e)
    {
        this->Close();
    }

    // Вспомогательные методы
    void RefreshDataGrid()
    {
        dataGridView->Rows->Clear();
        for each (NotebookEntry<int>^ entry in manager->GetAllEntries()) {
            dataGridView->Rows->Add(
                entry->GetId(),
                entry->GetFirstName(),
                entry->GetLastName(),
                entry->GetPhoneNumber(),
                entry->GetBirthDate(),
                entry->GetEmail(),
                entry->GetAddress(),
                entry->GetNotes()
            );
        }
    }

    void ClearInputFields()
    {
        firstNameTextBox->Clear();
        lastNameTextBox->Clear();
        phoneTextBox->Clear();
        birthDateTextBox->Clear();
        emailTextBox->Clear();
        addressTextBox->Clear();
        notesTextBox->Clear();
    }
};
} 
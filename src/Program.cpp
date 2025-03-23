#include "views/MainForm.h"

using namespace System;
using namespace System::Windows::Forms;

[STAThread]
int main(array<String^>^ args)
{
    // Включаем визуальные стили Windows
    Application::EnableVisualStyles();
    Application::SetCompatibleTextRenderingDefault(false);
    
    // Запускаем главную форму
    Application::Run(gcnew NBapp::MainForm());
    
    return 0;
} 
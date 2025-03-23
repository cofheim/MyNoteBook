#pragma once

using namespace System;

public ref class ValidationUtils {
public:
    // Проверка email
    static bool IsValidEmail(String^ email) {
        if (String::IsNullOrEmpty(email)) return true; // Email не обязателен
        
        // Простая проверка: должна содержать @ и точку после @
        return email->Contains("@") && email->IndexOf(".") > email->IndexOf("@");
    }

    // Проверка телефонного номера
    static bool IsValidPhoneNumber(String^ phone) {
        if (String::IsNullOrEmpty(phone)) return false; // Телефон обязателен
        
        // Проверяем, что в строке нет букв
        for (int i = 0; i < phone->Length; i++) {
            if (Char::IsLetter(phone[i])) {
                return false;
            }
        }
        
        // Удаляем все нецифровые символы для проверки
        String^ digitsOnly = "";
        for (int i = 0; i < phone->Length; i++) {
            if (Char::IsDigit(phone[i])) {
                digitsOnly += phone[i];
            }
        }
        
        // Телефон должен содержать 10-11 цифр
        return digitsOnly->Length >= 10 && digitsOnly->Length <= 11;
    }

    // Проверка даты рождения
    static bool IsValidBirthDate(String^ date) {
        if (String::IsNullOrEmpty(date)) return true; // Дата рождения не обязательна
        
        try {
            DateTime birthDate = DateTime::Parse(date);
            return birthDate <= DateTime::Now;
        }
        catch (...) {
            return false;
        }
    }

    // Проверка имени/фамилии - только буквы и дефис
    static bool IsValidName(String^ name) {
        if (String::IsNullOrEmpty(name)) return false; // Имя и фамилия обязательны
        
        // Проверяем, что в строке только буквы, пробелы и дефисы
        for (int i = 0; i < name->Length; i++) {
            if (!Char::IsLetter(name[i]) && name[i] != ' ' && name[i] != '-') {
                return false;
            }
        }
        
        // Проверяем длину имени
        return name->Length > 1 && name->Length < 50;
    }

    // Простой фильтр для ввода только букв (для имени и фамилии)
    static bool IsLetterInput(int keyChar) {
        // Разрешены буквы, пробел, дефис, backspace, delete и управляющие клавиши
        return Char::IsLetter((wchar_t)keyChar) || keyChar == ' ' || keyChar == '-' || 
               keyChar == 8 || keyChar == 127 || keyChar < 32;
    }
    
    // Простой фильтр для ввода телефона (только цифры и символы, но не буквы)
    static bool IsPhoneInput(int keyChar) {
        // Разрешены цифры, символы, но не буквы + управляющие клавиши
        return !Char::IsLetter((wchar_t)keyChar) || keyChar == 8 || keyChar == 127 || keyChar < 32;
    }

    // Форматирование телефонного номера
    static String^ FormatPhoneNumber(String^ phone) {
        if (String::IsNullOrEmpty(phone)) return phone;
        
        // Просто возвращаем номер как есть
        return phone;
    }
}; 
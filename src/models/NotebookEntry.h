#pragma once
#include <string>
#include <msclr\marshal_cppstd.h>

using namespace System;

// Шаблонный класс для хранения записи в записной книжке
template<typename T>
public ref class NotebookEntry {
private:
    T id;
    String^ firstName;
    String^ lastName;
    String^ phoneNumber;
    String^ birthDate;
    String^ email;
    String^ address;
    String^ notes;

public:
    // Конструктор по умолчанию
    NotebookEntry() {
        id = T();
        firstName = "";
        lastName = "";
        phoneNumber = "";
        birthDate = "";
        email = "";
        address = "";
        notes = "";
    }

    // Конструктор с параметрами
    NotebookEntry(T id, String^ firstName, String^ lastName, String^ phoneNumber,
                 String^ birthDate, String^ email, String^ address, String^ notes) {
        this->id = id;
        this->firstName = firstName;
        this->lastName = lastName;
        this->phoneNumber = phoneNumber;
        this->birthDate = birthDate;
        this->email = email;
        this->address = address;
        this->notes = notes;
    }

    // Геттеры
    T GetId() { return id; }
    String^ GetFirstName() { return firstName; }
    String^ GetLastName() { return lastName; }
    String^ GetPhoneNumber() { return phoneNumber; }
    String^ GetBirthDate() { return birthDate; }
    String^ GetEmail() { return email; }
    String^ GetAddress() { return address; }
    String^ GetNotes() { return notes; }

    // Сеттеры
    void SetId(T value) { id = value; }
    void SetFirstName(String^ value) { firstName = value; }
    void SetLastName(String^ value) { lastName = value; }
    void SetPhoneNumber(String^ value) { phoneNumber = value; }
    void SetBirthDate(String^ value) { birthDate = value; }
    void SetEmail(String^ value) { email = value; }
    void SetAddress(String^ value) { address = value; }
    void SetNotes(String^ value) { notes = value; }

    // Метод для получения полного имени
    String^ GetFullName() {
        return String::Format("{0} {1}", firstName, lastName);
    }

    // Метод для проверки валидности записи
    bool IsValid() {
        return !String::IsNullOrEmpty(firstName) &&
               !String::IsNullOrEmpty(lastName) &&
               !String::IsNullOrEmpty(phoneNumber);
    }
}; 
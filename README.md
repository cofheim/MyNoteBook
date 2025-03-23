# MyNoteBook

A simple contact management application created using C++ and Windows Forms.

## Features

- Adding, editing and deleting contacts
- Search for contacts by various criteria:
  - First Name
  - Last Name
  - Phone number
  - Email
  - Address
- Export contacts:
  - To Excel (new or existing file)
- Save and load contacts from files
- **JSON format support** for storing contacts
  - Automatic loading of contacts from JSON
  - Improved data storage
  - Better compatibility with modern applications

## Requirements

- Windows OS
- Visual Studio 2019 or newer
- .NET Framework 4.8
- Newtonsoft.Json library (included)
- Microsoft Excel (for Excel export functions)

## Build Instructions

1. Open `NBapp.sln` in Visual Studio
2. Build the solution (Ctrl+Shift+B)
3. Run the application (F5)

## JSON Support

The application uses JSON as the default format for storing contacts. At startup, the application automatically looks for the `contacts.json` file in the application directory and loads contacts from it. When adding or removing contacts, changes are automatically saved to this file.

## Export Features

### Export to Excel
Allows you to export the list of contacts to:
- New Excel file (.xlsx)
- Existing Excel file (as a new sheet)

## Screenshots

[Add screenshots here]

## License

[License information] 
# Student Management System - VB.NET Windows Forms

A complete, professional desktop application built with VB.NET and Microsoft Access Database.

## 🎯 Project Overview

This is a fully functional Student Management System featuring:
- ✅ Secure Authentication System
- ✅ Complete CRUD Operations for Students
- ✅ Advanced Search Functionality
- ✅ Microsoft Word Report Generation
- ✅ Professional UI with Error Handling
- ✅ Data Validation & Business Logic

## 📋 System Requirements

- Visual Studio 2019 or later
- .NET Framework 4.7.2 or higher
- Microsoft Office (for Word report generation)
- Windows 7 or later

## 🗂️ Project Structure

```
StudentManagementSystem/
├── Database/
│   └── StudentDB.accdb
├── Source Code/
│   ├── DatabaseConnection.vb
│   ├── LoginForm.vb
│   ├── MainForm.vb
│   ├── StudentManagementForm.vb
│   ├── StudentModule.vb
│   └── ReportModule.vb
├── App.config
├── Project.vbproj
└── README.md
```

## 🚀 Quick Start

1. Create new repository on GitHub
2. Clone the repository locally
3. Create new Windows Forms Application in Visual Studio
4. Copy all files from this repository
5. Add references for Office COM objects
6. Build and Run

## 🔐 Default Credentials

- **Username**: admin
- **Password**: admin

## 📚 Database Tables

### Utilisateurs (Users)
| Field | Type | Description |
|-------|------|-------------|
| ID | AutoNumber | Primary Key |
| Login | Text | Username |
| MotDePasse | Text | Password |

### Etudiants (Students)
| Field | Type | Description |
|-------|------|-------------|
| ID_Etudiant | AutoNumber | Primary Key |
| Nom | Text | Last Name |
| Prenom | Text | First Name |
| DateNaissance | Date | Date of Birth |
| Sexe | Text | Gender (M/F) |
| Filiere | Text | Field of Study |
| Telephone | Text | Phone Number |
| Adresse | Text | Address |

## ✨ Core Features

### 🔐 Login System
- Username/Password authentication
- Secure credential verification
- Error messages for invalid login
- Logout functionality

### 👥 Student Management
- **Add**: Create new student records with validation
- **Edit**: Modify existing records
- **Delete**: Remove student records with confirmation
- **Search**: Find by ID or Name (real-time search)
- **View**: Display all students in DataGridView
- **Refresh**: Reload data from database

### 📊 Reports
- Generate professional Word documents
- Include all students in formatted tables
- Automatic date and title
- Ready-to-print format

## 🛠️ Installation Steps

See `SETUP_INSTRUCTIONS.md` for detailed step-by-step guide.

## 📝 License

Free to use and modify for educational purposes.

## 👤 Author

Created by: gostaviaaa

## 🔗 Repository

https://github.com/gostaviaaa/StudentManagementSystem

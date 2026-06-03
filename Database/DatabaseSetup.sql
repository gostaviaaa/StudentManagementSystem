# DATABASE SETUP - SQL SCRIPTS
## Student Management System

---

## SQL QUERIES FOR DATABASE CREATION

Use these queries if you're creating tables in Microsoft Access programmatically.

### CREATE USERS TABLE (Utilisateurs)

```sql
CREATE TABLE Utilisateurs (
    ID AUTOINCREMENT PRIMARY KEY,
    Login TEXT NOT NULL,
    MotDePasse TEXT NOT NULL
);
```

### CREATE STUDENTS TABLE (Etudiants)

```sql
CREATE TABLE Etudiants (
    ID_Etudiant AUTOINCREMENT PRIMARY KEY,
    Nom TEXT NOT NULL,
    Prenom TEXT NOT NULL,
    DateNaissance DATETIME,
    Sexe TEXT,
    Filiere TEXT NOT NULL,
    Telephone TEXT,
    Adresse TEXT
);
```

---

## INSERT SAMPLE DATA

### Add Users

```sql
-- Default admin user
INSERT INTO Utilisateurs (Login, MotDePasse) VALUES ('admin', 'admin');

-- Additional users (optional)
INSERT INTO Utilisateurs (Login, MotDePasse) VALUES ('teacher1', 'pass123');
INSERT INTO Utilisateurs (Login, MotDePasse) VALUES ('teacher2', 'pass456');
```

### Add Sample Students

```sql
-- Student 1
INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Dupont', 'Jean', #1/15/2005#, 'M', 'Computer Science', '0612345678', '123 Rue de Paris, Paris');

-- Student 2
INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Martin', 'Marie', #5/20/2004#, 'F', 'Business Administration', '0698765432', '456 Rue de Lyon, Lyon');

-- Student 3
INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Bernard', 'Pierre', #8/10/2006#, 'M', 'Engineering', '0723456789', '789 Rue de Marseille, Marseille');

-- Student 4
INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Leblanc', 'Sophie', #3/25/2005#, 'F', 'Business Administration', '0687654321', '321 Rue de Toulouse, Toulouse');

-- Student 5
INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Moreau', 'Luc', #7/30/2004#, 'M', 'Computer Science', '0676543210', '654 Rue de Nice, Nice');
```

---

## VERIFICATION QUERIES

### Check Tables Exist

```sql
SELECT * FROM Utilisateurs;
SELECT * FROM Etudiants;
```

### Count Records

```sql
SELECT COUNT(*) AS UserCount FROM Utilisateurs;
SELECT COUNT(*) AS StudentCount FROM Etudiants;
```

### View All Students

```sql
SELECT * FROM Etudiants ORDER BY Nom, Prenom;
```

### Search by Field

```sql
-- Find student by name
SELECT * FROM Etudiants WHERE Nom LIKE '%Dupont%';

-- Find students by field of study
SELECT * FROM Etudiants WHERE Filiere = 'Computer Science';

-- Find students by gender
SELECT * FROM Etudiants WHERE Sexe = 'F';
```

---

## DATA MAINTENANCE QUERIES

### Delete All Students (with confirmation)

```sql
DELETE FROM Etudiants;
```

### Delete Specific Student

```sql
DELETE FROM Etudiants WHERE ID_Etudiant = 1;
```

### Update Student Information

```sql
UPDATE Etudiants 
SET Telephone = '0611111111', Adresse = 'New Address' 
WHERE ID_Etudiant = 1;
```

### Clear All Data (Reset Database)

```sql
DELETE FROM Etudiants;
DELETE FROM Utilisateurs;
```

### Reset AutoNumber

```sql
-- For MS Access, use the GUI or delete and recreate table
```

---

## DATABASE INFO

### Access Data Types Reference

| VB.NET Type | Access Type | Size | Example |
|-------------|------------|------|---------|
| String | Text | Up to 255 chars | "John" |
| Date | Date/Time | 8 bytes | #1/15/2005# |
| Integer | Number | 4 bytes | 25 |
| Decimal | Currency | 8 bytes | 1500.50 |
| Boolean | Yes/No | 1 byte | Yes/No |

### AutoNumber vs AUTOINCREMENT

- **AutoNumber**: GUI mode (default)
- **AUTOINCREMENT**: SQL mode
- Both create automatic incrementing IDs

### Date Format in Access SQL

- Format: `#MM/DD/YYYY#`
- Example: `#1/15/2005#` = January 15, 2005
- Always use # delimiters

---

## CONNECTION STRINGS

### For Access .accdb Files (2007+)

```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\To\StudentDB.accdb
```

### For Access .mdb Files (2003)

```
Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Path\To\StudentDB.mdb
```

### With Relative Path

```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\StudentDB.accdb
```

---

## BACKUP & RESTORE

### Backup Database

1. Locate `StudentDB.accdb` file
2. Copy to backup location
3. Name: `StudentDB_Backup_2026-06-03.accdb`

### Restore from Backup

1. Delete corrupted database
2. Copy backup file to original location
3. Rename if needed

### Compact Database (Reduces File Size)

In MS Access:
1. Open database
2. File → Info
3. Click "Compact & Repair Database"

---

## COMMON ACCESS ISSUES & SOLUTIONS

| Issue | Cause | Solution |
|-------|-------|----------|
| "Type Mismatch" | Wrong data type | Check field types match |
| "Syntax Error" | SQL syntax wrong | Review SQL statement |
| "Cannot add record" | Constraint violation | Check NOT NULL fields |
| "File locked" | Multiple connections | Close all connections |

---

## EXAMPLE: COMPLETE INITIAL SETUP

```sql
-- Drop existing tables (optional - CAUTION!)
-- DROP TABLE Etudiants;
-- DROP TABLE Utilisateurs;

-- Create Users Table
CREATE TABLE Utilisateurs (
    ID AUTOINCREMENT PRIMARY KEY,
    Login TEXT NOT NULL UNIQUE,
    MotDePasse TEXT NOT NULL
);

-- Create Students Table
CREATE TABLE Etudiants (
    ID_Etudiant AUTOINCREMENT PRIMARY KEY,
    Nom TEXT NOT NULL,
    Prenom TEXT NOT NULL,
    DateNaissance DATETIME,
    Sexe TEXT,
    Filiere TEXT NOT NULL,
    Telephone TEXT,
    Adresse TEXT
);

-- Insert Default User
INSERT INTO Utilisateurs (Login, MotDePasse) VALUES ('admin', 'admin');

-- Insert Sample Students
INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Dupont', 'Jean', #1/15/2005#, 'M', 'Computer Science', '0612345678', '123 Rue de Paris');

INSERT INTO Etudiants (Nom, Prenom, DateNaissance, Sexe, Filiere, Telephone, Adresse)
VALUES ('Martin', 'Marie', #5/20/2004#, 'F', 'Business', '0698765432', '456 Rue de Lyon');
```

---

## NOTES

- Always backup before running DELETE or UPDATE queries
- Test queries on a copy first
- Document any custom changes to schema
- Use parameterized queries in production for security

---

**For more help, see SETUP_INSTRUCTIONS.md**

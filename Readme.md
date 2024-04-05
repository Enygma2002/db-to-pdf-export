1. Create a DB
sqlite3 employees.db <<EOF
CREATE TABLE your_table (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    age INTEGER,
    email TEXT
);
EOF

2. Insert entries
sqlite3 employees.db <<EOF
INSERT INTO your_table (name, age, email) VALUES
('John Doe', 30, 'john.doe@example.com'),
('Jane Smith', 25, 'jane.smith@example.com'),
('Bob Johnson', 35, 'bob.johnson@example.com');
EOF

3. Insert a "name" field in the Word template:
Insert > Insert Field > Mail Merge > Merge Field > (e.g.) "MERGEFIELD name"

3.1 Tutorial and more details on using Mail Merge: https://pbpython.com/python-word-template.html

4. Install dependencies:
pip install docx-mailmerge docx2pdf

4.1 Alternatives:
4.1.1 LibreOffice on mac: brew install --cask libreoffice
4.1.2 Aspose: pip install aspose-words

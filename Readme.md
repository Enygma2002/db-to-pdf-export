# Description

Python script that exports data from an SQL DB (e.g. sqlite), runs it through a Word template and exports each result as PDF.

# Database

You can just use the provided sample `employees.sqlite` or create your own:

1. Create a DB
```
sqlite3 employees.db <<EOF
CREATE TABLE your_table (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    age INTEGER,
    email TEXT
);
EOF
```

2. Insert entries
```
sqlite3 employees.db <<EOF
INSERT INTO your_table (name, age, email) VALUES
('John Doe', 30, 'john.doe@example.com'),
('Jane Smith', 25, 'jane.smith@example.com'),
('Bob Johnson', 35, 'bob.johnson@example.com');
EOF
```

# Template

See `template.docx` as an example.

To insert a (e.g.) `name` field in the Word template:

`Insert > Insert Field > Mail Merge > Merge Field > "MERGEFIELD name"`

Tutorial and more details on using Mail Merge: https://pbpython.com/python-word-template.html

# Install dependencies

You can use the provided `.venv` environment (for `docx2pdf` or `LibreOffice`), or create your own:

1. Python dependencies:
    1. create venv: `python3 -m venv .venv`
    2. activate venv: `source .venv/bin/activate`
    3. install dependencies: `pip install docx-mailmerge`

2. Depending on the preferred pdf export option (see commented out code):
    1. **docx2pdf** (default)
       1. Depends on MS Word which must installed locally and licensed
       2. Produces optimum output (same as viewed in Word)
       3. Should work automatically for Windows (not tested)
       4. for Mac, due to security policies, requires initial access granting and then manual grant for each processed entry, making it **unusable for large number of entries**
       5. `pip install docx2pdf`
    2. **LibreOffice** (commented out)
       1. Depends on LibreOffice which must be installed locally
       2. 100% free with good enough results (slight font differences)
       3. Runs automatically, no intervention needed. Not very fast (2-3 sec/entry).
       4. : `brew install --cask libreoffice` (Mac)
    3. **Aspose** (commented out)
       1. Paid solution ($1200 license required) in addition to MS Word on which it relies
            * Without buying, produces unusable watermarks in the output
       2. Produces optimum ouput as well (same as viewed in Word)
       3. Works on Mac automatically as well, and very fast
       4. `pip install aspose-words`

# Execution

1. Ensure you've activated the venv: `source .venv/bin/activate`
2. Run the script: `python3 db-to-pdf-export.py`

Output located in the `output` folder (created at execution):
* `output/docx` - intermediary filled-in docx output, using the Word template
* `output/pdf` - final PDF output

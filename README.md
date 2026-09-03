# inpsPyPI

A versatile Python utility library providing helpers for data structures, ciphers, data conversion, validations, SQLite management, Excel spreadsheet operations, and common sorting algorithms.

---

## Installation

```bash
pip install inpsPyPI
```

### Dependencies
- `openpyxl`
- `unidecode`

---

## Features & Modules

- **Validation (`Check`)**: Email, Philippine mobile numbers, strings, and internet connectivity checks.
- **Ciphers (`Cipher`)**: Caesar, Keyword, Giovanni, and Transposition ciphers.
- **Conversion (`Convert`)**: Base64, Hex, Binary, Byte arrays, and data type conversions.
- **Data Structures**:
  - `Dictionarily`: Custom dictionary with key-sorting capabilities.
  - `Memory`: Dynamic list/storage manager.
  - `Stackily`: Lightweight LIFO stack implementation.
- **Database (`EasySQL`)**: Simplified SQLite CRUD interface.
- **Excel Operations (`excellent_reader`)**: Read and write column data across sheets.
- **File Management (`SimpleFileHandler`)**: Fast read, write, and append operations.
- **Algorithms (`sort`)**: 16+ classic and non-comparison sorting algorithms.
- **Flask Helpers (`bake`)**: Standardized JSON response formatting.

---

## Usage Guide

### 1. Validations (`Check`)

```python
from inpsPyPI.check import Check

# Email validation (by domain name and extension)
Check.Email.add_valid_domain_name("gmail")
Check.Email.add_valid_domain_extension("com")
Check.Email.should_use_full_domain(False)
Check.Email.is_valid("user@gmail.com")  # True

# Email validation (by full domain)
Check.Email.add_valid_domain("company.org")
Check.Email.should_use_full_domain(True)
Check.Email.is_valid("user@company.org")  # True

# Philippine mobile number check
Check.is_a_valid_philippine_mobile_number("+639171234567")  # True
Check.is_a_valid_philippine_mobile_number("09171234567")     # True

# String checks
Check.is_all_numbers("12345")  # True
Check.has_numbers("abc1")      # True
Check.has_symbols("hello!")    # True
Check.has_spaces("hello world") # True

# Network connectivity check
Check.is_connected()  # Returns True if connected to internet
```

---

### 2. Ciphers (`Cipher`)

```python
from inpsPyPI.cipher import Cipher

# Caesar Cipher
Cipher.caesar_cipher("HELLO WORLD", shift=3)  # "KHOOR ZRUOG"

# Keyword Cipher
Cipher.keyword_cipher("HELLO WORLD", keyword="KEYWORD")  # "AOGGJ UJNGW"

# Giovanni Cipher
Cipher.giovanni_cipher("HELLO WORLD", keyword="KEYWORD", key_letter="C")  # "RYCCH SHLCE"

# Transposition Cipher
Cipher.transposition_cipher("HELLO WORLD")  # "HLOOLELWRD"
```

---

### 3. Conversions (`Convert`)

```python
from inpsPyPI.convert import Convert

# String transformations
Convert.reverse("python")         # "nohtyp"
Convert.to_real_name("jOhN")      # "John"

# Base64
b64 = Convert.to_base64("hello")  # "aGVsbG8="
Convert.from_base64(b64)          # "hello"

# Hexadecimal
hex_val = Convert.to_hex("hello") # "68656C6C6F"
Convert.from_hex(hex_val)         # "hello"

# Binary
bin_val = Convert.to_binary("A")  # "01000001"
Convert.from_binary(bin_val)      # "A"

# Byte Array
byte_val = Convert.to_byte_array("hello")  # b"hello"
Convert.from_byte_array(byte_val)          # "hello"

# Numeric conversions
Convert.to_int("42")      # 42
Convert.to_float("3.14")  # 3.14
```

---

### 4. Custom Dictionary (`Dictionarily`)

```python
from inpsPyPI import Dictionarily

d = Dictionarily()
d.add("b", 2)
d.add("a", 1)
d.add(10, "ten")
d.add(2, "two")

# Sort alphabetically by stringified key
d.sort()

# Sort numbers before string keys
d.sort_numbers_first()
print(d.show())  # {2: 'two', 10: 'ten', 'a': 1, 'b': 2}
```

---

### 5. SQLite Helper (`EasySQL`)

```python
from inpsPyPI import EasySQL

# Connects to 'my_database.db'
sql = EasySQL("my_database")

# Create a table
sql.create_table("users", {"id": "INTEGER PRIMARY KEY", "name": "TEXT"})

# Insert records
sql.insert_to_table("users", {"id": 1, "name": "Alice"})
sql.insert_to_table("users", {"id": 2, "name": "Bob"})

# Fetch records
rows = sql.get_table_values("users")  # [(1, 'Alice'), (2, 'Bob')]

# Print table contents to stdout
sql.print_table("users")

# Delete specific row
sql.delete_from_table("users", {"id": 1})

# Clear all rows or drop table
sql.clear_table("users")
sql.delete_table("users")
```

---

### 6. Excel Reader & Writer (`excellent_reader`)

```python
from inpsPyPI import excellent_reader

# Read columns (skipping header rows, default=2)
sheet_data = excellent_reader.get_n_column_from_sheet_index("data.xlsx", index=0, column="A", skip_rows=2)
all_sheets_data = excellent_reader.get_first_column_from_all_sheets("data.xlsx", skip_rows=2)

# Write column values to a specific sheet
excellent_reader.set_n_column_from_sheet_index(
    "data.xlsx", 
    index=0, 
    column="B", 
    value=["Val1", "Val2"], 
    skip_rows=2
)

# Write single value to all sheets
excellent_reader.set_first_column_from_all_sheets("data.xlsx", value="Default", skip_rows=2)
```

---

### 7. Data Storage & Stacks (`Memory`, `Stackily`)

```python
from inpsPyPI import Memory, Stackily

# Memory (List wrapper)
mem = Memory()
mem.add("item1")
mem.add("item2")
mem.contains("item1")  # True
mem.get(0)             # "item1"
mem.remove_at(0)
mem.count()            # 1
mem.clear()

# Stackily (LIFO Stack)
stack = Stackily()
stack.push("a")
stack.push("b")
stack.peek()      # "b"
stack.pop()
stack.size()      # 1
stack.is_empty()  # False
stack.to_list()   # ['a']
```

---

### 8. Sorting Algorithms (`sort`)

```python
from inpsPyPI import sort

data = [5, 2, 9, 1, 5, 6]

sort.bubble_sort(list(data))
sort.quicksort(list(data))
sort.merge_sort(list(data))
sort.heapsort(list(data))
sort.insertion_sort(list(data))
sort.selection_sort(list(data))
sort.shellsort(list(data))
sort.timsort(list(data))
sort.cocktail_shaker_sort(list(data))
sort.odd_even_sort(list(data))
sort.introsort(list(data))
sort.counting_sort(list(data))
sort.pigeonhole_sort(list(data))
sort.patience_sorting(list(data))
sort.bead_sort(list(data))
sort.bucket_sort_uniform([0.5, 0.1, 0.9, 0.2])
```

---

### 9. File Operations & Response Baking

```python
from inpsPyPI.simple_file_handler import SimpleFileHandler
from inpsPyPI import bake

# Simple File Handler
SimpleFileHandler.write("output.txt", "Initial content")
SimpleFileHandler.append("output.txt", "\nAppended line")
content = SimpleFileHandler.read("output.txt")

# Flask Jsonify Helper
response = bake({"status": "success"})  # {'response_data': {'status': 'success'}}
```

---

## Running Unit Tests

To run all unit tests:

```bash
python -m unittest test_all.py
```

To run unit tests in buffer mode (suppress print outputs during testing):

```bash
python -m unittest -b test_all.py
```
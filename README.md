# inpsPyPI

`inpsPyPI` is a multipurpose Python utility library providing database wrappers, Excel automation tools, custom data structures, classical cryptography ciphers, sorting algorithms, type converters, and validation helpers.

---

## Installation

```bash
pip install inpsPyPI
```

---

## Table of Contents
- [Data Structures & Collections](#data-structures--collections)
  - [Dictionarily](#dictionarily)
  - [Memory](#memory)
  - [Stackily](#stackily)
- [Database & File I/O](#database--file-io)
  - [EasySQL](#easysql)
  - [Excellent Reader (Excel Helper)](#excellent-reader-excel-helper)
  - [SimpleFileHandler](#simplefilehandler)
- [Algorithms & Cryptography](#algorithms--cryptography)
  - [Sort](#sort)
  - [Cipher](#cipher)
- [Validation & Utilities](#validation--utilities)
  - [Check](#check)
  - [Convert](#convert)
  - [Flask Jsonify Helper (`bake`)](#flask-jsonify-helper-bake)
- [Running Unit Tests](#running-unit-tests)

---

## Data Structures & Collections

### Dictionarily

A wrapper around standard Python dictionaries with custom sorting capabilities.

```python
from inpsPyPI import Dictionarily

d = Dictionarily()
d.add("b", 2)
d.add("a", 1)
d.add(2, "two")
d.add(1, "one")

# Sort lexicographically by key/value representation
d.sort()

# Sort with numerical keys preceding string keys
d.sort_numbers_first()

# Retrieve internal dict
print(d.show())  # {1: 'one', 2: 'two', 'a': 1, 'b': 2}
```

---

### Memory

A dynamic in-memory list storage container.

```python
from inpsPyPI import Memory

mem = Memory()
mem.add("data1")
mem.add("data2")

count = mem.count()            # 2
exists = mem.contains("data1")  # True
item = mem.get(0)              # "data1"

mem.remove("data1")            # Removes first matching element
mem.remove_at(0)               # Removes element by index
mem.clear()                    # Resets internal storage
```

---

### Stackily

A LIFO (Last-In, First-Out) stack implementation.

```python
from inpsPyPI import Stackily

stack = Stackily()
stack.push("first")
stack.push("second")

top_item = stack.peek()        # "second" (does not remove)
stack.pop()                    # Removes "second"
size = stack.size()            # 1
empty = stack.is_empty()       # False
items_list = stack.to_list()   # ["first"]
```

---

## Database & File I/O

### EasySQL

A lightweight SQLite helper that abstracts SQL boilerplate with dictionaries.

```python
from inpsPyPI import EasySQL

# Connects to 'app_database.db'
sql = EasySQL("app_database")

# Create Table
sql.create_table("users", {
    "id": "INTEGER PRIMARY KEY",
    "name": "TEXT",
    "email": "TEXT"
})

# Insert Records
sql.insert_to_table("users", {"id": 1, "name": "Alice", "email": "alice@example.com"})
sql.insert_to_table("users", {"id": 2, "name": "Bob", "email": "bob@example.com"})

# Query Records
rows = sql.get_table_values("users")
# [(1, 'Alice', 'alice@example.com'), (2, 'Bob', 'bob@example.com')]

# Print Table Content Directly
sql.print_table("users")

# Delete Records Matching Conditions
sql.delete_from_table("users", {"id": 1})

# Table Management
sql.clear_table("users")   # DELETE FROM users
sql.delete_table("users")  # DROP TABLE IF EXISTS users
```

---

### Excellent Reader (Excel Helper)

Utilities for reading and writing column data in `.xlsx` files using `openpyxl`. Default `skip_rows=2` skips headers.

```python
from inpsPyPI import excellent_reader

file = "data.xlsx"

# --- Reading Data ---
# Read column A from the first worksheet (index 0)
first_col = excellent_reader.get_first_column_from_sheet_index(file, index=0, skip_rows=2)

# Read specific column ('B' or 2) from sheet at index 0
col_b = excellent_reader.get_n_column_from_sheet_index(file, index=0, column="B", skip_rows=2)

# Read column A from all sheets
all_first_cols = excellent_reader.get_first_column_from_all_sheets(file, skip_rows=2)

# Read column 'C' from all sheets
all_c_cols = excellent_reader.get_n_column_from_all_sheets(file, column="C", skip_rows=2)


# --- Writing Data ---
# Write single value or list of values to column A in sheet at index 0
excellent_reader.set_first_column_from_sheet_index(file, index=0, value=["Val1", "Val2"], skip_rows=2)

# Write to specific column
excellent_reader.set_n_column_from_sheet_index(file, index=0, column="B", value="StaticValue", skip_rows=2)

# Write to all sheets
excellent_reader.set_first_column_from_all_sheets(file, value="Updated", skip_rows=2)
excellent_reader.set_n_column_from_all_sheets(file, column="D", value=["Row1", "Row2"], skip_rows=2)
```

---

### SimpleFileHandler

Static file I/O operations with default UTF-8 encoding.

```python
from inpsPyPI.simple_file_handler import SimpleFileHandler

# Write (overwrites existing content)
SimpleFileHandler.write("sample.txt", "Hello World")

# Append
SimpleFileHandler.append("sample.txt", "\nNext line")

# Read
content = SimpleFileHandler.read("sample.txt")
```

---

## Algorithms & Cryptography

### Sort

A collection of sorting algorithms.

| Function | Description | Supported Input |
| :--- | :--- | :--- |
| `bubble_sort(arr)` | Bubble sort | Generic list |
| `cocktail_shaker_sort(arr)` | Bidirectional bubble sort | Generic list |
| `odd_even_sort(arr)` | Odd-even transposition sort | Generic list |
| `selection_sort(arr)` | Selection sort | Generic list |
| `insertion_sort(arr)` | Insertion sort | Generic list |
| `shellsort(arr)` | Shell sort | Generic list |
| `quicksort(arr)` | Quicksort (pivot at mid) | Generic list |
| `merge_sort(arr)` | Merge sort | Generic list |
| `heapsort(arr)` | Heapsort via `heapq` | Generic list |
| `introsort(arr)` | Hybrid introspective sort | Generic list |
| `timsort(arr)` | Standard Python Timsort | Generic list |
| `counting_sort(arr)` | Counting sort | Non-negative integers |
| `bucket_sort_uniform(arr)` | Uniform bucket sort | Floats in range $[0.0, 1.0)$ |
| `pigeonhole_sort(arr)` | Pigeonhole sort | Integers |
| `patience_sorting(arr)` | Patience sorting via piles | Generic list |
| `bogosort(arr)` | Randomized bogo sort | Generic list |
| `bead_sort(arr)` | Gravity / Bead sort | Non-negative integers |

```python
from inpsPyPI import sort

data = [5, 2, 9, 1, 5, 6]
sorted_data = sort.quicksort(data)

uniform_floats = [0.89, 0.56, 0.65, 0.12, 0.66]
sorted_floats = sort.bucket_sort_uniform(uniform_floats)
```

---

### Cipher

Classical text ciphers for uppercase alphabetic characters (spaces and non-alphabetic characters are preserved).

```python
from inpsPyPI.cipher import Cipher

# Caesar Cipher
Cipher.caesar_cipher("HELLO WORLD", 3)
# -> 'KHOOR ZRUOG'

# Keyword Cipher
Cipher.keyword_cipher("HELLO WORLD", "KEYWORD")
# -> 'AOGGJ UJNGW'

# Giovanni Cipher (Keyword cipher with rotational offset)
Cipher.giovanni_cipher("HELLO WORLD", "KEYWORD", "C")
# -> 'RYCCH SHLCE'

# Transposition Cipher (Interleaves even/odd indices, removes spaces)
Cipher.transposition_cipher("HELLO WORLD")
# -> 'HLOOLELWRD'
```

---

## Validation & Utilities

### Check

Input validation and network connectivity checkers.

```python
from inpsPyPI.check import Check

# --- Email Validation ---
# Split Domain Mode (matches name and extension separately)
Check.Email.add_valid_domain_name("gmail")
Check.Email.add_valid_domain_extension("com")
Check.Email.should_use_full_domain(False)
Check.Email.is_valid("user@gmail.com")  # True

# Full Domain Mode
Check.Email.add_valid_domain("company.co.uk")
Check.Email.should_use_full_domain(True)
Check.Email.is_valid("user@company.co.uk")  # True

# --- String Analysis ---
Check.is_a_valid_philippine_mobile_number("+639171234567")  # True
Check.is_a_valid_philippine_mobile_number("09171234567")     # True
Check.is_all_numbers("123456")                              # True
Check.has_numbers("User123")                                # True
Check.has_symbols("hello@world")                            # True
Check.has_spaces("Hello World")                             # True

# --- Network Status ---
has_internet = Check.is_connected()  # True if Google DNS (8.8.8.8:53) is reachable
```

---

### Convert

Type conversions and encoding transformations.

```python
from inpsPyPI.convert import Convert

# String & Format Conversions
Convert.reverse("python")         # "nohtyp"
Convert.to_real_name("jOHN")      # "John"

# Base64
b64_str = Convert.to_base64("hello")     # "aGVsbG8="
orig_str = Convert.from_base64(b64_str)  # "hello"

# Hexadecimal
hex_str = Convert.to_hex("hello")        # "68656C6C6F"
from_hex = Convert.from_hex(hex_str)     # "hello"

# Binary
bin_str = Convert.to_binary("hello")     # "0110100001100101011011000110110001101111"
from_bin = Convert.from_binary(bin_str)  # "hello"

# Byte Array
byte_arr = Convert.to_byte_array("hello")  # b'hello'
from_byte = Convert.from_byte_array(byte_arr)  # "hello"

# Numeric Parsing
int_val = Convert.to_int("100")       # 100
float_val = Convert.to_float("3.14")  # 3.14
double_val = Convert.to_double("3.14") # 3.14
long_val = Convert.to_long("100000")  # 100000
```

---

### Flask Jsonify Helper (`bake`)

Wraps any payload into a standardized dictionary response.

```python
from inpsPyPI import bake

bake("Success")
# Output: {'response_data': 'Success'}

bake({"status": 200, "user": "admin"})
# Output: {'response_data': {'status': 200, 'user': 'admin'}}
```

---

## Running Unit Tests

Run all tests via Python's standard `unittest` framework:

```bash
# Standard test run
python -m unittest test_all.py

# Buffer mode (suppresses print statements during tests)
python -m unittest -b test_all.py
```
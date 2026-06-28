import unittest
import os
import sqlite3
import io
from contextlib import redirect_stdout

# External dependency required by excellent_reader.py
import openpyxl

# Import the modules and components under test
from builder_for_flask_jsonify import bake, JSON_RESPONSE_TITLE
from check import Check
from cipher import Cipher
from convert import Convert
from dictionarily import Dictionarily
from easy_sql import EasySQL
from excellent_reader import (
    get_n_column_from_sheet_index,
    get_first_column_from_sheet_index,
    get_n_column_from_all_sheets,
    get_first_column_from_all_sheets,
    set_n_column_from_sheet_index,
    set_first_column_from_sheet_index,
    set_n_column_from_all_sheets,
    set_first_column_from_all_sheets,
)
from memory import Memory
from simple_file_handler import SimpleFileHandler
from sort import (
    bubble_sort,
    cocktail_shaker_sort,
    odd_even_sort,
    selection_sort,
    insertion_sort,
    shellsort,
    quicksort,
    merge_sort,
    heapsort,
    introsort,
    timsort,
    counting_sort,
    bucket_sort_uniform,
    pigeonhole_sort,
    patience_sorting,
    bogosort,
    bead_sort,
)
from stackily import Stackily
from test import test_method


class TestBuilderForFlaskJsonify(unittest.TestCase):
    def test_bake_string(self):
        result = bake("Hello World")
        self.assertEqual(result, {JSON_RESPONSE_TITLE: "Hello World"})

    def test_bake_dict(self):
        data = {"status": "success", "code": 200}
        result = bake(data)
        self.assertEqual(result, {JSON_RESPONSE_TITLE: data})

    def test_bake_integer(self):
        result = bake(12345)
        self.assertEqual(result, {JSON_RESPONSE_TITLE: 12345})


class TestCheck(unittest.TestCase):
    def setUp(self):
        # Reset Check.Email class variables to prevent test contamination
        Check.Email.valid_domain_names = []
        Check.Email.valid_domain_extensions = []
        Check.Email.valid_domains = []
        Check.Email._should_use_full_domain = False

    def test_email_partial_domain_validation(self):
        Check.Email.add_valid_domain_name("gmail")
        Check.Email.add_valid_domain_extension("com")
        Check.Email.should_use_full_domain(False)

        self.assertTrue(Check.Email.is_valid("user@gmail.com"))
        self.assertFalse(Check.Email.is_valid("user@yahoo.com"))
        self.assertFalse(Check.Email.is_valid("user@gmail"))  # Missing extension
        self.assertFalse(Check.Email.is_valid("invalid_email"))  # No @ symbol

    def test_email_full_domain_validation(self):
        Check.Email.add_valid_domain("yahoo.com")
        Check.Email.should_use_full_domain(True)

        self.assertTrue(Check.Email.is_valid("user@yahoo.com"))
        self.assertFalse(Check.Email.is_valid("user@gmail.com"))
        self.assertFalse(Check.Email.is_valid("user@yahoo.org"))
        self.assertFalse(Check.Email.is_valid("invalid_email"))

    def test_is_a_valid_philippine_mobile_number(self):
        # Valid patterns
        self.assertTrue(Check.is_a_valid_philippine_mobile_number("09171234567"))
        self.assertTrue(Check.is_a_valid_philippine_mobile_number("+639171234567"))
        self.assertTrue(Check.is_a_valid_philippine_mobile_number("639171234567"))
        self.assertTrue(Check.is_a_valid_philippine_mobile_number(" (0917) 123-4567 "))

        # Invalid patterns
        self.assertFalse(Check.is_a_valid_philippine_mobile_number("08171234567"))  # Wrong prefix
        self.assertFalse(Check.is_a_valid_philippine_mobile_number("091712345"))  # Too short
        self.assertFalse(Check.is_a_valid_philippine_mobile_number("091712345678"))  # Too long
        self.assertFalse(Check.is_a_valid_philippine_mobile_number("0917abc4567"))  # Non-digits

    def test_is_all_numbers(self):
        self.assertTrue(Check.is_all_numbers("123456"))
        self.assertFalse(Check.is_all_numbers("123a45"))
        # Python's all() returns True for empty sequences
        self.assertTrue(Check.is_all_numbers(""))

    def test_has_numbers(self):
        self.assertTrue(Check.has_numbers("abc1"))
        self.assertFalse(Check.has_numbers("abc"))
        self.assertFalse(Check.has_numbers(""))

    def test_has_symbols(self):
        self.assertTrue(Check.has_symbols("hello!"))
        self.assertFalse(Check.has_symbols("hello"))
        self.assertFalse(Check.has_symbols(""))

    def test_has_spaces(self):
        self.assertTrue(Check.has_spaces("hello world"))
        self.assertFalse(Check.has_spaces("helloworld"))
        self.assertFalse(Check.has_spaces(""))


class TestCipher(unittest.TestCase):
    def test_transposition_cipher(self):
        self.assertEqual(Cipher.transposition_cipher("HELLO WORLD"), "HLOOLELWRD")
        self.assertEqual(Cipher.transposition_cipher("1234"), "1324")

    def test_giovanni_cipher(self):
        self.assertEqual(Cipher.giovanni_cipher("HELLO WORLD", "KEYWORD", "C"), "RYCCH SHLCE")
        self.assertEqual(Cipher.giovanni_cipher("HELLO, WORLD!", "KEYWORD", "C"), "RYCCH, SHLCE!")

    def test_keyword_cipher(self):
        self.assertEqual(Cipher.keyword_cipher("HELLO WORLD", "KEYWORD"), "AOGGJ UJNGW")
        self.assertEqual(Cipher.keyword_cipher("HELLO, WORLD!", "KEYWORD"), "AOGGJ, UJNGW!")

    def test_caesar_cipher(self):
        self.assertEqual(Cipher.caesar_cipher("HELLO WORLD", 3), "KHOOR ZRUOG")
        self.assertEqual(Cipher.caesar_cipher("HELLO, WORLD!", 3), "KHOOR, ZRUOG!")


class TestConvert(unittest.TestCase):
    def test_reverse(self):
        self.assertEqual(Convert.reverse("python"), "nohtyp")
        self.assertEqual(Convert.reverse(""), "")

    def test_base64_conversions(self):
        encoded = Convert.to_base64("hello")
        self.assertEqual(encoded, "aGVsbG8=")
        decoded = Convert.from_base64("aGVsbG8=")
        self.assertEqual(decoded, "hello")

    def test_byte_array_conversions(self):
        self.assertEqual(Convert.to_byte_array("hello"), b"hello")
        self.assertEqual(Convert.from_byte_array(b"hello"), "hello")

    def test_hex_conversions(self):
        encoded = Convert.to_hex("hello")
        self.assertEqual(encoded, "68656C6C6F")
        self.assertEqual(Convert.from_hex("68656C6C6F"), "hello")

    def test_binary_conversions(self):
        encoded = Convert.to_binary("hello")
        self.assertEqual(encoded, "0110100001100101011011000110110001101111")
        self.assertEqual(Convert.from_binary("0110100001100101011011000110110001101111"), "hello")

    def test_numeric_conversions(self):
        self.assertEqual(Convert.to_int("100"), 100)
        self.assertEqual(Convert.to_double("3.14"), 3.14)
        self.assertEqual(Convert.to_long("100000"), 100000)
        self.assertEqual(Convert.to_float("3.14"), 3.14)


class TestDictionarily(unittest.TestCase):
    def test_add_and_show(self):
        d = Dictionarily()
        d.add("key1", "value1")
        self.assertEqual(d.show(), {"key1": "value1"})

    def test_sort(self):
        d = Dictionarily()
        d.add("b", 2)
        d.add("a", 1)
        d.sort()
        self.assertEqual(list(d.show().keys()), ["a", "b"])

    def test_sort_numbers_first(self):
        d = Dictionarily()
        d.add("b", 2)
        d.add("a", 1)
        d.add(2, "two")
        d.add(1, "one")
        d.sort_numbers_first()
        self.assertEqual(list(d.show().keys()), [1, 2, "a", "b"])


class TestEasySQL(unittest.TestCase):
    def setUp(self):
        self.db_name = "test_database"
        self.db_filename = f"{self.db_name}.db"
        self.sql = EasySQL(self.db_name)

    def tearDown(self):
        if os.path.exists(self.db_filename):
            try:
                os.remove(self.db_filename)
            except PermissionError:
                pass

    def test_database_crud_operations(self):
        # Create table
        self.sql.create_table("users", {"id": "INTEGER PRIMARY KEY", "name": "TEXT"})

        # Insert records
        self.sql.insert_to_table("users", {"id": 1, "name": "Alice"})
        self.sql.insert_to_table("users", {"id": 2, "name": "Bob"})

        # Get records
        rows = self.sql.get_table_values("users")
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0], (1, "Alice"))
        self.assertEqual(rows[1], (2, "Bob"))

        # Delete record with matching condition
        self.sql.delete_from_table("users", {"id": 1})
        rows = self.sql.get_table_values("users")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0], (2, "Bob"))

        # Clear records
        self.sql.clear_table("users")
        rows = self.sql.get_table_values("users")
        self.assertEqual(len(rows), 0)

        # Delete Table
        self.sql.delete_table("users")
        with self.assertRaises(sqlite3.OperationalError):
            self.sql.get_table_values("users")

    def test_print_table(self):
        self.sql.create_table("users", {"id": "INTEGER PRIMARY KEY", "name": "TEXT"})

        # Capture output on empty table
        f_empty = io.StringIO()
        with redirect_stdout(f_empty):
            self.sql.print_table("users")
        self.assertIn("Table 'users' is empty.", f_empty.getvalue())

        # Capture output on populated table
        self.sql.insert_to_table("users", {"id": 1, "name": "Alice"})
        f_populated = io.StringIO()
        with redirect_stdout(f_populated):
            self.sql.print_table("users")
        self.assertIn("(1, 'Alice')", f_populated.getvalue())


class TestExcellentReader(unittest.TestCase):
    def setUp(self):
        self.filename = "test_sheet.xlsx"
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "SheetA"
        ws2 = wb.create_sheet("SheetB")

        # Set up values to meet the skip_rows threshold (default is 2)
        ws1.append(["Header1"])
        ws1.append(["Header2"])
        ws1.append(["ValA3"])
        ws1.append(["ValA4"])

        ws2.append(["HeaderX"])
        ws2.append(["HeaderY"])
        ws2.append(["ValB3"])
        ws2.append(["ValB4"])

        wb.save(self.filename)

    def tearDown(self):
        if os.path.exists(self.filename):
            try:
                os.remove(self.filename)
            except PermissionError:
                pass

    def test_get_first_column_from_sheet_index(self):
        data = get_first_column_from_sheet_index(self.filename, 0, skip_rows=2)
        self.assertEqual(data, {"SHEETA": ["VALA3", "VALA4"]})

    def test_get_first_column_from_all_sheets(self):
        data = get_first_column_from_all_sheets(self.filename, skip_rows=2)
        self.assertEqual(data, {
            "SHEETA": ["VALA3", "VALA4"],
            "SHEETB": ["VALB3", "VALB4"]
        })

    def test_get_n_column_from_sheet_index(self):
        data = get_n_column_from_sheet_index(self.filename, 0, 'A', skip_rows=2)
        self.assertEqual(data, {"SHEETA": ["VALA3", "VALA4"]})

    def test_set_n_column_from_sheet_index_list(self):
        success = set_n_column_from_sheet_index(self.filename, 0, 'B', ["NewB3", "NewB4"], skip_rows=2)
        self.assertTrue(success)

        data = get_n_column_from_sheet_index(self.filename, 0, 'B', skip_rows=2)
        self.assertEqual(data, {"SHEETA": ["NEWB3", "NEWB4"]})

    def test_set_first_column_from_sheet_index_single_value(self):
        success = set_first_column_from_sheet_index(self.filename, 0, "Single", skip_rows=2)
        self.assertTrue(success)

        data = get_first_column_from_sheet_index(self.filename, 0, skip_rows=2)
        self.assertEqual(data, {"SHEETA": ["SINGLE", "SINGLE"]})

    def test_set_first_column_from_all_sheets(self):
        success = set_first_column_from_all_sheets(self.filename, "AllSheets", skip_rows=2)
        self.assertTrue(success)

        data = get_first_column_from_all_sheets(self.filename, skip_rows=2)
        self.assertEqual(data, {
            "SHEETA": ["ALLSHEETS", "ALLSHEETS"],
            "SHEETB": ["ALLSHEETS", "ALLSHEETS"]
        })

    def test_file_not_found(self):
        self.assertIsNone(get_first_column_from_sheet_index("nonexistent.xlsx", 0))
        self.assertFalse(set_first_column_from_sheet_index("nonexistent.xlsx", 0, "Val"))


class TestMemory(unittest.TestCase):
    def setUp(self):
        self.memory = Memory()

    def test_memory_lifecycle(self):
        self.assertEqual(self.memory.count(), 0)

        # Add
        self.memory.add("data1")
        self.memory.add("data2")
        self.assertEqual(self.memory.count(), 2)

        # Contains
        self.assertTrue(self.memory.contains("data1"))
        self.assertFalse(self.memory.contains("data3"))

        # Get
        self.assertEqual(self.memory.get(0), "data1")
        self.assertEqual(self.memory.get(1), "data2")

        # Remove element
        self.memory.remove("data1")
        self.assertEqual(self.memory.count(), 1)
        self.assertFalse(self.memory.contains("data1"))

        # Remove at index
        self.memory.add("data3")
        self.memory.remove_at(0)  # Removes "data2"
        self.assertEqual(self.memory.get(0), "data3")

        # Clear
        self.memory.clear()
        self.assertEqual(self.memory.count(), 0)


class TestSimpleFileHandler(unittest.TestCase):
    def setUp(self):
        self.filepath = "temp_text_file.txt"

    def tearDown(self):
        if os.path.exists(self.filepath):
            try:
                os.remove(self.filepath)
            except PermissionError:
                pass

    def test_file_operations(self):
        # Write
        SimpleFileHandler.write(self.filepath, "hello")
        self.assertTrue(os.path.exists(self.filepath))

        # Read
        content = SimpleFileHandler.read(self.filepath)
        self.assertEqual(content, "hello")

        # Append
        SimpleFileHandler.append(self.filepath, " world")
        content = SimpleFileHandler.read(self.filepath)
        self.assertEqual(content, "hello world")


class TestSort(unittest.TestCase):
    def test_sorting_algorithms(self):
        arr = [5, 2, 9, 1, 5, 6]
        expected = sorted(arr)

        self.assertEqual(bubble_sort(list(arr)), expected)
        self.assertEqual(cocktail_shaker_sort(list(arr)), expected)
        self.assertEqual(odd_even_sort(list(arr)), expected)
        self.assertEqual(selection_sort(list(arr)), expected)
        self.assertEqual(insertion_sort(list(arr)), expected)
        self.assertEqual(shellsort(list(arr)), expected)
        self.assertEqual(quicksort(list(arr)), expected)
        self.assertEqual(merge_sort(list(arr)), expected)
        self.assertEqual(heapsort(list(arr)), expected)
        self.assertEqual(introsort(list(arr)), expected)
        self.assertEqual(timsort(list(arr)), expected)
        self.assertEqual(pigeonhole_sort(list(arr)), expected)
        self.assertEqual(patience_sorting(list(arr)), expected)

    def test_counting_sort(self):
        arr = [5, 2, 9, 1, 5, 6]
        expected = sorted(arr)
        self.assertEqual(counting_sort(list(arr)), expected)
        self.assertEqual(counting_sort([]), [])

    def test_bucket_sort_uniform(self):
        # Bucket sort uniform expects values in [0, 1)
        arr = [0.5, 0.1, 0.9, 0.2, 0.7]
        expected = sorted(arr)
        self.assertEqual(bucket_sort_uniform(arr), expected)

    def test_bogosort(self):
        # Using a tiny/already-sorted list because bogosort is non-deterministic and can run indefinitely
        self.assertEqual(bogosort([2, 1]), [1, 2])

    def test_bead_sort(self):
        arr = [5, 2, 9, 1, 5, 6]
        expected = sorted(arr)
        self.assertEqual(bead_sort(list(arr)), expected)

        # Bead sort expects positive integers; raises ValueError on negative inputs
        with self.assertRaises(ValueError):
            bead_sort([-1, 2])


class TestStackily(unittest.TestCase):
    def setUp(self):
        self.stack = Stackily()

    def test_stack_operations(self):
        self.assertTrue(self.stack.is_empty())
        self.assertEqual(self.stack.size(), 0)

        # Push
        self.stack.push("item1")
        self.stack.push("item2")
        self.assertFalse(self.stack.is_empty())
        self.assertEqual(self.stack.size(), 2)

        # Peek
        self.assertEqual(self.stack.peek(), "item2")

        # Pop
        self.stack.pop()
        self.assertEqual(self.stack.size(), 1)
        self.assertEqual(self.stack.peek(), "item1")

        # To list
        self.assertEqual(self.stack.to_list(), ["item1"])


class TestTest(unittest.TestCase):
    def test_test_method(self):
        self.assertEqual(test_method(), "Test method")


if __name__ == "__main__":
    unittest.main()
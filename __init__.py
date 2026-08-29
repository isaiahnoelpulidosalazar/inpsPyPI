from .builder_for_flask_jsonify import bake
from .dictionarily import Dictionarily
from .easy_sql import EasySQL
from .memory import Memory
from .stackily import Stackily
from .test import test_method
from . import excellent_reader
from . import sort

__all__ = [
    "bake",
    "Dictionarily",
    "EasySQL",
    "Memory",
    "Stackily",
    "test_method",
    "excellent_reader",
    "sort",
]

# To run unit tests, use the following command:
# python -m unittest test_all.py

# To run unit tests without print statements, use the following command:
# python -m unittest -b test_all.py

# changes:
# - change cls to self in Check.Email
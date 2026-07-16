from enum import Enum

class StyleInheritMode(Enum):
    FROM_LEFT = 1
    FROM_RIGHT = 2
    FROM_NONE = 3


class MergeMode(Enum):
    LEFT_MERGE = 1
    RIGHT_MERGE = 2
    ISOLATE = 3
    ISOLATE_LEFT = 4
    ISOLATE_RIGHT = 5
    DIVIDE = 6


class NewlineMode(Enum):
    REPLACE = 1
    PRESERVE = 2


class CleanupMode(Enum):
    NONE = 1
    DELETE_EMPTY = 2
    DELETE_FULL_EMPTY = 3

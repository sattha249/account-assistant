"""
Utility functions for Farm Account Assistant.
Provides category normalization, whitespace handling, and category matching.
Designed for use across all menu functions.
"""

def normalize_category_name(name: str) -> str:
    """
    Normalizes category names by stripping leading/trailing whitespace
    and condensing multiple inner spaces into single spaces.
    
    Example:
        '  หย่านม   LY-F  ' -> 'หย่านม LY-F'
    """
    if name is None:
        return ""
    return " ".join(str(name).strip().split())


def category_key(name: str) -> str:
    """
    Generates a whitespace-insensitive key for category comparison.
    Completely removes all whitespace so that 'หย่านม LY-F' and 'หย่านมLY-F'
    yield the exact same key ('หย่านมLY-F').
    
    Example:
        'หย่านม LY-F' -> 'หย่านมLY-F'
        'หย่านมLY-F'  -> 'หย่านมLY-F'
    """
    if name is None:
        return ""
    return "".join(str(name).strip().split())


def is_same_category(name1: str, name2: str) -> bool:
    """
    Checks if two category strings are logically identical, ignoring whitespace variations.
    
    Example:
        is_same_category("หย่านม LY-F", "หย่านมLY-F") -> True
    """
    if name1 is None or name2 is None:
        return name1 == name2
    return category_key(name1) == category_key(name2)

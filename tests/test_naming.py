"""Safety and portability tests for user-controlled path components."""

import os

from config.naming import safe_child, safe_component, storage_key, unique_component


def test_safe_component_removes_traversal_and_windows_chars():
    value = safe_component(r"../Client:West?.xlsx")
    assert "/" not in value and "\\" not in value
    assert ":" not in value and "?" not in value
    assert ".." not in value


def test_reserved_windows_name_is_neutralized():
    assert safe_component("CON") == "_CON"


def test_storage_key_is_deterministic():
    assert storage_key("My Platform") == "my_platform"


def test_safe_child_stays_beneath_base(tmp_path):
    result = safe_child(str(tmp_path), "../outside")
    assert os.path.commonpath([str(tmp_path.resolve()), os.path.abspath(result)]) == \
        str(tmp_path.resolve())


def test_unique_component_handles_sanitization_and_case_collisions():
    used = set()
    assert unique_component("A/B", used) == "A_B"
    assert unique_component("A:B", used) == "A_B_2"
    assert unique_component("a_b", used) == "a_b_3"

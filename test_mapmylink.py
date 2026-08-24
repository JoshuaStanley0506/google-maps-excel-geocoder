import math

import pytest

from mapmylink import decimal_to_dms, extract_coordinates, is_valid_maps_link, validate_coordinates


@pytest.mark.parametrize(
    ("url", "expected"),
    [
        (
            "https://www.google.com/maps/place/Example/data=!3m1!4b1!4m6!3m5!1sabc!2sExample!8m2!3d12.9715987!4d77.5945627",
            ("12.9715987", "77.5945627", "place"),
        ),
        ("https://www.google.com/maps?q=12.9716,77.5946", ("12.9716", "77.5946", "query")),
        ("https://www.google.com/maps?ll=12.9716,77.5946", ("12.9716", "77.5946", "ll")),
        ("https://www.google.com/maps/@12.9716,77.5946,15z", ("12.9716", "77.5946", "viewport")),
        ("https://www.google.com/maps", ("", "", "")),
    ],
)
def test_extract_coordinates(url, expected):
    assert extract_coordinates(url) == expected


def test_place_coordinates_take_priority_over_viewport():
    url = "https://maps.google.com/@1.0,2.0,15z/data=!3d12.5!4d77.5"
    assert extract_coordinates(url) == ("12.5", "77.5", "place")


@pytest.mark.parametrize(
    ("value", "is_latitude", "expected"),
    [
        (12.9715987, True, '12°58\'17.8"N'),
        (-12.9715987, True, '12°58\'17.8"S'),
        (77.5945627, False, '77°35\'40.4"E'),
        (-77.5945627, False, '77°35\'40.4"W'),
    ],
)
def test_decimal_to_dms(value, is_latitude, expected):
    assert decimal_to_dms(value, is_latitude) == expected


def test_decimal_to_dms_invalid_value():
    assert decimal_to_dms("not-a-number") == ""
    assert decimal_to_dms("") == ""


def test_validate_coordinates():
    assert validate_coordinates(0, 0)
    assert validate_coordinates(90, 180)
    assert validate_coordinates(-90, -180)
    assert not validate_coordinates(90.1, 0)
    assert not validate_coordinates(0, 180.1)
    assert not validate_coordinates("bad", 0)


def test_is_valid_maps_link():
    assert is_valid_maps_link("https://maps.google.com")
    assert is_valid_maps_link("http://maps.google.com")
    assert not is_valid_maps_link("")
    assert not is_valid_maps_link(None)
    assert not is_valid_maps_link("nan")
    assert not is_valid_maps_link("maps.google.com")

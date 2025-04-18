import pytest
from src.extraction import extract_tables_from_pdf

def test_extract_invalid_path():
    with pytest.raises(Exception):
        extract_tables_from_pdf('inexiste.pdf')

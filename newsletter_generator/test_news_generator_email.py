import pytest
import pandas as pd
import numpy as np

from pytest_mock import mocker


from newsletter_generator.news_generator_email import (
    create_car_specs,
    create_highlights_dict,
    create_content_dict,
    create_header,
    create_highlight,
    create_content,
    create_newsletter,
    parse_excel_sheets,
    generate_newsletter,
)


@pytest.fixture
def sample_df():
    data = {
        "Brand": ["Toyota", "Honda"],
        "Model": ["Corolla", "Civic"],
        "year": [2010, 2012],
        "Km": [50000, 60000],
        "Address": ["123 Street", "456 Avenue"],
        "ID": [1, 2],
        "Link_to_folder": ["http://folder1", "http://folder2"],
        "Link_to_pic": ["http://pic1", "http://pic2"],
        "Comentarios": ["Good condition", "Excellent condition"],
        "Ativo": [1, 1],
        "Display_no": [1, 2],
    }
    return pd.DataFrame(data)

def test_create_car_specs(sample_df):
    car_specs = create_car_specs(sample_df, 0)
    assert car_specs == {
        "Brand": "Toyota",
        "Model": "Corolla",
        "year": "2010",
        "KM": "50000",
        "Address": "123 Street",
    }

def test_create_highlights_dict(sample_df):
    highlight_dict = create_highlights_dict(sample_df, 0)
    assert highlight_dict == {
        "NEWS_HIGHLIGHT_TITLE": "ID: 1",
        "HIGHLIGHT_LINK": "http://folder1",
        "HIGHLIGHT_IMAGE": "http://pic1".replace("open?id", "uc?export=view&id"),
        "HIGHLIGHT_TEXT": "Good condition",
        "HIGHLIGHT_FOLDER_LINK": "http://folder1",
    }

def test_create_content_dict(sample_df):
    content_dict = create_content_dict(sample_df, 1)
    assert content_dict == {
        "NEWS_HIGHLIGHT_TITLE": "Oferta: 1   ID: 2",
        "HIGHLIGHT_LINK": "http://folder2",
        "HIGHLIGHT_IMAGE": "http://pic2".replace("open?id", "uc?export=view&id"),
        "HIGHLIGHT_TEXT": "Excellent condition",
        "HIGHLIGHT_FOLDER_LINK": "http://folder2",
    }

def test_parse_excel_sheets(mocker):
    mock_excel = mocker.Mock()
    mock_excel.parse.side_effect = [
        pd.DataFrame({"Value": ["logo", "newsletter_logo", "date", "phone", "email"]}),
        pd.DataFrame({"Brand": ["Toyota"], "Model": ["Corolla"], "year": [2010], "Km": [50000], "Address": ["123 Street"], "Ativo": [1], "Display_no": [1]}),
    ]
    general_df, cars_df = parse_excel_sheets(mock_excel)
    assert not general_df.empty
    assert not cars_df.empty

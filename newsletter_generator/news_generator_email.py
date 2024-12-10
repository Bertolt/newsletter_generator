"""
Created 20/08/2018

This module generates a newsletter by processing data from an Excel file,
creating sections for header, highlights, and content, and combining them
into a final HTML file.
"""

__author__ = 'Ruben Bertelo'

import os
import time
import shutil
import logging
import pandas as pd
import numpy as np

logging.basicConfig(level=logging.INFO)


def create_car_specs(df, index):
    """
    Creates a dictionary with car specifications from a dataframe row.
    
    Args:
        df (pd.DataFrame): The dataframe containing the data.
        index (int): The index of the row to extract data from.
    
    Returns:
        dict: A dictionary with car specifications.
    """
    car_specs = {
        'Brand': df.loc[index]['Brand'],
        'Model': df.loc[index]['Model'],
        'year': df.loc[index]['year'],
        'KM': df.loc[index]['Km'],
        'Address': df.loc[index]['Address']
    }
    for f_key, f_value in car_specs.items():
        if isinstance(f_value, np.float64):
            car_specs[f_key] = str(int(f_value))
        else:
            car_specs[f_key] = str(f_value)
    return car_specs

def create_highlights_dict(df, index):
    """
    Creates a dictionary with highlight configuration values for the newsletter.
    
    Args:
        df (pd.DataFrame): The dataframe containing the data.
        index (int): The index of the row to extract data from.
    
    Returns:
        dict: A dictionary with highlight configuration values.
    """
    highlight_config_dict= {
        'NEWS_HIGHLIGHT_TITLE': f"ID: {int(df.at[index, 'ID'])}",
        'HIGHLIGHT_LINK': df.at[index, 'Link_to_folder'],
        'HIGHLIGHT_IMAGE': df.at[index, 'Link_to_pic'].replace("open?id", "uc?export=view&id"),
        'HIGHLIGHT_TEXT': df.at[index, 'Comentarios'],
        'HIGHLIGHT_FOLDER_LINK': df.at[index, 'Link_to_folder']
    }
    return highlight_config_dict    

def create_content_dict(df, index):
    """
    Creates a dictionary with content configuration values for the newsletter.
    
    Args:
        df (pd.DataFrame): The dataframe containing the data.
        index (int): The index of the row to extract data from.
    
    Returns:
        dict: A dictionary with content configuration values.
    """
    highlight_content_dict= {
        'NEWS_HIGHLIGHT_TITLE': f"Oferta: {index}   ID: {int(df.at[index, 'ID'])}",
        'HIGHLIGHT_LINK': df.at[index, 'Link_to_folder'],
        'HIGHLIGHT_IMAGE': df.at[index, 'Link_to_pic'].replace("open?id", "uc?export=view&id"),
        'HIGHLIGHT_TEXT': df.at[index, 'Comentarios'],
        'HIGHLIGHT_FOLDER_LINK': df.at[index, 'Link_to_folder']
    }
    return highlight_content_dict

def create_header(header_config):
    """
    Creates the newsletter header from a template and replaces placeholders with actual values.
    
    Args:
        header_config (dict): A dictionary containing header configuration values.
    """
    with open("header.html", "r", encoding="utf-8") as template, open("newsletter_header.html", "w", encoding="utf-8") as newsletter:
        for line in template:
            for key, value in header_config.items():
                line = line.replace(key, value)
            newsletter.write(line)


def create_highlight(highlight_config, car_specs):
    """
    Creates the newsletter highlight section from a template and replaces placeholders with actual values.
    
    Args:
        highlight_config (dict): A dictionary containing highlight configuration values.
        car_specs (dict): A dictionary containing car specifications.
    """
    with open("highlights.html", "r", encoding="utf-8") as template, open("newsletter_highlight.html", "w", encoding="utf-8") as newsletter:
        for line in template:
            for key, value in highlight_config.items():
                line = line.replace(key, value)
            for key, value in car_specs.items():
                line = line.replace(key, value)
            newsletter.write(line)


def create_content(content_dict, content_car_specs):
    """
    Creates the newsletter content section from a template and replaces placeholders with actual values.
    
    Args:
        content_dict (dict): A dictionary containing content configuration values.
        content_car_specs (dict): A dictionary containing car specifications.
    """
    with open("content.html", "r", encoding="utf-8") as template, open("newsletter_content.html", "a", encoding="utf-8") as newsletter:
        for line in template:
            for key, value in content_dict.items():
                line = line.replace(key, value)
            for key, value in content_car_specs.items():
                line = line.replace(key, value)
            newsletter.write(line)


def create_newsletter(newsletter_regex, contacts):
    """
    Creates the newsletter by combining header, highlight, and content sections from templates.
    
    Args:
        newsletter_regex (list): A list of regex patterns to identify sections in the template.
        contacts (dict): A dictionary containing contact information.
    """
    # Create new html from template
    with open("template.html", "r+", encoding="utf-8") as template:
        with open("newsletter.html", "w+", encoding="utf-8") as newsletter:
            for line in template:
                for reg in newsletter_regex:
                    if newsletter_regex[0] in line:
                        header_file = open('newsletter_header.html', 'r', encoding="utf-8")
                        header_data = header_file.read()
                        header_file.close()
                        line = line.replace(reg, header_data)
                    if newsletter_regex[1] in line:
                        highlight_file = open('newsletter_highlight.html', 'r', encoding="utf-8")
                        highlight_data = highlight_file.read()
                        highlight_file.close()
                        line = line.replace(reg, highlight_data)
                    if newsletter_regex[2] in line:
                        highlight_file = open('newsletter_content.html', 'r', encoding="utf-8")
                        highlight_data = highlight_file.read()
                        highlight_file.close()
                        line = line.replace(reg, highlight_data)
                # add contacts
                for f_key, f_value in contacts.items():
                    if f_key in line:
                        if f_value is not str:
                            f_value = str(f_value)
                        line = line.replace(f_key, f_value)
                newsletter.write(line)


def parse_excel_sheets(excel_file):
    """
    Parses the Excel file and returns dataframes for the 'General' and 'Cars' sheets.

    Args:
        excel_file (pd.ExcelFile): The Excel file to parse.

    Returns:
        tuple: A tuple containing two dataframes, 
        one for the 'General' sheet and one for the 'Cars' sheet.

    Raises:
        ValueError: If the required sheets are not found in the Excel file.
        Exception: For any other errors that occur during parsing.
    """
    try:
        df_general = excel_file.parse('General')
        df_cars = excel_file.parse('Cars').dropna()
        return df_general, df_cars
    except ValueError as ve:
        logging.error(f"Error: {ve}")
        raise
    except Exception as e:
        logging.error(f"An error occurred while parsing the Excel file: {e}")
        raise


def generate_newsletter():
    """
    Generates a newsletter by performing the following steps:
    1. Loads data from an Excel file and creates dataframes for each sheet.
    2. Configures header and contact information.
    3. Creates the newsletter header.
    4. Creates highlights and content sections for the newsletter.
    5. Generates the final newsletter HTML files.
    Raises:
        FileNotFoundError: If the Excel file or any required files are not found.
        KeyError: If required columns are missing in the Excel sheets.
        Exception: For any other errors that occur during the newsletter generation process.
    """
    cwd = os.getcwd()
    # Backup older newsletters
    for file in os.listdir(cwd):
        if file.startswith('newsletter'):
            shutil.move(os.path.join(cwd, file), os.path.join(cwd, 'old'))

    try:
        excel_file = pd.ExcelFile('file.xlsx')
    except FileNotFoundError:
        logging.error("Error: The file 'file.xlsx' was not found.")
        raise
    except Exception as e:
        logging.error(f"An error occurred while opening the Excel file: {e}")
        raise

    general_df, cars_df = parse_excel_sheets(excel_file)

    # Load values for header
    logo, newsletter_logo, newsletter_date = general_df['Value'].values[:3]
    newsletter_date = (
        newsletter_date if pd.notna(newsletter_date) else
        f"{time.strftime('%A')},  {time.strftime('%d')} de {time.strftime('%B')}  {time.strftime('%Y')}"
    )
    phone, email = general_df['Value'].values[3:5]

    header_config = {
        'LOGO': logo,
        'NEWSLETTER_IMAGE': newsletter_logo,
        'NEWSLETTER_DATE': newsletter_date
    }

    contacts_dict = {
        'TELEPHONE_NUMBER': phone,
        'EMAIL_LINK': email,
        'EMAIL_DISPLAY': email
    }

    newsletter_regex = ['<!-- HEADER_REG_EXP -->', '<!-- HIGHLIGHT_REG_EXP -->', '<!-- CONTENT_REG_EXP -->']

    # Filter and preprocess car data
    cars_df = cars_df[cars_df['Ativo'].astype(str) != '0'].sort_values(by=['Display_no']).reset_index(drop=True)

    create_header(header_config)
    create_highlight(create_highlights_dict(cars_df, 0), create_car_specs(cars_df, 0))

    for index in cars_df.index[1:]:
        create_content(create_content_dict(cars_df, index), create_car_specs(cars_df, index))

    create_newsletter(newsletter_regex, contacts_dict)

    timestamp = time.strftime("%d%m%Y%H%M%S")
    os.rename("newsletter.html", f'newsletter_{timestamp}.html')
    os.rename("newsletter_content.html", f'newsletter_content_{timestamp}.html')
    os.rename("newsletter_header.html", f'newsletter_header_{timestamp}.html')
    os.rename("newsletter_highlight.html", f'newsletter_highlight_{timestamp}.html')


def main():
    try:
        generate_newsletter()
    except Exception as e:
        logging.error(f"Oops! Something went wrong: {e}")


if __name__ == "__main__":
    main()

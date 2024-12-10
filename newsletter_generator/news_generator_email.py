"""
Created 20/08/2018
"""

__author__ = 'Ruben Bertelo'

import os
import time
import shutil

import pandas as pd
import numpy


def create_car_specs(df, x):
    """
    Creates a dictionary with car specifications from a dataframe row.
    
    Args:
        df (pd.DataFrame): The dataframe containing the data.
        x (int): The index of the row to extract data from.
    
    Returns:
        dict: A dictionary with car specifications.
    """
    car_specs = {
        'Brand': df.loc[x]['Brand'],
        'Model': df.loc[x]['Model'],
        'year': df.loc[x]['year'],
        'KM': df.loc[x]['Km'],
        'Address': df.loc[x]['Address']
    }
    for f_key in car_specs:
        if type(car_specs[f_key]) is not str and isinstance(car_specs[f_key], numpy.float64):  #type(carSpecs[f_key]) is 'numpy.float64'
            car_specs[f_key] = car_specs[f_key].astype(int)
            car_specs[f_key] = str(car_specs[f_key])
        else:
            car_specs[f_key] = str(car_specs[f_key])
    return car_specs

def create_highlights_dict(df, x):
    """
    Creates a dictionary with highlight configuration values for the newsletter.
    
    Args:
        df (pd.DataFrame): The dataframe containing the data.
        x (int): The index of the row to extract data from.
    
    Returns:
        dict: A dictionary with highlight configuration values.
    """
    highlight_config_dict = {
        'NEWS_HIGHLIGHT_TITLE': 'ID:' + str(df.loc[x]['ID'].astype(int)),
        'HIGHLIGHT_LINK': df.loc[x]['Link_to_folder'],
        'HIGHLIGHT_IMAGE': df.loc[x]['Link_to_pic'].replace("open?id","uc?export=view&id"),
        'HIGHLIGHT_TEXT': df.loc[x]['Comentarios'],
        'HIGHLIGHT_FOLDER_LINK': df.loc[x]['Link_to_folder']
    }
    return highlight_config_dict

def create_content_dict(df, x):
    """
    Creates a dictionary with content configuration values for the newsletter.
    
    Args:
        df (pd.DataFrame): The dataframe containing the data.
        x (int): The index of the row to extract data from.
    
    Returns:
        dict: A dictionary with content configuration values.
    """
    highlight_content_dict= {
        'NEWS_HIGHLIGHT_TITLE': 'Oferta:  '+str(x)+'   ID:'+str(df.loc[x]['ID'].astype(int)),
        'HIGHLIGHT_LINK': df.loc[x]['Link_to_folder'],
        'HIGHLIGHT_IMAGE': df.loc[x]['Link_to_pic'].replace("open?id", "uc?export=view&id"),
        'HIGHLIGHT_TEXT': df.loc[x]['Comentarios'],
        'HIGHLIGHT_FOLDER_LINK': df.loc[x]['Link_to_folder']
    }
    return highlight_content_dict

def create_header(header_config):
    """
    Creates the newsletter header from a template and replaces placeholders with actual values.
    
    Args:
        header_config (dict): A dictionary containing header configuration values.
    """
    # Create new header from header template
    with open("header.html", "r+", encoding="utf-8") as template:
        with open("newsletter_header.html", "w+", encoding="utf-8") as newsletter:
            for line in template:
                for f_key, f_value in header_config.items():
                    if f_key in line:
                        line = line.replace(f_key, f_value)
                newsletter.write(line)
    newsletter.close()
    template.close()

def create_higlight(highlight_config, car_specs):
    """
    Creates the newsletter highlight section from a template and replaces placeholders with actual values.
    
    Args:
        highlight_config (dict): A dictionary containing highlight configuration values.
        car_specs (dict): A dictionary containing car specifications.
    """
    # Create new header from header template
    with open("highlights.html", "r+", encoding="utf-8") as template:
        with open("newsletter_highlight.html", "w+", encoding="utf-8") as newsletter:
            for line in template:
                for f_key, f_value in highlight_config.items():
                    if f_key in line:
                        line = line.replace(f_key, f_value)
                for c_key, c_value in car_specs.items():
                    if c_key in line:
                        line = line.replace(c_key, c_value)
                newsletter.write(line)
    newsletter.close()
    template.close()

def create_content(content_dict, content_car_specs):
    """
    Creates the newsletter content section from a template and replaces placeholders with actual values.
    
    Args:
        content_dict (dict): A dictionary containing content configuration values.
        content_car_specs (dict): A dictionary containing car specifications.
    """
    # create content from content.html
    with open("content.html", "r+", encoding="utf-8") as template:
        with open("newsletter_content.html", "a", encoding="utf-8") as newsletter:
            for line in template:
                for f_key, f_value in content_dict.items():
                    if f_key in line:
                        line = line.replace(f_key, f_value)
                for c_key, c_value in content_car_specs.items():
                    if c_key in line:
                        line = line.replace(c_key, c_value)
                newsletter.write(line)
    newsletter.close()
    template.close()

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

    newsletter.close()
    template.close()

# ----- Variables ----- #
carSpecList = []

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
    files = os.listdir(cwd)

    for f in files:
        if f[0:10] == 'newsletter':
            shutil.move(cwd+'\\'+f, cwd+'\\old')

    # Load the excel and create dataframe for each sheet os.chdir('.\\files')
    xl = pd.ExcelFile('file.xlsx')
    df1 = xl.parse('General')
    df2 = xl.parse('Cars')
    # delete rows with blank Values
    df2 = df2.dropna()

    # Load values for header
    logo = df1['Value'].values[0]
    newsletter_logo = df1['Value'].values[1]

    newsletter_date = df1['Value'].values[2]

    if pd.isna(newsletter_date):
        newsletter_date = time.strftime("%A")+',  '+time.strftime("%d")+' de '+time.strftime("%B")+'  '+time.strftime("%Y")
    phone = df1['Value'].values[3]
    email = df1['Value'].values[4]

    header_config = {
        'LOGO': logo,
        'NEWSLETTER_IMAGE': newsletter_logo,
        'NEWSLETTER_DATE': newsletter_date
    }

    contacts_dict = {
        'TELEPHONE_NUMBER' : phone,
        'EMAIL_LINK' : email,
        'EMAIL_DISPLAY': email
    }

    newsletter_regex = ['<!-- HEADER_REG_EXP -->',
                        '<!-- HIGHLIGHT_REG_EXP -->',
                        '<!-- CONTENT_REG_EXP -->']

    # Load Values for Cars list
    # needs to pass to dict in order to get the x to delete teh row on the matrix df2
    car_config = df2.to_dict()
    df3 = df2
    # Create new matrix withou inactive adds
    for x in car_config['Ativo']:
        if car_config['Ativo'][x] == 0 or car_config['Ativo'][x] == '0':
            df3 = df3.drop([x])

    # Sort new the cars matrix
    df3 = df3.sort_values(by=['Display_no'], ascending=True)
    df3 = df3.dropna()
    df3 = df3.reset_index()

    # Create header
    create_header(header_config)

    # Create Highlights
    highlights_dict = create_highlights_dict(df3, 0)
    highlight_car_spec = create_car_specs(df3, 0)
    create_higlight(highlights_dict, highlight_car_spec)

    # Create content
    for x in df3.index.tolist()[1:]:
        content_dict = create_content_dict(df3, x)
        content_car_specs = create_car_specs(df3, x)
        create_content(content_dict, content_car_specs)

    # Create newsletter
    create_newsletter(newsletter_regex, contacts_dict)

    # Rename files
    today = time.strftime("%d%m%Y%H%M%S")
    os.rename("newsletter.html", 'newsletter_'+today+'.html')
    os.rename("newsletter_content.html", 'newsletter_content_'+today+'.html')
    os.rename("newsletter_header.html", 'newsletter_header_'+today+'.html')
    os.rename("newsletter_highlight.html", 'newsletter_highlight_'+today+'.html')

def main():
    try:
        generate_newsletter()
    except (RuntimeError, TypeError, NameError):
        print("Oops!  Something went wrong.  Try again...")

if __name__ == "__main__":
    main()

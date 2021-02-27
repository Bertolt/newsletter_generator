"""
Created 20/08/2018
Author Ruben Bertelo
"""
import os as os
import time

import pandas as pd
import numpy
import shutil


def create_car_specs(df,x):
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
    highlight_config_dict = {
        'NEWS_HIGHLIGHT_TITLE':   ID:'+str(df.loc[x]['ID'].astype(int)),
        'HIGHLIGHT_LINK': df.loc[x]['Link_to_folder'],
        'HIGHLIGHT_IMAGE': df.loc[x]['Link_to_pic'].replace("open?id","uc?export=view&id"), # replace the str so gdrive can display image on html
        'HIGHLIGHT_TEXT': df.loc[x]['Comentarios'],
        'HIGHLIGHT_FOLDER_LINK': df.loc[x]['Link_to_folder']
    }
    return highlight_config_dict

def create_content_dict(df, x):
    highlight_content_dict= {
        'NEWS_HIGHLIGHT_TITLE': 'Oferta:  '+str(x)+'   ID:'+str(df.loc[x]['ID'].astype(int)),
        'HIGHLIGHT_LINK': df.loc[x]['Link_to_folder'],
        'HIGHLIGHT_IMAGE': df.loc[x]['Link_to_pic'].replace("open?id", "uc?export=view&id"), # replace the str so gdrive can display image on html
        'HIGHLIGHT_TEXT': df.loc[x]['Comentarios'],
        'HIGHLIGHT_FOLDER_LINK': df.loc[x]['Link_to_folder']
    }
    return highlight_content_dict

def create_header(headerConfig):
    # Create new header from header template
    with open("header.html", "r+") as template:
        with open("newsletter_header.html", "w+") as newsletter:
            for line in template:
                for f_key, f_value in headerConfig.items():
                    if f_key in line:
                        line = line.replace(f_key, f_value)
                newsletter.write(line)
    newsletter.close()
    template.close()

def create_higlight(highlightConfig, CarSpecs):
    # Create new header from header template
    with open("highlights.html", "r+") as template:
        with open("newsletter_highlight.html", "w+") as newsletter:
            for line in template:
                for f_key, f_value in highlightConfig.items():
                    if f_key in line:
                        line = line.replace(f_key, f_value)
                for c_key, c_value in CarSpecs.items():
                    if c_key in line:
                        line = line.replace(c_key, c_value)
                newsletter.write(line)
    newsletter.close()
    template.close()

def create_content(contentDict, contentCarSpecs):
    # create content from content.html
    with open("content.html", "r+") as template:
        with open("newsletter_content.html", "a") as newsletter:
            for line in template:
                for f_key, f_value in contentDict.items():
                    if f_key in line:
                        line = line.replace(f_key, f_value)
                for c_key, c_value in contentCarSpecs.items():
                    if c_key in line:
                        line = line.replace(c_key, c_value)
                newsletter.write(line)
    newsletter.close()
    template.close()

def create_newsletter(newsletterRegExp, contacts):
    # Create new html from template
    with open("template.html", "r+") as template:
        with open("newsletter.html", "w+") as newsletter:
            for line in template:
                for reg in newsletterRegExp:
                    if newsletterRegExp[0] in line:
                        headerFile = open('newsletter_header.html', 'r')
                        headerData = headerFile.read()
                        headerFile.close()
                        line = line.replace(reg, headerData)
                    if newsletterRegExp[1] in line:
                        highlightFile = open('newsletter_highlight.html', 'r')
                        highlightData = highlightFile.read()
                        highlightFile.close()
                        line = line.replace(reg, highlightData)
                    if newsletterRegExp[2] in line:
                        highlightFile = open('newsletter_content.html', 'r')
                        highlightData = highlightFile.read()
                        highlightFile.close()
                        line = line.replace(reg, highlightData)
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
today = time.strftime("%d%m%Y%H%M%S")
newsletterDate = time.strftime("%A")+',  '+time.strftime("%d")+' de '+time.strftime("%B")+'  '+time.strftime("%Y")

def generate_newsletter():
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
    newsletterLogo = df1['Value'].values[1]
    # newsletterDate = df1['Value'].values[2]
    phone = df1['Value'].values[3]
    email = df1['Value'].values[4]

    headerConfig = {
        'LOGO': logo,
        'NEWSLETTER_IMAGE': newsletterLogo,
        'NEWSLETTER_DATE': newsletterDate
    }

    contactsDict = {
        'TELEPHONE_NUMBER' : phone,
        'EMAIL_LINK' : email,
        'EMAIL_DISPLAY': email
    }

    newsletterRegExp = ['<!-- HEADER_REG_EXP -->', '<!-- HIGHLIGHT_REG_EXP -->', '<!-- CONTENT_REG_EXP -->']
    bottomRegExp = ['TELEPHONE_NUMBER', 'EMAIL_LINK', 'EMAIL_DISPLAY']

    # Load Values for Cars list
    # needs to pass to dict in order to get the x to delete teh row on the matrix df2
    carConfig = df2.to_dict()
    df3 = df2
    # Create new matrix withou inactive adds
    for x in carConfig['Ativo']:
        if carConfig['Ativo'][x] == 0 or carConfig['Ativo'][x] == '0':
            df3 = df3.drop([x])

    # Sort new the cars matrix
    df3 = df3.sort_values(by=['Display_no'], ascending=True)
    df3 = df3.dropna()
    df3 = df3.reset_index()

    # Create header
    create_header(headerConfig)

    # Create Highlights
    highlightsDict = create_highlights_dict(df3, 0)
    highlightCarSpec = create_car_specs(df3, 0)
    create_higlight(highlightsDict, highlightCarSpec)

    # Create content
    for x in df3.index.tolist()[1:]:
        contentDict = create_content_dict(df3, x)
        contentCarSpecs = create_car_specs(df3, x)
        create_content(contentDict, contentCarSpecs)

    # Create newsletter
    create_newsletter(newsletterRegExp, contactsDict)

    # Rename files
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
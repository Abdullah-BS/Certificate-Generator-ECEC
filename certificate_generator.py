import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
import bidi.algorithm
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import io


# Code written by: Abdullah Binsalman 

def generate_cert(name, gender):
    if gender == 'male':
        image = Image.open("Template_boy.png")      #Template path
    else:
        image = Image.open("Template_girl.png")      #Template path
        
    d = ImageDraw.Draw(image)
    font = ImageFont.truetype("DINNextLTArabic-Regular-3.ttf", 100)     #Font - Font size

    reshaped_text = arabic_reshaper.reshape(name)
    text = bidi.algorithm.get_display(reshaped_text)
    
    text_bbox = d.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]

    image_width, image_height = image.size

    text_x = (image_width - text_width) / 2
    
    text_col = (15, 127, 78)     #Color

    d.text((text_x, 550),   #Position (550 default)
            text=text,
            fill=text_col,
            font=font,
            stroke_width=1,         #For bold
            stroke_fill=text_col)

    folder_name = "Certificates"        #Name of folder
    os.makedirs(folder_name, exist_ok=True)      #Create a folder in the current directory with the folder name
    
    saving_path = os.path.join(folder_name, f"{'_'.join(name.split())}.png") #Saved file name

    count = 2
    while os.path.exists(saving_path):
        saving_path = os.path.join(folder_name, f"{name.split().join('_')}_{count}.png")
        count += 1

    
    saving_path_pdf = saving_path.replace('.png', '.pdf')  # Convert to PDF file name
    image.convert('RGB').save(saving_path_pdf)

    base_directory = os.path.join(os.getcwd(), saving_path_pdf)

    return base_directory
    

## Application

df = pd.read_excel("names.xlsx")
for index, row in df.iterrows():
    
    name = row['name']
    email = row['email']
    gender = row['gender']
    saving_path = generate_cert(name, gender)
    
    df.loc[index, 'cert'] = saving_path
    
    with pd.ExcelWriter("names.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    
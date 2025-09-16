# Certificate Generator
This project is a Python-based certificate generator that creates personalized certificates from an Excel spreadsheet. It's designed to automate the process of generating and saving certificates for a list of names, emails, and genders.


## Features
- **Batch Generation**: Automatically generates certificates for all entries in a provided Excel file.


- **Gender-Specific Templates**: Uses different certificate templates for male and female recipients.


- **Custom Arabic Text**: Supports Arabic text for names, ensuring correct shaping and display using the arabic-reshaper and python-bidi libraries.


- **Personalization**: Dynamically inserts the recipient's name onto the certificate image.


- **PDF Output**: Saves each generated certificate as a PDF file for easy printing and sharing.

## Project Structure
The project is organized into a clear directory structure to manage the different components.

- **Certificates/:** This folder is automatically created by the script to store the generated certificates as PDF files.

- **DINNextLTArabic-Regular-3.ttf**: This is the font file used to render the Arabic names on the certificates.

- **Template_boy.png**: The certificate template used for male recipients.

- **Template_girl.png**: The certificate template used for female recipients.

- **certificate_generator.py**: The main Python script that contains all the logic for reading data, generating certificates, and saving the output.

- **names.xlsx**: An Excel spreadsheet where you enter the data for all certificate recipients. The script reads from this file to get names, emails, and genders.

## How to Use

1. **Prepare your data:** Open the names.xlsx file and populate it with the details of your recipients. The spreadsheet must include three columns with the exact headers: 
name, email, and gender.


2. **Run the script:** Open a terminal or command prompt, navigate to the project's root directory, and execute the Python script:
`python certificate_generator.py`


3. **Find your certificates:** Once the script completes its run, a new folder named Certificates will be created in the project directory. Inside, you will find all the generated certificates saved as PDF files, named after each recipient.

"""Author: Denzel Udemba
Date: 05/10/2023
Description: This code uses a basic word letter template addressed
to GA State representatives and adds variables in place of their names
districts, Rooms, and cities they represent. Data is sourced from a private .env file"""

import os
import shutil
from docx import Document
from dotenv import load_dotenv
from docx2pdf import convert


def main():
    # Load .env variables
    load_dotenv()

    # Get paths from .env variables
    source_file = os.getenv("SOURCE_FILE")
    target_directory = os.getenv("TARGET_DIRECTORY")
    pdf_directory = os.getenv("PDF_DIRECTORY")
    target_files = os.getenv("TARGET_FILE").split(",")
    date = os.getenv("__DATE__")
    senator_names = os.getenv("__SENATOR_NAME__").split(",")
    senator_last_names = os.getenv("__SEN_LNAME__").split(",")
    district_nos = os.getenv("__DISTRICT_NO__").split(",")
    room_nos = os.getenv("__ROOM_NO__").split(",")
    city_variables = os.getenv("__CITY__").split(",")

    # Loop through each target file
    for new_file, senator_name, senator_last_name, district_no, room_no, \
        city_variable in zip(target_files,
                             senator_names, senator_last_names, district_nos, room_nos, city_variables):
        # Copy source file to target directory
        shutil.copyfile(source_file, os.path.join(target_directory, new_file))

        # Open target file for editing
        document = Document(os.path.join(target_directory, new_file))

        # parse through new_file & change the following field if applicable
        for paragraph in document.paragraphs:
            if "__DATE__" in paragraph.text:
                paragraph.text = paragraph.text.replace("__DATE__", date)
                print(f"Date: {date} changed successfully")

        for paragraph in document.paragraphs:
            if "__SENATOR_NAME__" in paragraph.text:
                paragraph.text = paragraph.text.replace("__SENATOR_NAME__", senator_name)
                print(f"Senator Name: {senator_name} changed successfully")

        for paragraph in document.paragraphs:
            if "__SEN_LNAME__" in paragraph.text:
                paragraph.text = paragraph.text.replace("__SEN_LNAME__", senator_last_name)
                print(f"Senator Lastname: {senator_last_name} changed successfully")

        for paragraph in document.paragraphs:
            if "__DISTRICT_NO__" in paragraph.text:
                paragraph.text = paragraph.text.replace("__DISTRICT_NO__", district_no)
                print(f"{district_no} changed successfully")

        for paragraph in document.paragraphs:
            if "__ROOM_NO__" in paragraph.text:
                paragraph.text = paragraph.text.replace("__ROOM_NO__", room_no)
                print(f"{room_no} changed successfully")

        for paragraph in document.paragraphs:
            if "__CITY__" in paragraph.text:
                paragraph.text = paragraph.text.replace("__CITY__", city_variable)
                print(f"City Name: {city_variable} changed successfully")

        # Save updated file
        document.save(os.path.join(target_directory, f"{new_file}"))
        print(f"{new_file}: saved successfully to {os.path.join(target_directory, f'{new_file}')}")

        # Convert the file to PDF
        convert(os.path.join(target_directory, new_file), os.path.join(pdf_directory, f"{new_file}.pdf"))
        print(f"{new_file}: converted successfully to {os.path.join(pdf_directory, f'{new_file}.pdf')}")


if __name__ == '__main__':
    main()

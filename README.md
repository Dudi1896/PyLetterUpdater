#: PyLetterUpdater
![image](https://github.com/Dudi1896/PyLetterUpdater/assets/17666999/6b7e4c98-94e6-4aa9-9841-8e94e0fa3d08)

Simple program to modify a Base File word document and generate a multitude of unique new files (pdf and or docx) depending on what values are added to the base file
The code will generate as many new unique files as you want depending on how many variables you add to .env file
multiple .env file variables can be stacked as a list and must be delimeted by a comma

###:  How to Use
- Gather your generic word document that wil server as your base File Template
- create a .env file in the root directory of the program and file in your variables. look at example.env to understand the structure
- SOURCE_DIRECTORY={Directory of your Base File Template}
- TARGET_DIRECTORY={Directory specifiying the location of newely created word document(s)}
- PDF_DIRECTORY={Directory specifying the location of newely creaed pdf document(s)}

- The following variables below are preferential and dependant on the users requirements.
- Multiple variables can be stacked as a list that the LetterUpdater.py code will iterate through until completed. All stacked variables must be 
- deliniated by a comma. In this example 2023, Mike Jones, 45, 453B & Columbus will all be in File no.1 and Gabbie Edwards etc in File no.2 and so on
- the variables in the .env file must also match the variables specified for change in the Base File Template

__DATE__=2023,2024
__SENATOR_NAME__=Mike Jones,Gabbie Edwards
__DISTRICT_NO__=45,23
__ROOM_NO__=453B,323A
__CITY__=Columbus,Houston

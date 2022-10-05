# docufill
Docufill replaces words in a docx file based on a dictionary stored in plaintext.

Three folders are created upon the initial run of the script:

/templates - contains the template files (docx)

/dictionary - contains dictionary.txt

/output - docufilled templates are sent here

Simply insert a tag where you would like text replaced in an msword file eg. {{tag1}} and save the new template file to the templates folder

Open dictionary.txt and create a new keypair value for the tag where the value is the text you would like replaced

Run docufill.py


## Setting Up and Running
##### Create virtual environment
```
python -m venv venv
```
##### Activate the new virtual environment
```
.\venv\Scripts\activate
```
##### Install dependencies from requirements.txt
```
pip install -r requirements.txt
```
##### Run
```
python docufill.py
```

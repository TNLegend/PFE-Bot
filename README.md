

## AppPassword

first thing first get your google AppPassword for gmail and store it somewhere.
PS: don't share your appPassword with anyone.

Geneate it from here :

[gmail](https://myaccount.google.com/apppasswords)

[icloud](https://support.apple.com/fr-fr/HT204397)

## Installing requirements

Just run the script `install_requirements.bat` or type `pip install -r requirements.txt`

## Importing Resume Template

now launch the program type `1` to import your file.

**Make sure it is a word document**

You can convert pdf to document [here](https://www.ilovepdf.com/pdf_to_word)

## Configuration

Now after importing your document:

1-type the text to replace. Let's say you got already a resume and follow up letter for company `X` but you dont want to change the company name each time you want to send to a new company. The bot will do for you so you type the text you want to replace which is `X` in our case and it will replace that text with the company names provided in emails file (we'll talk about this file later) 

2- Select the txt file that has your emails. lines in this file should follow this format `<email>:<company_name>`

Exemple : `company1@gmail.com:company1`

3- type your email 

4- type your password/appPassword

5-Type email subject and emails Body

6-Now wait till it finishes sending

**PS: the bot needs the resume file in document format in order to be able to replace the company name ,it ll convert it to pdf and send the pdf to emails provided

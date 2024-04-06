# Green Got pdf to xlsx
---
Simple script that will convert the input pdf file to xlsx.
The pdf file must be an account statement from Green Got.

To run it:

on console
```sh
$ python -m venv venv
$ source ./venv/bin/activate
$ pip install -r requirements.txt
```
and then you can run it:
```sh
$ python main.py <YOUR FILE>
```
or you can install it globaly:
```sh
$ pip install -r requirement.txt
$ python main.py
```

## USAGE

The script expect one filename that won't be checked. The filename must point to the pdf file, including the extension.

It will generate an xlsx file with the same name (without .pdf extension of course)
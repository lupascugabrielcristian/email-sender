### Install
pyenv virtualenv 3.8.0 miniemail  
pyenv activate miniemail  
pip install --upgrade pip  
pip install openpyxl  
pip install requests  
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
pip install PyGObject  

### Prepare
Must have a Google Project active with Google Mail API enabled. That project must have a OAuth 2.0 Client ID user ( Desktop application). From that client credentials.json must be saved locally, in the running directory. In case the project is a "Test project", the email used for authentication window, must be added in the "OAuth consent screen", "Test users" section.  
Must create the relevant file structure as described in the xsls file, including the pdf files to be attached.

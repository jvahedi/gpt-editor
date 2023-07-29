@echo off
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org aspose-words
python %CD%\GPTDocsEdit.py %*
pause
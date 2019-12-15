# quick_translate
Translate files in the directory into English

Wiki needs to be updated.

Compilation notes (these steps necessary to ensure that package is small; otherwise, 200+ MB size)
1. Created virtual environment (virtualenv quick_translate). Deactivate base anaconda environment
2. In virtual env, installed all libraries (xlrd, python-pptx, pypiwin32, googletrans)
3. In virtual env, installed pyinstaller
4. Using pyinstaller -w -F (meaning not windowed and onefile), compiled script

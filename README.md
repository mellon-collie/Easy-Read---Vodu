# Easy-Read---Vodu
A Tkinter based app which produces reading material suited for anyone suffering from vestibular imbalance.

The uploaded files are the EasyRead.py, ICON.ico, settings.txt and this file.

- To create an executable, install pyinstaller for python-3x.

- cd to the directory containing project and open terminal

  test the code first by running 
  
  python3 EasyRead.py
 
  
  check if everything works, and then ,

  go to terminal and run,

  pyinstaller --onedir --windowed --icon=Icon.ico EasyRead.py

  once the .exe has been created follow these steps-

  1) Copy settings.txt and place it in dist/EasyRead

  2) Go back to the directory containing EasyRead.py and the other files.
     Right click on settings.txt and open Properties. Copy the file path of
     file. Open EasyRead.spec and within the Analysis instance declaration,
     go to the 'datas' attribute. The 'datas' attribute is currently assigned to
     an empty list. Now add a tuple ('copied_file_path_of_settings\settings.txt','.') to the list. Similarly do the follow the same steps for      Guide.txt -('copied_file_path_of_Guide\Guide.txt','.')
	
  3) Now, click on dist, and then click on EasyRead (basically dist/EasyRead)
  
  4) Create a folder called docx

  5) Open docx folder and create another folder called templates
  
  6) Open templates and create a word file called 'default.docx'

  7) Open the file 'default.docx' and place the cursor on the text field,
     and type  3-4 spaces. Save it and close the word file.

  8) Now, run and EasyRead.exe file in dist/EasyRead folder and check if it is      running.
  9) If the application falters at any point, either debug the EasyRead.py file        again, or go to the directory containing EasyRead.py. Delete EasyRead.spec,
     dist, build, __pycache__ folder and open terminal again.
     this time run 
     
     pyinstaller --debug EasyRead.py 

     This will create another similar set of folders and navigate to      dist/EasyRead folder and run the application. This time the application
     will open along with a debugging window which might help you to trace the
     error.
  10) If the application works fine, then go back to the dist folder, and compress the EasyRead folder.
  11) Once the zipping is done, create a fresh folder on Desktop and copy the zip file. Extract it and run EasyRead.exe again by navigating to EasyRead folder and to test if the application is running fine.
  12) If there are no issues, then you can distribute this zip file to your users.
      	
  

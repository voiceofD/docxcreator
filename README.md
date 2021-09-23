# docxcreator

This is a small program that builds some docx files!

Using this GUI Application, you can build MS Word (.docx) files quite easily. The purpose of this code is to demonstrate the proof of concept as to how documents can be created on the fly to meet various requirements. 


**The unique features of this program are:

This program allows you to easily perform the operations that most people do on MS Word: For example,
    
    --> *Create coverpages*
    
    --> *Define margins (in Cm/Inches)*
    
    --> *Upload images* 
    
    --> *Define tables*
    
    --> *Write paragraphs*
    
    --> *Write headings* 
    
    --> *Add page breaks, in addition to* 
    
    --> *Format text as per requirements (Bold, Italic, Underline, Alignment).*


This application also has a **record feature** which starts recording the instructions that you give, once you toggle the record button. It then repeats the recorded instructions a specified number of times, allowing you to 
    
    --> Seamlessly create multiple copies of a page without any hassles
    
    --> Simplify the process of creating templates
    
    --> Create templates for invoices and so on...
    

There are 3 versions of the program, which ensures its usage in most of the platforms, at any point of time.
1) A python commandline program (My personal favorite) that takes user input on cmd and outputs the file in the same directory as the program
2) A python GUI program using tkinter that has the same functionality as the commandline program
3) A windows executable (.exe) program that allows the user to run the GUI application without needing to install python on their systems

INSTRUCTIONS TO USE THE COMMANDLINE PROGRAM

Make sure that you install Python 3.x and add it to the PATH. Install the following dependencies **python-docx** (pip install python-docx) and **docxcompose** (pip install docxcompose). That's it! Run the python program (word.py) on the commandline and have fun!


INSTRUCTIONS TO USE THE STANDALONE GUI PROGRAM

Make sure that you install Python 3.x and add it to the PATH. Install the following dependencies **python-docx** (pip install python-docx) and **docxcompose** (pip install docxcompose). That's it! Run the python program (guiapp.py) on the commandline and have fun!

INSTRUCTIONS TO USE THE WINDOWS EXECUTABLE PROGRAM 

Download the contents of the GUI App folder. Navigate to the 'dist' folder and run **prototype.exe**. That's all you need to run the program. The created MS Word docs will be saved in the "Created Files" folder

POSSIBILITIES WITH THE PROJECT!

While this application is very rudimentary and fundamental in nature, anybody without a myopic vision can understand its potential. MS Word is an expensive piece of software and running it on several machines can cost a fortune! However, Python is free. It is easily installed on any platform and can even be containerised! And all one needs, is an image of this kind of application to create .docx files on the fly. 
What's more? You need not even run this applcation yourself! By feeding your instructions on a .csv/.txt file, you can just sit back and let your code do all the work for you!
You can host this on your systems, your servers, and even as a microservice! The possibilities are endless!

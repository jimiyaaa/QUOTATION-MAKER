I work as salesman for a material handling company, and we need to type in our quotation manually here, so i try to come up with a solution by making a program that can generate quotation.

How to run

    Extract ZIP (has to be new independent folder).
    Python 3 must have installed.
    Double click run.bat.
    Browser will run the app.

FILE EXPORT
I use libre office as the file generator and converter because is more stable. But, if you dont have libre installed, the program will try to use Microsoft office. Crashes sometime happened when it uses Microsoft office. Just restart the app it wont affect the database

DATA STORAGE
When you made a quotation, a database file "quotation_history.db" is created automatically in this folder. This file must be in the same folder as app.py to retain your quotation history. So dont move it anywhere.
Most importantly, The database is local to you computer only and is not uploaded to an online server.

DOCUMENT TEMPLATE
QT-new format.docx is the document template used for DOCX/PDF output. You may replace it with another compatible template. Mind the input placement (marked by yellow highlight)


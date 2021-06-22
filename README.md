# quickMSDV
Automated processing of spectral data

This Python program accompanies a spectral analysis program written by Steven Rofrano. 

His code (named MSDataView) extracts spectral information from Agilent data files (XML) and converts the data into a text file format.

However, MSDataView only runs one sample at a time and processing them manually can quickly become tedious, especially if you have 20-100 of these samples.

quickMSDV aims to solve that issue by iterating through a folder containing the relevant data files.

Modifications have to be made in the code to ensure proper pathing to the target folder.

--------------

Ongoing minor bug: Because our data is usually on a university NAS and not stored locally, opening folders in the path search process can have a delay outlasting the program's built-in sleep function.

Note: Because this is a PyAutoGUI based program, the user should not be using the computer during this processing.

Note: If MSDataView had a command line format, this whole process could probably be automated in Python for all relevant ".d" files in a folder.

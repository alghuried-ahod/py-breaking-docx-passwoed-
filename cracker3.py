"""
Ahod Alghuried, D15123616
Mobile Forinsic, Lab #3 

This Python program has been designed to break the password for docx files.

To get Windows Extensions you could download it from this link
( https://sourceforge.net/projects/pywin32/files/pywin32/Build%2520220/ ). 
 After download it write this command in the terminal ( pip install python-docx ).Then you will be able to import windows library. 
 
"""
import sys
import win32com.client
from win32com import client
# To make Word object
a = client.Dispatch("Word.Application")

filename= sys.argv[1]
# open the world list which is available online, to get all possible passwords from it 
password_file = open ( 'wordlist2.txt', 'r' )
passwords = password_file.readlines()
password_file.close()

passwords = [item.rstrip('\n') for item in passwords]
# write the correct password in this file 
results = open('results.txt', 'w')

for password in passwords:
	print(password)
	try:
		doc = a.Documents.Open(filename,True,True,True,PasswordDocument=password)
 		print "Success! Password is: " + password
		results.write(password)
		results.close()
		sys.exit()
	except:
		print "Incorrect password"
		pass
		
		
		
		

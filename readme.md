##HOW TO USE:
simply just grab an api key from the google maps api/geocode and create an excel workbook of the name
addresses in the same file as this file. within that sheet, the sheet name shall be called whatever you like
but should be changed in the sheet_name constant
similarily the same should be done with the address column and where your data will reside

the program will then grab each address and convert it to fit the url format and then fetch the corresponding latitude 
& longitutde to be returned
after such, then it will be outputed to a new excel file called output and saved.

'''
'''
if there is an error in retrieving the lat or long, the program
will instead put a 0,0. this is because the program doesn't account for errors on the request 
of which it responds with a 200 success, but results in an index out of bounds error 
because the return specifies the api key is wrong yet it is right.
for the few entries this problem occurs it can be done manually or programmed further to account for


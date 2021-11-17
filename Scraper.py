import requests
import openpyxl as opxl

#The vowels that can take a fada.
fadaList = ["á", "é","í","ó","ú"]

#Once we find a fada, we need to remove the fada
#Using keys, this dictionary streamlines the process
fadaDict = {
    "á":"a", 
    "é":"e",
    "í":"i",
    "ó":"o",
    "ú":"u"
}


#Determines the dialect we scrape from the server
diaDict = {
    "c":"_c.ogg",
    "u":"_u.ogg",
    "m":"_m.ogg"
}

#Allows user to choos which dialect they want.

diaChos = input("Dialect (m, c, u): ").lower()
print("Munster is reccommended \n")


#Open workbook (FILE PATH NEEDS TO BE INSERTED)
wb = opxl.load_workbook(r'')

#Set the workbook.active's name to "Sheet"
sheet = wb.active

#For each filled cell in the first column

for word in sheet["A"]:
    
    #The cell's value is read and saved as a string.
    #It is converted to lowercase for simplicity
    src = str(word.value).lower()
    
    #The word is saved in its original form, mainly to display to the user.
    srcOrigin = src
    
    #The length of the word will dictate how the string is iterated
    lenSrc = len(src)
    
    #Calculate how many fadas are in the string
    counter = 0
    for i in src:
        if i in fadaList:
            counter+=1
    
    #For each fada that's in the string, run this loop
    for i in range(counter):
        #For each index in the string
        for i in range(lenSrc):
            #If the letter from a given index has a fada in it...
            if src[i] in fadaList:
                
                #If the word has an á, that becomes a_x.
                #Formatted this way due to how the URLs work on focloir.ie
                src = src[0:i] + fadaDict[src[i]] + "_x" + src[i+1:lenSrc]
                lenSrc = len(src)
    
    #Add chosen dialect to URl
    if diaDict[diaChos] not in src:
        src = src + diaDict[diaChos]

    #Create URL 
    url = "https://www.focloir.ie/media/ei/sounds_ogg/" + src
    
    #Download file using requests module
    r = requests.get(url,headers={"User-Agent":"Mozilla/5.0"})
    
    #Checks the file size to verify it is above 0 bytes.
    #(To prevent saving null files)
    length = int(r.headers.get('Content-Length', 0))
    if length < 1:
        print(f"'{srcOrigin}' was not found on the server")
        continue
    
    #Save the file with a filename taken from the original input.
    with open(srcOrigin+diaDict[diaChos], "wb") as o:
        o.write(r.content)
        
        print(srcOrigin+diaDict[diaChos]+" saved!")



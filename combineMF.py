#!/usr/bin/env python
# coding: utf-8

# # Merge data to excel script
# 

# In[ ]:


#all modules imported
def combineMF():
    
    print('This program combines either Excel, CSV or Text files into one Excel File in the same Directory.\n')
        
    import pandas as pd
    import os
    
    print('\nPlease Enter the full folder path\n')

    def folder_Path():
        
        global folderPath
    
        folderPath = input()
        
        if os.path.isdir(folderPath)==False:
        
            print ("\nPlease Enter a valid path. You can copy folder address from status bar.")
        
            folder_Path()
        
    folder_Path() 
    
    os.chdir(folderPath)                                        # changing cwd to user input
    
    print('\nPlease Enter File Extension.\n')
    
    def file_Type():

        global fileType
        
        fileType = input()
            
    file_Type()  
    

    def files_List():
        
        global  filesList
        
        filesList = []                                              #creating the list container for full path

        for (dirPath, dirName, fileName) in os.walk(folderPath):    #iterating os.walk    
    
            for file in fileName:                                   #for each file in the os.walk, iterating over file name 
                                                                    #one by one
                if file.endswith('.'+ fileType):                    # filter criteria
                    filesList += [os.path.join(dirPath, file)]      # updating the list created previously for each file path
            continue
        
        if len(filesList)==0:
        
            print(f"\nThe file type '{fileType}' does not exist in the directory.\n")
            print("Please Enter 3 characters long valid file type without '.' dot.\n")
        
            file_Type()
        
            files_List()
    files_List()
        
    #Creating  and saving in a master file
    
    masterFile = pd.DataFrame()                                       #creating dataframe objec
    try:
        for xlfile in filesList:
     
            masterFile = masterFile.append(pd.read_excel(xlfile), ignore_index = True)  #appending each file to datafram object
        masterFile.to_excel('combined.xlsx')                                            # saving the final output in excel
                                                                                        #in the cwd
        print("\nWork Completed Successfully!\n")
            
    except:
            print (f"\nThis programs does not support the entered file type '{fileType}'.")
            print("Please try again by entering a valid file type which is 3 chraracters long.\n")
            
combineMF()
            


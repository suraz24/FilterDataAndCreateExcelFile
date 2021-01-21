from datetime import datetime
import pandas as pd

def printWelcomeMessage():
    print("========================================================================================")
    print("Hi, this application assists you to extract data from excel file for a particular date: ")
    print("The program needs the date and location of the file to be used")
    print("========================================================================================")

def main():
    try:
        printWelcomeMessage()
        
        # get all input values
        fileLocation = input("Please provide full file path for the source excel file without the extension (eg. C:/Files/myfile): ")
        userInput= input("Please enter a date in following format (yyyy-mm-dd): ")
        
        # parsed Data
        dateParseFormat = '%Y-%m-%d'
        parsedDate = datetime.strptime(userInput, dateParseFormat)

        # read and filter data
        originalFileName =  f'{fileLocation}.xlsx'
        data = pd.read_excel(originalFileName)
        if data.empty:
            raise Exception("Excel file is empty! Please check the file and try again")
        print('initial data')
        print(data)

        filter = data['Follow Up'] == parsedDate
        filteredData = data.loc[filter]
        if filteredData.empty:
            raise Exception("No records found for provided date")

        print("Filter Data is: ")
        print(filteredData)

        # create file
        print(f'Creating excel file for date for date: {parsedDate.date()}')
        destFileName = f'{parsedDate.date()}.xlsx'
        filteredData.to_excel(destFileName, index=False)

        print(f'Successfully created file called {destFileName}')

        print('Deleting copied data from original file')
        data.drop(index=data[filter].index, inplace=True, errors='ignore')

        print("left data in df: ")
        print(data)
        data.to_excel(originalFileName, index=False)
        print("Successfully updated original file")

        

    except Exception as e:
        print(e)
        print("Please run the program again and provide valid location and date")

main()
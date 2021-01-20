from datetime import datetime
import pandas as pd

def main():
    try:
        
        print("========================================================================================")
        print("Hi, this application assists you to extract data from excel file for a particular date: ")
        print("The program needs the date and location of the file to be used")
        print("========================================================================================")

        # get all input values
        fileLocation = input("Please provide full file path for the source excel file (eg. C:/Files/myfile.xlsx): ")
        userInput= input("Please enter a date in following format (yyyy-mm-dd): ")
        
        # parsed Data
        parsedDate = datetime.strptime(userInput, '%Y-%m-%d')
        print(f'Creating excel file for date for date: {parsedDate.date()}')

        # read and filter data
        data = pd.read_excel(fileLocation, skiprows=1)
        filteredData = data.loc[data['Follow Up'] == parsedDate]

        # create file
        fileName = f'{parsedDate.date()}.xlsx'
        filteredData.to_excel(fileName)

        print(f'Successfully created file called {fileName}')

    except Exception as e:
        print(e)
        print("Please run the program again and provide valid location and date")

main()
# UniStuttgart-PuppetGroup-gRPCGoogleSheets

Java project of the GoogleSheets gRPC API 

####GoogleSheets gRPC:

This is the gRPC implementation of a generic service in order to expose the main functionality of exisitng Google Sheets.

This gRPC API makes use of the exisitng Google Sheets V3 SDKs to connect to the requested Google Sheets and perform required functions.

####GoogleSheets proto  
Description of proto file being used in the gRPC API implementation of the GoogleSheets Service:

Service Name: GenericOps    

Method Name: spreadsheetList    
  No input to this method

Method Name: worksheetList   
Input to this method 

    spreadsheetId: The id of the target spreadsheeet    
    
Method Name: createWorksheet    
Input to this method 

    spreadsheetName: The name of the target spreadsheeet  
    worksheetName: The name of the worksheet to be created  
    row: The row count of the new worksheet 
    col: The coloum count of the new worksheet  
    
Method Name: updateWorksheet    
Input to this method   

    spreadsheetId: The id of the target spreadsheeet  
    oldWorksheetName: The name of the worksheet to be updated 
    newWorksheetName: The name to be updated to 
    newWorksheetRow: The new row count of the updated worksheet 
    newWorksheetCol: The new coloum count of the updated worksheet
    
Method Name: deleteWorksheet    
Input to this method 

    spreadsheetId: The id of the target spreadsheeet    
    worksheetName: The name of the worksheet to be deleted  
    
Method Name: getWorksheetContents   
Input to this method 

    spreadsheetId: The id of the target spreadsheeet    
    worksheetName: The name of the worksheet to get 
    
Method Name: fetchRowCol    
Input to this method  

    spreadsheetName: The name of the target spreadsheeet  
    worksheetName: The name of the worksheet to get 
    row: The row number to be fetched 
    col: The coloum number to be fetched  
    
Method Name: deleteRow     
Input to this method  

    spreadsheetName: The name of the target spreadsheeet  
    worksheetName: The name of the worksheet  
    row: The row number to be deleted 
    
NOTE:
In access the required spreadsheet, it would have to shared to the service account. The service account for our GoogleSheetsgRPC application is: grpcgooglesheets@ivory-plane-135618.iam.gserviceaccount.com

'=====================================================================================================================================
Option Explicit

'Variable declarations
Dim strUserName, strPass
Dim strPassWord
Dim strFlyFrom, strFlyTo
Dim strPassengerName
Dim strExecute
Dim intRowCount, intLoop
Dim intOrderNumber

'=====================================================================================================================================
'Import the test Data
DataTable.ImportSheet "C:\My_Doc\MY_UFT\Excel_Sheets\FlightAppTestDataUpdated.xlsx","FlightBookingData","Global"
intRowCount = DataTable.GlobalSheet.GetRowCount
'=====================================================================================================================================

For intLoop = 1 To intRowCount
	DataTable.GlobalSheet.SetCurrentRow(intLoop)	
	'Get the variables
	strUserName = trim(DataTable.Value("UserName","Global"))
	strPassWord = trim(DataTable.Value("PassWord","Global"))
	strFlyFrom = trim(DataTable.Value("FlyFrom","Global"))
	strFlyTo = trim(DataTable.Value("FlyTo","Global"))
	strPassengerName = trim(DataTable.Value("PassangerName","Global"))
	strExecute = trim(DataTable.Value("Execute","Global"))
	
	If ucase(strExecute) = "Y" Then
		'Launch application under test 
		Call LaunchAUT()	
		'Login
		Call Login(strUserName,strPassWord) 
 		'Entering Flight detail
 		Call EnterTicketDetails(strFlyFrom,strFlyTo,strPassengerName)
		'Fetch and place order message in datatable
        	DataTable.Value("OrderNumber","Global") = FetchOrderNumber()
	 	Call  LogOut()
             'Search and delete ticket
		Call LaunchAUT()
		Call Login(strUserName,strPassWord) 
		intOrderNumber = ReturnOrderNumber()
		Call DeleteOrder(intOrderNumber)
		Call  LogOut()
	End If
Next
'=====================================================================================================================================
'Export the results
DataTable.ExportSheet "C:\My_Doc\MY_UFT\TestResults\TestResult.xlsx","Global","FlightBookingData"
'=====================================================================================================================================

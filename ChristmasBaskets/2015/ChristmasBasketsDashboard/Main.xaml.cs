using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;

//Access Stuff
using System.Data;
using System.Data.OleDb;

//Excel Stuff
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ChristmasBasketsDashboard
{
    /// <Main>
    /// Interaction logic for Main.xaml
    /// </Main>
    public partial class Main : System.Windows.Window
    {
        //Define variables
        public static ChristmasBasketsAccessDatabase mChristmasBasketsAccessDatabase;
        public static string mSelectedYear;
        public static string mSelectedOrganization;
        public static bool mQuitClientImport;
        public static bool mQuitDelivererImport;
        public const int mNumberOfYearStatusStates = 12;
        public const int mMaxNumberGoogleMapsWaypoints = 8;
        public enum mSelectedYearStatusEnum { Step_1_Year_Created_In_Database = 0, Step_2_Clients_Imported, Step_2_a_Check_For_Client_Duplicates, Step_3_Green_Cards_Generated, Step_4_Deliverers_Imported, Step_5_Clients_Assigned_To_Deliverers, Step_6_Generated_Deliverer_Maps, Step_7_Day_Of_Event, Step_7_a_Generate_Unassigned_Clients_Map, Step_7_b_Generate_Client_Lists, Step_7_c_Generate_Food_Signs, Step_7_d_Generate_Box_Labels };
        public static bool [] mSelectedYearStatus = new bool[mNumberOfYearStatusStates] {false, false, false, false, false, false, false, false, false, false, false, false};

        //Define methods
        /// <Main>
        /// Constructor
        /// </Main>
        public Main()
        {
            InitializeComponent();

            //Initialize mChristmasBasketsAccessDatabase
            mChristmasBasketsAccessDatabase = null;

            //Initialize mQuitClientImport
            mQuitClientImport = false;

            //Initialize mQuitDelivererImport
            mQuitDelivererImport = false;

            //Initialize mSelectedYear
            mSelectedYear = "NONE";

            //Initialize SelectedYearLabel
            SelectedYearLabel.Content = mSelectedYear;

            //Initialize StampClientsWithSelectedYearButton
            string year = mSelectedYear.Replace("Year_", "");
            StampClientsWithSelectedYearButton.Content = "Set Client's Year__Last__Delivered__To = " + year;

            //Initialize StampDeliverersWithSelectedYearButton
            StampDeliverersWithSelectedYearButton.Content = "Set Deliverer's Year__Last__Delivered = " + year;
        }

        /// <SelectChristmasBasketsDatabase>
        /// Open a 2003 (.mdb) or 2007 (.acccb) Access database file
        /// </SelectChristmasBasketsDatabase>
        private void SelectChristmasBasketsDatabase()
        {
            //Define local variables
            string SelectedChristmasBasketsAccessDatabasePath = "";
            string[] SplitSelectedChristmasBasketsAccessDatabasePath;
            string DataBaseFileNameExtension = "";

            //Open a File Dialog Box to the user
            OpenFileDialog openFileDialogBox = new OpenFileDialog();
            openFileDialogBox.Title = "Find Christmas Baskets Master Database";

            //If the user selects a file and clicks OK...
            if (openFileDialogBox.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //Store the Access 2007 database path
                SelectedChristmasBasketsAccessDatabasePath = openFileDialogBox.FileName;

                //Determine the selected database's file name extension
                SplitSelectedChristmasBasketsAccessDatabasePath = SelectedChristmasBasketsAccessDatabasePath.Split('.');
                DataBaseFileNameExtension = SplitSelectedChristmasBasketsAccessDatabasePath.Last().ToString();

                //Determine what type of database it is 2003 or 2007
                if (DataBaseFileNameExtension == "mdb")
                {
                    //Access 2003 Database

                    //Change the Indicator Color
                    OpenDatabaseIndicator.Fill = new SolidColorBrush(Colors.LightGreen);

                    //Create a new ChristmasBasketsAccess2003Database object and initialize it
                    mChristmasBasketsAccessDatabase = new ChristmasBasketsAccess2003Database(SelectedChristmasBasketsAccessDatabasePath);

                    //Open the database
                    mChristmasBasketsAccessDatabase.OpenChristmasBasketsDatabase();
                }
                else if (DataBaseFileNameExtension == "accdb")
                {
                    //Access 2007 Database

                    //Change the Indicator Color
                    OpenDatabaseIndicator.Fill = new SolidColorBrush(Colors.LightGreen);

                    //Create a new ChristmasBasketsAccess2007Database object and initialize it
                    mChristmasBasketsAccessDatabase = new ChristmasBasketsAccess2007Database(SelectedChristmasBasketsAccessDatabasePath);

                    //Open the database
                    mChristmasBasketsAccessDatabase.OpenChristmasBasketsDatabase();
                }
                else
                {
                    //Change the Indicator Color
                    OpenDatabaseIndicator.Fill = new SolidColorBrush(Colors.Red);

                    //Database type not supported
                    System.Windows.MessageBox.Show("Database type ." + DataBaseFileNameExtension + " not supported", "Invalid database type selected");
                }
            }
        }

        /// <OpenDatabase_Click>
        /// Event Handler For OpenDatabase_Click
        /// </OpenDatabase_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenDatabase_Click(object sender, RoutedEventArgs e)
        {
            //Select the Christmas Baskets Database
            SelectChristmasBasketsDatabase();
        }

        /// <UpdateYearSelectedStatusIndicators>
        /// Update all Year Selected Status Indicators
        /// </UpdateYearSelectedStatusIndicators>
        private void UpdateYearSelectedStatusIndicators()
        {
            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database] == true)
            {
                Step_1_Year_Created_In_Database_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_1_Year_Created_In_Database_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_Clients_Imported] == true)
            {
                Step_2_Clients_Imported_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_2_Clients_Imported_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_a_Check_For_Client_Duplicates] == true)
            {
                Step_2_a_Check_For_Client_Duplicates_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_2_a_Check_For_Client_Duplicates_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_3_Green_Cards_Generated] == true)
            {
                Step_3_Green_Cards_Generated_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_3_Green_Cards_Generated_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_4_Deliverers_Imported] == true)
            {
                Step_4_Deliverers_Imported_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_4_Deliverers_Imported_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_5_Clients_Assigned_To_Deliverers] == true)
            {
                Step_5_Clients_Assigned_To_Deliverers_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_5_Clients_Assigned_To_Deliverers_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_6_Generated_Deliverer_Maps] == true)
            {
                Step_6_Generated_Deliverer_Maps_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_6_Generated_Deliverer_Maps_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map] == true &&
                mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists] == true &&
                mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs] == true &&
                mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels] == true)
            {
                Step_7_Day_Of_Event_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_7_Day_Of_Event_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map] == true)
            {
                Step_7_a_Generate_Unassigned_Clients_Map_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_7_a_Generate_Unassigned_Clients_Map_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists] == true)
            {
                Step_7_b_Generate_Client_Lists_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_7_b_Generate_Client_Lists_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs] == true)
            {
                Step_7_c_Generate_Food_Signs_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_7_c_Generate_Food_Signs_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }

            if (mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels] == true)
            {
                Step_7_d_Generate_Box_Labels_Indicator.Fill = new SolidColorBrush(Colors.LightGreen);
            }
            else
            {
                Step_7_d_Generate_Box_Labels_Indicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
            }
        }

        /// <SelectYearButton_Click>
        /// Event handler for SelectYearButton_Click
        /// </SelectYearButton_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectYearButton_Click(object sender, RoutedEventArgs e)
        {
            //Make sure we have an Access Database selected
            if (mChristmasBasketsAccessDatabase != null)
            {
                WindowSelectYear dialogSelectYear = new WindowSelectYear();
                dialogSelectYear.ShowDialog();
                UpdateYearSelectedStatusIndicators();
            }
            else
            {
                //Not connected to a database
                System.Windows.MessageBox.Show("Not connected to an Access Christmas Baskets Database", "Not connected to an Access Christmas Baskets Database");
            }
        }

        /// <CloseDatabase_Click>
        /// Event handler for CloseDatabase_Click
        /// </CloseDatabase_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CloseDatabase_Click(object sender, RoutedEventArgs e)
        {
            //Close the Access database if it is open
            if (mChristmasBasketsAccessDatabase != null)
            {
                //Close the Access Database
                mChristmasBasketsAccessDatabase.CloseChristmasBasketsDatabase();

                //Get rid of mChristmasBasketsAccessDatabase object
                mChristmasBasketsAccessDatabase = null;

                //Change the Indicator Colors
                OpenDatabaseIndicator.Fill = new SolidColorBrush(Colors.CornflowerBlue);
                Step_1_Year_Created_In_Database_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_2_Clients_Imported_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_2_a_Check_For_Client_Duplicates_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_3_Green_Cards_Generated_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_4_Deliverers_Imported_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_5_Clients_Assigned_To_Deliverers_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_6_Generated_Deliverer_Maps_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_7_Day_Of_Event_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_7_a_Generate_Unassigned_Clients_Map_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_7_b_Generate_Client_Lists_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_7_c_Generate_Food_Signs_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
                Step_7_d_Generate_Box_Labels_Indicator.Fill = new SolidColorBrush(Colors.LightGray);
            }

            //Update Main.mSelectedYearStatus
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_Clients_Imported] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_3_Green_Cards_Generated] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_4_Deliverers_Imported] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_5_Clients_Assigned_To_Deliverers] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_6_Generated_Deliverer_Maps] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_Day_Of_Event] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels] = false;

            //Update mSelected Year
            mSelectedYear = "NONE";

            //Update SelectedYearLabel
            Window_MouseEnter(null, null);
        }

        /// <Window_MouseEnter>
        /// Event Handler for Window_MouseEnter
        /// </Window_MouseEnter>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            //Update SelectedYearLabel
            SelectedYearLabel.Content = mSelectedYear;

            //Initialize StampClientsWithSelectedYearButton
            string year = mSelectedYear.Replace("Year_", "");
            StampClientsWithSelectedYearButton.Content = "Set Client's Year__Last__Delivered__To = " + year;

            //Initialize StampDeliverersWithSelectedYearButton
            StampDeliverersWithSelectedYearButton.Content = "Set Deliverer's Year__Last__Delivered = " + year;
        }

        /// <ExportClientsToExcel_Click>
        /// Event Handler for ExportClientsToExcel_Click
        /// </ExportClientsToExcel_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportClientsToExcel_Click(object sender, RoutedEventArgs e)
        {
            //Define local variables
            string worksheetName = mSelectedOrganization + " " + mSelectedYear.Replace("Year_", "") + " Clients";
            DataSet dataToExport = new DataSet();
            dataToExport.Tables.Add();
            bool columnsImported = false;

            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            if (mSelectedOrganization == "NONE")
            {
                System.Windows.MessageBox.Show("No organization selected for Excel Export");
                return;
            }

            //Create the select selectByClient_IDQueryText
            string selectByClient_IDQueryText = "SELECT Client_ID FROM " + mSelectedYear;

            //Create the selectedYearDataSet
            DataSet selectedYearDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClient_IDQueryText, mSelectedYear);

            //Process each record in the selectedYear table
            foreach (DataRow dataRow in selectedYearDataSet.Tables[0].Rows)
            {
                //Create the selectByClientIDAndOrganizationQueryText
                string selectByClientIDAndOrganizationQueryText = "SELECT Client_ID, Last_Name, First_Name, Middle_Name, Title, Address_Number, Street_Address, City, Zipcode, Phone, Organization FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString() + " AND Organization = '" + mSelectedOrganization + "'";

                if (mSelectedOrganization == "LOA and RCSS")
                {
                    //LOA and RCSS
                    selectByClientIDAndOrganizationQueryText = "SELECT Client_ID, Last_Name, First_Name, Middle_Name, Title, Address_Number, Street_Address, City, Zipcode, Phone, Organization FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString() + " AND (Organization = 'LOA' OR Organization = 'RCSS')";
                }
                else
                {
                    //Single organization
                    selectByClientIDAndOrganizationQueryText = "SELECT Client_ID, Last_Name, First_Name, Middle_Name, Title, Address_Number, Street_Address, City, Zipcode, Phone, Organization FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString() + " AND Organization = '" + mSelectedOrganization + "'";
                }

                //Perform the selectByClientIDAndOrganizationQuery and store the results in a Data Table
                System.Data.DataSet selectByClientIDAndOrganizationDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClientIDAndOrganizationQueryText, "Clients");

                //Check to see we got results back
                if (selectByClientIDAndOrganizationDataSet != null)
                {
                    //Add columns if not already added
                    if (columnsImported == false)
                    {
                        //Add each column requested from the query
                        foreach (DataColumn dataColumn in selectByClientIDAndOrganizationDataSet.Tables[0].Columns)
                        {
                            dataToExport.Tables[0].Columns.Add(dataColumn.ColumnName);
                        }

                        //Update columnsImported
                        columnsImported = true;
                    }

                    //Copy the data row
                    DataRow record = selectByClientIDAndOrganizationDataSet.Tables[0].Rows[0];
                    dataToExport.Tables[0].ImportRow(record);
                }
            }

            //Export the query to Excel
            ExcelMethods.ExportToExcel(dataToExport, worksheetName);
        }

        /// <SelectOrganizationComboBox_SelectionChanged>
        /// Event handler for SelectOrganizationComboBox_SelectionChanged
        /// </SelectOrganizationComboBox_SelectionChanged>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectOrganizationComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //Check for Select Organization field
            if (SelectOrganizationComboBox.SelectedIndex == 0)
            {
                //Store the selected organization
                mSelectedOrganization = "NONE";
            }
            else if (SelectOrganizationComboBox.SelectedIndex == 3)
            {
                mSelectedOrganization = "LOA and RCSS";
            }
            else
            {
                //Get the curent selected item from the SelectOrganizationComboBox and parse it by spaces
                string[] parsedSelectedOrganization = SelectOrganizationComboBox.SelectedItem.ToString().Split(' ');

                //Store the selected organization
                mSelectedOrganization = parsedSelectedOrganization[1];
            }
        }

        /// <ImportSelectedYearClientsFromExcelButton_Click>
        /// Event Handler for ImportSelectedYearClientsFromExcelButton_Click
        /// </ImportSelectedYearClientsFromExcelButton_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportSelectedYearClientsFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if(mSelectedYear != "NONE")
            {
                string selectQueryText = "SELECT * FROM [Sheet1$]";

                //Open an Excel Workbook to Import Clients
                DataSet clientsToImport = ExcelMethods.OpenExcelWorkbookAndExtractWorksheetInformation(selectQueryText);

                //Temp stats variables
                int fullhit = 0;
                int existsinselectedyear = 0;
                int duplicatesinselectedyeartable = 0;
                int doesnotexitinselectedyear = 0;
                int duplicatesinclienttable = 0;
                int nameandorghit = 0;
                int nohit = 0;

                if (clientsToImport != null)
                {
                    foreach (DataRow record in clientsToImport.Tables[0].Rows)
                    {
                        //Process each client from the excel spreadsheet
                        string clientSelectQuery = "SELECT * FROM Clients WHERE Last_Name = '" + record["Last_Name"].ToString().Replace("'", "''") + "' AND " +
                                                   "First_Name = '" + record["First_Name"].ToString().Replace("'", "''") + "' AND " +
                                                   "Middle_Name = '" + record["Middle_Name"].ToString().Replace("'", "''") + "' AND " +
                                                   "Title = '" + record["Title"].ToString().Replace("'", "''") + "' AND " +
                                                   "Address_Number = '" + record["Address_Number"].ToString().Replace("'", "''") + "' AND " +
                                                   "Street_Address = '" + record["Street_Address"].ToString().Replace("'", "''") + "' AND " +
                                                   "City = '" + record["City"].ToString().Replace("'", "''") + "' AND " +
                                                   "Zipcode = '" + record["Zipcode"].ToString().Replace("'", "''") + "' AND " +
                                                   "Phone = '" + record["Phone"].ToString().Replace("'", "''") + "'";

                        DataSet Clients = mChristmasBasketsAccessDatabase.PerformSelectQuery(clientSelectQuery, "Clients");

                        if (Clients != null)
                        {
                            //Client Already Exists in Clients Database

                            //See if only one instance of the client was found in the Clients table
                            if (Clients.Tables[0].Rows.Count == 1)
                            {
                                //Only 1 record was found in teh Clients table

                                fullhit++;
                                //See if the Client Already Exists in the mSelected year Table
                                string selectedYearClientSelectQuery = "SELECT * FROM " + mSelectedYear + " WHERE Client_ID = " + Clients.Tables[0].Rows[0]["Client_ID"];

                                DataSet selectedYearClient = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectedYearClientSelectQuery, mSelectedYear);

                                if (selectedYearClient != null)
                                {
                                    //Make sure there was only 1 record returned from the selected year table
                                    if (selectedYearClient.Tables[0].Rows.Count == 1)
                                    {
                                        //Only 1 record was found in the selected year table
                                        existsinselectedyear++;
                                    }
                                    else
                                    {
                                        //Duplicate records found in the selected year table
                                        duplicatesinselectedyeartable++;

                                        //Remove Duplicates from the mSelectedYear Table
                                        for (int i = 1; i < selectedYearClient.Tables[0].Rows.Count; i++)
                                        {
                                            string deleteCommand = "DELETE FROM " + mSelectedYear + " WHERE Box_Number = " + selectedYearClient.Tables[0].Rows[i]["Box_Number"].ToString() + " AND Client_ID = " + selectedYearClient.Tables[0].Rows[i]["Client_ID"].ToString();
                                            mChristmasBasketsAccessDatabase.ExecuteNonQuery(deleteCommand);
                                        }
                                    }
                                }
                                else
                                {
                                    //Client does not exist in the selected year table
                                    doesnotexitinselectedyear++;

                                    //Insert the client to the mSelectedYear table
                                    string insertCommand = "INSERT INTO " + mSelectedYear + " (Box_Number,Client_ID) VALUES (-1," + Clients.Tables[0].Rows[0]["Client_ID"] + ")";
                                    mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommand);
                                }

                                //Object clean up
                                selectedYearClient = null;
                            }
                            else
                            {
                                //Duplicates clients exist in the Client table from all field
                                //search - very unlikely

                                ImportClientsFromExcelHelper excelImportHelper = new ImportClientsFromExcelHelper();
                                excelImportHelper.SetDatabaseClientsTable(Clients.Tables[0]);
                                excelImportHelper.SetMode(ClientMode.DatabaseClientsTableDuplicates);
                                excelImportHelper.SetExcelClientsTable(record, Clients.Tables[0].Columns);
                                excelImportHelper.BindData();
                                excelImportHelper.ShowDialog();
                                excelImportHelper.UnBindData();

                                duplicatesinclienttable++;
                            }
                        }
                        else
                        {
                            //Either client does not exist or a field miss-match

                            //See if we can get a first name, last name, and organization match on the client
                            //trying to be imported from the excel spreadsheet
                            string selectQuery = "SELECT * FROM Clients WHERE Last_Name = '" + record["Last_Name"].ToString().Replace("'", "''") + "' AND " +
                                                   "First_Name = '" + record["First_Name"].ToString().Replace("'", "''") + "' AND " +
                                                   "Organization = '" + record["Organization"].ToString().Replace("'", "''") + "'";

                            DataSet clientDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectQuery, "Clients");

                            if (clientDataSet != null)
                            {
                                //Hit in Clients Table by First_Name, Last_Name, and Organization
                                //System.Windows.MessageBox.Show(record["Last_Name"].ToString() + ", " + record["First_Name"].ToString() + " Hit in Clients Table by First_Name, Last_Name, and Organization");
                                nameandorghit++;

                                //See if there was only 1 record returned from the Clients table
                                if (clientDataSet.Tables[0].Rows.Count == 1)
                                {
                                    //Only 1 record found in the Clients Table

                                    //////////////////////////////////////////////
                                    //HANDLE CLIENT INFO UPDATE IN CLIENTS TABLE//
                                    //////////////////////////////////////////////

                                    ///////////////////////////////////////////////////////////////
                                    //HANDLE ADDING CLIENT, WHO'S INFORMATION WAS UPDATED, TO THE//
                                    //SELECTED YEAR TABLE                                        //
                                    ///////////////////////////////////////////////////////////////

                                    ImportClientsFromExcelHelper excelImportHelper = new ImportClientsFromExcelHelper();
                                    excelImportHelper.SetDatabaseClientsTable(clientDataSet.Tables[0]);
                                    excelImportHelper.SetMode(ClientMode.DatabaseClientsTableUpdateClient);
                                    excelImportHelper.SetExcelClientsTable(record, clientDataSet.Tables[0].Columns);
                                    excelImportHelper.BindData();
                                    excelImportHelper.ShowDialog();
                                    excelImportHelper.UnBindData();

                                }
                                else
                                {
                                    //Duplicate records found in the Clients Table

                                    /////////////////////
                                    //HANDLE DUPLICATES//
                                    /////////////////////

                                    //////////////////////////////////////////////
                                    //HANDLE CLIENT INFO UPDATE IN CLIENTS TABLE//
                                    //////////////////////////////////////////////

                                    ///////////////////////////////////////////////////////////////
                                    //HANDLE ADDING CLIENT, WHO'S INFORMATION WAS UPDATED, TO THE//
                                    //SELECTED YEAR TABLE                                        //
                                    ///////////////////////////////////////////////////////////////

                                    ImportClientsFromExcelHelper excelImportHelper = new ImportClientsFromExcelHelper();
                                    excelImportHelper.SetDatabaseClientsTable(clientDataSet.Tables[0]);
                                    excelImportHelper.SetMode(ClientMode.DatabaseClientsTableDuplicates);
                                    excelImportHelper.SetExcelClientsTable(record, clientDataSet.Tables[0].Columns);
                                    excelImportHelper.BindData();
                                    excelImportHelper.ShowDialog();
                                    excelImportHelper.UnBindData();

                                    duplicatesinclienttable++;
                                }

                            }
                            else
                            {
                                //Still no hit
                                //System.Windows.MessageBox.Show(record["Last_Name"].ToString() + ", " + record["First_Name"].ToString() + " Still No Hit in Clients Table");
                                nohit++;

                                System.Data.DataTable temp = new System.Data.DataTable();
                                temp.Columns.Add("Client_ID");
                                temp.Columns.Add("Last_Name");
                                temp.Columns.Add("First_Name");
                                temp.Columns.Add("Middle_Name");
                                temp.Columns.Add("Title");
                                temp.Columns.Add("Address_Number");
                                temp.Columns.Add("Street_Address");
                                temp.Columns.Add("City");
                                temp.Columns.Add("Zipcode");
                                temp.Columns.Add("Phone");
                                temp.Columns.Add("Organization");
                                temp.Columns.Add("Directions");
                                temp.Columns.Add("Instructions");
                                temp.Columns.Add("Deliverer_ID");
                                temp.Columns.Add("Year_Last_Delivered_To");

                                ImportClientsFromExcelHelper excelImportHelper = new ImportClientsFromExcelHelper();
                                excelImportHelper.SetDatabaseClientsTable(temp);
                                excelImportHelper.SetMode(ClientMode.DatabaseClientsTableAddClient);
                                excelImportHelper.SetExcelClientsTable(record, temp.Columns);
                                excelImportHelper.BindData();
                                excelImportHelper.ShowDialog();
                                excelImportHelper.UnBindData();

                            }

                            //Object clean up
                            clientDataSet = null;
                        }

                        //Object clean up
                        Clients = null;

                        //Quit the Client Import Process
                        if (Main.mQuitClientImport)
                        {
                            //Update Main.mQuitClientImport
                            Main.mQuitClientImport = false;

                            //Quit the For loop
                            break;
                        }
                    }
                }

                //Temp Stats messagebox
                System.Windows.MessageBox.Show("Full hit = " + fullhit.ToString() +
                                                "\nExists In Selected Year Table = " + existsinselectedyear.ToString() +
                                                "\nDuplicates In Selected Year Table = " + duplicatesinselectedyeartable.ToString() +
                                                "\nDoes not exist In Selected Year Table = " + doesnotexitinselectedyear.ToString() +
                                                "\nDuplicates In Clients Table = " + duplicatesinclienttable.ToString() +
                                                "\nName and Organization hit = " + nameandorghit.ToString() +
                                                "\nNo hit = " + nohit.ToString());

                //Object clean up
                clientsToImport = null;

                //Update the Box Numbers
                SetSelectedYearBoxNumbers();
            }
            else
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
            }

        }

        /// <SetSelectedYearBoxNumbersButton_Click>
        /// Event handler for Set Selected Year Box Numbers Button_Click
        /// </SetSelectedYearBoxNumbersButton_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetSelectedYearBoxNumbersButton_Click(object sender, RoutedEventArgs e)
        {
            SetSelectedYearBoxNumbers();
        }

        /// <SetSelectedYearBoxNumbers>
        /// Set the Selected Year's Box Numbers
        /// </SetSelectedYearBoxNumbers>
        public void SetSelectedYearBoxNumbers()
        {
            if (mSelectedYear != "NONE")
            {
                //Get the mSelectedYear list of clients
                string selectQuery = "SELECT * FROM " + mSelectedYear;
                DataSet Clients = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectQuery, mSelectedYear);

                int boxNumber = 1;

                //Update the Box_Number field for all clients in mSelectedYear
                for (int i = 0; i < Clients.Tables[0].Rows.Count; i++)
                {
                    string updateCommand = "UPDATE " + mSelectedYear + " SET Box_Number = " + boxNumber.ToString() + " WHERE Client_ID = " + Clients.Tables[0].Rows[i]["Client_ID"].ToString();
                    mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCommand);
                    boxNumber++;
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
            }
        }

        /// <GenerateWhiteCardWithMap>
        /// Generate a white card with a map for a single client
        /// </GenerateWhiteCardWithMap>
        private void GenerateWhiteCardWithMap(string fileName, string boxNumber, DataRow clientInfo, DataRow delivererInfo)
        {
            //Define local variables
            TextWriter textWriter = new StreamWriter(fileName);

            //Write the header part of the .htm page
            textWriter.WriteLine("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">");
            textWriter.WriteLine("<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:v=\"urn:schemas-microsoft-com:vml\">");
            textWriter.WriteLine("   <head>");
            textWriter.WriteLine("      <meta http-equiv=\"content-type\\\" content=\"text/html; charset=UTF-8\"/>");

            //Deliverer, Client, and Box Number Data goes here
            if (delivererInfo != null)
            {
                textWriter.WriteLine("      <title>Deliverer:  " + delivererInfo["First_Name"] + " " + delivererInfo["Last_Name"] + " (" + delivererInfo["Deliverer_ID"] + ") - Client:  " + clientInfo["First_Name"] + " " + clientInfo["Last_Name"] + " (" + clientInfo["Client_ID"] + ") - Box:  " + boxNumber + "</title>");
            }
            else
            {
                textWriter.WriteLine("      <title>Deliverer:  UNASSIGNED - Client:  " + clientInfo["First_Name"] + " " + clientInfo["Last_Name"] + " (" + clientInfo["Client_ID"] + ") - Box:  " + boxNumber + "</title>");
            }

            //Write the hava script part of the .htm page
            textWriter.WriteLine("      <script src=\"http://maps.google.com/maps?file=api&amp;v=2.x&amp;key=ABQIAAAAzr2EBOXUKnm_jVnk0OJI7xSosDVG8KKPE1-m51RBrvYughuyMxQ-i1QfUnH94QxWIa6N4U6MouMmBA\" type=\"text/javascript\"></script>");
            textWriter.WriteLine("      <script type=\"text/javascript\">");
            textWriter.WriteLine("");
            textWriter.WriteLine("      var map = null;");
            textWriter.WriteLine("      var geocoder = null;");

            //Client address goes here
            textWriter.WriteLine("      var address = \"" + clientInfo["Address_Number"] + " " + clientInfo["Street_Address"] + " " + clientInfo["City"] + ", VA  " + clientInfo["Zipcode"] + "\";");

            textWriter.WriteLine("");
            textWriter.WriteLine("      function initialize()");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("         if (GBrowserIsCompatible())");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("             //Create a new map object");
            textWriter.WriteLine("             map = new GMap2(document.getElementById(\"map_canvas\"));");
            textWriter.WriteLine("");
            textWriter.WriteLine("             //Create a new geocoder object");
            textWriter.WriteLine("		       geocoder = new GClientGeocoder();");
            textWriter.WriteLine("");
            textWriter.WriteLine("		       //Show the address");
            textWriter.WriteLine("		       showAddress(address);");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("");
            textWriter.WriteLine("      function showAddress(address)");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("         if (geocoder) {");
            textWriter.WriteLine("             geocoder.getLatLng(");
            textWriter.WriteLine("                   address,");
            textWriter.WriteLine("                   function(point) {");
            textWriter.WriteLine("                      if (!point) {");
            textWriter.WriteLine("                         alert(address + \" not found\");");
            textWriter.WriteLine("                      } else {");
            textWriter.WriteLine("                         //Center the map");
            textWriter.WriteLine("                         map.setCenter(point, 15);");
            textWriter.WriteLine("");
            textWriter.WriteLine("                         //Create a new marker");
            textWriter.WriteLine("                         var marker = new GMarker(point);");
            textWriter.WriteLine("");
            textWriter.WriteLine("                         //Add the marker to the map");
            textWriter.WriteLine("                         map.addOverlay(marker);");
            textWriter.WriteLine("");
            textWriter.WriteLine("                         //Show an info balloon on the marker");
            textWriter.WriteLine("                         marker.openInfoWindowHtml(address);");
            textWriter.WriteLine("                   }");
            textWriter.WriteLine("                }");
            textWriter.WriteLine("             );");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("      </script>");
            textWriter.WriteLine("   </head>");
            textWriter.WriteLine("");

            //Write the body part of the .htm page
            textWriter.WriteLine("   <body onload=\"initialize()\" onunload=\"GUnload()\">");
            textWriter.WriteLine("      <form onload=\"initialize()\">");
            textWriter.WriteLine("");
            textWriter.WriteLine("         <!-- General Instructions -->");
            textWriter.WriteLine("         <h2>General Instructions:</h2>");
            textWriter.WriteLine("         <p>If the box cannot be delivered, call Dick Stanfield at 540-353-7977 for assistance.  If there is still");
            textWriter.WriteLine("            a problem, please return the box to Timber Truss with this sheet by NOON TODAY.</p>");
            textWriter.WriteLine("");

            //Deliverer Data goes here
            textWriter.WriteLine("         <!-- Deliverer -->");

            if (delivererInfo != null)
            {
                textWriter.WriteLine("         <h2>Deliverer:   " + delivererInfo["First_Name"] + " " + delivererInfo["Last_Name"] + "</h2>");
            }
            else
            {
                textWriter.WriteLine("         <h2>Deliverer:   UNASSIGNED</h2>");
            }
            textWriter.WriteLine("");

            //Box Number Data goes here
            textWriter.WriteLine("         <!-- Box Number -->");
            textWriter.WriteLine("         <h2>Box Number:  " + boxNumber + "</h2>");
            textWriter.WriteLine("");

            //Client Table
            textWriter.WriteLine("         <!-- Client Information Table-->");
            textWriter.WriteLine("         <h2 align = \"center\">Client Information</h2>");
            textWriter.WriteLine("         <table align = \"center\">");
            textWriter.WriteLine("            <!-- Table Headings -->");
            textWriter.WriteLine("            <tr>");
            textWriter.WriteLine("               <th align = \"left\" width = \"70\">Client ID</th>");
            textWriter.WriteLine("               <th align = \"left\" width = \"325\">Client Name</th>");
            textWriter.WriteLine("               <th align = \"left\" width = \"325\">Street Address</th>");
            textWriter.WriteLine("               <th align = \"left\" width = \"150\">City</th>");
            textWriter.WriteLine("               <th align = \"left\" width = \"50\">Zip Code</th>");
            textWriter.WriteLine("               <th align = \"left\" width = \"70\">Phone</th>");
            textWriter.WriteLine("               <th align = \"left\" width = \"80\">Organization</th>");
            textWriter.WriteLine("            </tr>");
            textWriter.WriteLine("            <!-- Client Data -->");
            textWriter.WriteLine("            <tr>");

            //Client Data goes here
            textWriter.WriteLine("               <td align = \"left\" width = \"70\">" + clientInfo["Client_ID"] + "</td>");
            textWriter.WriteLine("               <td align = \"left\" width = \"325\">" + clientInfo["First_Name"] + " " + clientInfo["Last_Name"] + "</td>");
            textWriter.WriteLine("               <td align = \"left\" width = \"325\">" + clientInfo["Address_Number"] + " " + clientInfo["Street_Address"] + "</td>");
            textWriter.WriteLine("               <td align = \"left\" width = \"150\">" + clientInfo["City"] + "</td>");
            textWriter.WriteLine("               <td align = \"left\" width = \"50\">" + clientInfo["Zipcode"] + "</td>");
            textWriter.WriteLine("               <td align = \"left\" width = \"70\">" + clientInfo["Phone"] + "</td>");
            textWriter.WriteLine("               <td align = \"left\" width = \"80\">" + clientInfo["Organization"] + "</td>");
            textWriter.WriteLine("            </tr>");
            textWriter.WriteLine("         </table>");
            textWriter.WriteLine("");
            textWriter.WriteLine("         <!-- Client Specific Notes -->");
            textWriter.WriteLine("         <h3>Client Specific Notes:</h3>");

            //Client Data goes here
            textWriter.WriteLine("         <p>" + clientInfo["Directions"] + "  " + clientInfo["Instructions"] + "</p>");
            textWriter.WriteLine("         <!-- Google Map -->");
            textWriter.WriteLine("         <h3>Map:</h3>");
            textWriter.WriteLine("         <div id=\"map_canvas\" style=\"width: 1000px; height: 800px\"></div>");
            //textWriter.WriteLine("");
            //textWriter.WriteLine("         <!-- White Space -->");
            //textWriter.WriteLine("");
            //textWriter.WriteLine("         <!-- Deliverer Notes -->");
            //textWriter.WriteLine("         <h2 align=\"center\">DO NOT LEAVE THIS SHEET WITH CLIENTS</h2>");
            textWriter.WriteLine("      </form>");
            textWriter.WriteLine("   </body>");
            textWriter.WriteLine("</html>");

            //Close the file stream
            textWriter.Close();
        }

        private void ExcelListofDeliverersForEventButtons_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }
            //Define local variables
            string selectedYear = mSelectedYear.Replace("Year_", "");
            string worksheetName = selectedYear + " Deliverers";
            DataSet dataToExport = new DataSet();
            dataToExport.Tables.Add();
            bool columnsImported = false;

            //Create the select selectByClient_IDQueryText
            string selectByDeliverer_IDQueryText = "SELECT * FROM " + mSelectedYear + "_Deliverers";

            //Create the selectedYearDataSet
            DataSet selectedYearDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByDeliverer_IDQueryText, mSelectedYear);

            //Process each record in the selectedYear table
            foreach (DataRow dataRow in selectedYearDataSet.Tables[0].Rows)
            {
                //Add columns if not already added
                if (columnsImported == false)
                {
                    //Add each column requested from the query
                    foreach (DataColumn dataColumn in selectedYearDataSet.Tables[0].Columns)
                    {
                        dataToExport.Tables[0].Columns.Add(dataColumn.ColumnName);
                    }

                    //Update columnsImported
                    columnsImported = true;
                }

                //Copy the data row
                dataToExport.Tables[0].ImportRow(dataRow);
            }

            //Export the query to Excel
            ExcelMethods.ExportToExcel(dataToExport, worksheetName);
        }

        private void ExcelListofClientsForEventButtons_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }
            //Define local variables
            string worksheetName = mSelectedYear.Replace("Year_", "") + " Clients";
            DataSet dataToExport = new DataSet();
            dataToExport.Tables.Add();
            bool columnsImported = false;
            int currentRow = 0;

            //Create the select selectByClient_IDQueryText
            string selectByClient_IDQueryText = "SELECT * FROM " + mSelectedYear;

            //Create the selectedYearDataSet
            DataSet selectedYearDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClient_IDQueryText, mSelectedYear);

            //Process each record in the selectedYear table
            foreach (DataRow dataRow in selectedYearDataSet.Tables[0].Rows)
            {
                //Create the selectByClientIDAndOrganizationQueryText
                string selectByClientIDAndOrganizationQueryText = "SELECT Client_ID, Last_Name, First_Name, Middle_Name, Title, Address_Number, Street_Address, City, Zipcode, Phone, Directions, Deliverer_ID, Organization FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString();

                //Perform the selectByClientIDAndOrganizationQuery and store the results in a Data Table
                System.Data.DataSet selectByClientIDAndOrganizationDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClientIDAndOrganizationQueryText, "Clients");

                //Check to see we got results back
                if (selectByClientIDAndOrganizationDataSet != null)
                {
                    //Add columns if not already added
                    if (columnsImported == false)
                    {
                        ///////////////////////////////////////////////////////////////////////////////////////
                        ///////////////////////////////////////////////////////////////////////////////////////
                        ///////////////////////         COME BACK HERE JERK             ///////////////////////
                        ///////////////////////////////////////////////////////////////////////////////////////
                        ///////////////////////////////////////////////////////////////////////////////////////

                        //IMPROVEMENT - Add Deliverer Column to Client Row for DAY of Event Spreadsheet
                        //Code below shows where we previously added Box_Number Column, repace with Deliverer Column
                        //Add Box_Number column
                        //dataToExport.Tables[0].Columns.Add("Box_Number");

                        //Add each column requested from the query
                        foreach (DataColumn dataColumn in selectByClientIDAndOrganizationDataSet.Tables[0].Columns)
                        {
                            dataToExport.Tables[0].Columns.Add(dataColumn.ColumnName);
                        }

                        //Update columnsImported
                        columnsImported = true;
                    }

                    //Copy the data row
                    DataRow record = selectByClientIDAndOrganizationDataSet.Tables[0].Rows[0];
                    dataToExport.Tables[0].ImportRow(record);
                   
                    ///////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////         COME BACK HERE JERK             ///////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////
                    ///////////////////////////////////////////////////////////////////////////////////////
                    
                    //IMPROVEMENT - Add Deliverer First and Last Name to Client Row for DAY of Event Spreadsheet
                    //Code below shows where we previously added Box_Number Info, repace with Deliverer Info
                    //dataToExport.Tables[0].Rows[currentRow]["Box_Number"] = dataRow["Box_Number"];
                    currentRow++;
                }
            }

            //Export the query to Excel
            ExcelMethods.ExportToExcel(dataToExport, worksheetName);
        }

        /// <ImportDeliverersFromExcel_Click>
        /// Import Deliverers for the Selected Year from Excel
        /// </ImportDeliverersFromExcel_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportDeliverersFromExcel_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            string selectQueryText = "SELECT * FROM [Sheet1$]";

            //Open an Excel Workbook to Import Deliverers
            DataSet deliverersToImport = ExcelMethods.OpenExcelWorkbookAndExtractWorksheetInformation(selectQueryText);

            //Temp stats variables
            int fullhit = 0;
            int existsinselectedyear = 0;
            int duplicatesinselectedyeartable = 0;
            int doesnotexitinselectedyear = 0;
            int duplicatesindeliverertable = 0;
            int firstandlastnamehit = 0;
            int nohit = 0;

            int currentRow = 0;

            if (deliverersToImport != null)
            {
                foreach (DataRow record in deliverersToImport.Tables[0].Rows)
                {
                    //Process each deliverer from the excel spreadsheet
                    string deliverersSelectQuery = "SELECT * FROM Deliverers WHERE Last_Name = '" + record["Last_Name"].ToString().Replace("'", "''") + "' AND " +
                                               "First_Name = '" + record["First_Name"].ToString().Replace("'", "''") + "'";

                    DataSet Deliverers = mChristmasBasketsAccessDatabase.PerformSelectQuery(deliverersSelectQuery, "Deliverers");

                    if (Deliverers != null)
                    {
                        //Deliverer Already Exists in Deliverers Database

                        //See if only one instance of the client was found in the Deliverers table
                        if (Deliverers.Tables[0].Rows.Count == 1)
                        {
                            //Only 1 record was found in the Deliverers table

                            fullhit++;
                            //See if the Deliverer Already Exists in the mSelected year Table
                            string selectedYearDelivererSelectQuery = "SELECT * FROM " + mSelectedYear + "_Deliverers" + " WHERE Deliverer_ID = " + Deliverers.Tables[0].Rows[0]["Deliverer_ID"];

                            DataSet selectedYearDeliverer = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectedYearDelivererSelectQuery, mSelectedYear);

                            if (selectedYearDeliverer != null)
                            {
                                //Make sure there was only 1 record returned from the selected year table
                                if (selectedYearDeliverer.Tables[0].Rows.Count == 1)
                                {
                                    //Only 1 record was found in the selected year table
                                    existsinselectedyear++;
                                }
                                else
                                {
                                    //Duplicate records found in the selected year table
                                    duplicatesinselectedyeartable++;

                                    //Remove Duplicates from the mSelectedYear Table
                                    for (int i = 1; i < selectedYearDeliverer.Tables[0].Rows.Count; i++)
                                    {
                                        string deleteCommand = "DELETE FROM " + mSelectedYear + "_Deliverers" + " WHERE Deliverer_ID = " + selectedYearDeliverer.Tables[0].Rows[i]["Deliverer_ID"].ToString();
                                        mChristmasBasketsAccessDatabase.ExecuteNonQuery(deleteCommand);
                                    }
                                }
                            }
                            else
                            {
                                //Deliverer does not exist in the selected year table
                                doesnotexitinselectedyear++;

                                //Insert the deliverer to the mSelectedYear table
                                string insertCommand = "INSERT INTO " + mSelectedYear + "_Deliverers" + " (Deliverer_ID) VALUES (" + Deliverers.Tables[0].Rows[0]["Deliverer_ID"] + ")";
                                mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommand);
                            }

                            //Check for null values
                            foreach (DataColumn column in deliverersToImport.Tables[0].Columns)
                            {
                                if (deliverersToImport.Tables[0].Rows[currentRow][column.ColumnName].ToString() == "")
                                {
                                    deliverersToImport.Tables[0].Rows[currentRow][column.ColumnName] = 0;
                                }
                            }

                            //Update information from the Deliverer
                            string updateCommand = "UPDATE Deliverers SET Last_Name = '" + deliverersToImport.Tables[0].Rows[currentRow]["Last_Name"].ToString() + "', " +
                                                   "First_Name = '" + deliverersToImport.Tables[0].Rows[currentRow]["First_Name"] + "', " +
                                                   "Home_Phone = '" + deliverersToImport.Tables[0].Rows[currentRow]["Home_Phone"].ToString() + "', " +
                                                   "Work_Phone = '" + deliverersToImport.Tables[0].Rows[currentRow]["Work_Phone"].ToString() + "', " +
                                                   "Capacity = '" + deliverersToImport.Tables[0].Rows[currentRow]["Capacity"] + "', " +
                                                   "Comments = '" + deliverersToImport.Tables[0].Rows[currentRow]["Comments"] + "', " +
                                                   "Occupation_Status = '" + deliverersToImport.Tables[0].Rows[currentRow]["Occupation_Status"] + "'" + 
                                                   " WHERE Deliverer_ID = " + Deliverers.Tables[0].Rows[0]["Deliverer_ID"].ToString();

                            mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCommand);

                            //Object clean up
                            selectedYearDeliverer = null;
                        }
                        else
                        {
                            //Duplicates clients exist in the Deliverer table from all field
                            //search - very unlikely

                            ImportDeliverersFromExcelHelper excelImportHelper = new ImportDeliverersFromExcelHelper();
                            excelImportHelper.SetDatabaseDeliverersTable(Deliverers.Tables[0]);
                            excelImportHelper.SetMode(DelivererMode.DatabaseDeliverersTableDuplicates);
                            excelImportHelper.SetExcelDeliverersTable(record, Deliverers.Tables[0].Columns);
                            excelImportHelper.BindData();
                            excelImportHelper.ShowDialog();
                            excelImportHelper.UnBindData();

                            duplicatesindeliverertable++;
                        }
                    }
                    else
                    {
                        //Either deliverer does not exist or a field miss-match

                        //See if we can get a first name and last name on the deliverer
                        //trying to be imported from the excel spreadsheet
                        string selectQuery = "SELECT * FROM Deliverers WHERE Last_Name = '" + record["Last_Name"].ToString().Replace("'", "''") + "' AND " +
                                               "First_Name = '" + record["First_Name"].ToString().Replace("'", "''") + "'";

                        DataSet delivererDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectQuery, "Deliverers");

                        if (delivererDataSet != null)
                        {
                            //Hit in Deliverers Table by First_Name and Last_Name
                            //System.Windows.MessageBox.Show(record["Last_Name"].ToString() + ", " + record["First_Name"].ToString() + " Hit in Clients Table by First_Name, Last_Name, and Organization");
                            firstandlastnamehit++;

                            //See if there was only 1 record returned from the Clients table
                            if (delivererDataSet.Tables[0].Rows.Count == 1)
                            {
                                //Only 1 record found in the Clients Table

                                //////////////////////////////////////////////
                                //HANDLE CLIENT INFO UPDATE IN CLIENTS TABLE//
                                //////////////////////////////////////////////

                                ///////////////////////////////////////////////////////////////
                                //HANDLE ADDING CLIENT, WHO'S INFORMATION WAS UPDATED, TO THE//
                                //SELECTED YEAR TABLE                                        //
                                ///////////////////////////////////////////////////////////////

                                ImportDeliverersFromExcelHelper excelImportHelper = new ImportDeliverersFromExcelHelper();
                                excelImportHelper.SetDatabaseDeliverersTable(delivererDataSet.Tables[0]);
                                excelImportHelper.SetMode(DelivererMode.DatabaseDeliverersTableUpdateDeliverer);
                                excelImportHelper.SetExcelDeliverersTable(record, delivererDataSet.Tables[0].Columns);
                                excelImportHelper.BindData();
                                excelImportHelper.ShowDialog();
                                excelImportHelper.UnBindData();

                            }
                            else
                            {
                                //Duplicate records found in the Clients Table

                                /////////////////////
                                //HANDLE DUPLICATES//
                                /////////////////////

                                //////////////////////////////////////////////
                                //HANDLE CLIENT INFO UPDATE IN CLIENTS TABLE//
                                //////////////////////////////////////////////

                                ///////////////////////////////////////////////////////////////
                                //HANDLE ADDING CLIENT, WHO'S INFORMATION WAS UPDATED, TO THE//
                                //SELECTED YEAR TABLE                                        //
                                ///////////////////////////////////////////////////////////////

                                ImportDeliverersFromExcelHelper excelImportHelper = new ImportDeliverersFromExcelHelper();
                                excelImportHelper.SetDatabaseDeliverersTable(delivererDataSet.Tables[0]);
                                excelImportHelper.SetMode(DelivererMode.DatabaseDeliverersTableDuplicates);
                                excelImportHelper.SetExcelDeliverersTable(record, delivererDataSet.Tables[0].Columns);
                                excelImportHelper.BindData();
                                excelImportHelper.ShowDialog();
                                excelImportHelper.UnBindData();

                                duplicatesindeliverertable++;
                            }

                        }
                        else
                        {
                            //Still no hit
                            //System.Windows.MessageBox.Show(record["Last_Name"].ToString() + ", " + record["First_Name"].ToString() + " Still No Hit in Clients Table");
                            nohit++;
                            System.Data.DataTable temp = new System.Data.DataTable();
                            temp.Columns.Add("Deliverer_ID");
                            temp.Columns.Add("Last_Name");
                            temp.Columns.Add("First_Name");
                            temp.Columns.Add("Home_Phone");
                            temp.Columns.Add("Work_Phone");
                            temp.Columns.Add("Assigned_Previous_Delivery_Year");
                            temp.Columns.Add("Capacity");
                            temp.Columns.Add("Assigned");
                            temp.Columns.Add("Comments");
                            temp.Columns.Add("Clients");
                            temp.Columns.Add("Occupation_Status");
                            temp.Columns.Add("Year_Last_Delivered");

                            ImportDeliverersFromExcelHelper excelImportHelper = new ImportDeliverersFromExcelHelper();
                            excelImportHelper.SetDatabaseDeliverersTable(temp);
                            excelImportHelper.SetMode(DelivererMode.DatabaseDeliverersTableAddDeliverer);
                            excelImportHelper.SetExcelDeliverersTable(record, temp.Columns);
                            excelImportHelper.BindData();
                            excelImportHelper.ShowDialog();
                            excelImportHelper.UnBindData();

                        }

                        //Object clean up
                        delivererDataSet = null;
                    }

                    //Object clean up
                    Deliverers = null;

                    //Quit the Deliverer Import Process
                    if (Main.mQuitDelivererImport)
                    {
                        //Update Main.mQuitDelivererImport
                        Main.mQuitDelivererImport = false;

                        //Quit the For loop
                        break;
                    }

                    currentRow++;
                }
            }

            //Temp Stats messagebox
            System.Windows.MessageBox.Show("Full hit = " + fullhit.ToString() +
                                           "\nExists In Selected Year Table = " + existsinselectedyear.ToString() +
                                           "\nDuplicates In Selected Year Table = " + duplicatesinselectedyeartable.ToString() +
                                           "\nDoes not exist In Selected Year Table = " + doesnotexitinselectedyear.ToString() +
                                           "\nDuplicates In Deliverer Table = " + duplicatesindeliverertable.ToString() +
                                           "\nFirst Name and Last Name hit = " + firstandlastnamehit.ToString() +
                                           "\nNo hit = " + nohit.ToString());

            //Object clean up
            deliverersToImport = null;
        }

        /// <StampClientssWithSelectedYearButton_Click>
        /// Update Client's Year_Last_Delivered_To to Main.mSelectedYear without the "Year_"
        /// </StampClientssWithSelectedYearButton_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StampClientssWithSelectedYearButton_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            //Parse out year from mSelectedYear
            string selectedYear = mSelectedYear.Replace("Year_", "");

            //Get the mSelectedYear list of clients
            string selectQuery = "SELECT * FROM " + mSelectedYear;
            DataSet Clients = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectQuery, mSelectedYear);

            //Update the Year_Last_Delivered_To field = year for all clients in mSelectedYear
            foreach (DataRow dataRow in Clients.Tables[0].Rows)
            {
                string updateCommand = "UPDATE Clients SET Year_Last_Delivered_To = " + selectedYear + " WHERE Client_ID = " + dataRow["Client_ID"].ToString();
                mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCommand);
            }
        }

        private void StampDeliverersWithSelectedYearButton_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }
            //Parse out year from mSelectedYear
            string selectedYear = mSelectedYear.Replace("Year_", "");

            //Create the selectByYear_Last_DeliveredQueryText
            string selectByYear_Last_DeliveredQueryText = "SELECT * FROM "+ mSelectedYear + "_Deliverers";

            //Create the deliverersDataSet
            DataSet deliverersDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByYear_Last_DeliveredQueryText, "Deliverers");

            //Update the Year_Last_Delivered_To field = year for all clients in mSelectedYear
            foreach (DataRow dataRow in deliverersDataSet.Tables[0].Rows)
            {
                string updateCommand = "UPDATE Deliverers SET Year_Last_Delivered = " + selectedYear + " WHERE Deliverer_ID = " + dataRow["Deliverer_ID"].ToString();
                mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCommand);
            }
        }

        private void GenerateDelivererAssignmentExcelSheet()
        {
            //Define local variables
            string worksheetName = mSelectedYear.Replace("Year_", "") + " Deliverers";
            DataSet dataToExport = new DataSet();
            dataToExport.Tables.Add();
            bool columnsImported = false;
            int currentRow = 0;

            //Create the select selectByDeliverer_IDQueryText
            string selectByDeliverer_IDQueryText = "SELECT * FROM " + mSelectedYear + "_Deliverers";

            //Create the selectedYearDataSet
            DataSet selectedYearDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByDeliverer_IDQueryText, mSelectedYear + "_Deliverers");

            //Process each record in the selectedYear_Deliverers table
            foreach (DataRow dataRow in selectedYearDataSet.Tables[0].Rows)
            {
                //Create the selectByDelivererIDQueryText
                string selectByDelivererIDQueryText = "SELECT Deliverer_ID, Last_Name, First_Name, Home_Phone, Capacity, Comments, Assigned, Occupation_Status, Client_History FROM Deliverers WHERE Deliverer_ID = " + dataRow["Deliverer_ID"].ToString();

                //Perform the selectByDelivererIDQueryText and store the results in a Data Table
                System.Data.DataSet selectByDelivererIDDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByDelivererIDQueryText, "Deliverers");

                //Check to see we got results back
                if (selectByDelivererIDDataSet != null)
                {
                    //Add columns if not already added
                    if (columnsImported == false)
                    {
                        //Add each column requested from the query
                        foreach (DataColumn dataColumn in selectByDelivererIDDataSet.Tables[0].Columns)
                        {
                            dataToExport.Tables[0].Columns.Add(dataColumn.ColumnName, dataColumn.DataType);
                        }

                        //Update columnsImported
                        columnsImported = true;
                    }

                    //Copy the data row
                    DataRow record = selectByDelivererIDDataSet.Tables[0].Rows[0];
                    dataToExport.Tables[0].ImportRow(record);
                    currentRow++;
                }
            }

            //Export the query to Excel
            ExcelMethods.ExportToExcel(dataToExport, worksheetName);
        }


        /// <GenerateUnassignedClientsMapButton_Click>
        /// Genereate the unassigned clients map for the day of the event
        /// </GenerateUnassignedClientsMapButton_Click>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateUnassignedClientsMapButton_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            //Define local variables
            string selectedYear = mSelectedYear.Replace("Year_", "");
            string fileName = selectedYear + " Unassigned Clients Map.htm";

            //Create the select selectByClient_IDQueryText
            string selectByClient_IDQueryText = "SELECT Client_ID FROM " + mSelectedYear;

            //Get the list of clients

            //Create the clientIDList
            DataSet clientIDList = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClient_IDQueryText, mSelectedYear);

            int clientCount = clientIDList.Tables[0].Rows.Count;
            List<string> clientList = new List<string>();
            List<string> addressList = new List<string>();
            int index = 0;
            int count = 0;

            //Create the list of assigned clients from the Year_****_Deliverers list
            string selectAssignedClientsByClient_IDQueryText = "SELECT Clients FROM " + mSelectedYear + "_Deliverers";
            DataSet assignedClientIDDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectAssignedClientsByClient_IDQueryText, mSelectedYear);
            List<string> assignedClientIDList = new List<string>();

            //Split apart the clients by the ',' delimeter from each deliverer and add it to the list
            foreach (DataRow dataRow in assignedClientIDDataSet.Tables[0].Rows)
            {
                string[] parsedAssignedClients = dataRow["Clients"].ToString().Split(',');

                foreach (string clientID in parsedAssignedClients)
                {
                    //Verify there are no clients assigned more than once
                    if (assignedClientIDList.Contains(clientID))
                    {
                        //There is a client assigned more than once
                        System.Windows.MessageBox.Show("Duplicate clientID:  " + clientID);
                    }
                    else
                    {
                        //Add the client to the list
                        assignedClientIDList.Add(clientID);
                    }
                }
            }

            //Process each record in the clientIDList table
            foreach (DataRow dataRow in clientIDList.Tables[0].Rows)
            {
                //Create the selectByClientIDQueryText
                string selectByClientIDQueryText = "SELECT Client_ID, Last_Name, First_Name, Address_Number, Street_Address, City, Zipcode FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString();

                //Perform the selectByClientIDAndOrganizationQuery and store the results in a Data Table
                System.Data.DataSet selectByClientIDDataSet = mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClientIDQueryText, "Clients");

                //Check to see we got results back
                if (selectByClientIDDataSet != null)
                {
                    //Get the Client_ID from the DataRow
                    string clientID = selectByClientIDDataSet.Tables[0].Rows[0]["Client_ID"].ToString();

                    //Check to see if the client has been assigned, clientID of 0 means unassigned
                    if (!assignedClientIDList.Contains(clientID))
                    {
                        //Client has not been assigned so add it to the list of clients and adddress

                        //Format the strings properly
                        string id = selectByClientIDDataSet.Tables[0].Rows[0]["Client_ID"].ToString();
                        string firstName = selectByClientIDDataSet.Tables[0].Rows[0]["First_Name"].ToString();
                        string lastName = selectByClientIDDataSet.Tables[0].Rows[0]["Last_Name"].ToString();
                        string addressNumber = selectByClientIDDataSet.Tables[0].Rows[0]["Address_Number"].ToString();
                        string streetAddress = selectByClientIDDataSet.Tables[0].Rows[0]["Street_Address"].ToString();
                        string city = selectByClientIDDataSet.Tables[0].Rows[0]["City"].ToString();
                        string zipcode = selectByClientIDDataSet.Tables[0].Rows[0]["Zipcode"].ToString();

                        clientList.Add(id + " - " + lastName + ", " + firstName);
                        addressList.Add(addressNumber + " " + streetAddress + " " + city + ", VA  " + zipcode);
                        index++;
                    }
                }

                //Increment count
                count++;

                if (count == clientCount)
                {
                    break;
                }
            }

            //Create the final client and address lists to pass to the map generation method
            string[] clients = new string[index];
            string[] addresses = new string[index];

            for (int i = 0; i < index; i++)
            {
                clients[i] = clientList[i];
                addresses[i] = addressList[i];
            }

            //Generate the Deliverer Assignment Map
            //GenerateAssignmentMap(fileName, selectedYear, clients, addresses);
        }

        private void CheckForDuplicatesButton_Click(object sender, RoutedEventArgs e)
        {
            //First and Last Name the same
            //Check for identical phone number
            bool columnsImported = false;
            DataSet dataToExport = new DataSet();
            dataToExport.Tables.Add();
            DataSet duplicatesRecorded = new DataSet();
            duplicatesRecorded.Tables.Add();
            string worksheetName = mSelectedYear.Replace("Year_", "") + " Duplicate Clients";
            string worksheetPath = "";
            string selectQueryText = "SELECT * FROM [Sheet1$]";

            //Open an Excel Workbook to Import All Clients
            DataSet clientsToCheck = ExcelMethods.OpenExcelWorkbookAndExtractWorksheetInformationAndSaveWorksheetPath(selectQueryText, ref worksheetPath);

            //Process each record in the imported clients table
            foreach (DataRow dataRow in clientsToCheck.Tables[0].Rows)
            {
                //First look for first and last name match

                //Create the selectDuplicateClientsFirstNameLastNameQuery
                string selectDuplicateClientsFirstNameLastNameQuery = "SELECT * FROM [Sheet1$] WHERE First_Name = '" + dataRow["First_Name"].ToString().Replace("'", "''") + "' AND Last_Name = '" + dataRow["Last_Name"].ToString().Replace("'", "''") + "'";
                string selectDuplicateClientsPhoneQuery = "SELECT * FROM [Sheet1$] WHERE Phone = '" + dataRow["Phone"].ToString().Replace("'", "''") + "'";
                string selectDuplicateClientsAddressQuery = "SELECT * FROM [Sheet1$] WHERE Address_Number = '" + dataRow["Address_Number"].ToString().Replace("'", "''") + "' AND " +
                                                            "Street_Address = '" + dataRow["Street_Address"].ToString().Replace("'", "''") + "' AND " +
                                                            "City = '" + dataRow["City"].ToString().Replace("'", "''") +"'";

                //Find Duplicates using First Name and Last Name
                FindDuplicates(ref columnsImported, selectDuplicateClientsFirstNameLastNameQuery, worksheetPath, ref dataToExport, ref duplicatesRecorded);
                
                //Find Duplicates using Phone Number
                FindDuplicates(ref columnsImported, selectDuplicateClientsPhoneQuery, worksheetPath, ref dataToExport, ref duplicatesRecorded);

                //Find Duplicates using Address
                FindDuplicates(ref columnsImported, selectDuplicateClientsAddressQuery, worksheetPath, ref dataToExport, ref duplicatesRecorded);
            }

            //Export the query to Excel
            ExcelMethods.ExportToExcel(dataToExport, worksheetName);
        }

        private void FindDuplicates(ref bool columnsImported, string selectQuery, string worksheetPath, ref DataSet dataToExport, ref DataSet duplicatesRecorded)
        {
            //Initialize local variables
            bool duplicateRecorded = false;

            //Perform the selectDuplicateClientsFirstNameLastNameQuery and store the results in a Data Table
            DataSet duplicateResults = ExcelMethods.OpenExcelWorkbookAndExtractWorksheetInformation(selectQuery, worksheetPath);

            //Check to see we got results back
            if (duplicateResults.Tables[0].Rows.Count > 1)
            {
                //Add columns if not already added
                if (columnsImported == false)
                {
                    //Add each column requested from the query
                    foreach (DataColumn dataColumn in duplicateResults.Tables[0].Columns)
                    {
                        dataToExport.Tables[0].Columns.Add(dataColumn.ColumnName);
                        duplicatesRecorded.Tables[0].Columns.Add(dataColumn.ColumnName);
                    }

                    //Update columnsImported
                    columnsImported = true;
                }

                foreach (DataRow record in duplicateResults.Tables[0].Rows)
                {
                    //See if the duplicate has been recorded already
                    duplicateRecorded = false;

                    foreach (DataRow duplicateRecord in duplicatesRecorded.Tables[0].Rows)
                    {
                        if(duplicateRecord["Last_Name"].ToString() == record["Last_Name"].ToString() &&
                            duplicateRecord["First_Name"].ToString() == record["First_Name"].ToString() &&
                            duplicateRecord["Middle_Name"].ToString() == record["Middle_Name"].ToString() &&
                            duplicateRecord["Title"].ToString() == record["Title"].ToString() &&
                            duplicateRecord["Address_Number"].ToString() == record["Address_Number"].ToString() &&
                            duplicateRecord["Street_Address"].ToString() == record["Street_Address"].ToString() &&
                            duplicateRecord["City"].ToString() == record["City"].ToString() &&
                            duplicateRecord["Zipcode"].ToString() == record["Zipcode"].ToString() &&
                            duplicateRecord["Phone"].ToString() == record["Phone"].ToString() &&
                            duplicateRecord["Organization"].ToString() == record["Organization"].ToString())
                        {
                            //Set duplicateRecorded flag to true
                            duplicateRecorded = true;
                        }

                    }

                    if(!duplicateRecorded)
                    {

                        //Duplicate has not been recorded so record it

                        //Copy the data row for excel export
                        dataToExport.Tables[0].ImportRow(record);

                        //Put record in the duplicate recorded list
                        duplicatesRecorded.Tables[0].ImportRow(record);
                    }
                }
            }
        }

        private void AssignClientsToDeliverersButton_Click(object sender, RoutedEventArgs e)
        {
            //Make sure a year has been selected
            if (Main.mSelectedYear != "NONE")
            {
                var delivererAssignment = new ChristmasBasketsDashboard.WindowDelivererAssignmentDash();
                delivererAssignment.Show();
            }
            else
            {
                System.Windows.MessageBox.Show("Open the database and select a year jerk!");
            }
        }

        private void AutoAssignClientHistoryToDeliverersButton_Click(object sender, RoutedEventArgs e)
        {
            //////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////
            //////////////          COME BACK HERE JERK!!!!!!!!!        //////////////////
            //////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////

            //Make sure a year has been selected
            if (Main.mSelectedYear != "NONE")
            {
                //Create the select selectCurrentYearDeliverersQuery
                string selectCurrentYearDeliverersQuery = "SELECT * FROM " + Main.mSelectedYear + "_Deliverers";

                //Get the list of deliverers

                //Create the deliverers
                DataSet deliverers = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectCurrentYearDeliverersQuery, Main.mSelectedYear + "_Deliverers");

                if (deliverers != null)
                {
                    //Process each record in the clientIDList table
                    foreach (DataRow deliverer in deliverers.Tables[0].Rows)
                    {
                        
                        //Create the selectIndividualDelivererQuery
                        string selectIndividualDelivererQuery = "SELECT * FROM Deliverers WHERE Deliverer_ID=" + deliverer["Deliverer_ID"].ToString();

                        //Perform the selectIndividualDelivererQuery and store the results in a Data Table
                        System.Data.DataSet selectIndividualDelivererDataSet = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectIndividualDelivererQuery, "Deliverers");

                        //Check to see we got results back
                        if (selectIndividualDelivererDataSet != null)
                        {
                            //Format the strings properly
                            string clientHistoryString = selectIndividualDelivererDataSet.Tables[0].Rows[0]["Client_History"].ToString();
                            string[] clientHistory = clientHistoryString.Split(',');
                            

                            //Create the selectedYearDelivererDataQuery
                            string selectedYearDelivererTableName = mSelectedYear + "_Deliverers";
                            string selectedYearDelivererDataQuery = "SELECT * FROM " + selectedYearDelivererTableName + " WHERE Deliverer_ID=" + deliverer["Deliverer_ID"].ToString();

                            //Perform the selectedYearDelivererData and store the results in a Data Table
                            System.Data.DataSet selectedYearDelivererData = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectedYearDelivererDataQuery, selectedYearDelivererTableName);

                            //Get assignedClients for "Year_XXXX_Deliverers" table
                            string assignedClients = selectedYearDelivererData.Tables[0].Rows[0]["Clients"].ToString();

                            //Get current deliverer assigned count and capacity
                            string assignedCount = selectIndividualDelivererDataSet.Tables[0].Rows[0]["Assigned"].ToString();
                            string capacity = selectIndividualDelivererDataSet.Tables[0].Rows[0]["Capacity"].ToString();
                            int capacityNumber = 0;
                            int assignedCountNumber = 0;

                            if (assignedCount != "")
                            {
                                assignedCountNumber = Convert.ToInt32(assignedCount);
                            }

                            if (capacity != "")
                            {
                                capacityNumber = Convert.ToInt32(capacity);
                            }

                            foreach (string clientId in clientHistory)
                            {
                                //Verify we have not assigned the deliverer their full capacity and we have a client ID to process
                                if (clientId != ""  && (assignedCountNumber < capacityNumber))
                                {
                                    //We have not assigned the deliverer their full capacity and we have a client ID to process

                                    //Perform the selectIndividualDelivererQuery and store the results in a Data Table
                                    selectIndividualDelivererDataSet = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectIndividualDelivererQuery, "Deliverers");

                                    //See if Client exists in this year's clients
                                    string clientExistInCurrentYearQuery = "SELECT * FROM " + Main.mSelectedYear + " WHERE Client_ID=" + clientId;

                                    //Perform the clientExistInCurrentYearQuery and store the results in a Data Table
                                    System.Data.DataSet clientExistsInCurrentYearDataSet = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(clientExistInCurrentYearQuery, Main.mSelectedYear);

                                    if (clientExistsInCurrentYearDataSet != null)
                                    {
                                        //Client exists in selected year clients
                                        string clientAssignedStatus = clientExistsInCurrentYearDataSet.Tables[0].Rows[0]["Assigned_Status"].ToString();

                                        //See if the client has already been assigned
                                        if (clientAssignedStatus != "true")
                                        {
                                            //Client has not been assigned so assign the client to the current deliverer

                                            //Update Year_**** Client assigned status
                                            string updateClientTableAssignedStatus = "UPDATE " + Main.mSelectedYear + " SET Assigned_Status='true' WHERE Client_ID=" + clientId;
                                            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateClientTableAssignedStatus);

                                            //Append client to Clients


                                            if (assignedClients == "")
                                            {
                                                //First element
                                                assignedClients = clientId + ",";
                                            }
                                            else
                                            {
                                                //Not first element
                                                assignedClients += clientId + ",";
                                            }

                                            //Update Year_****_Deliverers Deliverer Clients field
                                            string updateCurrentYearDeliverersTableClients = "UPDATE " + Main.mSelectedYear + "_Deliverers SET Clients='" + assignedClients + "' WHERE Deliverer_ID=" + deliverer["Deliverer_ID"].ToString();
                                            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCurrentYearDeliverersTableClients);

                                            //Increment current deliverer assigned count
                                            assignedCountNumber++;

                                            //Update Deliverers Deliverer Assigned Count
                                            string updateDeliverersTableAssignedCount = "UPDATE Deliverers SET Assigned=" + assignedCountNumber.ToString() + " WHERE Deliverer_ID=" + deliverer["Deliverer_ID"].ToString();
                                            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateDeliverersTableAssignedCount);
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Open the database and select a year jerk!");
            }
        }

        private void GenerateDelivererPacketsButton_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            //Initialize variables
            string delivererIDQuery = "SELECT * FROM " + Main.mSelectedYear + "_Deliverers";
            List<Deliverer> deliverers = new List<Deliverer>();
            List<string> generatedFiles = new List<string>();

            //Get all the deliverers assigned to the year
            DataSet delivererIDs = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(delivererIDQuery, Main.mSelectedYear + "_Deliverers");

            //Process each record in the imported deliverers table
            foreach (DataRow dataRow in delivererIDs.Tables[0].Rows)
            {
                //Get the detailed information from the master Deliverer's Table in the database
                string getDelivererInfoQuery = "SELECT * FROM Deliverers WHERE Deliverer_ID = " + dataRow["Deliverer_ID"];

                DataSet delivererBeingProcessed = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(getDelivererInfoQuery, "Deliverers");

                DataRow delivererInfoFromDatabase = delivererBeingProcessed.Tables[0].Rows[0];

                Deliverer delivererToAdd = new Deliverer();

                delivererToAdd.DelivererID = Convert.ToInt32(delivererInfoFromDatabase["Deliverer_ID"]);
                delivererToAdd.FirstName = delivererInfoFromDatabase["First_Name"].ToString();
                delivererToAdd.LastName = delivererInfoFromDatabase["Last_Name"].ToString();
                delivererToAdd.Capacity = Convert.ToInt32(delivererInfoFromDatabase["Capacity"]);
                delivererToAdd.HelpStatus = delivererInfoFromDatabase["Help_Status"].ToString();
                delivererToAdd.Room = delivererInfoFromDatabase["Room"].ToString();
                delivererToAdd.WorkPhone = delivererInfoFromDatabase["Work_Phone"].ToString();
                delivererToAdd.HomePhone = delivererInfoFromDatabase["Home_Phone"].ToString();
                delivererToAdd.OccupationStatus = delivererInfoFromDatabase["Occupation_Status"].ToString();
                delivererToAdd.Comments = delivererInfoFromDatabase["Comments"].ToString();
                delivererToAdd.ClientHistory = delivererInfoFromDatabase["Client_History"].ToString();
                delivererToAdd.YearLastDelivered = delivererInfoFromDatabase["Year_Last_Delivered"].ToString();
                delivererToAdd.Assigned = Convert.ToInt32(delivererInfoFromDatabase["Assigned"]);

                //Get Clients from Year_XXXX_Clients Table
                string getDelivererClientsQuery = "SELECT Clients FROM " + Main.mSelectedYear + "_Deliverers WHERE Deliverer_ID = " + dataRow["Deliverer_ID"];

                DataSet delivererClientsBeingProcessed = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(getDelivererClientsQuery, Main.mSelectedYear + "_Deliverers");

                DataRow delivererClientsFromDatabase = delivererClientsBeingProcessed.Tables[0].Rows[0];

                delivererToAdd.Clients = delivererClientsFromDatabase["Clients"].ToString();

                deliverers.Add(delivererToAdd);
            }

            //Generate all deliverer packets
            foreach(Deliverer deliverer in deliverers)
            {
                //Create the fileName
                string [] fileNames = null;
                string [] fullAddresses = null;
                string [] clients = null;
                string [] names = null;
                string [] streetAddresses = null;
                string [] cities = null;
                string [] zipCodes = null;
                string [] phoneNumbers = null;
                string [] organizations = null;
                string [] comments = null;

                //Prepare Deliverer Packet Arguments
                PrepareDelivererPacketArguments(deliverer, ref fullAddresses, ref clients, ref names, ref streetAddresses, ref cities, ref zipCodes, ref phoneNumbers, ref organizations, ref comments);

                //Determine how many clients are to process
                int numberOfFilesToGenerate = clients.Length / mMaxNumberGoogleMapsWaypoints;
                numberOfFilesToGenerate += (clients.Length % mMaxNumberGoogleMapsWaypoints == 0 ? 0 : 1);

                //Allocate number of fileNames
                fileNames = new string[numberOfFilesToGenerate];

                //Add each filename to generated files list
                for (int i = 1; i < (numberOfFilesToGenerate + 1); i++)
                {
                    string nameOfFile = "Deliverer(" + deliverer.DelivererID + ")_" + deliverer.LastName + "_" + deliverer.FirstName + "_Part(" + i.ToString() + ").htm";
                    fileNames[i - 1] = nameOfFile;
                    generatedFiles.Add(nameOfFile);
                }

                GenerateDelivererPacket(deliverer, ref fileNames, ref fullAddresses, ref clients, ref names, ref streetAddresses, ref cities, ref zipCodes, ref phoneNumbers, ref organizations, ref comments);
            }

            //Since we've created the .htm client white card files with maps
            //open them in firefox to print them

            //Get the total number of deliverers to display in firefox
            int deliverersToDisplayInFirefox = generatedFiles.Count;

            //Create a variable for the client to start from
            int startingDeliverer = 0;

            //Have a variable to use for number to print in a single session
            int sessionDelivererLimit = 100;

            while (deliverersToDisplayInFirefox != 0)
            {
                //Check for the case where deliverersToDisplayInFirefox < sessionDelivererLimit
                if (deliverersToDisplayInFirefox < sessionDelivererLimit)
                {
                    sessionDelivererLimit = deliverersToDisplayInFirefox;
                }

                //Create a new process
                System.Diagnostics.Process process = new System.Diagnostics.Process();

                //The process will be firefox
                process.StartInfo.FileName = "firefox.exe";

                //Create firefox's command argument line
                //Open all of the files we created in seperate tabs
                for (int i = startingDeliverer; i < (startingDeliverer + sessionDelivererLimit); i++)
                {
                    process.StartInfo.Arguments += "\"" + generatedFiles[i] + "\" ";
                }

                //Start Firefox
                process.Start();

                //Close our handle to Firefox
                process.Close();

                //Update deliverersToDisplayInFirefox
                deliverersToDisplayInFirefox -= sessionDelivererLimit;

                //Update startingDeliverer
                startingDeliverer += sessionDelivererLimit;
            }   
        }

        private void PrepareDelivererPacketArguments(Deliverer deliverer, ref string[] fullAddresses, ref string[] clients, ref string[] names, ref string[] streetAddresses, ref string[] cities, ref string[] zipCodes, ref string[] phoneNumbers, ref string[] organizations, ref string[] comments)
        {
            //Get the list of clients

            //Create the clientIDList
            string [] clientIDList = deliverer.Clients.Split(',');

            int index = 0;

            //Create new arrays
            fullAddresses = new string[deliverer.Assigned];
            clients = new string[deliverer.Assigned];
            names = new string[deliverer.Assigned];
            streetAddresses = new string[deliverer.Assigned];
            cities = new string[deliverer.Assigned];
            zipCodes = new string[deliverer.Assigned];
            phoneNumbers = new string[deliverer.Assigned];
            organizations = new string[deliverer.Assigned];
            comments = new string[deliverer.Assigned];

            //Process each record in the clientIDList table
            foreach (string clientIDToProcess in clientIDList)
            {
                if (clientIDToProcess != "")
                {
                    //Create the selectByClientIDQueryText
                    string selectByClientIDQueryText = "SELECT * FROM Clients WHERE Client_ID = " + clientIDToProcess;

                    //Perform the selectByClientIDDataSet and store the results in a Data Table
                    System.Data.DataSet selectByClientIDDataSet = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClientIDQueryText, "Clients");

                    //Check to see we got results back
                    if (selectByClientIDDataSet != null)
                    {
                        //Format the strings properly
                        string id = selectByClientIDDataSet.Tables[0].Rows[0]["Client_ID"].ToString();
                        string firstName = selectByClientIDDataSet.Tables[0].Rows[0]["First_Name"].ToString();
                        string lastName = selectByClientIDDataSet.Tables[0].Rows[0]["Last_Name"].ToString();
                        string addressNumber = selectByClientIDDataSet.Tables[0].Rows[0]["Address_Number"].ToString();
                        string streetAddress = selectByClientIDDataSet.Tables[0].Rows[0]["Street_Address"].ToString();
                        string city = selectByClientIDDataSet.Tables[0].Rows[0]["City"].ToString();
                        string phone = selectByClientIDDataSet.Tables[0].Rows[0]["Phone"].ToString();
                        string zipcode = selectByClientIDDataSet.Tables[0].Rows[0]["Zipcode"].ToString();
                        string organization = selectByClientIDDataSet.Tables[0].Rows[0]["Organization"].ToString();
                        string comment = selectByClientIDDataSet.Tables[0].Rows[0]["Directions"].ToString();

                        fullAddresses[index] = addressNumber + " " + streetAddress + " " + city + ", VA  " + zipcode;
                        clients[index] = id;
                        names[index] = firstName + " " + lastName;
                        streetAddresses[index] = addressNumber + " " + streetAddress;
                        cities[index] = city;
                        zipCodes[index] = zipcode;
                        phoneNumbers[index] = phone;
                        organizations[index] = organization;
                        comments[index] = comment;

                        index++;
                    }

                    if (index == deliverer.Assigned)
                    {
                        break;
                    }
                }
            }
        }

        private void GenerateDelivererPacket(Deliverer deliverer, ref string  [] fileNames, ref string [] fullAddresses, ref string []  clients, ref string []  names, ref string []  streetAddresses, ref string []  cities, ref string []  zipCodes, ref string [] phoneNumbers, ref string []  organizations, ref string []  comments)
        {
            //See if there is just one file to generate
            if (fileNames.Length == 0)
            {
                //Nothing to Do
                return;
            }
            else
            {
                //Initialize variables for previous end addresses
                string previousEndAddress = "";
                int remainingNumberOfClients = clients.Length;
                int numberClientsInTable = 0;
                int index = 0;

                for(int i = 0; i < fileNames.Length; i++)
                {
                    //Define local variables
                    TextWriter textWriter = new StreamWriter(fileNames[i]);

                    //////////////////////////////////////////////////////////////////////////////////////////
                    //////////////////////////////////////////////////////////////////////////////////////////
                    ////////////////////////        COME BACK HERE JERK!!!!      /////////////////////////////
                    //////////////////////////////////////////////////////////////////////////////////////////
                    //////////////////////////////////////////////////////////////////////////////////////////

                    //Write the header part of the .htm page
                    textWriter.WriteLine("<!DOCTYPE html>");
                    textWriter.WriteLine("<html>");
                    textWriter.WriteLine("   <head>");
                    textWriter.WriteLine("      <meta name=\"viewport\" content=\"initial-scale=1.0, user-scalable=no\"/>");
                    textWriter.WriteLine("      <title>Deliverer(" + deliverer.DelivererID + "):  " + deliverer.FirstName + " " + deliverer.LastName + "</title>");

                    //Write the java script part of the .htm page
                    textWriter.WriteLine("      <script type=\"text/javascript\" src=\"http://maps.google.com/maps/api/js?sensor=false\"></script>");
                    textWriter.WriteLine("      <script type=\"text/javascript\">");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("      var map;");
                    textWriter.WriteLine("      var directionsDisplay = new google.maps.DirectionsRenderer();");
                    textWriter.WriteLine("      var directionsService = new google.maps.DirectionsService();");

                    //If this is the first file we are generating start the route at GE in Salem
                    if(i == 0)
                    {
                        textWriter.WriteLine("      var start = \"1501 Roanoke Blvd Salem, VA 24153\";");
                    }
                    else
                    {
                        //This is not the first file we are generating so start the route at the previousEndAddress
                        textWriter.WriteLine("      var start = \""+ previousEndAddress + "\";");
                    }

                    //Generate addresses arrays
                    string javaAddresses = "      var addresses = [";

                    //See if we are generating a single file, if so send them back to GE in Salem as the end destination
                    if (fileNames.Length == 1)
                    {
                        textWriter.WriteLine("      var end = \"1501 Roanoke Blvd Salem, VA 24153\";");

                        //See if we have 1 or more than 1 client to process
                        if (clients.Count() == 1)
                        {
                            //One client
                            javaAddresses += "\"" + fullAddresses[0] + "\"];";
                        }
                        else
                        {
                            //More than one client
                            for (int j = 0; j < clients.Count(); j++)
                            {
                                //Insert clients
                                if (j == (clients.Count() - 1))
                                {
                                    //Last Record
                                    javaAddresses += "\"" + fullAddresses[j] + "\"];";
                                }
                                else
                                {
                                    //Not the Last Record
                                    javaAddresses += "\"" + fullAddresses[j] + "\",";
                                }
                            }

                        }

                        //Set numberClientsInTable
                        numberClientsInTable = clients.Count();
                    }
                    else
                    {
                        //More than one file to produce so choose an end destination from the clients list
                        if (remainingNumberOfClients <= mMaxNumberGoogleMapsWaypoints)
                        {
                            if (remainingNumberOfClients == 0)
                            {
                                //Last Record
                                textWriter.WriteLine("      var end = \"1501 Roanoke Blvd Salem, VA 24153\";");
                                javaAddresses += "];";
                                numberClientsInTable = 0;
                            }
                            else
                            {
                                //Write remainingNumberOfClients
                                for (int k = 0; k < remainingNumberOfClients; k++)
                                {
                                    index = (i * mMaxNumberGoogleMapsWaypoints) + k + 1;

                                    //Insert clients
                                    if ((k == (mMaxNumberGoogleMapsWaypoints - 2)) || (k == (remainingNumberOfClients - 1)))
                                    {
                                        //Last Record
                                        javaAddresses += "\"" + fullAddresses[index] + "\"];";
                                    }
                                    else
                                    {
                                        //Not the Last Record
                                        javaAddresses += "\"" + fullAddresses[index] + "\",";
                                    }
                                }

                                //Set var end as GE in Salem
                                textWriter.WriteLine("      var end = \"1501 Roanoke Blvd Salem, VA 24153\";");

                                //Cache away number of Clients to print in Deliverer Table
                                numberClientsInTable = remainingNumberOfClients;

                                //Subtract remainingNumberOfClients from remainingNumberOfClients to arrive at remainingNumberOfClients = 0
                                remainingNumberOfClients -= remainingNumberOfClients;
                            }
                        }
                        else
                        {
                            //Write 8 clients as waypoints + 1 client as the var end address
                            
                            //Write remainingNumberOfClients
                            for (int k = 0; k < mMaxNumberGoogleMapsWaypoints; k++)
                            {
                                index = (i * mMaxNumberGoogleMapsWaypoints) + k;

                                //Insert clients
                                if (k == (mMaxNumberGoogleMapsWaypoints - 1))
                                {
                                    //Last Record
                                    javaAddresses += "\"" + fullAddresses[index] + "\"];";
                                }
                                else
                                {
                                    //Not the Last Record
                                    javaAddresses += "\"" + fullAddresses[index] + "\",";
                                }
                            }

                            //Update previousEndAddress
                            previousEndAddress = fullAddresses[index + 1];

                            //Cache away number of Clients to print in Deliverer Table
                            numberClientsInTable = mMaxNumberGoogleMapsWaypoints + 1;

                            //Update var end
                            textWriter.WriteLine("      var end = \"" + previousEndAddress + "\";");

                            //Subtract 8 clients (waypoints) + 1 client (var end) from remaining
                            remainingNumberOfClients -= (mMaxNumberGoogleMapsWaypoints + 1);
                        }
                    }

                    //Insert addresses array
                    textWriter.WriteLine(javaAddresses);
                    textWriter.WriteLine("");

                    textWriter.WriteLine("      var unassignedClientImage = \"UnassignedClient.png\";");
                    textWriter.WriteLine("      var mapCenterAddress = \"2147 Dale Avenue Southeast Roanoke, VA 24013\";");

                    //initialize function
                    textWriter.WriteLine("      function initialize()");
                    textWriter.WriteLine("      {");
                    textWriter.WriteLine("         //Create map options");
                    textWriter.WriteLine("         var myOptions = {mapTypeId: google.maps.MapTypeId.ROADMAP};");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("         //Create map and direction display");
                    textWriter.WriteLine("         map = new google.maps.Map(document.getElementById(\"map_canvas\"),myOptions);");
                    textWriter.WriteLine("         directionsDisplay.setMap(map);");
                    textWriter.WriteLine("         directionsDisplay.setPanel(document.getElementById(\"directions_panel\"));");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("         //Calculate the proper route");
                    textWriter.WriteLine("         calcRoute();");
                    textWriter.WriteLine("      }");
                    textWriter.WriteLine("");

                    //calcRoute function
                    textWriter.WriteLine("      //Calculate Client Route");
                    textWriter.WriteLine("      function calcRoute()");
                    textWriter.WriteLine("      {");
                    textWriter.WriteLine("         //Build waypoints object");
                    textWriter.WriteLine("         var waypts = [];");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("         for (var i = 0; i < addresses.length; i++)");
                    textWriter.WriteLine("         {");
                    textWriter.WriteLine("            waypts.push({location:addresses[i],stopover:true});");
                    textWriter.WriteLine("         }");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("         //Build Request");
                    textWriter.WriteLine("         var request = {");
                    textWriter.WriteLine("         origin: start,");
                    textWriter.WriteLine("         destination: end,");
                    textWriter.WriteLine("         waypoints: waypts,");
                    textWriter.WriteLine("         optimizeWaypoints: true,");
                    textWriter.WriteLine("         travelMode: google.maps.DirectionsTravelMode.DRIVING};");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("         //Show directions");
                    textWriter.WriteLine("         directionsService.route(request,");
                    textWriter.WriteLine("         function(response, status)");
                    textWriter.WriteLine("         {");
                    textWriter.WriteLine("            if (status == google.maps.DirectionsStatus.OK)");
                    textWriter.WriteLine("            {");
                    textWriter.WriteLine("               directionsDisplay.setDirections(response);");
                    textWriter.WriteLine("            }");
                    textWriter.WriteLine("            else");
                    textWriter.WriteLine("            {");
                    textWriter.WriteLine("               //Alert - Route Generation not successful");
                    textWriter.WriteLine("               alert(\"calcRoute - Route Generation was not successful for the following reason: (\" + status + \")\");");
                    textWriter.WriteLine("            }");
                    textWriter.WriteLine("         });");
                    textWriter.WriteLine("      }");
                    textWriter.WriteLine("");

                    //Final script tags
                    textWriter.WriteLine("      </script>");
                    textWriter.WriteLine("   </head>");
                    textWriter.WriteLine("");

                    //Write the body part of the .htm page
                    textWriter.WriteLine("   <body onload=\"initialize()\">");
                    textWriter.WriteLine("      <!-- General Instructions -->");
                    textWriter.WriteLine("      <h2>General Instructions:</h2>");
                    textWriter.WriteLine("      <div style=\"width:700px;height:85px;border:6px outset red;\">");
                    textWriter.WriteLine("      <font size = 3> Each delivery consists of two boxes; a red and a green.");
                    textWriter.WriteLine("      The recipient needs to sign a USDA form upon delivery.");
                    textWriter.WriteLine("      If a box cannot be delivered, call Dick Stanfield at 540-353-7977 for assistance.");
                    textWriter.WriteLine("      Undelivered boxes should be returned to the west gate of GE at 1501 Roanoke Blvd Salem, VA.");
                    textWriter.WriteLine("      If returning a box, be sure to include the corresponding USDA form.</font></div>");
                    textWriter.WriteLine("      <!-- Deliverer -->");
                    textWriter.WriteLine("      <h2>Deliverer:   " + deliverer.LastName + ", " + deliverer.FirstName + " - Number of Clients (" + clients.Length + ") - Number Of Boxes (" + clients.Length * 2 + ")" + "</h2>");
                    
                    textWriter.WriteLine("");
                    textWriter.WriteLine("      <!-- Google Map -->");
                    textWriter.WriteLine("      <h2>Suggested Route:</h2>");
                    textWriter.WriteLine("      <div id=\"map_canvas\" style=\"width:800px; height:800px;\"></div>");
                    textWriter.WriteLine("      <div style=\"page-break-before:always;\"></div>");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("      <!-- Google Directions -->");
                    textWriter.WriteLine("      <h2>Directions:  Leaving From GE</h2>");
                    textWriter.WriteLine("      <div id=\"directions_panel\"></div>");
                    textWriter.WriteLine("      <div style=\"page-break-before:always;\"></div>");
                    textWriter.WriteLine("");
                    textWriter.WriteLine("      <!-- Client Information Table-->");
                    textWriter.WriteLine("      <h2 align = \"left\">Client Information</h2>");
                    textWriter.WriteLine("      <table align = \"left\" width=\"100%\" border=\"1\">");
                    textWriter.WriteLine("      <!-- Table Headings -->");
                    textWriter.WriteLine("      <tr>");
                    textWriter.WriteLine("         <th width = \"10%\">Client ID</th>");
                    textWriter.WriteLine("         <th width = \"20%\">Client Name</th>");
                    textWriter.WriteLine("         <th width = \"30%\">Street Address</th>");
                    textWriter.WriteLine("         <th width = \"10%\">City</th>");
                    textWriter.WriteLine("         <th width = \"10%\">Zip</th>");
                    textWriter.WriteLine("         <th width = \"10%\">Phone</th>");
                    textWriter.WriteLine("         <th width = \"10%\">Organization</th>");
                    textWriter.WriteLine("      </tr>");

                    //Put each deliverer's client in the table
                    for (int m = 0; m < numberClientsInTable; m++)
                    {
                        if (remainingNumberOfClients == 0 && numberClientsInTable != 9)
                        {
                           index = (i * mMaxNumberGoogleMapsWaypoints) + m + 1;
                        }
                        else
                        {
                           index = (i * mMaxNumberGoogleMapsWaypoints) + m;
                        }
                        

                        textWriter.WriteLine("      <!-- Client Data -->");
                        textWriter.WriteLine("      <tr>");
                        textWriter.WriteLine("         <td>" + clients[index] + "</td>");
                        textWriter.WriteLine("         <td>" + names[index] + "</td>");
                        textWriter.WriteLine("         <td>" + streetAddresses[index] + "</td>");
                        textWriter.WriteLine("         <td>" + cities[index] + "</td>");
                        textWriter.WriteLine("         <td>" + zipCodes[index] + "</td>");
                        textWriter.WriteLine("         <td>" + phoneNumbers[index] + "</td>");
                        textWriter.WriteLine("         <td>" + organizations[index] + "</td>");
                        textWriter.WriteLine("      </tr>");
                    }

                    textWriter.WriteLine("");
                    textWriter.WriteLine("      </table>");
                    textWriter.WriteLine("      <!-- Client Specific Notes -->");
                    textWriter.WriteLine("      <h2 align = \"left\">Client Specific Notes</h2>");
                    textWriter.WriteLine("      <table align = \"left\" width=\"100%\" border=\"1\">");
                    textWriter.WriteLine("      <!-- Table Headings -->");
                    textWriter.WriteLine("      <tr>");
                    textWriter.WriteLine("         <th width = \"10%\">Client ID</th>");
                    textWriter.WriteLine("         <th width = \"20%\">Client Name</th>");
                    textWriter.WriteLine("         <th width = \"70%\">Comments</th>");
                    textWriter.WriteLine("      </tr>");

                    //Put each deliverer's client comments in the table
                    for (int n = 0; n < numberClientsInTable; n++)
                    {
                        if (remainingNumberOfClients == 0 && numberClientsInTable != 9)
                        {
                            index = (i * mMaxNumberGoogleMapsWaypoints) + n + 1;
                        }
                        else
                        {
                            index = (i * mMaxNumberGoogleMapsWaypoints) + n;
                        }

                        textWriter.WriteLine("      <!-- Client Data -->");
                        textWriter.WriteLine("      <tr>");
                        textWriter.WriteLine("         <td>" + clients[index] + "</td>");
                        textWriter.WriteLine("         <td>" + names[index] + "</td>");
                        textWriter.WriteLine("         <td>" + comments[index] + "</td>");
                        textWriter.WriteLine("      </tr>");
                    }

                    textWriter.WriteLine("      </table>");
                    textWriter.WriteLine("   </body>");
                    textWriter.WriteLine("</html>");

                    //Close the file stream
                    textWriter.Close();
                }
            }
        }

        private void GenerateUnassignedClientPacketsButton_Click(object sender, RoutedEventArgs e)
        {
            if (mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            List<string> generatedFiles = new List<string>();
            string getUnassignedClientsQuery = "SELECT * FROM " + Main.mSelectedYear + " WHERE Assigned_Status = 'false'";

            DataSet unassignedClients = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(getUnassignedClientsQuery, Main.mSelectedYear);

            if (unassignedClients != null)
            {
                if (unassignedClients.Tables[0].Rows.Count != 0)
                {
                    foreach (DataRow unassignedClientRow in unassignedClients.Tables[0].Rows)
                    {
                        Deliverer deliverer = new Deliverer();

                        deliverer.LastName = "Unassigned";
                        deliverer.FirstName = "Client";
                        deliverer.DelivererID = 0;
                        deliverer.Assigned = 1;
                        deliverer.Clients = unassignedClientRow["Client_ID"].ToString();

                        //Create the fileName
                        string[] fileNames = new string[1];
                        string[] fullAddresses = null;
                        string[] clients = null;
                        string[] names = null;
                        string[] streetAddresses = null;
                        string[] cities = null;
                        string[] zipCodes = null;
                        string[] phoneNumbers = null;
                        string[] organizations = null;
                        string[] comments = null;

                        PrepareDelivererPacketArguments(deliverer, ref fullAddresses, ref clients, ref names, ref streetAddresses, ref cities, ref zipCodes, ref phoneNumbers, ref organizations, ref comments);

                        //Define FileName
                        fileNames[0] = "Deliverer(" + deliverer.DelivererID + ")_" + deliverer.LastName + "_" + deliverer.FirstName + "_Client_ID(" + unassignedClientRow["Client_ID"].ToString() + ").htm";

                        GenerateDelivererPacket(deliverer, ref fileNames, ref fullAddresses, ref clients, ref names, ref streetAddresses, ref cities, ref zipCodes, ref phoneNumbers, ref organizations, ref comments);

                        //Add the file to the list of files that are generated

                        generatedFiles.Add(fileNames[0]);


                        //Since we've created the .htm client white card files with maps
                        //open them in firefox to print them

                        //Get the total number of deliverers to display in firefox
                        int deliverersToDisplayInFirefox = generatedFiles.Count;

                        //Create a variable for the client to start from
                        int startingDeliverer = 0;

                        //Have a variable to use for number to print in a single session
                        int sessionDelivererLimit = 100;

                        while (deliverersToDisplayInFirefox != 0)
                        {
                            //Check for the case where deliverersToDisplayInFirefox < sessionDelivererLimit
                            if (deliverersToDisplayInFirefox < sessionDelivererLimit)
                            {
                                sessionDelivererLimit = deliverersToDisplayInFirefox;
                            }

                            //Create a new process
                            System.Diagnostics.Process process = new System.Diagnostics.Process();

                            //The process will be firefox
                            process.StartInfo.FileName = "firefox.exe";

                            //Create firefox's command argument line
                            //Open all of the files we created in seperate tabs
                            for (int i = startingDeliverer; i < (startingDeliverer + sessionDelivererLimit); i++)
                            {
                                process.StartInfo.Arguments += "\"" + generatedFiles[i] + "\" ";
                            }

                            //Start Firefox
                            process.Start();

                            //Close our handle to Firefox
                            process.Close();

                            //Update deliverersToDisplayInFirefox
                            deliverersToDisplayInFirefox -= sessionDelivererLimit;

                            //Update startingDeliverer
                            startingDeliverer += sessionDelivererLimit;
                        }
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("All clients are assigned to deliverers");
                }
            }
        }
    }




    #region AccessDatabaseClasses

    /// <ChristmasBasketsDatabase>
    /// Abstract ChristmasBasketsDatabase Access database class
    /// </ChristmasBasketsDatabase>
    public class ChristmasBasketsAccessDatabase
    {
        //Define variables
        public System.Data.OleDb.OleDbConnection mChristmasBasketsAccessDatabase;
        public string mChristmasBasketsAccessDatabasePath;

        //Define Methods

        /// <OpenChristmasBasketsDatabase>
        /// Open the Christmas Baskets Access Database
        /// </OpenChristmasBasketsDatabase>
        public void OpenChristmasBasketsDatabase()
        {
            //Open the Christmas Baskets Access Database
            mChristmasBasketsAccessDatabase.Open();
        }

        /// <CloseChristmasBasketsDatabase>
        /// Close the Christmas Baskets Access Database
        /// </CloseChristmasBasketsDatabase>
        public void CloseChristmasBasketsDatabase()
        {
            //Close the Christmas Baskets Access Database
            mChristmasBasketsAccessDatabase.Close();
        }
        /// <GetSchema>
        /// Use this to get the list of tables in the database
        /// </GetSchema>
        /// <param name="collectionName"></param>
        /// <param name="restrictionValues"></param>
        public System.Data.DataTable GetSchema(string collectionName, string [] restrictionValues)
        {
            //Returns a Datatable
            return mChristmasBasketsAccessDatabase.GetSchema(collectionName, restrictionValues);
        }

        /// <PerformSelectQuery>
        /// Perform a Select Query on the datablase
        /// </PerformSelectQuery>
        /// <param name="selectQuery"></param>
        /// <param name="sourceTableName"></param>
        /// <returns></returns>
        public System.Data.DataSet PerformSelectQuery(string selectQueryCommandText, string sourceTableName)
        {
            //Create a OleDb adapter
            OleDbDataAdapter adapter = new OleDbDataAdapter();

            //Create a selectQueryCommand
            OleDbCommand selectQueryCommand = new OleDbCommand(selectQueryCommandText, mChristmasBasketsAccessDatabase);
            
            //Create a data set to fill from the database
            DataSet selectQueryDataSet = new DataSet();

            //Set the adapter's select command to the select query command object
            adapter.SelectCommand = selectQueryCommand;

            //Perform the select query command
            int numberOfRecordsFound = adapter.Fill(selectQueryDataSet, sourceTableName);

            //Make sure we found something
            if (numberOfRecordsFound > 0)
            {
                //Return the results DataSet
                return selectQueryDataSet;
            }
            else
            {
                //Return null since we found no records
                return null;
            }
        }

        /// <ExecuteNonQuery>
        /// Execute a Non Query on the database
        /// </ExecuteNonQuery>
        /// <param name="nonQueryCommandText"></param>
        /// <returns></returns>
        public int ExecuteNonQuery(string nonQueryCommandText)
        {
            //Create a nonQueryCommand
            OleDbCommand nonQueryCommand = new OleDbCommand(nonQueryCommandText, mChristmasBasketsAccessDatabase);

            //Perform the nonQueryCommand and return a status values
            return nonQueryCommand.ExecuteNonQuery();
        }
    }

    /// <ChristmasBasketsAccess2003Database>
    /// ChristmasBasketsDatabase Access 2003 database class
    /// </ChristmasBasketsAccess2003Database>
    public class ChristmasBasketsAccess2003Database : ChristmasBasketsAccessDatabase
    {
        //Define Methods

        /// <ChristmasBasketsAccess2007Database>
        /// Constructor
        /// </ChristmasBasketsAccess2007Database>
        /// <param name="ChristmasBasketsAccessDatabasePathToSet"></param>
        public ChristmasBasketsAccess2003Database(string ChristmasBasketsAccessDatabasePathToSet)
        {
            //Initialize ChristmasBasketsAccessDatabasePath
            mChristmasBasketsAccessDatabasePath = ChristmasBasketsAccessDatabasePathToSet;

            //Try and open the Christmas Baskets Access Database
            mChristmasBasketsAccessDatabase = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @mChristmasBasketsAccessDatabasePath);

        }
    }

    /// <ChristmasBasketsAccess2007Database>
    /// ChristmasBasketsDatabase Access 2007 database class
    /// </ChristmasBasketsAccess2007Database>
    public class ChristmasBasketsAccess2007Database : ChristmasBasketsAccessDatabase
    {
        //Define Methods

        /// <ChristmasBasketsAccess2007Database>
        /// Constructor
        /// </ChristmasBasketsAccess2007Database>
        /// <param name="ChristmasBasketsAccessDatabasePathToSet"></param>
        public ChristmasBasketsAccess2007Database(string ChristmasBasketsAccessDatabasePathToSet)
        {
            //Initialize ChristmasBasketsAccessDatabasePath
            mChristmasBasketsAccessDatabasePath = ChristmasBasketsAccessDatabasePathToSet;

            //Try and open the Christmas Baskets Access Database
            mChristmasBasketsAccessDatabase = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @mChristmasBasketsAccessDatabasePath);
        }
    }

    #endregion


    #region Client Class

    public class Client
    {
        public int Client_ID;
        public string Last_Name;
        public string First_Name;
        public string Middle_Name;
        public string Title;
        public string Address_Number;
        public string Street_Address;
        public string City;
        public string Zipcode;
        public string Phone;
        public string Organization;
        public string Directions;
        public string Instructions;
        public string Deliverer_ID;
        public string Year_Last_Delivered;

        /// <Client>
        /// Default Constructor
        /// </Client>
        public Client()
        {
            Client_ID = -1;
            Last_Name = "";
            First_Name = "";
            Middle_Name = "";
            Title = "";
            Address_Number = "";
            Street_Address = "";
            City = "";
            Zipcode = "";
            Phone = "";
            Organization = "";
            Directions = "";
            Instructions = "";
            Deliverer_ID = "";
            Year_Last_Delivered = "";
        }

        /// <Client>
        /// Constructor for Client coming from an excel spreadsheet
        /// </Client>
        /// <param name="Last_NameToSet"></param>
        /// <param name="First_NameToSet"></param>
        /// <param name="Middle_NameToSet"></param>
        /// <param name="TitleToSet"></param>
        /// <param name="Address_NumberToSet"></param>
        /// <param name="Street_AddressToSet"></param>
        /// <param name="CityToSet"></param>
        /// <param name="ZipcodeToSet"></param>
        /// <param name="PhoneToSet"></param>
        /// <param name="OrganizationToSet"></param>
        public Client(string Last_NameToSet, string First_NameToSet, string Middle_NameToSet, string TitleToSet, string Address_NumberToSet, string Street_AddressToSet, string CityToSet, string ZipcodeToSet, string PhoneToSet, string OrganizationToSet)
        {
            Client_ID = -1;

            Last_Name = Last_NameToSet;
            First_Name = First_NameToSet;
            Middle_Name = Middle_NameToSet;
            Title = TitleToSet;
            Address_Number = Address_NumberToSet;
            Street_Address = Street_AddressToSet;
            City = CityToSet;
            Zipcode = ZipcodeToSet;
            Phone = PhoneToSet;
            Organization = OrganizationToSet;

            Directions = "";
            Instructions = "";
            Deliverer_ID = "";
            Year_Last_Delivered = "";
        }

        /// <Client>
        /// Constructor for Client coming from the access database
        /// </Client>
        /// <param name="Client_IDToSet"></param>
        /// <param name="Last_NameToSet"></param>
        /// <param name="First_NameToSet"></param>
        /// <param name="Middle_NameToSet"></param>
        /// <param name="TitleToSet"></param>
        /// <param name="Address_NumberToSet"></param>
        /// <param name="Street_AddressToSet"></param>
        /// <param name="CityToSet"></param>
        /// <param name="ZipcodeToSet"></param>
        /// <param name="PhoneToSet"></param>
        /// <param name="OrganizationToSet"></param>
        /// <param name="DirectionsToSet"></param>
        /// <param name="InstructionsToSet"></param>
        /// <param name="Deliverer_IDToSet"></param>
        /// <param name="Year_Last_DeliveredToSet"></param>
        public Client(int Client_IDToSet, string Last_NameToSet, string First_NameToSet, string Middle_NameToSet, string TitleToSet, string Address_NumberToSet, string Street_AddressToSet, string CityToSet, string ZipcodeToSet, string PhoneToSet, string OrganizationToSet, string DirectionsToSet, string InstructionsToSet, string Deliverer_IDToSet, string Year_Last_DeliveredToSet)
        {
            Client_ID = Client_IDToSet;
            Last_Name = Last_NameToSet;
            First_Name = First_NameToSet;
            Middle_Name = Middle_NameToSet;
            Title = TitleToSet;
            Address_Number = Address_NumberToSet;
            Street_Address = Street_AddressToSet;
            City = CityToSet;
            Zipcode = ZipcodeToSet;
            Phone = PhoneToSet;
            Organization = OrganizationToSet;
            Directions = DirectionsToSet;
            Instructions = InstructionsToSet;
            Deliverer_ID = Deliverer_IDToSet;
            Year_Last_Delivered = Year_Last_DeliveredToSet;
        }

    }

    #endregion


    #region  ExcelMethods

    static class ExcelMethods
    {

        public static void ExportToExcel(DataSet dataSet, string worksheetName)
        {
            // Create the Excel Application object
            ApplicationClass excelApp = new ApplicationClass();

            // Create a new Excel Workbook
            Workbook excelWorkbook = excelApp.Workbooks.Add(Type.Missing);

            int sheetIndex = 0;

            // Copy each DataTable
            foreach (System.Data.DataTable dt in dataSet.Tables)
            {

                // Copy the DataTable to an object array
                object[,] rawData = new object[dt.Rows.Count + 1, dt.Columns.Count];

                // Copy the column names to the first row of the object array
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    rawData[0, col] = dt.Columns[col].ColumnName;
                }

                // Copy the values to the object array
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        rawData[row + 1, col] = dt.Rows[row].ItemArray[col];
                    }
                }

                // Calculate the final column letter
                string finalColLetter = string.Empty;
                string colCharset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                int colCharsetLen = colCharset.Length;

                if (dt.Columns.Count > colCharsetLen)
                {
                    finalColLetter = colCharset.Substring(
                        (dt.Columns.Count - 1) / colCharsetLen - 1, 1);
                }

                finalColLetter += colCharset.Substring(
                        (dt.Columns.Count - 1) % colCharsetLen, 1);

                // Create a new Sheet
                Worksheet excelSheet = (Worksheet)excelWorkbook.Sheets.Add(
                    excelWorkbook.Sheets.get_Item(++sheetIndex),
                    Type.Missing, 1, XlSheetType.xlWorksheet);

                excelSheet.Name = dt.TableName;

                // Fast data export to Excel
                string excelRange = string.Format("A1:{0}{1}", finalColLetter, dt.Rows.Count + 1);

                Range cell = excelSheet.get_Range(excelRange, Type.Missing);
                
                //Make all cells text formatted
                cell.NumberFormat = "@";

                excelSheet.get_Range(excelRange, Type.Missing).Value2 = rawData;

                // Mark the first row as BOLD
                ((Range)excelSheet.Rows[1, Type.Missing]).Font.Bold = true;

                //AutoFit the columns
                excelSheet.Columns.AutoFit();
            }

            // Save and Close the Workbook
            SaveExcelWorkbook(excelWorkbook, worksheetName);

            //Set the workbook to null
            excelWorkbook = null;

            // Release the Application object
            excelApp.Quit();
            excelApp = null;

            // Collect the unreferenced objects
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }

        public static void SaveExcelWorkbook(Workbook excelWorkbook, string worksheetName)
        {
            //Define local variables
            SaveFileDialog saveExcelFileDialog = new SaveFileDialog();

            //Set the file name
            saveExcelFileDialog.FileName = worksheetName;

            //Customize filter
            saveExcelFileDialog.Filter = "Excel Office 2003 file (*.xls)|*.xls|Excel Office 2007 file (*.xlsx)|*.xlsx";

            //Customize filter index
            saveExcelFileDialog.FilterIndex = 1;

            //Customize restore director
            saveExcelFileDialog.RestoreDirectory = true;

            //Show the dialog
            DialogResult dialogResult = saveExcelFileDialog.ShowDialog();

            if (dialogResult != DialogResult.Cancel)
            {
                //Save the worksheet
                object missingOption = System.Reflection.Missing.Value;

                //Check for Excel Office 2007 file format
                if (saveExcelFileDialog.FileName.IndexOf(".xlsx") > -1)
                {
                    //Save the worksheet as an Excel Office 2007 .xlsx file
                    excelWorkbook.SaveAs(saveExcelFileDialog.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                                         missingOption, missingOption, missingOption, missingOption,
                                         Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingOption,
                                         missingOption, missingOption, missingOption, missingOption);
                }
                //Check for Excel Office 2003 file format
                else if (saveExcelFileDialog.FileName.IndexOf(".xls") > -1)
                {
                    //Save the worksheet as an Excel Office 2003 .xls file   
                    excelWorkbook.SaveAs(saveExcelFileDialog.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8,
                                         missingOption, missingOption, missingOption, missingOption,
                                         Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingOption,
                                         missingOption, missingOption, missingOption, missingOption);
                }
            }
        }

        public static DataSet OpenExcelWorkbookAndExtractWorksheetInformation(string selectQueryText)
        {
            //Define local variables
            OpenFileDialog openExcelFileDialog = new OpenFileDialog();
            DataSet toReturn = new DataSet();
            string connectionString = "";

            //Customize filter
            openExcelFileDialog.Filter = "Excel Office 2003 file (*.xls)|*.xls|Excel Office 2007 file (*.xlsx)|*.xlsx";

            //Customize filter index
            openExcelFileDialog.FilterIndex = 1;

            //Customize restore director
            openExcelFileDialog.RestoreDirectory = true;

            //Show the dialog
            DialogResult dialogResult = openExcelFileDialog.ShowDialog();

            if (dialogResult != DialogResult.Cancel)
            {
                //Open the worksheet
                object missingOption = System.Reflection.Missing.Value;

                //Check for Excel Office 2007 and 2010 file format
                if (openExcelFileDialog.FileName.IndexOf(".xlsx") > -1)
                {
                    //Save the worksheet as an Excel Office 2007 .xlsx file
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        "Data Source=" + openExcelFileDialog.FileName + ";" +
                                        "Extended Properties=Excel 12.0;";

                }
                //Check for Excel Office 2003 file format
                else if (openExcelFileDialog.FileName.IndexOf(".xls") > -1)
                {
                    //Save the worksheet as an Excel Office 2003 .xls file   
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        "Data Source=" + openExcelFileDialog.FileName + ";" +
                                        "Extended Properties=Excel 12.0;";
                }

                // Put user code to initialize the page here
                // Create connection string variable. Modify the "Data Source"
                // parameter as appropriate for your environment.


                // Create connection object by using the preceding connection string.
                OleDbConnection objConn = new OleDbConnection(connectionString);

                // Open connection with the database.
                objConn.Open();

                // The code to follow uses a SQL SELECT command to display the data from the worksheet.

                // Create new OleDbCommand to return data from worksheet.
                OleDbCommand objCmdSelect = new OleDbCommand(selectQueryText, objConn);

                // Create new OleDbDataAdapter that is used to build a DataSet
                // based on the preceding SQL SELECT statement.
                OleDbDataAdapter objAdapter = new OleDbDataAdapter();

                // Pass the Select command to the adapter.
                objAdapter.SelectCommand = objCmdSelect;


                // Fill the DataSet with the information from the worksheet.
                objAdapter.Fill(toReturn);

                // Clean up objects.
                objConn.Close();
            }

            return toReturn;

        }

        public static DataSet OpenExcelWorkbookAndExtractWorksheetInformation(string selectQueryText, string worksheetPath)
        {
            //Define local variables
            DataSet toReturn = new DataSet();
            string connectionString = "";

            if (worksheetPath != "")
            {
                //Open the worksheet
                object missingOption = System.Reflection.Missing.Value;

                //Check for Excel Office 2007 and 2010 file format
                if (worksheetPath.IndexOf(".xlsx") > -1)
                {
                    //Save the worksheet as an Excel Office 2007 .xlsx file
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        "Data Source=" + worksheetPath + ";" +
                                        "Extended Properties=Excel 12.0;";

                }
                //Check for Excel Office 2003 file format
                else if (worksheetPath.IndexOf(".xls") > -1)
                {
                    //Save the worksheet as an Excel Office 2003 .xls file   
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        "Data Source=" + worksheetPath + ";" +
                                        "Extended Properties=Excel 12.0;";
                }

                // Put user code to initialize the page here
                // Create connection string variable. Modify the "Data Source"
                // parameter as appropriate for your environment.


                // Create connection object by using the preceding connection string.
                OleDbConnection objConn = new OleDbConnection(connectionString);

                // Open connection with the database.
                objConn.Open();

                // The code to follow uses a SQL SELECT command to display the data from the worksheet.

                // Create new OleDbCommand to return data from worksheet.
                OleDbCommand objCmdSelect = new OleDbCommand(selectQueryText, objConn);

                // Create new OleDbDataAdapter that is used to build a DataSet
                // based on the preceding SQL SELECT statement.
                OleDbDataAdapter objAdapter = new OleDbDataAdapter();

                // Pass the Select command to the adapter.
                objAdapter.SelectCommand = objCmdSelect;


                // Fill the DataSet with the information from the worksheet.
                objAdapter.Fill(toReturn);

                // Clean up objects.
                objConn.Close();
            }

            return toReturn;

        }


        public static DataSet OpenExcelWorkbookAndExtractWorksheetInformationAndSaveWorksheetPath(string selectQueryText, ref string worksheetPath)
        {
            //Define local variables
            OpenFileDialog openExcelFileDialog = new OpenFileDialog();
            DataSet toReturn = new DataSet();
            string connectionString = "";

            //Customize filter
            openExcelFileDialog.Filter = "Excel Office 2003 file (*.xls)|*.xls|Excel Office 2007 file (*.xlsx)|*.xlsx";

            //Customize filter index
            openExcelFileDialog.FilterIndex = 1;

            //Customize restore director
            openExcelFileDialog.RestoreDirectory = true;

            //Show the dialog
            DialogResult dialogResult = openExcelFileDialog.ShowDialog();

            if (dialogResult != DialogResult.Cancel)
            {
                //Open the worksheet
                object missingOption = System.Reflection.Missing.Value;

                //Check for Excel Office 2007 and 2010 file format
                if (openExcelFileDialog.FileName.IndexOf(".xlsx") > -1)
                {
                    //Save the worksheet as an Excel Office 2007 .xlsx file
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        "Data Source=" + openExcelFileDialog.FileName + ";" +
                                        "Extended Properties=Excel 12.0;";

                    //Save path to be passed back by reference
                    worksheetPath = openExcelFileDialog.FileName;

                }
                //Check for Excel Office 2003 file format
                else if (openExcelFileDialog.FileName.IndexOf(".xls") > -1)
                {
                    //Save the worksheet as an Excel Office 2003 .xls file   
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                        "Data Source=" + openExcelFileDialog.FileName + ";" +
                                        "Extended Properties=Excel 12.0;";

                    //Save path to be passed back by reference
                    worksheetPath = openExcelFileDialog.FileName;
                }

                // Put user code to initialize the page here
                // Create connection string variable. Modify the "Data Source"
                // parameter as appropriate for your environment.


                // Create connection object by using the preceding connection string.
                OleDbConnection objConn = new OleDbConnection(connectionString);

                // Open connection with the database.
                objConn.Open();

                // The code to follow uses a SQL SELECT command to display the data from the worksheet.

                // Create new OleDbCommand to return data from worksheet.
                OleDbCommand objCmdSelect = new OleDbCommand(selectQueryText, objConn);

                // Create new OleDbDataAdapter that is used to build a DataSet
                // based on the preceding SQL SELECT statement.
                OleDbDataAdapter objAdapter = new OleDbDataAdapter();

                // Pass the Select command to the adapter.
                objAdapter.SelectCommand = objCmdSelect;


                // Fill the DataSet with the information from the worksheet.
                objAdapter.Fill(toReturn);

                // Clean up objects.
                objConn.Close();
            }

            return toReturn;

        }


    }

    #endregion
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Data;
using System.Data.OleDb;

namespace ChristmasBasketsDashboard
{
    public enum ClientMode
    {
        None = 0,
        DatabaseClientsTableDuplicates = 1,
        DatabaseClientsTableUpdateClient = 2,
        DatabaseClientsTableAddClient = 3
    }

    /// <summary>
    /// Interaction logic for ImportClientsFromExcelHelper.xaml
    /// </summary>
    public partial class ImportClientsFromExcelHelper : Window
    {
        public System.Data.DataTable mDatabaseClientsTable;
        public System.Data.DataTable mExcelClientsTable;
        public System.Data.DataTable mDatabaseSelectedYearClientsTable;
        public string mActionLabelText;
        public string mTopTableLabelText;
        public string mBottomTableLabelText;
        public ClientMode mMode;

        public ImportClientsFromExcelHelper()
        {
            InitializeComponent();
        }

        public void SetDatabaseClientsTable(DataTable databaseClientsTableToSet)
        {
            //Get a handle to the DataTable
            mDatabaseClientsTable = databaseClientsTableToSet;
        }

        public void SetExcelClientsTable(DataRow excelClientDataRow, DataColumnCollection columns)
        {
            //Create a DataTable
            DataTable excelClient = new DataTable();

            //Copy column names into DataTable
            foreach (DataColumn column in columns)
            {
                excelClient.Columns.Add(column.ColumnName);
            }
            
            //Import the client into the DataTable
            excelClient.ImportRow(excelClientDataRow);

            //Get a handle to the DataTable
            mExcelClientsTable = excelClient;
        }

        public void SetDatabaseSelectedYearClientsTable(DataTable databaseSelectedYearClientsTableToSet)
        {
            //Get a handle to the DataTable
            mDatabaseSelectedYearClientsTable = databaseSelectedYearClientsTableToSet;
        }

        public void SetMode(ClientMode modeToSet)
        {
            //Create a new mode
            mMode = new ClientMode();

            //Store mode
            mMode = modeToSet;
        }

        public void SetActionLabelText(string actionLabelTextToSet)
        {
            //Store the new action label text
            mActionLabelText = actionLabelTextToSet;

            //Display the new text.
            actionLabel.Content = mActionLabelText;
        }

        public void SetTopTableLabelText(string topTableLabelTextToSet)
        {
            //Store the new top table label text
            mTopTableLabelText = topTableLabelTextToSet;

            //Display the new text.
            TopTableLabel.Content = mTopTableLabelText;
        }

        public void SetBottomTableLabelText(string bottomTableLabelTextToSet)
        {
            //Store the new bottom table label text
            mBottomTableLabelText = bottomTableLabelTextToSet;

            //Display the new text.
            BottomTableLabel.Content = mBottomTableLabelText;
        }


        public void BindData()
        {
            ///////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////
            //////////////////      BASED ON MODE                  ////////////////
            ///////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////


            if (mMode == ClientMode.DatabaseClientsTableDuplicates)
            {
                SetActionLabelText("Handle Duplicate Records found in the Clients Database Table Then Add To " + Main.mSelectedYear + " Table");
                SetTopTableLabelText("Excel Client Being Imported");
                SetBottomTableLabelText("Database Clients Table Matches");

                //Bind the TopTable DataGrid to mExcelClientsTable
                TopTable.DataSource = mExcelClientsTable;

                //Bind the BottomTable DataGrid to mDatabaseClientsTable
                BottomTable.DataSource = mDatabaseClientsTable;

                DeleteSelectedClientsButton.Content = "Delete Selected Client from Database Clients Table";
                AddSelectedClientsButton.Content = "Add Selected Client to Database " + Main.mSelectedYear + " Table";
                UpdateClientInformationButton.Content = "Update Selected Client in Database Clients Table";
            }
            else if (mMode == ClientMode.DatabaseClientsTableUpdateClient)
            {
                SetActionLabelText("Update Single Record found in the Clients Database Table Then Add To " + Main.mSelectedYear + " Table");
                SetTopTableLabelText("Excel Client Being Imported");
                SetBottomTableLabelText("Database Client Table Match");

                //Bind the TopTable DataGrid to mExcelClientsTable
                TopTable.DataSource = mExcelClientsTable;

                //Bind the BottomTable DataGrid to mDatabaseClientsTable
                BottomTable.DataSource = mDatabaseClientsTable;

                DeleteSelectedClientsButton.Content = "Delete Selected Client from Database Clients Table";
                AddSelectedClientsButton.Content = "Add Selected Client to Database " + Main.mSelectedYear + " Table";
                UpdateClientInformationButton.Content = "Update Selected Client in Database Clients Table";
                
                //Hide this button
                AddSelectedClientToClientsAndSelectedTableButton.Visibility = Visibility.Hidden;
            }
            else if (mMode == ClientMode.DatabaseClientsTableAddClient)
            {
                SetActionLabelText("Add a Single Record to Clients Database Table Then Add To " + Main.mSelectedYear + " Table");
                SetTopTableLabelText("Excel Client Being Imported");
                SetBottomTableLabelText("Database Client Table Match");

                //Bind the TopTable DataGrid to mExcelClientsTable
                TopTable.DataSource = mExcelClientsTable;

                //Bind the BottomTable DataGrid to mDatabaseClientsTable
                BottomTable.DataSource = mDatabaseClientsTable;

                DeleteSelectedClientsButton.Content = "Delete Selected Client from Database Clients Table";
                AddSelectedClientsButton.Content = "Add Selected Client to Database " + Main.mSelectedYear + " Table"; ;
                UpdateClientInformationButton.Content = "Update Selected Client in Database Clients Table";
                AddSelectedClientToClientsAndSelectedTableButton.Content = "Add Selected Client to Database Clients and " + Main.mSelectedYear + " Tables";

                //Show this button
                AddSelectedClientToClientsAndSelectedTableButton.Visibility = Visibility.Visible;
            }
        }

        public void UnBindData()
        {
            //UnBind the BottomTable DataGrid
            BottomTable.DataSource = null;

            //UnBind the TopTable DataGrid
            TopTable.DataSource = null;
        }

        private void DeleteSelectedClientButton_Click(object sender, RoutedEventArgs e)
        {
            if (mMode == ClientMode.DatabaseClientsTableDuplicates || mMode == ClientMode.DatabaseClientsTableUpdateClient || mMode == ClientMode.DatabaseClientsTableAddClient || mMode == ClientMode.DatabaseClientsTableAddClient)
            {
                int rowToDelete = -1;
                int selectedRowCount = 0;
                string clientID = "";
                string firstName = "";
                string lastName = "";

                //See see if only 1 row is selected and get that row number
                for (int i = 0; i < mDatabaseClientsTable.Rows.Count; i++)
                {
                    if (BottomTable.IsSelected(i))
                    {
                        rowToDelete = i;
                        selectedRowCount++;
                    }
                }

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 client");
                    return;
                }

                //See if we should delete the client into the Clients table
                if (rowToDelete != -1)
                {
                    //Get information to display to the user for confirmation
                    clientID = mDatabaseClientsTable.Rows[rowToDelete]["Client_ID"].ToString();
                    lastName = mDatabaseClientsTable.Rows[rowToDelete]["Last_Name"].ToString();
                    firstName = mDatabaseClientsTable.Rows[rowToDelete]["First_Name"].ToString();

                    //Update Access Database
                    string deleteCommand = "DELETE FROM Clients WHERE Client_ID = " + mDatabaseClientsTable.Rows[rowToDelete]["Client_ID"];
                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(deleteCommand);

                    //Update data grid source
                    mDatabaseClientsTable.Rows.RemoveAt(rowToDelete);
                }
                
                //Update data grid display
                TopTable.Refresh();
                BottomTable.Refresh();

                if (rowToDelete != -1)
                {
                    //Show status message box
                    System.Windows.MessageBox.Show("Client:  " + clientID + " - " + lastName + ", " + firstName + " DELETED from Clients Table");
                }
            }
        }

        private void AddClientToSelectedYearTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (mMode == ClientMode.DatabaseClientsTableDuplicates || mMode == ClientMode.DatabaseClientsTableUpdateClient || mMode == ClientMode.DatabaseClientsTableAddClient)
            {
                int rowToInsert = -1;
                int selectedRowCount = 0;
                string clientID = "";
                string firstName = "";
                string lastName = "";

                //See see if only 1 row is selected and get that row number
                for (int i = 0; i < mDatabaseClientsTable.Rows.Count; i++)
                {
                    if (BottomTable.IsSelected(i))
                    {
                        rowToInsert = i;
                        selectedRowCount++;
                    }
                }

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 client");
                    return;
                }

                //See if we should insert the client into the mSelectedYear table
                if (rowToInsert != -1)
                {
                    //Get information to display to the user for confirmation
                    clientID = mDatabaseClientsTable.Rows[rowToInsert]["Client_ID"].ToString();
                    lastName = mDatabaseClientsTable.Rows[rowToInsert]["Last_Name"].ToString();
                    firstName = mDatabaseClientsTable.Rows[rowToInsert]["First_Name"].ToString();

                    //Update Access Database
                    string insertCommand = "INSERT INTO " + Main.mSelectedYear + " (Box_Number,Client_ID) VALUES (-1," + mDatabaseClientsTable.Rows[rowToInsert]["Client_ID"] + ")";
                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommand);
                }

                //Update data grid display
                TopTable.Refresh();
                BottomTable.Refresh();

                if (rowToInsert != -1)
                {
                    //Show status message box
                    System.Windows.MessageBox.Show("Client:  " + clientID + " - " + lastName + ", " + firstName + " ADDED in " + Main.mSelectedYear + " Table");

                    //Close the Import Helper
                    this.Close();
                }
            }
        }

        private void UpdateClientInformationButton_Click(object sender, RoutedEventArgs e)
        {
            if (mMode == ClientMode.DatabaseClientsTableDuplicates || mMode == ClientMode.DatabaseClientsTableUpdateClient || mMode == ClientMode.DatabaseClientsTableAddClient)
            {
                int rowToUpdate = -1;
                int selectedRowCount = 0;
                string clientID = "";
                string firstName = "";
                string lastName = "";

                //See see if only 1 row is selected and get that row number
                for (int i = 0; i < mDatabaseClientsTable.Rows.Count; i++)
                {
                    if (BottomTable.IsSelected(i))
                    {
                        rowToUpdate = i;
                        selectedRowCount++;
                    }
                }

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 client");
                    return;
                }

                //See if we should update the client in the Clients table
                if (rowToUpdate != -1)
                {
                    //Get information to display to the user for confirmation
                    clientID = mDatabaseClientsTable.Rows[rowToUpdate]["Client_ID"].ToString();
                    lastName = mDatabaseClientsTable.Rows[rowToUpdate]["Last_Name"].ToString();
                    firstName = mDatabaseClientsTable.Rows[rowToUpdate]["First_Name"].ToString();

                    //Check for null values
                    foreach (DataColumn column in mDatabaseClientsTable.Columns)
                    {
                        if (mDatabaseClientsTable.Rows[rowToUpdate][column.ColumnName].ToString() == "")
                        {
                            mDatabaseClientsTable.Rows[rowToUpdate][column.ColumnName] = "null";
                        }
                    }

                    string updateCommand = "UPDATE Clients SET Last_Name = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Last_Name"].ToString() + "', " +
                                               "First_Name = '" + mDatabaseClientsTable.Rows[rowToUpdate]["First_Name"].ToString() + "', " +
                                               "Middle_Name = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Middle_Name"].ToString() + "', " +
                                               "Title = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Title"].ToString() + "', " +
                                               "Address_Number = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Address_Number"].ToString() + "', " +
                                               "Street_Address = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Street_Address"].ToString() + "', " +
                                               "City = '" + mDatabaseClientsTable.Rows[rowToUpdate]["City"].ToString() + "', " +
                                               "Zipcode = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Zipcode"].ToString() + "', " +
                                               "Phone = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Phone"].ToString() + "', " +
                                               "Organization = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Organization"].ToString() + "', " +
                                               "Directions = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Directions"].ToString() + "', " +
                                               "Instructions = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Instructions"].ToString() + "', " +
                                               "Year_Last_Delivered_To = '" + mDatabaseClientsTable.Rows[rowToUpdate]["Year_Last_Delivered_To"].ToString() + "'" +
                                               " WHERE Client_ID = " + mDatabaseClientsTable.Rows[rowToUpdate]["Client_ID"].ToString();

                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCommand);
                }

                //Update data grid display
                TopTable.Refresh();
                BottomTable.Refresh();

                if (rowToUpdate != -1)
                {
                    //Show Status message box
                    System.Windows.MessageBox.Show("Client:  " + clientID + " - " + lastName + ", " + firstName + " UPDATED in Clients Table");
                }
            }
        }

        private void AddSelectedClientToClientsAndSelectedYearTableButton_Click(object sender, RoutedEventArgs e)
        {
            int rowToAdd = -1;
            int selectedRowCount = 0;
            string clientID = "";
            string firstName = "";
            string lastName = "";

            if (mMode == ClientMode.DatabaseClientsTableAddClient)
            {

                //Check for null values
                for (int i = 0; i < mExcelClientsTable.Rows.Count; i++)
                {
                    if (TopTable.IsSelected(i))
                    {
                        rowToAdd = i;
                        selectedRowCount++;
                    }
                }

                //Add the record into the Clients Table

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 client");
                    return;
                }

                //See if we should add the client to the Clients table
                if (rowToAdd != -1)
                {

                    //Check for null values
                    foreach (DataColumn column in mExcelClientsTable.Columns)
                    {
                        if (mExcelClientsTable.Rows[rowToAdd][column.ColumnName].ToString() == "")
                        {
                            mExcelClientsTable.Rows[rowToAdd][column.ColumnName] = "null";
                        }
                    }

                    //Update Access Database
                    string insertCommand = "INSERT INTO Clients (Last_Name,First_Name,Middle_Name,Title,Address_Number,Street_Address,City,Zipcode,Phone,Organization,Directions,Instructions,Year_Last_Delivered_To)" +
                                           " VALUES (" +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Last_Name"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["First_Name"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Middle_Name"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Title"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Address_Number"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Street_Address"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["City"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Zipcode"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Phone"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Organization"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Directions"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Instructions"].ToString() + "'," +
                                           "'" + mExcelClientsTable.Rows[rowToAdd]["Year_Last_Delivered_To"].ToString() + "')";
                    
                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommand);

                    //Figure out what the Client_ID was assigned
                    string selectCommand = "SELECT * FROM Clients WHERE " +
                                            "Last_Name = '" + mExcelClientsTable.Rows[rowToAdd]["Last_Name"].ToString() + "' AND " +
                                            "First_Name = '" + mExcelClientsTable.Rows[rowToAdd]["First_Name"].ToString() + "' AND " +
                                            "Middle_Name = '" + mExcelClientsTable.Rows[rowToAdd]["Middle_Name"].ToString() + "' AND " +
                                            "Title = '" + mExcelClientsTable.Rows[rowToAdd]["Title"].ToString() + "' AND " +
                                            "Address_Number = '" + mExcelClientsTable.Rows[rowToAdd]["Address_Number"].ToString() + "' AND " +
                                            "Street_Address = '" + mExcelClientsTable.Rows[rowToAdd]["Street_Address"].ToString() + "' AND " +
                                            "City = '" + mExcelClientsTable.Rows[rowToAdd]["City"].ToString() + "' AND " +
                                            "Zipcode = '" + mExcelClientsTable.Rows[rowToAdd]["Zipcode"].ToString() + "' AND " +
                                            "Phone = '" + mExcelClientsTable.Rows[rowToAdd]["Phone"].ToString() + "' AND " +
                                            "Organization = '" + mExcelClientsTable.Rows[rowToAdd]["Organization"].ToString() + "' AND " +
                                            "Directions = '" + mExcelClientsTable.Rows[rowToAdd]["Directions"].ToString() + "' AND " +
                                            "Instructions = '" + mExcelClientsTable.Rows[rowToAdd]["Instructions"].ToString() + "' AND " +
                                            "Year_Last_Delivered_To = '" + mExcelClientsTable.Rows[rowToAdd]["Year_Last_Delivered_To"].ToString() + "'";

                    
                    DataSet Client = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectCommand, "Clients");

                    if (Client != null)
                    {
                        //Get the record information to display to the user
                        clientID = Client.Tables[0].Rows[0]["Client_ID"].ToString();
                        lastName = mExcelClientsTable.Rows[rowToAdd]["Last_Name"].ToString();
                        firstName = mExcelClientsTable.Rows[rowToAdd]["First_Name"].ToString();

                        System.Windows.MessageBox.Show("Client:  " + clientID + " - " + lastName + ", " + firstName + " ADDED in Clients Table");

                        //Copy Record over into mDatabaseClientsTable
                        mDatabaseClientsTable.Rows.Add();
                        foreach (DataColumn column in mExcelClientsTable.Columns)
                        {
                            mDatabaseClientsTable.Rows[0][column.ColumnName] = mExcelClientsTable.Rows[0][column.ColumnName];
                        }

                        mDatabaseClientsTable.Rows[0]["Client_ID"] = clientID;

                        //Insert the record to the Table
                        string insertCommandTwo = "INSERT INTO " + Main.mSelectedYear + " (Box_Number,Client_ID) VALUES (-1," + clientID + ")";
                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommandTwo);

                        //Show status message box
                        System.Windows.MessageBox.Show("Client:  " + clientID + " - " + lastName + ", " + firstName + " ADDED in " + Main.mSelectedYear + " Table");

                        //Close the Import Helper
                        this.Close();
                    }

                    TopTable.Refresh();
                    BottomTable.Refresh();
                }
            }
        }

        private void QuitClientImport_Click(object sender, RoutedEventArgs e)
        {
            //Stop Client Import
            Main.mQuitClientImport = true;

            //Close this form
            this.Close();
        }

        private void Done_Click(object sender, RoutedEventArgs e)
        {
            //Continue Client Import
            Main.mQuitClientImport = false;

            //Close this form
            this.Close();
        }
    }
}
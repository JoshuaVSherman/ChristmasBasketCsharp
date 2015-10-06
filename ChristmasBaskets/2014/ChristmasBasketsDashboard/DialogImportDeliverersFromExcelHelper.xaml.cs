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
    public enum DelivererMode
    {
        None = 0,
        DatabaseDeliverersTableDuplicates = 1,
        DatabaseDeliverersTableUpdateDeliverer = 2,
        DatabaseDeliverersTableAddDeliverer = 3
    }

    /// <summary>
    /// Interaction logic for ImportDeliverersFromExcelHelper.xaml
    /// </summary>
    public partial class ImportDeliverersFromExcelHelper : Window
    {
        public System.Data.DataTable mDatabaseDeliverersTable;
        public System.Data.DataTable mExcelDeliverersTable;
        public System.Data.DataTable mDatabaseSelectedYearDeliverersTable;
        public string mActionLabelText;
        public string mTopTableLabelText;
        public string mBottomTableLabelText;
        public DelivererMode mMode;

        public ImportDeliverersFromExcelHelper()
        {
            InitializeComponent();
        }

        public void SetDatabaseDeliverersTable(DataTable databaseDeliverersTableToSet)
        {
            //Get a handle to the DataTable
            mDatabaseDeliverersTable = databaseDeliverersTableToSet;
        }

        public void SetExcelDeliverersTable(DataRow excelDelivererDataRow, DataColumnCollection columns)
        {
            //Create a DataTable
            DataTable excelDeliverer = new DataTable();

            //Copy column names into DataTable
            foreach (DataColumn column in columns)
            {
                excelDeliverer.Columns.Add(column.ColumnName);
            }
            
            //Import the Deliverer into the DataTable
            excelDeliverer.ImportRow(excelDelivererDataRow);

            //Get a handle to the DataTable
            mExcelDeliverersTable = excelDeliverer;
        }

        public void SetDatabaseSelectedYearDeliverersTable(DataTable databaseSelectedYearDeliverersTableToSet)
        {
            //Get a handle to the DataTable
            mDatabaseSelectedYearDeliverersTable = databaseSelectedYearDeliverersTableToSet;
        }

        public void SetMode(DelivererMode modeToSet)
        {
            //Create a new mode
            mMode = new DelivererMode();

            //Store mode
            mMode = modeToSet;
        }

        public void SetActionLabelText(string actionLabelTextToSet)
        {
            //Store the new action label text
            mActionLabelText = actionLabelTextToSet;

            //Display the new text.
            DelivererActionLabel.Content = mActionLabelText;
        }

        public void SetTopTableLabelText(string topTableLabelTextToSet)
        {
            //Store the new top table label text
            mTopTableLabelText = topTableLabelTextToSet;

            //Display the new text.
            DelivererTopTableLabel.Content = mTopTableLabelText;
        }

        public void SetBottomTableLabelText(string bottomTableLabelTextToSet)
        {
            //Store the new bottom table label text
            mBottomTableLabelText = bottomTableLabelTextToSet;

            //Display the new text.
            DelivererBottomTableLabel.Content = mBottomTableLabelText;
        }


        public void BindData()
        {
            ///////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////
            //////////////////      BASED ON MODE                  ////////////////
            ///////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////


            if (mMode == DelivererMode.DatabaseDeliverersTableDuplicates)
            {
                SetActionLabelText("Handle Duplicate Records found in the Deliverers Database Table Then Add To " + Main.mSelectedYear + " Table");
                SetTopTableLabelText("Excel Deliverer Being Imported");
                SetBottomTableLabelText("Database Deliverers Table Matches");

                //Bind the TopTable DataGrid to mExcelDeliverersTable
                DelivererTopTable.DataSource = mExcelDeliverersTable;

                //Bind the BottomTable DataGrid to mDatabaseDeliverersTable
                DelivererBottomTable.DataSource = mDatabaseDeliverersTable;

                DeleteSelectedDeliverersButton.Content = "Delete Selected Deliverer from Database Deliverers Table";
                AddSelectedDeliverersButton.Content = "Add Selected Deliverer to Database " + Main.mSelectedYear + " Table";
                UpdateDelivererInformationButton.Content = "Update Selected Deliverer in Database Deliverers Table";
            }
            else if (mMode == DelivererMode.DatabaseDeliverersTableUpdateDeliverer)
            {
                SetActionLabelText("Update Single Record found in the Deliverers Database Table Then Add To " + Main.mSelectedYear + " Table");
                SetTopTableLabelText("Excel Deliverer Being Imported");
                SetBottomTableLabelText("Database Deliverer Table Match");

                //Bind the TopTable DataGrid to mExcelDeliverersTable
                DelivererTopTable.DataSource = mExcelDeliverersTable;

                //Bind the BottomTable DataGrid to mDatabaseDeliverersTable
                DelivererBottomTable.DataSource = mDatabaseDeliverersTable;

                DeleteSelectedDeliverersButton.Content = "Delete Selected Deliverer from Database Deliverers Table";
                AddSelectedDeliverersButton.Content = "Add Selected Deliverer to Database " + Main.mSelectedYear + " Table";
                UpdateDelivererInformationButton.Content = "Update Selected Deliverer in Database Deliverers Table";
                
                //Hide this button
                AddSelectedDelivererToDeliverersAndSelectedTableButton.Visibility = Visibility.Hidden;
            }
            else if (mMode == DelivererMode.DatabaseDeliverersTableAddDeliverer)
            {
                SetActionLabelText("Add a Single Record to Deliverers Database Table Then Add To " + Main.mSelectedYear + " Table");
                SetTopTableLabelText("Excel Deliverer Being Imported");
                SetBottomTableLabelText("Database Deliverer Table Match");

                //Bind the TopTable DataGrid to mExcelDeliverersTable
                DelivererTopTable.DataSource = mExcelDeliverersTable;

                //Bind the BottomTable DataGrid to mDatabaseDeliverersTable
                DelivererBottomTable.DataSource = mDatabaseDeliverersTable;

                DeleteSelectedDeliverersButton.Content = "Delete Selected Deliverer from Database Deliverers Table";
                AddSelectedDeliverersButton.Content = "Add Selected Deliverer to Database " + Main.mSelectedYear + " Table"; ;
                UpdateDelivererInformationButton.Content = "Update Selected Deliverer in Database Deliverers Table";
                AddSelectedDelivererToDeliverersAndSelectedTableButton.Content = "Add Selected Deliverer to Database Deliverers and " + Main.mSelectedYear + " Tables";

                //Show this button
                AddSelectedDelivererToDeliverersAndSelectedTableButton.Visibility = Visibility.Visible;
            }
        }

        public void UnBindData()
        {
            //UnBind the BottomTable DataGrid
            DelivererBottomTable.DataSource = null;

            //UnBind the TopTable DataGrid
            DelivererTopTable.DataSource = null;
        }

        private void DeleteSelectedDelivererButton_Click(object sender, RoutedEventArgs e)
        {
            if (mMode == DelivererMode.DatabaseDeliverersTableDuplicates || mMode == DelivererMode.DatabaseDeliverersTableUpdateDeliverer || mMode == DelivererMode.DatabaseDeliverersTableAddDeliverer || mMode == DelivererMode.DatabaseDeliverersTableAddDeliverer)
            {
                int rowToDelete = -1;
                int selectedRowCount = 0;
                string DelivererID = "";
                string firstName = "";
                string lastName = "";

                //See see if only 1 row is selected and get that row number
                for (int i = 0; i < mDatabaseDeliverersTable.Rows.Count; i++)
                {
                    if (DelivererBottomTable.IsSelected(i))
                    {
                        rowToDelete = i;
                        selectedRowCount++;
                    }
                }

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 Deliverer");
                    return;
                }

                //See if we should delete the Deliverer into the Deliverers table
                if (rowToDelete != -1)
                {
                    //Get information to display to the user for confirmation
                    DelivererID = mDatabaseDeliverersTable.Rows[rowToDelete]["Deliverer_ID"].ToString();
                    lastName = mDatabaseDeliverersTable.Rows[rowToDelete]["Last_Name"].ToString();
                    firstName = mDatabaseDeliverersTable.Rows[rowToDelete]["First_Name"].ToString();

                    //Update Access Database
                    string deleteCommand = "DELETE FROM Deliverers WHERE Deliverer_ID = " + mDatabaseDeliverersTable.Rows[rowToDelete]["Deliverer_ID"];
                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(deleteCommand);

                    //Update data grid source
                    mDatabaseDeliverersTable.Rows.RemoveAt(rowToDelete);
                }
                
                //Update data grid display
                DelivererTopTable.Refresh();
                DelivererBottomTable.Refresh();

                if (rowToDelete != -1)
                {
                    //Show status message box
                    System.Windows.MessageBox.Show("Deliverer:  " + DelivererID + " - " + lastName + ", " + firstName + " DELETED from Deliverers Table");
                }
            }
        }

        private void AddDelivererToSelectedYearTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (mMode == DelivererMode.DatabaseDeliverersTableDuplicates || mMode == DelivererMode.DatabaseDeliverersTableUpdateDeliverer || mMode == DelivererMode.DatabaseDeliverersTableAddDeliverer)
            {
                int rowToInsert = -1;
                int selectedRowCount = 0;
                string DelivererID = "";
                string firstName = "";
                string lastName = "";

                //See see if only 1 row is selected and get that row number
                for (int i = 0; i < mDatabaseDeliverersTable.Rows.Count; i++)
                {
                    if (DelivererBottomTable.IsSelected(i))
                    {
                        rowToInsert = i;
                        selectedRowCount++;
                    }
                }

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 Deliverer");
                    return;
                }

                //See if we should insert the Deliverer into the mSelectedYear table
                if (rowToInsert != -1)
                {
                    //Get information to display to the user for confirmation
                    DelivererID = mDatabaseDeliverersTable.Rows[rowToInsert]["Deliverer_ID"].ToString();
                    lastName = mDatabaseDeliverersTable.Rows[rowToInsert]["Last_Name"].ToString();
                    firstName = mDatabaseDeliverersTable.Rows[rowToInsert]["First_Name"].ToString();

                    //Update Access Database
                    string insertCommand = "INSERT INTO " + Main.mSelectedYear + "_Deliverers" + " (Deliverer_ID) VALUES (" + mDatabaseDeliverersTable.Rows[rowToInsert]["Deliverer_ID"] + ")";
                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommand);
                }

                //Update data grid display
                DelivererTopTable.Refresh();
                DelivererBottomTable.Refresh();

                if (rowToInsert != -1)
                {
                    //Show status message box
                    System.Windows.MessageBox.Show("Deliverer:  " + DelivererID + " - " + lastName + ", " + firstName + " ADDED in " + Main.mSelectedYear +"_Deliverers" + " Table");

                    //Close the Import Helper
                    this.Close();
                }
            }
        }

        private void UpdateDelivererInformationButton_Click(object sender, RoutedEventArgs e)
        {
            if (mMode == DelivererMode.DatabaseDeliverersTableDuplicates || mMode == DelivererMode.DatabaseDeliverersTableUpdateDeliverer || mMode == DelivererMode.DatabaseDeliverersTableAddDeliverer)
            {
                int rowToUpdate = -1;
                int selectedRowCount = 0;
                string DelivererID = "";
                string firstName = "";
                string lastName = "";

                //See see if only 1 row is selected and get that row number
                for (int i = 0; i < mDatabaseDeliverersTable.Rows.Count; i++)
                {
                    if (DelivererBottomTable.IsSelected(i))
                    {
                        rowToUpdate = i;
                        selectedRowCount++;
                    }
                }

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 Deliverer");
                    return;
                }

                //See if we should update the Deliverer in the Deliverers table
                if (rowToUpdate != -1)
                {
                    //Get information to display to the user for confirmation
                    DelivererID = mDatabaseDeliverersTable.Rows[rowToUpdate]["Deliverer_ID"].ToString();
                    lastName = mDatabaseDeliverersTable.Rows[rowToUpdate]["Last_Name"].ToString();
                    firstName = mDatabaseDeliverersTable.Rows[rowToUpdate]["First_Name"].ToString();

                    //Check for null values
                    foreach (DataColumn column in mDatabaseDeliverersTable.Columns)
                    {
                        if (mDatabaseDeliverersTable.Rows[rowToUpdate][column.ColumnName].ToString() == "")
                        {
                            mDatabaseDeliverersTable.Rows[rowToUpdate][column.ColumnName] = "null";
                        }
                    }

                    string updateCommand = "UPDATE Deliverers SET Last_Name = '" + mDatabaseDeliverersTable.Rows[rowToUpdate]["Last_Name"].ToString() + "', " +
                                               "First_Name = '" + mDatabaseDeliverersTable.Rows[rowToUpdate]["First_Name"].ToString() + "', " +
                                               "Home_Phone = '" + mDatabaseDeliverersTable.Rows[rowToUpdate]["Home_Phone"].ToString() + "', " +
                                               "Work_Phone = '" + mDatabaseDeliverersTable.Rows[rowToUpdate]["Work_Phone"].ToString() + "', " +
                                               "Capacity = " + mDatabaseDeliverersTable.Rows[rowToUpdate]["Capacity"] + ", " +
                                               "Assigned = " + mDatabaseDeliverersTable.Rows[rowToUpdate]["Assigned_" + Main.mSelectedYear.ToString()].ToString() + ", " +
                                               "Occupation_Status = '" + mDatabaseDeliverersTable.Rows[rowToUpdate]["Occupation_Status"].ToString() + "', " + 
                                               "Comments = '" + mDatabaseDeliverersTable.Rows[rowToUpdate]["Comments"].ToString() + "'" +
                                               " WHERE Deliverer_ID = " + mDatabaseDeliverersTable.Rows[rowToUpdate]["Deliverer_ID"].ToString();

                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateCommand);
                }

                //Update data grid display
                DelivererTopTable.Refresh();
                DelivererBottomTable.Refresh();

                if (rowToUpdate != -1)
                {
                    //Show Status message box
                    System.Windows.MessageBox.Show("Deliverer:  " + DelivererID + " - " + lastName + ", " + firstName + " UPDATED in Deliverers Table");
                }
            }
        }

        private void AddSelectedDelivererToDeliverersAndSelectedYearTableButton_Click(object sender, RoutedEventArgs e)
        {
            int rowToAdd = -1;
            int selectedRowCount = 0;
            string DelivererID = "";
            string firstName = "";
            string lastName = "";

            if (mMode == DelivererMode.DatabaseDeliverersTableAddDeliverer)
            {

                //Check for null values
                for (int i = 0; i < mExcelDeliverersTable.Rows.Count; i++)
                {
                    if (DelivererTopTable.IsSelected(i))
                    {
                        rowToAdd = i;
                        selectedRowCount++;
                    }
                }

                //Add the record into the Deliverers Table

                //See see if only 1 row is selected
                if (selectedRowCount > 1)
                {
                    System.Windows.MessageBox.Show("Only Select 1 Deliverer");
                    return;
                }

                //See if we should add the Deliverer to the Deliverers table
                if (rowToAdd != -1)
                {
                    //Check for null values
                    foreach (DataColumn column in mExcelDeliverersTable.Columns)
                    {
                        if (mExcelDeliverersTable.Rows[rowToAdd][column.ColumnName].ToString() == "")
                        {
                            if(column.ColumnName == "Capacity_" + Main.mSelectedYear || column.ColumnName == "Assigned_" + Main.mSelectedYear)
                            {
                                mExcelDeliverersTable.Rows[rowToAdd][column.ColumnName] = "0";
                            }
                            else
                            {
                                mExcelDeliverersTable.Rows[rowToAdd][column.ColumnName] = "null";
                            }
                        }
                    }

                    //Update Access Database
                    string insertCommand = "INSERT INTO Deliverers (Last_Name,First_Name,Home_Phone,Work_Phone,Capacity,Assigned,Occupation_Status,Comments)" +
                                           " VALUES (" +
                                           "'" + mExcelDeliverersTable.Rows[rowToAdd]["Last_Name"].ToString() + "'," +
                                           "'" + mExcelDeliverersTable.Rows[rowToAdd]["First_Name"].ToString() + "'," +
                                           "'" + mExcelDeliverersTable.Rows[rowToAdd]["Home_Phone"].ToString() + "'," +
                                           "'" + mExcelDeliverersTable.Rows[rowToAdd]["Work_Phone"].ToString() + "'," +
                                           "" + mExcelDeliverersTable.Rows[rowToAdd]["Capacity"] + "," +
                                           "" + mExcelDeliverersTable.Rows[rowToAdd]["Assigned"] + "," +
                                           "'" + mExcelDeliverersTable.Rows[rowToAdd]["Occupation_Status"].ToString() + "'," + 
                                           "'" + mExcelDeliverersTable.Rows[rowToAdd]["Comments"].ToString() + "')";

                    Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommand);

                    //Figure out what the Deliverer_ID was assigned
                    string selectCommand = "SELECT * FROM Deliverers WHERE " +
                                            "Last_Name = '" + mExcelDeliverersTable.Rows[rowToAdd]["Last_Name"].ToString() + "' AND " +
                                            "First_Name = '" + mExcelDeliverersTable.Rows[rowToAdd]["First_Name"].ToString() + "'";

                    
                    DataSet Deliverer = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectCommand, "Deliverers");

                    if (Deliverer != null)
                    {
                        //Get the record information to display to the user
                        DelivererID = Deliverer.Tables[0].Rows[0]["Deliverer_ID"].ToString();
                        lastName = mExcelDeliverersTable.Rows[rowToAdd]["Last_Name"].ToString();
                        firstName = mExcelDeliverersTable.Rows[rowToAdd]["First_Name"].ToString();

                        System.Windows.MessageBox.Show("Deliverer:  " + DelivererID + " - " + lastName + ", " + firstName + " ADDED in Deliverers Table");

                        //Copy Record over into mDatabaseDeliverersTable
                        mDatabaseDeliverersTable.Rows.Add();
                        foreach (DataColumn column in mExcelDeliverersTable.Columns)
                        {
                            mDatabaseDeliverersTable.Rows[0][column.ColumnName] = mExcelDeliverersTable.Rows[0][column.ColumnName];
                        }

                        mDatabaseDeliverersTable.Rows[0]["Deliverer_ID"] = DelivererID;

                        //Insert the record to the Table
                        string insertCommandTwo = "INSERT INTO " + Main.mSelectedYear + "_Deliverers" + " (Deliverer_ID) VALUES (" + DelivererID + ")";
                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertCommandTwo);

                        //Show status message box
                        System.Windows.MessageBox.Show("Deliverer:  " + DelivererID + " - " + lastName + ", " + firstName + " ADDED in " + Main.mSelectedYear + "_Deliverers" + " Table");

                        //Close the Import Helper
                        this.Close();
                    }

                    DelivererTopTable.Refresh();
                    DelivererBottomTable.Refresh();
                }
            }
        }

        private void QuitDelivererImport_Click(object sender, RoutedEventArgs e)
        {
            //Stop Deliverer Import
            Main.mQuitDelivererImport = true;

            //Close this form
            this.Close();
        }

        private void Done_Click(object sender, RoutedEventArgs e)
        {
            //Continue Deliverer Import
            Main.mQuitDelivererImport = false;

            //Close this form
            this.Close();
        }
    }
}
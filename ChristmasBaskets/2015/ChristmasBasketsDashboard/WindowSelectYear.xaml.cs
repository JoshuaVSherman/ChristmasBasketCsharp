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
    /// <summary>
    /// Interaction logic for WindowSelectYear.xaml
    /// </summary>
    public partial class WindowSelectYear : Window
    {
        public WindowSelectYear()
        {
            InitializeComponent();

            //Define local variables
            DataTable userTables = null;

            // We only want user tables, not system tables
            string[] restrictions = new string[4];
            restrictions[3] = "Table";

            //See if we have an Access database
            if (Main.mChristmasBasketsAccessDatabase != null)
            {
                // Get list of user tables
                userTables = Main.mChristmasBasketsAccessDatabase.GetSchema("Tables", restrictions);
            }

            // Add list of table names to listBox
            for (int i = 0; i < userTables.Rows.Count; i++)
            {
                //Only add the Tables formatted as Year_
                if (userTables.Rows[i][2].ToString().Contains("Year_") && (!userTables.Rows[i][2].ToString().Contains("Deliverers") && !userTables.Rows[i][2].ToString().Contains("Food") ))
                {
                    YearsListBox.Items.Add(userTables.Rows[i][2].ToString());
                }
            }

            //Highlight the selected item if it exists already
            if (Main.mSelectedYear != "NONE")
            {
                // Find the string in ListBox2.
                int index = YearsListBox.Items.IndexOf(Main.mSelectedYear);

                // If the item was  found in ListBoxOfYears select it in ListBoxOfYears.
                if (index != -1)
                {
                    YearsListBox.SelectedIndex = index;
                    YearsListBox.ScrollIntoView(Main.mSelectedYear);
                }

            }
        }

        private void SelectYear_Click(object sender, RoutedEventArgs e)
        {
            //Store the Selected Year into Main.mSelectedYear
            Main.mSelectedYear = YearsListBox.SelectedValue.ToString();

            //Define getSelectedYearStatusQuery
            string getSelectedYearStatusQuery = "SELECT * FROM Status WHERE Year_ID = '" + Main.mSelectedYear +"'";

            //Get Selected Year's Status from the Status Table database record
            DataSet selectedYearStatus = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(getSelectedYearStatusQuery, Main.mSelectedYear + "_Status");

            if (selectedYearStatus != null)
            {
                //Make sure there was only 1 record returned from the selected year table
                if (selectedYearStatus.Tables[0].Rows.Count == 1)
                {
                    //Only 1 record was found in the Status year table
                    
                    //Parse data from table
                    DataRow dataRow = selectedYearStatus.Tables[0].Rows[0];

                    //Update all the Status Indicators
                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_2_Clients_Imported.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_Clients_Imported] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_Clients_Imported] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_2_a_Check_For_Client_Duplicates.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_a_Check_For_Client_Duplicates] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_a_Check_For_Client_Duplicates] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_3_Green_Cards_Generated.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_3_Green_Cards_Generated] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_3_Green_Cards_Generated] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_4_Deliverers_Imported.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_4_Deliverers_Imported] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_4_Deliverers_Imported] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_5_Clients_Assigned_To_Deliverers.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_5_Clients_Assigned_To_Deliverers] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_5_Clients_Assigned_To_Deliverers] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_6_Generated_Deliverer_Maps.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_6_Generated_Deliverer_Maps] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_6_Generated_Deliverer_Maps] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_7_Day_Of_Event.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_Day_Of_Event] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_Day_Of_Event] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs] = false;
                    }

                    if ((string)dataRow[Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels.ToString()] == "1")
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels] = true;
                    }
                    else
                    {
                        Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels] = false;
                    }
                }
            }

            //Close this window
            this.Close();
        }

        private void DeleteYear_Click(object sender, RoutedEventArgs e)
        {
            //Define local variables
            string yearToDelete = YearsListBox.SelectedItem.ToString();
            
            //Create the deleteYearCommandText - specifies what tables to delete from the database
            string deleteYearCommandText = "DROP TABLE " + yearToDelete + ", " +
                                                           yearToDelete + "_Deliverers," +
                                                           yearToDelete + "_Food";
            
            //Delete the yearToDelete, yearToDelete_Deliverers, yearToDelete_Food tables from the database
            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(deleteYearCommandText);

            //Remove the Year from the List
            YearsListBox.Items.Remove(yearToDelete);

            //Create command to remove Year_ID record from the Status Table
            string deleteYearIDRecordFromStatusTableCommand = "DELETE FROM Status WHERE Year_ID = '" + yearToDelete + "'";

            //Remove Year_ID record from the Status Table
            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(deleteYearIDRecordFromStatusTableCommand);

            //Update Main.mSelectedYearStatus
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_Clients_Imported] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_2_a_Check_For_Client_Duplicates] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_3_Green_Cards_Generated] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_4_Deliverers_Imported] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_5_Clients_Assigned_To_Deliverers] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_6_Generated_Deliverer_Maps] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_Day_Of_Event] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_a_Generate_Unassigned_Clients_Map] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_b_Generate_Client_Lists] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_c_Generate_Food_Signs] = false;
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_7_d_Generate_Box_Labels] = false;
        }

        private void YearsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (YearsListBox.SelectedItem == null)
            {
                //NO item in ListBoxOfYears is selected

                //Update SelectYear Button Text
                SelectYear.Content = "Select Year_****";

                //Update DeleteYear Button Text
                DeleteYear.Content = "Delete Year_****";
            }
            else
            {
                //Item in ListBoxOfYears is selected

                //Update SelectYear Button Text
                SelectYear.Content = "Select " + YearsListBox.SelectedItem.ToString();

                //Update DeleteYear Button Text
                DeleteYear.Content = "Delete " + YearsListBox.SelectedItem.ToString();
            }
        }

        private void YearToCreateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (YearsListBox.SelectedItem == null)
            {
                //NO item in ListBoxOfYears is selected
                if (YearToCreateTextBox.Text == "")
                {
                    //Nothing in the YearToCreateTextBox

                    //Update CreateYear Button Text
                    CreateYear.Content = "Create Year_****";
                }
                else
                {
                    //Something in the YearToCreateTextBox

                    //Update CreateYear Button Text
                    CreateYear.Content = "Create Year_" + YearToCreateTextBox.Text;
                }
            }
            else
            {
                //Item in ListBoxOfYears is selected
                if (YearToCreateTextBox.Text == "")
                {
                    //Nothing in the YearToCreateTextBox

                    //Update CreateYear Button Text
                    CreateYear.Content = "Create Year_****";
                }
                else
                {
                    //Something in the YearToCreateTextBox

                    //Update CreateYear Button Text
                    CreateYear.Content = "Create Year_" + YearToCreateTextBox.Text;
                }
            }
        }

        private void CreateYear_Click(object sender, RoutedEventArgs e)
        {
            //Define local variables
            string yearToCreate = "Year_" + YearToCreateTextBox.Text;

            //Create the createYearCommandText - Creates a Year_selectedYear table in the database
            string createClientYearCommandText = "CREATE TABLE " + yearToCreate + " (Box_Number INTEGER, Client_ID INTEGER, Assigned_Status TEXT)";

            //Create the Year_yearToCreate table in the database
            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(createClientYearCommandText);

            //Create the createDelivererYearCommandText - Creates a Year_yearToCreate_Deliverers table in the database
            string createDelivererYearCommandText = "CREATE TABLE " + yearToCreate + "_Deliverers" + " (Deliverer_ID INTEGER, Clients TEXT)";

            //Create the Year_yearToCreate_Deliverers table in the database
            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(createDelivererYearCommandText);

            //Create the createFoodYearCommandText - Creates a Year_selectedYear_Food table in the database
            string createFoodYearCommandText = "CREATE TABLE " + yearToCreate + "_Food" + " (Food_Name TEXT, Food_Number_To_Put_In_Box INTEGER)";

            //Create the Year_yearToCreate_Food table in the database
            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(createFoodYearCommandText);

            //Update the YearsListBox
            YearsListBox.Items.Add(yearToCreate);
            YearsListBox.SelectedItem = yearToCreate;
            YearsListBox.ScrollIntoView(yearToCreate);

            //Create insertStatusYearCommand to insert and initialize yearToCreate_Status's yearToCreate Record database item
            string insertStatusYearCommand = "INSERT INTO Status (Year_ID, " +
                                                                 "Step_1_Year_Created_In_Database, " +
                                                                 "Step_2_Clients_Imported, " +
                                                                 "Step_2_a_Check_For_Client_Duplicates, " +
                                                                 "Step_3_Green_Cards_Generated, " +
                                                                 "Step_4_Deliverers_Imported, " +
                                                                 "Step_5_Clients_Assigned_To_Deliverers, " +
                                                                 "Step_6_Generated_Deliverer_Maps, " +
                                                                 "Step_7_Day_Of_Event, " +
                                                                 "Step_7_a_Generate_Unassigned_Clients_Map, " +
                                                                 "Step_7_b_Generate_Client_Lists, " +
                                                                 "Step_7_c_Generate_Food_Signs, " +
                                                                 "Step_7_d_Generate_Box_Labels) " +
                                                                 "VALUES ('" + yearToCreate + "', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0')";

            //insert and initialize yearToCreate_Status's yearToCreate Record database item
            Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(insertStatusYearCommand);

            //Update Main.mSelectedYearStatus
            Main.mSelectedYearStatus[(int)Main.mSelectedYearStatusEnum.Step_1_Year_Created_In_Database] = true;
        }

    }
}

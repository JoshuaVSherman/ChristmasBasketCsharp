using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Shapes;
using System.ComponentModel;

namespace ChristmasBasketsDashboard
{
    /// <summary>
    /// Interaction logic for WindowAssignClientsToDeliverer.xaml
    /// </summary>
    public partial class WindowAssignClientsToDeliverer : Window
    {
        public ICollectionView currentDelivererCollectionView { get; private set; }
        public List<Deliverer> currentDeliverer;
        string googleMapControlHtmFile = @"/GoogleMapControl.htm";
        List<string> zipcodesList;

        public WindowAssignClientsToDeliverer(Deliverer selectedDeliverer)
        {
            //Initialize variables
            currentDeliverer = new List<Deliverer>();
            zipcodesList = new List<string>();

            InitializeComponent();

            //Update ZipcodeListBox
            UpdateZipcodeListBox();

            //Update Data Grid
            UpdateDelivererInfoDataGrid(selectedDeliverer);

            //Generate Deliverer Assignment Map
            GenerateDelivererAssignmentMap(googleMapControlHtmFile, selectedDeliverer, "None");

            //Initialize Map Browser - Jerk
            InitializeMapBrowser(googleMapControlHtmFile);
        }

        private void InitializeMapBrowser(string filename)
        {
            string currentDirectory = Directory.GetCurrentDirectory();
            string fullPathFileToLoad = "file://127.0.0.1/" + currentDirectory.Replace(":", "$") + filename;
            Uri fullPathFileToLoadUri = new Uri(fullPathFileToLoad);

            //Load file
            mapBrowser.Navigate(fullPathFileToLoadUri);
        }

        private void UpdateDelivererInfoDataGrid(Deliverer selectedDeliverer)
        {
            //Initialize variables
            string delivererIDQuery = "SELECT * FROM Deliverers WHERE Deliverer_ID = " + selectedDeliverer.DelivererID;

            //Clear currentDeliverer
            currentDeliverer.Clear();

            //Get all the selectedDelivererInfo
            DataSet selectedDelivererInfo = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(delivererIDQuery, Main.mSelectedYear + "_Deliverers");

            //See if we got a hit in the database
            if(selectedDelivererInfo.Tables[0].Rows.Count == 1)
            {
                //Single record found in the database
                DataRow delivererInfoFromDatabase = selectedDelivererInfo.Tables[0].Rows[0];

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

                if (delivererInfoFromDatabase["Assigned"].ToString() == "")
                {
                    delivererInfoFromDatabase["Assigned"] = 0;
                }
                else
                {
                    delivererToAdd.Assigned = Convert.ToInt32(delivererInfoFromDatabase["Assigned"]);
                }

                //Get Clients from Year_XXXX_Clients Table
                string getDelivererClientsQuery = "SELECT Clients FROM " + Main.mSelectedYear + "_Deliverers WHERE Deliverer_ID = " + delivererInfoFromDatabase["Deliverer_ID"];

                DataSet delivererClientsBeingProcessed = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(getDelivererClientsQuery, Main.mSelectedYear + "_Deliverers");

                DataRow delivererClientsFromDatabase = delivererClientsBeingProcessed.Tables[0].Rows[0];

                delivererToAdd.Clients = delivererClientsFromDatabase["Clients"].ToString();

                currentDeliverer.Add(delivererToAdd);
            }
            else
            {
                MessageBox.Show("Deliverer_ID " + selectedDeliverer.DelivererID + " not found in the database!");
            }

            DataContext = this;

            currentDelivererCollectionView = CollectionViewSource.GetDefaultView(currentDeliverer);
        }

        void UpdateZipcodeListBox()
        {
            //Initialize variables
            string clientsQuery = "SELECT * FROM "+ Main.mSelectedYear;

            //Clear currentDeliverer
            currentDeliverer.Clear();

            //Get all the clients from Main.mSelectedYear
            DataSet clients = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(clientsQuery, Main.mSelectedYear);

            //See if we got a hit in the database
            if (clients.Tables[0].Rows.Count > 0)
            {

                foreach (DataRow client in clients.Tables[0].Rows)
                {
                    string zipcodeQuery = "SELECT * FROM Clients WHERE Client_ID = " + client["Client_ID"].ToString();

                    DataSet specificClient = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(zipcodeQuery, "Clients");

                    if (specificClient.Tables[0].Rows.Count == 1)
                    {
                        DataRow specificClientRow = specificClient.Tables[0].Rows[0];

                        //See if the zipcode is unique
                        if (!zipcodesList.Contains(specificClientRow["Zipcode"]))
                        {
                            //Uniques so add to list
                            zipcodesList.Add(specificClientRow["Zipcode"].ToString());
                        }
                    }
                    else
                    {
                        //Client not found
                    }
                }
            }
            else
            {
            }

            //Update actual listbox
            ZipcodeListbox.Items.Clear();

            //Add the "All" selection to the zipcode list
            zipcodesList.Insert(0, "All");

            //Sort in numerical order
            zipcodesList.Sort();

            foreach(string zipcode in zipcodesList)
            {
                ZipcodeListbox.Items.Add(zipcode);
            }
        }

        private void AssignClientButton_Click(object sender, RoutedEventArgs e)
        {
            //Initialize variables
            bool updateClients = false;
            bool updateClientHistory = false;
            bool mustUpdateDelivererInfoDataGrid = false;
            bool updateGoogleMap = false;

            if (ClientIDTextBox.Text != "")
            {
                foreach (Deliverer deliverer in currentDeliverer)
                {
                    updateClients = false;
                    updateClientHistory = false;
                    mustUpdateDelivererInfoDataGrid = false;

                    //Check Clients in Year_XXXX_Clients Table
                    if (deliverer.Clients == null)
                    {
                        //ClientID not in Clients and first element
                        deliverer.Clients += ClientIDTextBox.Text + ",";

                        updateClients = true;
                    }
                    else
                    {
                        string[] clients = deliverer.Clients.Split(',');
                        bool clientAlreadyAssigned = false;

                        foreach(string client in clients)
                        {
                            if(client != "")
                            {
                                if (client == ClientIDTextBox.Text)
                                {
                                    clientAlreadyAssigned = true;
                                    break;
                                }
                            }
                        }

                        if (clientAlreadyAssigned)
                        {
                            MessageBox.Show(ClientIDTextBox.Text + " already assigned to Deliverer's Clients Field");
                        }
                        else
                        {
                            //ClientID not in Clients and not first element so add comma
                            deliverer.Clients += ClientIDTextBox.Text + ",";

                            updateClients = true;
                        }
                    }

                    //Check Client_History in Deliverer's Table
                    if (deliverer.ClientHistory == null)
                    {
                        //ClientID not in ClientHistory and first element
                        deliverer.ClientHistory += ClientIDTextBox.Text + ",";

                        updateClientHistory = true;
                    }
                    else
                    {

                        //Determine if client is already in ClientHistory
                        string[] clientList = deliverer.ClientHistory.Split(',');
                        bool clientAlreadyInClientHistory = false;

                        foreach (string currentClient in clientList)
                        {
                            if (currentClient != "")
                            {
                                if (currentClient == ClientIDTextBox.Text)
                                {
                                    clientAlreadyInClientHistory = true;
                                    break;
                                }
                            }
                        }

                        //Determine if client is assigned to the current deliverer which is the target of the .htm file being investigatedv
                        if (clientAlreadyInClientHistory)
                        {
                            MessageBox.Show(ClientIDTextBox.Text + " already assigned to Deliverer's Client_History Field");
                        }
                        else
                        {
                            //ClientID not in Client_History and not first element so add comma
                            deliverer.ClientHistory += ClientIDTextBox.Text + ",";

                            updateClientHistory = true;
                        }

                    }

                    if (updateClients)
                    {
                        deliverer.Assigned += 1;

                        //Push back to database and refresh data table - Year_xxxx_Deliverers database table
                        string updateClientsCommand = "UPDATE " + Main.mSelectedYear + "_Deliverers SET Clients = '" + deliverer.Clients + "' " +
                                               " WHERE Deliverer_ID = " + deliverer.DelivererID;

                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateClientsCommand);

                        //Update Client's "Assigned_Status" value in the Year_xxxx database table
                        string updateAssignedStatusCommand = "UPDATE " + Main.mSelectedYear + " SET Assigned_Status = 'true' WHERE Client_ID = " + ClientIDTextBox.Text;

                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateAssignedStatusCommand);

                        mustUpdateDelivererInfoDataGrid = true;
                        updateGoogleMap = true;
                    }

                    if (updateClientHistory)
                    {
                        //Push back to database and refresh data table - Deliverers database table
                        string updateClientHistoryCommand = "UPDATE Deliverers SET Client_History = '" + deliverer.ClientHistory + "', " +
                                                   "Assigned = '" + deliverer.Assigned + "'" +
                                                   " WHERE Deliverer_ID = " + deliverer.DelivererID;

                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateClientHistoryCommand);

                        mustUpdateDelivererInfoDataGrid = true;
                    }

                    if (mustUpdateDelivererInfoDataGrid)
                    {
                        currentDelivererCollectionView = CollectionViewSource.GetDefaultView(currentDeliverer);
                        DelivererInfoDataGrid.Items.Refresh();
                    }

                    if (updateGoogleMap)
                    {
                        this.mapBrowser.InvokeScript("assignClient", ClientIDTextBox.Text);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter a ClientID value to Assign to Deliverer");
            }
        }

        private void DelivererInfoDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            // Get the DataRow corresponding to the DataGridRow that is loading.
            Deliverer deliverer = e.Row.Item as Deliverer;

            //Clear Row background color
            e.Row.Background = new SolidColorBrush(Colors.White);

            if (deliverer != null)
            {
                if (deliverer.Capacity == deliverer.Assigned)
                {
                    // Set the background color of the DataGrid row as deliver has been assigned the full amount
                    //of clients
                    e.Row.Background = new SolidColorBrush(Colors.LightGreen);
                }
            }
        }

        private void RemoveClientButton_Click(object sender, RoutedEventArgs e)
        {
            //Initialize variables
            bool updateClients = false;
            bool updateClientHistory = false;
            bool mustUpdateDelivererInfoDataGrid = false;
            bool updateGoogleMap = false;

            if (ClientIDTextBox.Text != "")
            {

                foreach (Deliverer deliverer in currentDeliverer)
                {
                    updateClients = false;
                    updateClientHistory = false;
                    mustUpdateDelivererInfoDataGrid = false;

                    //Check Clients in Year_XXXX_Clients Table
                    if (deliverer.Clients != null)
                    {
                        string[] clients = deliverer.Clients.Split(',');
                        bool clientAlreadyAssigned = false;

                        foreach (string client in clients)
                        {
                            if (client != "")
                            {
                                if (client == ClientIDTextBox.Text)
                                {
                                    clientAlreadyAssigned = true;
                                    break;
                                }
                            }
                        }

                        if (clientAlreadyAssigned)
                        {
                            //ClientID is in Clients so remove it
                            deliverer.Clients = deliverer.Clients.Replace(ClientIDTextBox.Text + ",", "");

                            updateClients = true;
                        }
                        else
                        {
                            MessageBox.Show(ClientIDTextBox.Text + " NOT assigned to Deliverer's Clients Field");
                        }

                    }

                    //Check Client_History in Deliverer's Table
                    if (deliverer.ClientHistory != null)
                    {









                        //Determine if client is already in ClientHistory
                        string[] clientList = deliverer.ClientHistory.Split(',');
                        bool clientAlreadyInClientHistory = false;

                        foreach (string currentClient in clientList)
                        {
                            if (currentClient != "")
                            {
                                if (currentClient == ClientIDTextBox.Text)
                                {
                                    clientAlreadyInClientHistory = true;
                                    break;
                                }
                            }
                        }

                        //Determine if client is assigned to the current deliverer which is the target of the .htm file being investigatedv
                        if (clientAlreadyInClientHistory)
                        {
                            //ClientID is in Client_History so remove it
                            deliverer.ClientHistory = deliverer.ClientHistory.Replace(ClientIDTextBox.Text + ",", "");

                            updateClientHistory = true;
                        }
                        else
                        {
                            MessageBox.Show(ClientIDTextBox.Text + " NOT assigned to Deliverer's Client_History Field");
                        }
                    }

                    if (updateClients)
                    {
                        deliverer.Assigned -= 1;

                        //Push back to database and refresh data table - Year_xxxx_Deliverers database table
                        string updateClientsCommand = "UPDATE " + Main.mSelectedYear + "_Deliverers SET Clients = '" + deliverer.Clients + "' " +
                                               " WHERE Deliverer_ID = " + deliverer.DelivererID;

                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateClientsCommand);

                        //Update Client's "Assigned_Status" value in the Year_xxxx database table
                        string updateAssignedStatusCommand = "UPDATE " + Main.mSelectedYear + " SET Assigned_Status = 'false' WHERE Client_ID = " + ClientIDTextBox.Text;

                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateAssignedStatusCommand);

                        mustUpdateDelivererInfoDataGrid = true;
                        updateGoogleMap = true;

                    }

                    if (updateClientHistory)
                    {
                        //Push back to database and refresh data table - Deliverers database table
                        string updateClientHistoryCommand = "UPDATE Deliverers SET Client_History = '" + deliverer.ClientHistory + "', " +
                                                   "Assigned = '" + deliverer.Assigned + "'" +
                                                   " WHERE Deliverer_ID = " + deliverer.DelivererID;

                        Main.mChristmasBasketsAccessDatabase.ExecuteNonQuery(updateClientHistoryCommand);

                        mustUpdateDelivererInfoDataGrid = true;
                    }

                    if (mustUpdateDelivererInfoDataGrid)
                    {
                        DelivererInfoDataGrid.Items.Refresh();
                    }

                    if (updateGoogleMap)
                    {
                        this.mapBrowser.InvokeScript("unassignClient", ClientIDTextBox.Text);
                    }
                }
            }
            else
            {
                MessageBox.Show("Enter a ClientID value to Remove from Deliverer");
            }
        }

        private void RefreshMapButton_Click(object sender, RoutedEventArgs e)
        {
            if (RefreshMapListBox.SelectedIndex >= 0  && RefreshMapListBox.SelectedIndex <=4)
            {
                //Generate new googleMapControlHtmFile
                foreach(Deliverer deliverer in currentDeliverer)
                {
                    GenerateDelivererAssignmentMap(googleMapControlHtmFile, deliverer, ZipcodeListbox.SelectedItem.ToString());
                }

                //Refresh Map Browser - Jerk
                InitializeMapBrowser(googleMapControlHtmFile);
            }
            else
            {
                MessageBox.Show("Invalid option");
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            //Clear the ClientIDTextBox contents
            ClientIDTextBox.Text = "";
        }

        private void GenerateDelivererAssignmentMap(string filename, Deliverer currentDeliverer, string clientZipcode)
        {
            //Define Map Generation Arguments
            string[] clients = null;
            string[] addresses = null;
            string[] clientsAssigned = null;
            string[] clientIDs = null;
            string showMarkerMode = "0";

            string currentDirectory = Directory.GetCurrentDirectory();
            string fullFilePath = currentDirectory + filename;

            //PrepareMapGenerationArguments
            PrepareMapGenerationArguments(ref clients, ref addresses, ref clientsAssigned, ref clientIDs, ref showMarkerMode, ref currentDeliverer, ref clientZipcode);


            //Update TotalClientsToDisplayValueLabel
            if (clients != null)
            {
                TotalClientsToDisplayValueLabel.Content = clients.Length.ToString();
            }
            else 
            {
                TotalClientsToDisplayValueLabel.Content = "0";
            }

            //Generate the Deliverer Assignment Map
            //Define local variables
            TextWriter textWriter = new StreamWriter(fullFilePath);

            //Write the header part of the .htm page
            textWriter.WriteLine("<!DOCTYPE html>");
            textWriter.WriteLine("<html>");
            textWriter.WriteLine("   <head>");
            textWriter.WriteLine("      <meta name=\"viewport\" content=\"initial-scale=1.0, user-scalable=no\"/>");
            textWriter.WriteLine("      <style type=\"text/css\">");
            textWriter.WriteLine("         html { height: 100% }");
            textWriter.WriteLine("         body { height: 100%; margin: 0; padding: 0 }");
            textWriter.WriteLine("         #map_canvas { height: 100% }");
            textWriter.WriteLine("      </style>");

            //Write the java script part of the .htm page
            textWriter.WriteLine("      <!--[if IE]> <script type=\"text/javascript\" src=\"ie-set_timeout.js\"></script> <![endif]-->");
            textWriter.WriteLine("      <script type=\"text/javascript\" src=\"http://maps.google.com/maps/api/js?sensor=false\"></script>");
            textWriter.WriteLine("      <script type=\"text/javascript\">");
            textWriter.WriteLine("");
            textWriter.WriteLine("      var geocoder;");
            textWriter.WriteLine("      var map;");
            textWriter.WriteLine("      var markers = [];");
            textWriter.WriteLine("      var showMarkerMode = " + showMarkerMode + ";	//0 = All Clients, 1 = Only Assigned Clients, 2 = Only Unassigned Clients, 3 = Current Deliverer Assigned and Unassigned Clients, 4 = Only Current Deliverer Assigned Clients");
            textWriter.WriteLine("      var assignedClientToCurrentDelivererImage = \"AssignedClientToCurrentDeliverer.png\";");
            textWriter.WriteLine("      var assignedClientToOtherDelivererImage = \"AssignedClientToOtherDeliverer.png\";");
            textWriter.WriteLine("      var unassignedClientImage = \"UnassignedClient.png\";");
            textWriter.WriteLine("      var mapCenterAddress = \"2147 Dale Avenue Southeast Roanoke, VA 24013\";");

            //Generate addresses and clients arrays
            string javaClients =            "      var clients = [";
            string javaAddresses =          "      var addresses = [";
            string javaClientsAssigned =    "      var clientsAssigned = [";
            string javaClientIDs =          "      var clientIDs = [";

            if (clientZipcode == "None")
            {
                //No clients wanted to be shown
                javaClients += "];";
                javaAddresses += "];";
                javaClientsAssigned += "];";  //0 - Client unassigned, 1 - client assigned to current deliverer the htm file represents, 2 - client assigned to another deliverer other than current deliverer htm file represents";
                javaClientIDs += "];";
            }
            else
            {
                //See if we have 1 or more than 1 client to process
                if (clients.Count() == 1)
                {
                    //One client
                    javaClients += "\"" + clients[0] + "\"];";
                    javaAddresses += "\"" + addresses[0] + "\"];";
                    javaClientsAssigned += "\"" + clientsAssigned[0] + "\"];";
                    javaClientIDs += "\"" + clientIDs[0] + "\"];";
                }
                else
                {
                    //More than one client
                    for (int i = 0; i < clients.Count(); i++)
                    {
                        //Insert clients
                        if (i == (clients.Count() - 1))
                        {
                            //Last Record
                            javaClients += "\"" + clients[i] + "\"];";
                            javaAddresses += "\"" + addresses[i] + "\"];";
                            javaClientsAssigned += "\"" + clientsAssigned[i] + "\"];  //0 - Client unassigned, 1 - client assigned to current deliverer the htm file represents, 2 - client assigned to another deliverer other than current deliverer htm file represents";
                            javaClientIDs += "\"" + clientIDs[i] + "\"];";
                        }
                        else
                        {
                            //Not the Last Record
                            javaClients += "\"" + clients[i] + "\",";
                            javaAddresses += "\"" + addresses[i] + "\",";
                            javaClientsAssigned += "\"" + clientsAssigned[i] + "\",";
                            javaClientIDs += "\"" + clientIDs[i] + "\",";
                        }
                    }
                }
            }
            //Insert addresses array
            textWriter.WriteLine(javaAddresses);

            //Insert clients array
            textWriter.WriteLine(javaClients);

            //Insert clientsAssigned array
            textWriter.WriteLine(javaClientsAssigned);

            //Insert clientIDs array
            textWriter.WriteLine(javaClientIDs);
            textWriter.WriteLine("");

            //initialize function
            textWriter.WriteLine("      function initialize()");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("         //Create map options");
            textWriter.WriteLine("         var myOptions = {mapTypeId: google.maps.MapTypeId.ROADMAP};");
            textWriter.WriteLine("");
            textWriter.WriteLine("         //Create map and geocoder");
            textWriter.WriteLine("         map = new google.maps.Map(document.getElementById(\"map_canvas\"),myOptions);");
            textWriter.WriteLine("         geocoder = new google.maps.Geocoder();");
            textWriter.WriteLine("         //Show all addresses");
            textWriter.WriteLine("         for (i in clients)");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("            address = addresses[i];");
            textWriter.WriteLine("            client = clients[i];");
            textWriter.WriteLine("            clientAssigned = clientsAssigned[i];");
            textWriter.WriteLine("            delay = i * 1000;");
            textWriter.WriteLine("            setTimeout(showAddress, delay, address, client, clientAssigned);");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("");
            textWriter.WriteLine("         //Center the Map");
            textWriter.WriteLine("         centerMap(mapCenterAddress, 12);");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("");

            //showAddress function
            textWriter.WriteLine("      //Show a single address with client info");
            textWriter.WriteLine("      function showAddress(address, client, assigned)");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("          geocoder.geocode( { 'address': address},");
            textWriter.WriteLine("                              function(results, status)");
            textWriter.WriteLine("                              {");
            textWriter.WriteLine("                                 //Make sure we got a good result");
            textWriter.WriteLine("                                 if (status == google.maps.GeocoderStatus.OK)");
            textWriter.WriteLine("                                 {");
            textWriter.WriteLine("                                    //Temporary working marker");
            textWriter.WriteLine("                                    var marker;");
            textWriter.WriteLine("");
            textWriter.WriteLine("                                    //Create and display marker based on assigned status");
            textWriter.WriteLine("                                    if(assigned == \"0\")");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       //Client is not assigned to any deliverer");
            textWriter.WriteLine("                                       marker = new google.maps.Marker({map: map, position: results[0].geometry.location, title: client + \"\\n\" + address, icon: unassignedClientImage});");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("                                    else if(assigned == \"1\")");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       //Client is assigned to current deliverer");
            textWriter.WriteLine("                                       marker = new google.maps.Marker({map: map, position: results[0].geometry.location, title: client + \"\\n\" + address, icon: assignedClientToCurrentDelivererImage});");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("                                    else if(assigned == \"2\")");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       //Client is assigned to another deliverer other than current");
            textWriter.WriteLine("                                       marker = new google.maps.Marker({map: map, position: results[0].geometry.location, title: client + \"\\n\" + address, icon: assignedClientToOtherDelivererImage});");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("");
            textWriter.WriteLine("                                    //Determine if we should show the Marker based on showMarkerMode");
            textWriter.WriteLine("");
            textWriter.WriteLine("                                    //Show All Clients");
            textWriter.WriteLine("                                    if(showMarkerMode == 0)");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       marker.setMap(map);");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("                                    //Show Only Assigned Clients");
            textWriter.WriteLine("                                    else if(showMarkerMode == 1 || showMarkerMode == 4)");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       //If the client is not assigned - do not show on the map");
            textWriter.WriteLine("                                       if(assigned == \"0\")");
            textWriter.WriteLine("                                       {");
            textWriter.WriteLine("                                          marker.setMap(null);");
            textWriter.WriteLine("                                       }");
            textWriter.WriteLine("                                       else");
            textWriter.WriteLine("                                       {");
            textWriter.WriteLine("                                          marker.setMap(map);");
            textWriter.WriteLine("                                       }");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("                                    //Show Only Unassigned Clients");
            textWriter.WriteLine("                                    else if(showMarkerMode == 2)");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       //If the client is assigned - do not show on the map");
            textWriter.WriteLine("                                       if(assigned != \"0\")");
            textWriter.WriteLine("                                       {");
            textWriter.WriteLine("                                          marker.setMap(null);");
            textWriter.WriteLine("                                       }");
            textWriter.WriteLine("                                       else");
            textWriter.WriteLine("                                       {");
            textWriter.WriteLine("                                          marker.setMap(map);");
            textWriter.WriteLine("                                       }");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("                                    //Show Only Unassigned Clients and Clients that below to current Deliverer");
            textWriter.WriteLine("                                    else if(showMarkerMode == 3)");
            textWriter.WriteLine("                                    {");
            textWriter.WriteLine("                                       //If the client is assigned to the current deliverer or Unassigned - show on the map");
            textWriter.WriteLine("                                       if(assigned == \"0\"  || assigned == \"1\")");
            textWriter.WriteLine("                                       {");
            textWriter.WriteLine("                                          marker.setMap(map);");
            textWriter.WriteLine("                                       }");
            textWriter.WriteLine("                                       else");
            textWriter.WriteLine("                                       {");
            textWriter.WriteLine("                                          marker.setMap(null);");
            textWriter.WriteLine("                                       }");
            textWriter.WriteLine("                                    }");
            textWriter.WriteLine("");
            textWriter.WriteLine("                                    //Add marker to marker array");
            textWriter.WriteLine("                                    markers.push(marker);");
            textWriter.WriteLine("                                 }");
            textWriter.WriteLine("                                 else");
            textWriter.WriteLine("                                 {");
            textWriter.WriteLine("                                    //Alert - Geocode not successful");
            textWriter.WriteLine("                                    alert(\"showAddress - Geocode of (\" + address + \") was not successful for the following reason: (\" + status + \")\");");
            textWriter.WriteLine("                                 }");
            textWriter.WriteLine("                              });");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("");
              
            //centerMap function
            textWriter.WriteLine("      //Show a single address with client info");
            textWriter.WriteLine("      function centerMap(address, zoom)");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("         geocoder.geocode( { 'address': address},");
            textWriter.WriteLine("                           function(results, status)");
            textWriter.WriteLine("                           {");
            textWriter.WriteLine("                              //Make sure we got a good result");
            textWriter.WriteLine("                              if (status == google.maps.GeocoderStatus.OK)");
            textWriter.WriteLine("                              {");
            textWriter.WriteLine("                                 //Center Map");
            textWriter.WriteLine("                                 map.setCenter(results[0].geometry.location);");
            textWriter.WriteLine("");
            textWriter.WriteLine("                                 //Set Map Zoom");
            textWriter.WriteLine("                                 map.setZoom(zoom);");
            textWriter.WriteLine("                              }");
            textWriter.WriteLine("                              else");
            textWriter.WriteLine("                              {");
            textWriter.WriteLine("                                 //Alert - Geocode not successful");
            textWriter.WriteLine("                                 alert(\"centerMap - Geocode of (\" + address + \") was not successful for the following reason: (\" + status + \")\");");
            textWriter.WriteLine("                              }");
            textWriter.WriteLine("                           });");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("");

            //assignClient function
            textWriter.WriteLine("      //Assign Client");
            textWriter.WriteLine("      function assignClient(clientID)");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("         var clientIndex = -1;");
            textWriter.WriteLine("");
            textWriter.WriteLine("         //Find index for Client in the ClientIDs array");
            textWriter.WriteLine("         for(i in clientIDs)");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("            if(clientIDs[i] == clientID)");
            textWriter.WriteLine("            {");
            textWriter.WriteLine("               clientIndex = i;");
            textWriter.WriteLine("               break;");
            textWriter.WriteLine("            }");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("         //See if clientIndex was found");
            textWriter.WriteLine("         if(clientIndex > -1)");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("            if(markers)");
            textWriter.WriteLine("            {");
            textWriter.WriteLine("               //Alter Marker");
            textWriter.WriteLine("               markers[clientIndex].setIcon(assignedClientToCurrentDelivererImage);");
            textWriter.WriteLine("");
            textWriter.WriteLine("               //Determine if we should show the Marker based on showMarkerMode");
            textWriter.WriteLine("");
            textWriter.WriteLine("               //Show All Clients");
            textWriter.WriteLine("               if(showMarkerMode == 0)");
            textWriter.WriteLine("               {");
            textWriter.WriteLine("               //We are assigning a client and we should display All Clients - show clinet on the map");
            textWriter.WriteLine("               markers[clientIndex].setMap(map);");
            textWriter.WriteLine("               }");
            textWriter.WriteLine("               //Show Only Assigned Clients");
            textWriter.WriteLine("               else if(showMarkerMode == 1  || showMarkerMode == 4)");
            textWriter.WriteLine("               {");
            textWriter.WriteLine("                  //We are assigning a client and we should display only Assigned Clients - show client on the map");
            textWriter.WriteLine("                  markers[clientIndex].setMap(map);");
            textWriter.WriteLine("               }");
            textWriter.WriteLine("               //Show Only Unassigned Clients");
            textWriter.WriteLine("               else if(showMarkerMode == 2)");
            textWriter.WriteLine("               {");
            textWriter.WriteLine("                  //We are assigning a client and we should display only UnassignedClients - do not show client on the map");
            textWriter.WriteLine("                  markers[clientIndex].setMap(null);");
            textWriter.WriteLine("               }");
            textWriter.WriteLine("");
            textWriter.WriteLine("               //Update clients Assigned");
            textWriter.WriteLine("               clientsAssigned[clientIndex] = \"1\";");
            textWriter.WriteLine("            }");
            textWriter.WriteLine("            else");
            textWriter.WriteLine("            {");
            textWriter.WriteLine("               //Alert - markers[] is null");
            textWriter.WriteLine("               alert(\"assignClient - markers[] is null\");");
            textWriter.WriteLine("            }");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("         else");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("         //Alert - clientID does not exits");
            textWriter.WriteLine("         alert(\"assignClient - ClientID (\" + clientID + \") not found in ClientIDs array\");");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("");

            //unassignClient function            
            textWriter.WriteLine("      //Unassign Client");
            textWriter.WriteLine("      function unassignClient(clientID)");
            textWriter.WriteLine("      {");
            textWriter.WriteLine("         var clientIndex = -1;");
            textWriter.WriteLine("");
            textWriter.WriteLine("         //Find index for Client in the ClientIDs array");
            textWriter.WriteLine("         for(i in clientIDs)");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("            if(clientIDs[i] == clientID)");
            textWriter.WriteLine("            {");
            textWriter.WriteLine("               clientIndex = i;");
            textWriter.WriteLine("               break;");
            textWriter.WriteLine("            }");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("         //See if clientIndex was found");
            textWriter.WriteLine("         if(clientIndex > -1)");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("            if(markers)");
            textWriter.WriteLine("            {");
            textWriter.WriteLine("               //Alter Marker");
            textWriter.WriteLine("               markers[clientIndex].setIcon(unassignedClientImage);");
            textWriter.WriteLine("");
            textWriter.WriteLine("               //Determine if we should show the Marker based on showMarkerMode");
            textWriter.WriteLine("");
            textWriter.WriteLine("               //Show All Clients");
            textWriter.WriteLine("               if(showMarkerMode == 0)");
            textWriter.WriteLine("               {");
            textWriter.WriteLine("               //We are unassigning a client and we should display All Clients - show clinet on the map);");
            textWriter.WriteLine("               markers[clientIndex].setMap(map);");
            textWriter.WriteLine("               }");
            textWriter.WriteLine("               //Show Only Assigned Clients");
            textWriter.WriteLine("               else if(showMarkerMode == 1  || showMarkerMode == 4)");
            textWriter.WriteLine("               {");
            textWriter.WriteLine("                  //We are unassigning a client and we should display only Assigned Clients - do not show client on the map");
            textWriter.WriteLine("                  markers[clientIndex].setMap(null);");
            textWriter.WriteLine("               }");
            textWriter.WriteLine("               //Show Only Unassigned Clients");
            textWriter.WriteLine("               else if(showMarkerMode == 2)");
            textWriter.WriteLine("               {");
            textWriter.WriteLine("                  //We are unassigning a client and we should display only UnassignedClients - show client on the map");
            textWriter.WriteLine("                  markers[clientIndex].setMap(map);");
            textWriter.WriteLine("               }");
            textWriter.WriteLine("");
            textWriter.WriteLine("               //Update clients Assigned");
            textWriter.WriteLine("               clientsAssigned[clientIndex] = \"0\";");
            textWriter.WriteLine("            }");
            textWriter.WriteLine("            else");
            textWriter.WriteLine("            {");
            textWriter.WriteLine("               //Alert - markers[] is null");
            textWriter.WriteLine("               alert(\"unassignClient - markers[] is null\");");
            textWriter.WriteLine("            }");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("         else");
            textWriter.WriteLine("         {");
            textWriter.WriteLine("         //Alert - clientID does not exits");
            textWriter.WriteLine("         alert(\"unassignClient - ClientID (\" + clientID + \") not found in ClientIDs array\");");
            textWriter.WriteLine("         }");
            textWriter.WriteLine("      }");
            textWriter.WriteLine("      </script>");
            textWriter.WriteLine("   </head>");
            textWriter.WriteLine("");

            //Write the body part of the .htm page
            textWriter.WriteLine("   <body onload=\"initialize()\">");
            textWriter.WriteLine("      <div id=\"map_canvas\" style=\"height:90%;top:30px\"></div>");
            textWriter.WriteLine("   </body>");
            textWriter.WriteLine("</html>");

            //Close the file stream
            textWriter.Close();
        }

        private void PrepareMapGenerationArguments(ref string[] clients, ref string[] addresses, ref string[] clientsAssigned, ref string[] clientIDs, ref string showMarkerMode, ref Deliverer currentDeliverer, ref string clientZipcode)
        {
            if (Main.mSelectedYear == "NONE")
            {
                System.Windows.MessageBox.Show("Open a database and select a year.");
                return;
            }

            //Get showMarkerMode
            showMarkerMode = RefreshMapListBox.SelectedIndex.ToString();

            //Define local variables
            int clientCountByShowMarkerModeAndZipcode = GetClientCountByShowMarkerModeAndZipcode(showMarkerMode, clientZipcode, currentDeliverer);

            //Create the select selectByClient_IDQueryText
            string selectByClient_IDQueryText = "SELECT * FROM " + Main.mSelectedYear;

            //Get the list of clients

            //Create the clientIDList
            DataSet clientIDList = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClient_IDQueryText, Main.mSelectedYear);

            int index = 0;

            if (clientZipcode == "None")
            {
                //Return as we do not want any clients as the zipcode is None
                //Create new arrays
                return;
            }

            //Create new arrays
            clients = new string[clientCountByShowMarkerModeAndZipcode];
            addresses = new string[clientCountByShowMarkerModeAndZipcode];
            clientsAssigned = new string[clientCountByShowMarkerModeAndZipcode];
            clientIDs = new string[clientCountByShowMarkerModeAndZipcode];

            //Process each record in the clientIDList table
            foreach (DataRow dataRow in clientIDList.Tables[0].Rows)
            {
                //Create the selectByClientIDQueryText
                string selectByClientIDQueryText = "";

                if (clientZipcode == "All")
                {
                    //All zipcodes
                    selectByClientIDQueryText = "SELECT Client_ID, Last_Name, First_Name, Address_Number, Street_Address, City, Zipcode FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString();
                }
                else
                {
                    //Specific zipcode
                    selectByClientIDQueryText = "SELECT Client_ID, Last_Name, First_Name, Address_Number, Street_Address, City, Zipcode FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString() + " AND Zipcode = \"" + clientZipcode + "\"";
                }

                //Perform the selectByClientIDAndOrganizationQuery and store the results in a Data Table
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
                    string zipcode = selectByClientIDDataSet.Tables[0].Rows[0]["Zipcode"].ToString();
                    string assignedStatus = "";

                    string[] clientList = currentDeliverer.Clients.Split(',');
                    bool clientAlreadyAssigned = false;

                    foreach (string client in clientList)
                    {
                        if (client != "")
                        {
                            if (client == id)
                            {
                                clientAlreadyAssigned = true;
                                break;
                            }
                        }
                    }

                    //Determine if client is assigned to the current deliverer which is the target of the .htm file being investigatedv
                    if (clientAlreadyAssigned)
                    {
                        //Deliverer is not assigned to this deliverer
                        assignedStatus = "1";
                    }
                    else
                    {
                        if (dataRow["Assigned_Status"].ToString() == "true")
                        {
                            //Client is assigned to another deliverer other than the current deliverer
                            assignedStatus = "2";
                        }
                        else
                        {
                            //Client is not assigned to any deliverer
                            assignedStatus = "0";
                        }
                    }

                    bool addClient = false;

                    //Determine if we add the client

                    //Show Clients Assigned to Current Deliverer, Unassigned, and Assigned to Other Deliverers
                    if (showMarkerMode == "0")
                    {
                        addClient = true;
                    }
                    //Show Clients Assigned to Current Deliverer and Other Deliverers
                    else if (showMarkerMode == "1" && (assignedStatus == "1" || assignedStatus == "2"))
                    {
                        addClient = true;
                    }
                    //Show Unassigned Clients
                    else if (showMarkerMode == "2" && (assignedStatus == "0"))
                    {
                        addClient = true;
                    }
                    //Show Clients Assigned to Current Deliverer and Unassigned Clients
                    else if (showMarkerMode == "3" && (assignedStatus == "1" || assignedStatus == "0"))
                    {
                        addClient = true;
                    }
                    //Show Clients only assigned to Current Deliverer
                    else if (showMarkerMode == "4" && assignedStatus == "1")
                    {
                        addClient = true;
                    }

                    //Determine if we add the client
                    if (addClient)
                    {
                        clients[index] = id + " - " + lastName + ", " + firstName;
                        addresses[index] = addressNumber + " " + streetAddress + " " + city + ", VA  " + zipcode;
                        clientsAssigned[index] = assignedStatus;
                        clientIDs[index] = id;

                        index++;
                    }
                }

                if (index == clientCountByShowMarkerModeAndZipcode)
                {
                    break;
                }
            }
        }

        int GetClientCountByShowMarkerModeAndZipcode(string showMarkerMode, string zipcode, Deliverer deliverer)
        {
            //Define local variables
            int numberOfClientsWithModeAndZipcode = 0;

            if (zipcode == "None")
            {
                numberOfClientsWithModeAndZipcode = 0;
            }
            else
            {
                //Create the select selectByClient_IDQueryText
                string selectByClient_IDQueryText = "SELECT * FROM " + Main.mSelectedYear;

                //Get the list of clients

                //Create the clientIDList
                DataSet clientIDList = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(selectByClient_IDQueryText, Main.mSelectedYear);

                //Process each record in the clientIDList table and count clients with specific zipcode
                foreach (DataRow dataRow in clientIDList.Tables[0].Rows)
                {
                    string query = "";

                    if (zipcode == "All")
                    {
                        //All zipcodes
                        query = "SELECT * FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString();
                    }
                    else
                    {
                        //Single zipcode
                        query = "SELECT * FROM Clients WHERE Client_ID = " + dataRow["Client_ID"].ToString() + " AND Zipcode = \"" + zipcode + "\"";
                    }
                    
                    
                    DataSet client = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(query, "Clients");

                    if(client != null)
                    {
                        string id = client.Tables[0].Rows[0]["Client_ID"].ToString();
                        string assignedStatus = "";

                        string[] delivererClientList = deliverer.Clients.Split(',');
                        bool clientAssignedToCurrentDeliverer = false;

                        foreach (string currentClient in delivererClientList)
                        {
                            if (currentClient != "")
                            {
                                if (id == currentClient)
                                {
                                    clientAssignedToCurrentDeliverer = true;
                                    break;
                                }
                            }
                        }

                        //Determine if client is assigned to the current deliverer which is the target of the .htm file being investigatedv
                        if (clientAssignedToCurrentDeliverer)
                        {
                            //Deliverer is assigned to this deliverer
                            assignedStatus = "1";
                        }
                        else
                        {
                            if (dataRow["Assigned_Status"].ToString() == "true")
                            {
                                //Client is assigned to another deliverer other than the current deliverer
                                assignedStatus = "2";
                            }
                            else
                            {
                                //Client is not assigned to any deliverer
                                assignedStatus = "0";
                            }
                        }

                        //Determine if we add the client

                        //Show Clients Assigned to Current Deliverer, Unassigned, and Assigned to Other Deliverers
                        if (showMarkerMode == "0")
                        {
                            numberOfClientsWithModeAndZipcode++;
                        }
                        //Show Clients Assigned to Current Deliverer and Other Deliverers
                        else if (showMarkerMode == "1" && (assignedStatus == "1" || assignedStatus == "2"))
                        {
                            numberOfClientsWithModeAndZipcode++;
                        }
                        //Show Unassigned Clients
                        else if (showMarkerMode == "2" && (assignedStatus == "0"))
                        {
                            numberOfClientsWithModeAndZipcode++;
                        }
                        //Show Clients Assigned to Current Deliverer and Unassigned Clients
                        else if (showMarkerMode == "3" && (assignedStatus == "1" || assignedStatus == "0"))
                        {
                            numberOfClientsWithModeAndZipcode++;
                        }
                        //Show Only Clients assigned to the Current Deliverer
                        else if (showMarkerMode == "4" && assignedStatus == "1")
                        {
                            numberOfClientsWithModeAndZipcode++;
                        }
                    }
                    
                }
            }

            return numberOfClientsWithModeAndZipcode;
        }

    }
}

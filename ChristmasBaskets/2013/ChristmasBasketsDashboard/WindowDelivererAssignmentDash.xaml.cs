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
    /// Interaction logic for WindowDelivererAssignmentDash.xaml
    /// </summary>
    public partial class WindowDelivererAssignmentDash : Window
    {
        public ICollectionView deliverersCollectionView { get; private set; }
        public List<Deliverer> deliverers;

        public double TotalBoxes;
        public double TotalDelivererAssignedBoxes;
        public double TotalDelivererCapacityForBoxes;
        public double TotalUnassignedBoxes;
        public double PercentageAssignedBoxes;
        public double PercentageCapacityBoxes;

        public WindowDelivererAssignmentDash()
        {
            //Initialize Variables

            ////////////FIX THIS JERK!!!!!////// - Get the total boxes for the year from the database
            TotalBoxes = 396;
            TotalDelivererAssignedBoxes = 0;
            TotalDelivererCapacityForBoxes = 0;
            TotalUnassignedBoxes = 0;
            PercentageAssignedBoxes = 0;
            PercentageCapacityBoxes = 0;
            deliverers = new List<Deliverer>();
            
            InitializeComponent();

            UpdateDeliverersDataGrid();

            UpdateValueLabelsAndProgressBars();
        }

        private void UpdateDeliverersDataGrid()
        {
            //Initialize variables
            string delivererIDQuery = "SELECT * FROM " + Main.mSelectedYear + "_Deliverers";

            //Clear current data in deliverers
            deliverers.Clear();

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

                if (delivererInfoFromDatabase["Assigned"].ToString() == "")
                {
                    delivererToAdd.Assigned = 0;
                }
                else
                {
                    delivererToAdd.Assigned = Convert.ToInt32(delivererInfoFromDatabase["Assigned"]);
                }

                //Get Clients from Year_XXXX_Clients Table
                string getDelivererClientsQuery = "SELECT Clients FROM " + Main.mSelectedYear + "_Deliverers WHERE Deliverer_ID = " + dataRow["Deliverer_ID"];

                DataSet delivererClientsBeingProcessed = Main.mChristmasBasketsAccessDatabase.PerformSelectQuery(getDelivererClientsQuery, Main.mSelectedYear + "_Deliverers");

                DataRow delivererClientsFromDatabase = delivererClientsBeingProcessed.Tables[0].Rows[0];

                delivererToAdd.Clients = delivererClientsFromDatabase["Clients"].ToString();

                deliverers.Add(delivererToAdd);
            }

            DataContext = this;
            
            deliverersCollectionView = CollectionViewSource.GetDefaultView(deliverers);
            
            DeliverersDataGrid.Items.Refresh();
            UpdateValueLabelsAndProgressBars();
        }

        private void AssignClientsToDeliverers_Click(object sender, RoutedEventArgs e)
        {
            //Get selected Deliverer from DataGrid
            Deliverer selectedDeliverer = DeliverersDataGrid.SelectedItem as Deliverer;

           if (selectedDeliverer != null)
           {
               //MessageBox.Show("Assigning Clients to " + selectedDeliverer.LastName + ", " + selectedDeliverer.FirstName + "!");

               WindowAssignClientsToDeliverer delivererAssignment = new WindowAssignClientsToDeliverer(selectedDeliverer);
               delivererAssignment.ShowDialog();
           }
           else
           {
               MessageBox.Show("Select a deliverer row in the data grid above fool!");
           }
        }

        private void UpdateValueLabelsAndProgressBars()
        {
            //Initialize values
            TotalDelivererAssignedBoxes = 0;
            TotalDelivererCapacityForBoxes = 0;

            foreach (Deliverer deliverer in deliverers)
            {
                TotalDelivererAssignedBoxes += deliverer.Assigned;
                TotalDelivererCapacityForBoxes += deliverer.Capacity;
            }

            //Update Totals
            TotalUnassignedBoxes = TotalBoxes - TotalDelivererAssignedBoxes;
            PercentageAssignedBoxes = (TotalDelivererAssignedBoxes / TotalBoxes) * 100;
            PercentageCapacityBoxes = (TotalDelivererCapacityForBoxes / TotalBoxes) * 100;

            //Update value labels and progress bars on the main form
            TotalBoxesValueLabel.Content = TotalBoxes;
            TotalAssignedBoxesValueLabel.Content = TotalDelivererAssignedBoxes;
            TotalDelivererBoxCapacityValueLabel.Content = TotalDelivererCapacityForBoxes;
            TotalUnassignedBoxesValueLabel.Content = TotalUnassignedBoxes;
            TotalPercentageBoxesAssignedValueLabel.Content = PercentageAssignedBoxes;
            TotalPercentageDelivererBoxCapacityValueLabel.Content = PercentageCapacityBoxes;
            TotalPercentageBoxesAssignedProgressBar.Value = PercentageAssignedBoxes;
            TotalPercentageDelivererBoxCapacityProgressBar.Value = PercentageCapacityBoxes;
        }

        private void DeliverersDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            // Get the DataRow corresponding to the DataGridRow that is loading.
            Deliverer deliverer = e.Row.Item as Deliverer;

            //Clear Row background color
            e.Row.Background = new SolidColorBrush(Colors.White);

            if (deliverer != null)
            {
                if (deliverer.Capacity == deliverer.Assigned)
                {
                    //Set the background color of the DataGrid row as deliver has been assigned the full amount
                    //of clients
                    e.Row.Background = new SolidColorBrush(Colors.LightGreen);
                }
            }
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            UpdateDeliverersDataGrid();
            UpdateValueLabelsAndProgressBars();
        }

        private void RefreshDataGridButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateDeliverersDataGrid();
            UpdateValueLabelsAndProgressBars();
        }
    }

    public class Deliverer : INotifyPropertyChanged
    {
        private int delivererID;
        private string firstName;
        private string lastName;
        private int capacity;
        private string helpStatus;
        private string room;
        private string workPhone;
        private string homePhone;
        private string yearLastDelivererd;
        private string occupationStatus;
        private string comments;
        private string clientHistory;
        private string clients;
        private int assigned;

        public int DelivererID
        {
            get { return delivererID; }
            set
            {
                delivererID = value;
                NotifyPropertyChanged("DelivererID");
            }
        }


        public string FirstName
        {
            get { return firstName; }
            set
            {
                firstName = value;
                NotifyPropertyChanged("FirstName");
            }
        }

        public string LastName
        {
            get { return lastName; }
            set
            {
                lastName = value;
                NotifyPropertyChanged("LastName");
            }
        }

        public int Capacity
        {
            get { return capacity; }
            set
            {
                capacity = value;
                NotifyPropertyChanged("Capacity");
            }
        }

        public string HelpStatus
        {
            get { return helpStatus; }
            set
            {
                helpStatus = value;
                NotifyPropertyChanged("HelpStatus");
            }
        }

        public string Room
        {
            get { return room; }
            set
            {
                room = value;
                NotifyPropertyChanged("Room");
            }
        }

        public string WorkPhone
        {
            get { return workPhone; }
            set
            {
                workPhone = value;
                NotifyPropertyChanged("WorkPhone");
            }
        }

        public string HomePhone
        {
            get { return homePhone; }
            set
            {
                homePhone = value;
                NotifyPropertyChanged("HomePhone");
            }
        }

        public string YearLastDelivered
        {
            get { return yearLastDelivererd; }
            set
            {
                yearLastDelivererd = value;
                NotifyPropertyChanged("YearLastDelivered");
            }
        }

        public string Comments
        {
            get { return comments; }
            set
            {
                comments = value;
                NotifyPropertyChanged("Comments");
            }
        }

        public string OccupationStatus
        {
            get { return occupationStatus; }
            set
            {
                occupationStatus = value;
                NotifyPropertyChanged("OccupationStatus");
            }
        }

        public string ClientHistory
        {
            get { return clientHistory; }
            set
            {
                clientHistory = value;
                NotifyPropertyChanged("ClientHistory");
            }
        }

        public string Clients
        {
            get { return clients; }
            set
            {
                clients = value;
                NotifyPropertyChanged("Clients");
            }
        }

        public int Assigned
        {
            get { return assigned; }
            set
            {
                assigned = value;
                NotifyPropertyChanged("Assigned");
            }
        }

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Private Helpers

        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion
    }
}


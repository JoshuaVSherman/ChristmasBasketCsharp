   M S B A M L   `   `   `  ����  Christmas Basket Dashboard>    8ChristmasBasketsDashboard.WindowAssignClientsToDeliverer   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$$  WindowAssignClientsToDeliverer��5   	   $	��753��60   $
��1302��6=    � Icon$; 5/Christmas%20Basket%20Dashboard;component/present.ico��6J   .��5      � $
��1287��6   .+�5   
   �� -   5        $��
rectangle3��5   N   $	��605��6   $
��Left=�6!   $��9,89,0,0q�6<   
���   �$	��Top=�6o   $
��1024��6�   �� 5   
   -   5        $��
rectangle2��5   Q   $	��470��6   $
��Left=�6!   $��1044,12,0,0q�6<   
���   �$	��Top=�6r   $	��239��6�   �� 5   
   -   5        $��
rectangle1��5   R   $	��168��6   $
��Left=�6!   $��1058,494,0,0q�6<   
���   �$	��Top=�6s   $	��207��6�   �� 5	   
   -   5        $��ClientIDTextBox��5	   C   $
��Left=�6   $��1088,538,0,0q�6-   $��92��6Z   $��23��6e   $	��Top=�6q   �� 5
   
   -   5        $��AssignClientButton��5
   f   $��Assign Client��6   $��23��6)   $
��Left=�65   $��1088,575,0,0q�6P   $	��Top=�6�   $��92��6�   6�   �� 5   
   -   5        $��RemoveClientButton��5   f   $��Remove Client��6   $��23��6)   $
��Left=�65   $��1088,616,0,0q�6P   $	��Top=�6�   $��92��6�   6�   �� 5   
   -   5        $��ClientIDLabel��5   u   $#��Client ID to Assign or Remove��6   $��34��68   $
��Left=�6D   $��1069,500,0,0q�6_   $	��Top=�6�   $	��170��6�   &   System.Windows.Controls.DataGrid  5   
   -   5        $��DelivererInfoDataGrid��5   i      AutoGenerateColumns . $��71��6.   $
��Left=�6:   $��
128,12,0,0q�6U   $	��Top=�6�   $	��905��6�   ��6�   �� * currentDelivererCollectionView+6�   �� 5   
   -	   5        $��
RefreshMap��5   X   $��Refresh Map��6   $
��Left=�6'   $��1049,368,0,0q�6B   $	��226��6j   $��23��6�   $	��Top=�6�   �� 5   
   -
   5        $��RefreshMapListBox��5   N   $��90��6   $
��Left=�6   $��1049,40,0,0q�69   $	��Top=�6g   $	��226��6   ��6�   ��  � DataContext 6�   �� .@�5      �� $��All Clients��6   �� 5      $��Assigned Clients��6   �� 5      $��Unassigned Clients��6   �� 5      $.��(Current Deliverer and Unassigned Clients��6   �� 5      $��Current Deliverer��6   5      �� 5   
   -   5        $��RefreshMapModeLabel��5   g   $��Refresh Map Mode��6   $��28��6+   $
��Left=�67   $��1111,16,0,0q�6R   $	��Top=�6�   $	��110��6�   �� 5   
   -   5        $��ClientInfoLabel��5   b   $��Deliverer Info��6   $��34��6)   $
��Left=�65   $��9,32,0,0q�6P   $	��Top=�6y   $	��113��6�   {�.  $
��Bold"�6�   $��16(�6�   $��])�6�   �� 5   
   -   5        $��ClearButton��5   ]   $��Clear��6   $��23��6!   $��Right=�6-   $��
0,538,40,0q�6I   $	��Top=�6p   $��43��6�   6�   (  "System.Windows.Controls.WebBrowser  5   
   -   5        $��
mapBrowser��5   Q   $	��571��6   $
��Left=�6"   $��
20,108,0,0q�6=   $	��Top=�6c   $
��1000��6{   �� 5   
   -   5        $��AvailableZipcodesLabel��5   j   $��Available Zipcodes��6   $��28��6-   $
��Left=�69   $��1111,136,0,0q�6T   $	��Top=�6�   $	��110��6�   �� 5   
   -   5        $��ZipcodeListbox��5   P   $	��175��6   $
��Left=�6   $��1049,170,0,0q�6:   $	��Top=�6f   $	��226��6~   �� 5   
   -   5        $��ClientsToDisplayLabel��5   i   $��Clients Displayed��6   $��28��6,   $
��Left=�68   $��1050,411,0,0q�6S   $	��Top=�6�   �� 5   
   -   5        $ ��TotalClientsToDisplayLabel��5   p   $��Total Clients to Display��6   $��28��63   $
��Left=�6?   $��1050,438,0,0q�6Z   $	��Top=�6�   �� 5   
   -   5        $ ��ClientsToDisplayValueLabel��5   Y   $��0��6   $��28��6   $
��Left=�6(   $��1204,411,0,0q�6C   $	��Top=�6{   �� 5   
   -   5        $%��TotalClientsToDisplayValueLabel��5   Y   $��0��6   $��28��6   $
��Left=�6(   $��1204,438,0,0q�6C   $	��Top=�6�   5       5!      
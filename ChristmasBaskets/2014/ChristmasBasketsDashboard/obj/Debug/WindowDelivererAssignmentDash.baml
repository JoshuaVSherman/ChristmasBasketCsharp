   M S B A M L   `   `   `  ����  Christmas Basket Dashboard=    7ChristmasBasketsDashboard.WindowDelivererAssignmentDash   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   -   5          � Title$  Deliverer Assignment Dash��5   	   $	��668��6+   $
��1192��68   .��5      � .+�5   
   �� -   5        $��Delivers_For_Year_Label��5   m   $��Deliverers for Year_XXXX��6   $��31��63   $
��Left=�6?   $��	10,10,0,0q�6Z   $	��Top=�6�   $	��345��6�   &   System.Windows.Controls.DataGrid  5   
   -   5        $��DeliverersDataGrid��5   �      AutoGenerateColumns . $	��241��6.   $
��Left=�6;   $��10,46,0,342q�6�   $
��1033��6�   ��6V   �� *deliverersCollectionView+6�   �� 5   
   -   5        $��TotalBoxesLabel��5   a   $��Total Boxes��6   $��28��6&   $
��Left=�62   $��
10,326,0,0q�6M   $	��Top=�6x   $	��175��6�   �� 5	   
   -   5        $��TotalBoxesValueLabel��5	   j   $��0��6   $
��Bold"�6   $��28��6.   $
��Left=�6:   $��206,326,0,0q�6U   $	��Top=�6�   $��59��6�   �� 5
   
   -   5        $$��TotalDelivererBoxCapacityLabel��5
   r   $"��Total Deliverer Box Capacity��6   $��28��67   $
��Left=�6C   $��
10,454,0,0q�6^   $	��Top=�6�   $	��175��6�   �� 5   
   -   5        $)��#TotalDelivererBoxCapacityValueLabel��5   X   $��0��6   $��28��6   $
��Left=�6(   $��206,454,0,0q�6C   $	��Top=�6�   $��59��6�   �� 5   
   -   5        $��TotalAssignedBoxesLabel��5   j   $��Total Assigned Boxes��6   $��28��6/   $
��Left=�6;   $��
10,360,0,0q�6V   $	��Top=�6�   $	��175��6�   �� 5   
   -	   5        $"��TotalAssignedBoxesValueLabel��5   �   $��0��6   $
��Bold"�6   ���	#FF359C35$��28��6E   $
��Left=�6Q   $��206,360,0,0q�6l   $	��Top=�6�   $��59��6�   �� 5   
   -
   5        $��TotalUnassignedBoxesLabel��5   l   $��Total Unassigned Boxes��6   $��34��61   $
��Left=�6=   $��
10,0,0,199q�6X   $��Bottom=�6�   $	��175��6�   �� 5   
   -   5        $$��TotalUnassignedBoxesValueLabel��5   �   $��0��6   $
��Bold"�6   ���	#FFFD0606$��28��6E   $
��Left=�6Q   $��206,394,0,0q�6l   $	��Top=�6�   $��59��6�   �� 5   
   -   5        $'��!TotalPercentageBoxesAssignedLabel��5   v   $%��Total Percentage Boxes Assigned��6   $��28��6:   $
��Left=�6F   $��308,326,0,0q�6a   $	��Top=�6�   $	��195��6�   �� 5   
   -   5        $,��&TotalPercentageBoxesAssignedValueLabel��5   {   $��0��6   $
��Bold"�6   
���  ��$��28��6?   $
��Left=�6K   $��561,326,0,0q�6f   $	��Top=�6�   $��51��6�   �� 5   
   -   5        $.��(TotalPercentageDelivererBoxCapacityLabel��5   ~   $-��'Total Percentage Deliverer Box Capacity��6   $��28��6B   $
��Left=�6N   $��308,454,0,0q�6i   $	��Top=�6�   $	��231��6�   �� 5   
   -   5        $3��-TotalPercentageDelivererBoxCapacityValueLabel��5   X   $��0��6   $��28��6   $
��Left=�6(   $��561,454,0,0q�6C   $	��Top=�6�   $��51��6�   � 5   
   -   5        $-��'TotalPercentageBoxesAssignedProgressBar��5   R   $��28��6   $
��Left=�6"   $��308,360,0,0q�6=   $	��Top=�6�   $	��304��6�   � 5   
   -   5        $4��.TotalPercentageDelivererBoxCapacityProgressBar��5   R   $��28��6   $
��Left=�6"   $��308,488,0,0q�6=   $	��Top=�6�   $	��304��6�   �� 5   
   -   5        $��AssignClientsToDeliverers��5   t   $"��Assign Clients To Deliverers��6   $��37��68   $
��Left=�6D   $��889,351,0,0q�6_   $	��Top=�6�   $	��185��6�   6�   �� 5   
   -   5        $��RefreshDataGridButton��5   k   $��Refresh Data Grid��6   $	��101��6-   $
��Left=�6:   $��1052,127,0,0q�6U   $	��Top=�6�   $	��108��6�   6�   5      5      
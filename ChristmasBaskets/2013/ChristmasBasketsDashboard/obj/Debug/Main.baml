   M S B A M L   `   `   `  ����  Christmas Basket Dashboard$    ChristmasBasketsDashboard.Main   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      -   5          � Title$!  Christmas Baskets Dashboard��5      $	��519��6)   $	��833��66    � Icon$; 5/Christmas%20Basket%20Dashboard;component/present.ico��6a   .��5      � $	��770��6   $	��438��6    � ColumnDefinitions 5   
   �� 5      $��0* �6   �� 5      $��0* �6   �� 5	      $��0* �6   �� 5
      $
��584* �6   .+�5   
   �� -   5        $��tabControl1��5      $��3��6(   .@�5      �� -   5        $��tabItem1��5   3   $��Connect To Database��6   .��5      � .+�5      �� -   5        $��OpenDatabase��5   Z   $��25��6   $
��Left=�6)   $��10,49.98,0,0q�6D   $	��Top=�6n   $	��130��6�   .��6�   Open Database6�   �� 5      -   5        $��OpenDatabaseIndicator��5   t   
����d�$��10��66   $
��Left=�6B   $��152,57.98,0,0q�6]   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $��SelectYearButton��5   Z   $��25��6   $
��Left=�6)   $��10,89.98,0,0q�6D   $	��Top=�6r   $	��130��6�   .��6�   Select Year6�   �� 5      -   5        $��CloseDatabase��5   [   $��25��6   $
��Left=�6)   $��10,131.98,0,0q�6D   $	��Top=�6p   $	��130��6�   .��6�   Close Database6�   �� 5      -   5        $��SelectedYearLabel��5   Z   $��25��6   $
��Left=�6(   $��152,89.98,0,0q�6C   $	��Top=�6s   $	��121��6�   .��6�   No Year Selected6�   �� 5      -	   5        $��OpenDatabaseLabel��5   d   $��18(�6   $��37��6*   $
��Left=�66   $��	10,10,0,0q�6Q   $	��Top=�6}   $	��145��6�   .��6�   Open Database6�   �� 5      -
   5        $��SelectedYearStatusLabel��5   �   $��Selected Year Status��6   $��18(�6;   $��37��6I   $
��Left=�6U   $��
379,10,0,0q�6p   $	��Top=�6�   $	��178��6�   �� 5      -   5        $/��)Step_1_Year_Created_In_Database_Indicator��5   l   
�������$��10��61   $
��Left=�6=   $��
379,58,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $'��!Step_2_Clients_Imported_Indicator��5   l   
�������$��10��61   $
��Left=�6=   $��
379,82,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $,��&Step_3_Green_Cards_Generated_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��379,129,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $*��$Step_4_Deliverers_Imported_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��379,155,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $5��/Step_5_Clients_Assigned_To_Deliverers_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��379,181,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $/��)Step_6_Generated_Deliverer_Maps_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��379,208,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $#��Step_7_Day_Of_Event_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��379,239,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $8��2Step_7_a_Generate_Unassigned_Clients_Map_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��422,263,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $.��(Step_7_b_Generate_Client_Lists_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��422,290,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5      -   5        $,��&Step_7_c_Generate_Food_Signs_Indicator��5   m   
�������$��10��61   $
��Left=�6=   $��422,318,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5       -   5        $,��&Step_7_d_Generate_Box_Labels_Indicator��5    m   
�������$��10��61   $
��Left=�6=   $��422,346,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5!      -   5        $5��/Step_1_Year_Created_In_Database_Indicator_Label��5!   �   $'��!Step 1 - Year Created In Database��6   $��25��6H   $
��Left=�6T   $��
417,49,0,0q�6o   $	��Top=�6�   $	��314��6�   �� 5"      -   5        $-��'Step_2_Clients_Imported_Indicator_Label��5"   {   $��Step 2 - Clients Imported��6   $��25��6@   $
��Left=�6L   $��
417,72,0,0q�6g   $	��Top=�6�   $	��314��6�   �� 5#      -   5        $2��,Step_3_Green_Cards_Generated_Indicator_Label��5#   �   $$��Step 3 - Green Cards Generated��6   $��25��6E   $
��Left=�6Q   $��417,121,0,0q�6l   $	��Top=�6�   $	��314��6�   �� 5$      -   5        $0��*Step_4_Deliverers_Imported_Indicator_Label��5$      $"��Step 4 - Deliverers Imported��6   $��25��6C   $
��Left=�6O   $��417,147,0,0q�6j   $	��Top=�6�   $	��314��6�   �� 5%      -   5        $;��5Step_5_Clients_Assigned_To_Deliverers_Indicator_Label��5%   �   $-��'Step 5 - Clients Assigned To Deliverers��6   $��25��6N   $
��Left=�6Z   $��417,173,0,0q�6u   $	��Top=�6�   $	��314��6�   �� 5&      -   5        $5��/Step_6_Generated_Deliverer_Maps_Indicator_Label��5&   �   $'��!Step 6 - Generated Deliverer Maps��6   $��25��6H   $
��Left=�6T   $��417,200,0,0q�6o   $	��Top=�6�   $	��314��6�   �� 5'      -   5        $-��'Step_7_Day_Of_Event_Indicator_Indicator��5'   x   $��Step 7 - Day Of Event��6   $��25��6<   $��Right=�6H   $��
0,232,90,0q�6d   $	��Top=�6�   $	��253��6�   �� 5(      -   5        $>��8Step_7_a_Generate_Unassigned_Clients_Map_Indicator_Label��5(   �   $/��)Step 7a - Generate Unassigned Clients Map��6   $��25��6P   $
��Left=�6\   $��471,255,0,0q�6w   $	��Top=�6�   $	��260��6�   �� 5)      -   5        $4��.Step_7_b_Generate_Client_Lists_Indicator_Label��5)   �   $%��Step 7b - Generate Client Lists��6   $��25��6F   $
��Left=�6R   $��471,282,0,0q�6m   $	��Top=�6�   $	��260��6�   �� 5*      -   5        $2��,Step_7_c_Generate_Food_Signs_Indicator_Label��5*   f   $#��Step 7c - Generate Food Signs��6   $��25��6D   $��471,310,29,0q�6P   $	��Top=�6�   �� 5+      -    5        $2��,Step_7_d_Generate_Box_Labels_Indicator_Label��5+   �   $#��Step 7d - Generate Box Labels��6   $��25��6D   $
��Left=�6P   $��471,338,0,0q�6k   $	��Top=�6�   $	��260��6�   �� 5,      -!   5        $4��.Step_2_a_Check_For_Client_Duplicates_Indicator��5,   m   
�������$��10��61   $
��Left=�6=   $��422,105,0,0q�6X   
���   �$	��Top=�6�   $��32��6�   �� 5-      -"   5        $0��*Step_2_a_Check_For_Client_Duplicates_Label��5-   �   $+��%Step 2a - Check for Client Duplicates��6   $��25��6L   $��Right=�6X   $��	0,97,29,0q�6t   $	��Top=�6�   $	��260��6�   5.      5/      �� 50      -#   5        $��tabItem2��50   9   $��Import Clients From Excel��6   .��51      � .+�52      �� -$   5        $.��(ImportSelectedYearClientsFromExcelButton��52   X   $��35��6   $
��Left=�6)   $��
13,180,0,0q�6D   $	��Top=�6�   $	��260��6�   .��6�   )'Import Selected Year Clients From Excel6  �� 53      -%   5        $!��ImportClientsFromExcelLabel��53   e   $��18(�6   $��37��6*   $
��Left=�66   $��
10,137,0,0q�6Q   $	��Top=�6�   $	��391��6�   .��6�   .,Import Clients Into Selected Year From Excel6�   �� 54      -&   5        $'��!ClearSelectedYearBoxNumbersButton��54   K   $
��Left=�6   $��	13,0,0,58q�68   $	��240��6t   $��36��6�   $��Bottom=�6�   .��6�   Set Selected Year Box Numbers6�   �� 55      -'   5        $$��SetSelectedYearBoxNumbersLabel��55   Z   $��18(�6   $
��Left=�6*   $��10,267,0,87q�6E   $	��263��6�   .��6�   Set Selected Year Box Numbers6�   �� 56      -(   5        $��CheckForDuplicatesLabel��56   �   $��Check for Duplicates��6   $��18(�6;   $��37��6I   $
��Left=�6U   $��	13,18,0,0q�6p   $	��Top=�6�   $	��391��6�   �� 57      -)   5        $��CheckForDuplicatesButton��57   v   $��Check For Duplicates��6   $��35��6<   $
��Left=�6H   $��	13,61,0,0q�6c   $	��Top=�6�   $	��260��6�   6�   �� 58      -*   5        $!��CheckForDuplicatesTipsLabel��58   [   $��51��6    $
��Left=�6,   $��
293,61,0,0q�6G   $r��lPut LOA and RCSS Lists in one Excel file, format all cells as TEXT, use this file for the duplicate checking��6~   $
��Wrap=�6�   $	��Top=�6  $	��394��6  $��14(�6*  59      5:      �� 5;      -+   5        $��tabItem3��5;   +   $��Green Cards��6   .��5<      � .+�5=      �� -,   5        $��ExportClientsToExcel��5=   W   $��38��6   $
��Left=�6)   $��	11,76,0,0q�6D   $	��Top=�6s   $	��240��6�   .��6�   '%Export Selected Year Clients To Excel6�   �� 5>      --   5        $ ��SelectOrganizationComboBox��5>   Z   $��26��6   $
��Left=�6+   $��
11,127,0,0q�6F    �� SelectedIndex$ 0��6|   $	��Top=�6�   $	��240��6�   .@�5?      �� .��6'   Select Organization6<   �� 5@      .��6'   LOA (League Of Old Americans)6F   �� 5A      .��6'   '%RCSS (Roanoke County Social Services)6N   �� 5B      .��6'   LOA and RCSS65   5C      �� 5D      -.   5        $��ExportClientsToExcelLabel��5D   d   $��18(�6   $��37��6*   $
��Left=�66   $��	10,10,0,0q�6Q   $	��Top=�6�   $	��240��6�   .��6�   Export Clients To Excel6�   �� 5E      -/   5        $��label1��5E   d   $��12(�6   $��26��6*   $
��Left=�66   $��	11,43,0,0q�6Q   $	��Top=�6r   $	��262��6�   .��6�   .,Give to LOA and RCSS and use for Green Cards6�   5F      5G      �� 5H      -0   5        $��tabItem4��5H   4   $��Deliverer Assignment��6   .��5I      � .+�5J      �� -1   5        $4��.UpdateClientsFromCurrentYearDelivererListLabel��5J   =   $��18(�6   $��	22,20,0,0q�6*   $��37.04��6s   $	��Top=�6�   $
��Left=�6�   $	��165��6�   .��6�   Import Deliverers6�   �� 5K      -2   5        $(��"StampClientsWithSelectedYearButton��5K   L   $��Right=�6   $��	0,86,33,0q�69   $	��278��6v   $��43��6�   $	��Top=�6�   .��6�   1/Stamp Client's Year__Last__Delivered__To = ****6	  �� 5L      -3   5        $��button1��5L   W   $��43��6   $
��Left=�6)   $��	22,63,0,0q�6D   $	��Top=�6f   $	��165��6~   .��6�   Import Deliverers From Excel6�   �� 5M      -4   5        $��label2��5M   @   $��18(�6   $��0,35.92,33,0q�6*   $��37.04��6N   $	��Top=�6]   $��Right=�6u   $	��271��6�   .��6�   !Stamp Clients with Current Year6�   �� 5N      -5   5        $��label3��5N   @   $��18(�6   $��0,173,13,196q�6*   $��Right=�6N   $	��298��6j   .��6v   $"Stamp Deliverers with Current Year6�   �� 5O      -6   5        $+��%StampDeliverersWithSelectedYearButton��5O   1   $��
0,0,33,147q�6   $��Right=�6�   $	��278��6�   $��43.04��6�   $��Bottom=�6�   .��6�   0.Stamp Deliverer's Year__Last__Delivered = ****6  �� 5P      -7   5        $��label7��5P   �   $"��Assign Clients to Deliverers��6   $��18(�6C   $
��Left=�6Q   $��22,273,0,96q�6l   $	��301��6�   �� 5Q      -8   5        $%��AssignClientsToDeliverersButton��5Q   ~   $"��Assign Clients to Deliverers��6   $��43��6D   $
��Left=�6P   $��	22,0,0,40q�6k   $��Bottom=�6�   $	��161��6�   6�   �� 5R      -9   5        $��AutoAssignLabel��5R   �   $.��(Auto Assign Client History To Deliverers��6   $��18(�6O   $��37.04��6]   $
��Left=�6l   $��
22,148,0,0q�6�   $	��Top=�6�   $	��517��6�   �� 5S      -:   5        $/��)AutoAssignClientHistoryToDeliverersButton��5S   �   $.��(Auto Assign Client History To Deliverers��6   $��44��6P   $
��Left=�6\   $��
22,191,0,0q�6w   $	��Top=�6�   $	��243��6�   6�   5T      5U      �� 5V      -;   5        $��tabItem5��5V   .   $��Deliverer Maps��6   .��5W      � .+�5X      �� -<   5        $#��GenerateDelivererPacketsLabel��5X   d   $��18(�6   $��37��6*   $
��Left=�66   $��	10,10,0,0q�6Q   $	��Top=�6�   $	��299��6�   .��6�   Generate Deliverer Packets6�   �� 5Y      -=   5        $#��PrepareDelivererPacketsButton��5Y   |   $ ��Generate Deliverer Packets��6   $��37��6B   $
��Left=�6N   $��	10,53,0,0q�6i   $	��Top=�6�   $	��226��6�   6�   �� 5Z      ->   5        $*��$GenerateUnassignedClientPacketsLabel��5Z   �   $(��"Generate Unassigned Client Packets��6   $��18(�6I   $��37��6W   $
��Left=�6c   $��	6,129,0,0q�6~   $	��Top=�6�   $	��303��6�   �� 5[      -?   5        $��button2��5[   �   $(��"Generate Unassigned Client Packets��6   $��37��6J   $
��Left=�6V   $��
10,172,0,0q�6q   $	��Top=�6�   $	��226��6�   6�   5\      5]      �� 5^      -@   5        $��tabItem6��5^   ,   $��Day of Event��6   .��5_      �  5`      �� 5a      $
��379* �6+   �� 5b      $
��381* �6+   .+�5d      �� -A   5        $'��!ExcelListofClientsForEventButtons��5d   W   $��42��6   $
��Left=�6)   $��	10,52,0,0q�6D   $	��Top=�6�   $	��234��6�   .��6�   !Excel List of Clients For Event6�   �� 5e      -B   5        $��DayOfEventLists��5e   d   $��18(�6   $��37��6*   $
��Left=�66   $��	10,10,0,0q�6Q   $	��Top=�6{   $	��299��6�   .��6�   Day of Event Lists6�   �� 5f      -C   5        $*��$ExcelListofDeliverersForEventButtons��5f   X   $��42��6   $
��Left=�6)   $��
10,114,0,0q�6D   $	��Top=�6�   $	��234��6�   .��6�   $"Excel List of Deliverers For Event6�   �� 5g      -D   5        $��label5��5g   [   $��18(�6   $
��Left=�6*   $��10,192,0,177q�6E   $	��301��6i   .��6u   !Generate Unassigned Clients Map6�   �� 5h      -E   5        $(��"GenerateUnassignedClientsMapButton��5h   [   $��43.04��6   $
��Left=�6,   $��
10,0,0,128q�6G   $��Bottom=�6�   $	��293��6�   .��6�   !Generate Unassigned Clients Map6�   5i      5j      �� 5k      -F   5        $��tabItem7��5k   '   $��Reports��6   .��5l      � 5m      5n      5o      5p      
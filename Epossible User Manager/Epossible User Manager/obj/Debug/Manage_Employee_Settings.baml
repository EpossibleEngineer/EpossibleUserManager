   M S B A M L   `   `   `  ����Q  MEpossible User Manager, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null?$clr-namespace:Epossible_User_ManagerEpossible_User_Manager  5   	   5    /Epossible_User_Manager.Manage_Employee_Settings   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   6local$clr-namespace:Epossible_User_Manager     5   	   X TXceed.Wpf.Toolkit, Version=3.4.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f44xctk)http://schemas.xceed.com/wpf/xaml/toolkit  5   	     � Title$  Manage_Employee_Settings��5	   	   $	��550��6*   $	��800��67   .��5
      � .+�5   
     Xceed.Wpf.Toolkit.Wizard     FinishButtonClosesWindow . .@�5      "  Xceed.Wpf.Toolkit.WizardPage  -   5          Name$ 	IntroPage��5         Title$' !Welcome to Manage Employee Wizard��5   $      Description$? 9This Wizard will assist you on managing employee accounts��5   $     5      -   5        $ Page1��5         PageType$ Interior=�6-   $ Page 1��5   $   $ Employee management Page��5   $      NextPage 5   $   ��  �� ElementName$ Page2��   PreviousPage 5   $   �� $ 	IntroPage��.��5      � .+�5      �� $7��1Please select desired option from the list below:��6   $
��Left=�6X   $��
52,164,0,0q�6s   $	��Top=�6�   $	��660��6�   	� 5      -   5        	 	�Name$	 ResetPassRadioBTN��5   "   $��Reset Employee's Password��6=   $
��Left=�6a   $��
52,208,0,0q�6|   $	��Top=�6�   	� 5      -   5        $	 ChgDeptRadioBTN��5   "   $*��$Change Employee's Primary Department��6;   $
��Left=�6j   $��
52,254,0,0q�6�   $	��Top=�6�   	� 5      -   5        $	 ChgGroupRadioBTN��5   "   $(��"Change Employee's Group Membership��6<   $
��Left=�6i   $��
52,297,0,0q�6�   $	��Top=�6�   	� 5      $%��Disable Employee's User Account��6"   $
��Left=�6L   $��
52,344,0,0q�6g   $	��Top=�6{   �� 5      $��Employee's First Name��6   $
��Left=�6<   $��	52,52,0,0q�6W   $	��Top=�6j   �� 5      $
��Left=�6   $��23��69   $��
191,55,0,0q�6E   
 �� TextWrapping$

 Wrap=�6Y   $	��Top=�6m   $	��120��6�   �� 5      $��Employee's Last Name��6   $
��Left=�6;   $��
342,52,0,0q�6V   $	��Top=�6j   �� 5      $
��Left=�6   $��23��69   $��
498,55,0,0q�6E   $

 Wrap=�6Y   $	��Top=�6m   $	��120��6�   �� 5      $��Submit��6   $
��Left=�6.   $��
668,60,0,0q�6I   $	��Top=�6]   $��75��6u   �� 5      $S��MPlease enter employee's first and last name to find the correct user account.��6   $
��Left=�6t   $��52,9,0,0q�6�   $	��Top=�6�   �� 5       $��Selected User Account:��6   $
��Left=�6=   $��
52,102,0,0q�6X   $	��Top=�6l   �� 5!      $��User account��6   $
��Left=�63   $��225,102,0,0q�6N   $	��Top=�6c   $��86��6{   5#      5$        5%      -   5        $ Page2��5%      $ Interior=�6-   $ Page 2��5&   $   $, &This is the second page in the process��5'   $     5(      -   5        $ LastPage��5(      $ Interior=�60   $ 	Last Page��5)   $   $* $This is the last page in the process��5*   $      	CanFinish$
 TrueX�5+   $   5,      5-      5.      
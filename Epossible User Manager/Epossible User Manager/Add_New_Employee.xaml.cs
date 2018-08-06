using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;


namespace Epossible_User_Manager
{
    /// <summary>
    /// Interaction logic for Add_New_Employee.xaml
    /// </summary>
    public partial class Add_New_Employee : Window
    {
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string GroupMembership { get; set; }
        public string Email { get; set; }
        public string Department { get; set; }
        public string OULDAPPath { get; set; }
        public string DepGroupDN { get; set; }
        public string UserDN { get; set; }
        public string DepGroupLDAPpath { get; set; }
        public string UserLDAPpath { get; set; }
        public string DomainName { get; set; }
        public string DomainDN { get; set; }
        public string GroupList { get; set; }
        public string AllGroupsLDAPpath { get; set; }

        public Add_New_Employee()
        {
            InitializeComponent();

            //Discover domain Name and extract the Domain DN
            DomainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            var index = DomainName.IndexOf(".");
            var domain = DomainName.Substring(0, index);
            var suffix = DomainName.Substring(index + 1);
            DomainDN = "dc=" + domain + ",dc=" + suffix;

        }

        private void SaveBTN_Click(object sender, RoutedEventArgs e)
        {
            //Print User info on last page before user clicking create user
            Firstname = FNTextbox.Text;
            Lastname = LNTextbox.Text;
            UserName = Firstname + "." + Lastname;
            Password = PassTextBox.Text;
            Department = DeptComboBox.Text;
            GroupMembership = Department + "group";
            Email = UserName + "@domainsufix";
            SummaryLabel.Content = string.Format("Employee first name is {0}\n\nEmployee last name is {1}\n\nEmployee UserName is {2}\n\nEmployee password is {3}\n\nEmployee department is {4}\n\nEmployee is a member of the following groups: {5}\n\nEmployee email address is {6}", Firstname, Lastname, UserName, Password, Department, GroupMembership, Email);
            MessageBox.Show("Information successfully saved.");

            AllGroupsLDAPpath = "LDAP://ou=Domain Groups," + DomainDN;
            //Find all groups in Domain Groups OU
            SearchGroup(AllGroupsLDAPpath);
        }

        private void SubmitBTN1_Click(object sender, RoutedEventArgs e)
        {

        }

        private void CreateUser_Click(object sender, RoutedEventArgs e)
        {
            Firstname = FNTextbox.Text;
            Lastname = LNTextbox.Text;
            UserName = Firstname + "." + Lastname;
            Password = PassTextBox.Text;
            CreateUserAccount(Firstname + "." + Lastname, Password);
        }
        public void CreateUserAccount(string userName, string userPassword)
        {
            Firstname = FNTextbox.Text;
            Lastname = LNTextbox.Text;
            Password = PassTextBox.Text;

                        
            try
            {

                OULDAPPath = "LDAP://ou=" + DeptComboBox.Text + "," + DomainDN;
                DepGroupLDAPpath = "LDAP://cn=" + DeptComboBox.Text + "group,ou=Domain Groups," + DomainDN;
                DepGroupDN = "cn=" + DeptComboBox.Text + "group,ou = " + DeptComboBox.Text + "," + DomainDN;
                UserLDAPpath = "LDAP://cn=" + userName + ",ou=" + DeptComboBox.Text + "," + DomainDN;
                UserDN = "cn=" + userName + ",ou=" + DeptComboBox.Text + "," + DomainDN;
                


                DirectoryEntry dirEntry = new DirectoryEntry(OULDAPPath);

                var newUser = dirEntry.Children.Add("CN=" + userName, "user");

                //Add properties to Desktop user
                if (DSCheckbox.IsChecked == true)
                {
                    newUser.Properties["samAccountName"].Add(userName);
                    newUser.Properties["DisplayName"].Add(Firstname + " " + Lastname);
                    newUser.Properties["GivenName"].Add(Firstname);
                    newUser.Properties["sn"].Add(Lastname);
                    newUser.Properties["HomeDirectory"].Add(@"\\ad\home\" + userName);
                    newUser.Properties["HomeDrive"].Add(@"Z:");
                    newUser.Properties["ProfilePath"].Add(@"\\ad\profile\" + userName);
                    newUser.Properties["mail"].Add(userName + "@emailsuffix");
                    newUser.Properties["UserPrincipalName"].Add(userName +"@"+DomainName);
                }

                //Add properties to Laptop user
                if (LTCheckbox.IsChecked == true)
                {
                    newUser.Properties["samAccountName"].Add(userName);
                    newUser.Properties["DisplayName"].Add(Firstname + " " + Lastname);
                    newUser.Properties["GivenName"].Add(Firstname);
                    newUser.Properties["sn"].Add(Lastname);
                    newUser.Properties["HomeDirectory"].Add(@"\\ad\home\" + userName);
                    newUser.Properties["HomeDrive"].Add(@"Z:");
                    newUser.Properties["mail"].Add(userName + "@emailsuffix");
                    newUser.Properties["UserPrincipalName"].Add(userName + "@" + DomainName);
                }

                newUser.CommitChanges();

                //Set user password
                newUser.Invoke("SetPassword", new object[] { userPassword });
                
                //Enable user account
                newUser.Properties["userAccountControl"].Value = 0x200;
                newUser.CommitChanges();
                dirEntry.Close();
                newUser.Close();

                //Add new user to group
                AddToGroup(UserDN, DepGroupLDAPpath);
                
                System.Windows.MessageBox.Show("User was created successfully. Please click finish.");
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                System.Windows.MessageBox.Show("Error creating user");

            }

        }
        public void AddToGroup(string userDn, string groupldappath)
        {
            try
            {
                //This method would add the user to the AD group
                DirectoryEntry dirEntry = new DirectoryEntry(groupldappath);
                dirEntry.Properties["member"].Add(userDn);
                dirEntry.CommitChanges();
                dirEntry.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException)
            {


            }
        }
        
        public void SearchGroup(string direntryldappath)
        {
            DirectoryEntry entry = new DirectoryEntry(direntryldappath);

            DirectorySearcher dSearch = new DirectorySearcher(entry);
            dSearch.Filter = "(&(objectClass=group))";
            dSearch.SearchScope = SearchScope.Subtree;

            SearchResultCollection results = dSearch.FindAll();
            
            foreach (SearchResult found in results)
            {
               
               
                string group = (found.Properties["Name"][0].ToString());
                var chbox = new CheckBox();
                chbox.Content = group;
                GroupListView.Items.Add(chbox);
                                       
            }
           
        }

    }
}

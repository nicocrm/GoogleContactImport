using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.GData.Client;
using System.Data.OleDb;
using Google.GData.Contacts;
using Google.GData.Extensions;

namespace ContactImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            ImportAllContacts(txtEmail.Text, txtPassword.Text);
            MessageBox.Show("DONE");
        }

        private static void ImportAllContacts(String username, String password)
        {
            Service service = new Service("cp", "exampleCo-exampleApp-1");
            // Set your credentials:
            service.setUserCredentials(username, password);
            service.ProtocolMajor = 3;
            service.ProtocolMinor = 0;

            //ContactsRequest cr = new ContactsRequest();
            OleDbConnection cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Documents and Settings\\Administrator\\My Documents\\BlackBerry\\Backup;Extended Properties=\"text;HDR=Yes;FMT=Delimited\";");
            cn.Open();
            OleDbCommand cmd = cn.CreateCommand();
            cmd.CommandText = "select * from [MB_Export_02205018.csv]";
            using (var rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    if (rdr[0].ToString() == "")
                        continue;
                    ContactEntry entry = new ContactEntry();
                    entry.Name = new Google.GData.Extensions.Name { FullName = rdr["Name"].ToString(), GivenName = rdr["First Name"].ToString(), FamilyName = rdr["Last Name"].ToString() };
                    entry.GroupMembership.Add(new GroupMembership { HRef = "https://www.google.com/m8/feeds/groups/" + username + "/base/6" });
                    if (rdr["Company"].ToString() != "")
                        entry.Organizations.Add(new Organization { Name = rdr["Company"].ToString(), Title = rdr["Job"].ToString(), Rel = "http://schemas.google.com/g/2005#work" });
                    if (rdr["Email"].ToString() != "")
                        entry.Emails.Add(new Google.GData.Extensions.EMail { Label = "Home", Address = rdr["Email"].ToString() });
                    if (rdr["Mobile"].ToString() != "")
                        entry.Phonenumbers.Add(new PhoneNumber { Label = "Mobile", Value = rdr["Mobile"].ToString() });
                    if (rdr["Work Phone 1"].ToString() != "")
                        entry.Phonenumbers.Add(new PhoneNumber { Label = "Work", Value = rdr["Work Phone 1"].ToString() });
                    if (rdr["Work Phone 2"].ToString() != "")
                        entry.Phonenumbers.Add(new PhoneNumber { Label = "Work 2", Value = rdr["Work Phone 2"].ToString() });
                    if (rdr["Home Phone 1"].ToString() != "")
                        entry.Phonenumbers.Add(new PhoneNumber { Label = "Home", Value = rdr["Home Phone 1"].ToString() });
                    if (rdr["Home Address 1"].ToString() != "")
                    {
                        entry.PostalAddresses.Add(new StructuredPostalAddress
                        {
                            Label = "Home",
                            Street = rdr["Home Address 1"].ToString(),
                            City = rdr["Home City"].ToString(),
                            Region = rdr["Home State"].ToString(),
                            Postcode = rdr["Home ZIP"].ToString()
                        });
                    }
                    if (rdr["Work Address 1"].ToString() != "")
                    {
                        entry.PostalAddresses.Add(new StructuredPostalAddress
                        {
                            Label = "Work",
                            Street = rdr["Work Address 1"].ToString(),
                            City = rdr["Work City"].ToString(),
                            Region = rdr["Work State"].ToString(),
                            Postcode = rdr["Work ZIP"].ToString()
                        });
                    }
                    if (rdr["Birthday"].ToString() != "")
                    {
                        DateTime bd;
                        if (DateTime.TryParse(rdr["Birthday"].ToString(), out bd))
                        {
                            entry.Birthday = bd.ToString("yyyy-MM-dd");
                        }
                    }
                    Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
                    AtomEntry insertedEntry = service.Insert(feedUri, entry);
                }
            }
        }
    }
}

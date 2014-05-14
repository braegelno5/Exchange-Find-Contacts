using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace FindContact
{
    class ContactAdpator : IContactDB
    {
        public void AddContact(Contact contact)
        {
            SqlConnection connection = DBConnection();
            using(connection){
                try
                {
                    SqlCommand cmd = new SqlCommand("INSERT INTO CONTACTS (ContactID, DisplayName, Alias, AssistantName, BusinessHomePage, CompanyName, " +
                     "CompleteName, Department, DirectoryId, FileAs, Generation, GivenName, " +
                     "JobTitle, Manager, Mileage, NickName, Notes, OfficeLocation, PhoneticFirstName, " +
                     "PhoneticFullName, PhoneticLastName, Profession, SpouseName, Surname)" +
                     "VALUES(@ID, @DisplayName, @Alias, @AssistantName, @BusinessHomePage, @CompanyName, " +
                     "@CompleteName,@Department, @DirectoryId, @FileAs, " +
                     "@Generation, @GivenName, @JobTitle, @Manager, @Mileage, @NickName, " +
                     "@Notes, @OfficeLocation, @PhoneticFirstName, @PhoneticFullName, @PhoneticLastName, @Profession, " +
                     "@SpouseName, @Surname)");
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = connection;
                    CreateContactParameters(contact, cmd);
                    cmd.ExecuteNonQuery();
                }
                catch (SqlException e)
                {
                    Console.WriteLine(e.StackTrace + " " + e.Message);
                }
            }
        }
        public void UpdateContact(Contact contact)
        {
            SqlConnection connection = DBConnection();
            using (connection)
            {
                SqlCommand cmd = new SqlCommand("UPDATE Contacts SET ContactID =  @ID, DisplayName = @DisplayName, Alias = @Alias, AssistantName = @AssistantName, BusinessHomePage = @BusinessHomePage, " +
                " CompanyName = @CompanyName, CompleteName = @CompleteName, Department = @Department, DirectoryId = @DirectoryId, FileAs = @FileAs, Generation = @Generation, GivenName = @GivenName, @JobTitle = JobTitle, " +
                " Manager = @Manager, Mileage = @Mileage, Nickname = @Nickname, Notes = @Notes, OfficeLocation = @OfficeLocation, PhoneticFirstName = @PhoneticFirstName, PhoneticFullName = @PhoneticFullName, Profession = @Profession, "+
                " SpouseName = @SpouseName, Surname = @Surname WHERE CONTACTID = @ID");
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                CreateContactParameters(contact, cmd);
                cmd.ExecuteNonQuery();
            }
        }
        public void DeleteContact(ItemId id)
        {
            Console.WriteLine(id);
            SqlConnection connection = DBConnection();
            using (connection)
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM Contacts WHERE CONTACTID = @ID");
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                cmd.Parameters.AddWithValue("@ID", id.ToString());
                cmd.ExecuteNonQuery();
            }

        }
        public SqlConnection DBConnection()
        {
            String connstr = ConfigurationManager.ConnectionStrings["ConnStr"].ToString();
            SqlConnection cnn;
            cnn = new SqlConnection(connstr);
            try{
                cnn.Open();
            }
            catch(SqlException e)
            {
                Console.WriteLine("Can't connect: " + e.Message);
            }
            return cnn;
        }
        /* Adds any changes to the contacts to the audit table, regardless of changes/updates/deletions. */
        public void PushUpdatesToAuditTable(String action, Contact contacts) 
        {
            SqlConnection connection = DBConnection();
            using (connection)
            {
                SqlCommand cmd = new SqlCommand("INSERT INTO CONTACTLOG (ContactID, DisplayName, Alias, AssistantName, BusinessHomePage, CompanyName, " +
                "CompleteName, Department, DirectoryId, FileAs, Generation, GivenName, " +
                "JobTitle, Manager, Mileage, NickName, Notes, OfficeLocation, PhoneticFirstName, " +
                "PhoneticFullName, PhoneticLastName, Profession, SpouseName, Surname, Action, ActionDate)" +
                "VALUES(@ID, @DisplayName, @Alias, @AssistantName, @BusinessHomePage, @CompanyName, " +
                "@CompleteName,@Department, @DirectoryId, @FileAs, " +
                "@Generation, @GivenName, @JobTitle, @Manager, @Mileage, @NickName, " +
                "@Notes, @OfficeLocation, @PhoneticFirstName, @PhoneticFullName, @PhoneticLastName, @Profession, " +
                "@SpouseName, @Surname, @Action, @ActionDate)");
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                CreateContactParameters(contacts, cmd);
                cmd.Parameters.AddWithValue("@Action", action);
                cmd.Parameters.AddWithValue("@ActionDate",DateTime.Now);
                cmd.ExecuteNonQuery();
            }
        }
        /* Get all contacts in the users contacts folder - Needs to be called once. */
        public void GetAllContacts(ExchangeService service)
        {
            Contacts contacts = new Contacts();
            ContactsFolder contactsFolder = ContactsFolder.Bind(service, WellKnownFolderName.Contacts, new PropertySet(BasePropertySet.IdOnly, FolderSchema.TotalCount));
            ItemView view = new ItemView(contactsFolder.TotalCount);
            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ContactSchema.Id);
            view.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ContactSchema.DisplayName);
            FindItemsResults<Item> contactItems = service.FindItems(WellKnownFolderName.Contacts, view);

            // Display the list of contacts. 
            foreach (Item item in contactItems)
            {
                if (item is Contact)
                {
                    Contact contact = item as Contact;
                    contacts.id = contact.Id.ToString();
                    contacts.DisplayName = contact.DisplayName.ToString();
                   // AddContact(service, contact);
                }
            }
        }
        /* The method is not really used - It's just here because I can use it for debugging */
        public void DeleteAllContacts()
        {
            SqlConnection connection = DBConnection();
            using (connection)
            {
                SqlCommand cmd = new SqlCommand("DELETE FROM Contacts");
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                cmd.ExecuteNonQuery();
            }
        }
        private void CreateContactParameters(Contact contact, SqlCommand cmd)
        {
            cmd.Parameters.AddWithValue("@ID", contact.Id.ToString());
            cmd.Parameters.AddWithValue("@DisplayName", contact.DisplayName.ToString());
            cmd.Parameters.AddWithValue("@Alias", ValueOrDbNull(contact.Alias));
            cmd.Parameters.AddWithValue("@AssistantName", ValueOrDbNull(contact.AssistantName));
            cmd.Parameters.AddWithValue("@BusinessHomePage", ValueOrDbNull(contact.BusinessHomePage));
           
            cmd.Parameters.AddWithValue("@Companies", ValueOrDbNull(contact.Companies[0].ToString()));
            cmd.Parameters.AddWithValue("@CompanyName", ValueOrDbNull(contact.CompanyName));
            cmd.Parameters.AddWithValue("@CompleteName", ValueOrDbNull(contact.CompleteName.FullName));
            cmd.Parameters.AddWithValue("@Department", ValueOrDbNull(contact.Department));
            cmd.Parameters.AddWithValue("@DirectoryId", ValueOrDbNull(contact.DirectoryId));
            cmd.Parameters.AddWithValue("@EmailAddress1", ValueOrDbNull(contact.EmailAddresses[EmailAddressKey.EmailAddress1].ToString()));
            cmd.Parameters.AddWithValue("@EmailAddress2", ValueOrDbNull(contact.EmailAddresses[EmailAddressKey.EmailAddress2].ToString()));
            cmd.Parameters.AddWithValue("@EmailAddress3", ValueOrDbNull(contact.EmailAddresses[EmailAddressKey.EmailAddress3].ToString()));
            cmd.Parameters.AddWithValue("@FileAs", ValueOrDbNull(contact.FileAs));
            cmd.Parameters.AddWithValue("@Generation", ValueOrDbNull(contact.Generation));
            cmd.Parameters.AddWithValue("@GivenName", ValueOrDbNull(contact.GivenName));
            cmd.Parameters.AddWithValue("@ImAddress1", ValueOrDbNull(contact.ImAddresses[ImAddressKey.ImAddress1].ToString()));
            cmd.Parameters.AddWithValue("@ImAddress2", ValueOrDbNull(contact.ImAddresses[ImAddressKey.ImAddress1].ToString()));
            cmd.Parameters.AddWithValue("@ImAddress3", ValueOrDbNull(contact.ImAddresses[ImAddressKey.ImAddress1].ToString()));
            cmd.Parameters.AddWithValue("@JobTitle", ValueOrDbNull(contact.JobTitle));
            cmd.Parameters.AddWithValue("@Manager", ValueOrDbNull(contact.Manager));
            cmd.Parameters.AddWithValue("@MiddleName", ValueOrDbNull(contact.MiddleName));
            cmd.Parameters.AddWithValue("@Mileage", ValueOrDbNull(contact.Mileage));
            cmd.Parameters.AddWithValue("@NickName", ValueOrDbNull(contact.NickName));
            cmd.Parameters.AddWithValue("@Notes", ValueOrDbNull(contact.Notes));
            cmd.Parameters.AddWithValue("@OfficeLocation", ValueOrDbNull(contact.OfficeLocation));
            cmd.Parameters.AddWithValue("@BusinessPhone", ValueOrDbNull(contact.PhoneNumbers[PhoneNumberKey.BusinessPhone].ToString()));
            cmd.Parameters.AddWithValue("@HomePhone", ValueOrDbNull(contact.PhoneNumbers[PhoneNumberKey.HomePhone].ToString()));
            cmd.Parameters.AddWithValue("@BusinessPhone", ValueOrDbNull(contact.PhoneNumbers[PhoneNumberKey.BusinessPhone].ToString()));
            cmd.Parameters.AddWithValue("@PhoneticFirstName", ValueOrDbNull(contact.PhoneticFirstName));
            cmd.Parameters.AddWithValue("@PhoneticFullName", ValueOrDbNull(contact.PhoneticFullName));
            cmd.Parameters.AddWithValue("@PhoneticLastName", ValueOrDbNull(contact.PhoneticLastName));
            cmd.Parameters.AddWithValue("@Profession", ValueOrDbNull(contact.Profession));
            cmd.Parameters.AddWithValue("@SpouseName", ValueOrDbNull(contact.SpouseName));
            cmd.Parameters.AddWithValue("@Surname", ValueOrDbNull(contact.Surname));
        }

        private object ValueOrDbNull(String val)
        {
            if (val == null) return DBNull.Value;
            return val;
        }
        /* Update the Audit Table */
        public void PushUpdatesToAuditTable(String action, ItemId itemId)
        {
            SqlConnection connection = DBConnection();
            using (connection)
            {
                SqlCommand cmd = new SqlCommand("INSERT INTO CONTACTLOG(ContactID, DisplayName, Alias, AssistantName, BusinessHomePage, CompanyName, " +
                "CompleteName, Department, DirectoryId, FileAs, Generation, GivenName, " +
                "JobTitle, Manager, Mileage, NickName, Notes, OfficeLocation, PhoneticFirstName, " +
                "PhoneticFullName, PhoneticLastName, Profession, SpouseName, Surname) (SELECT ContactID, DisplayName, Alias, AssistantName, BusinessHomePage, CompanyName, " +
                "CompleteName, Department, DirectoryId, FileAs, Generation, GivenName, " +
                "JobTitle, Manager, Mileage, NickName, Notes, OfficeLocation, PhoneticFirstName, " +
                "PhoneticFullName, PhoneticLastName, Profession, SpouseName, Surname FROM CONTACTS WHERE ContactID = @ID)");
                cmd.CommandType = CommandType.Text;
                cmd.Connection = connection;
                cmd.Parameters.AddWithValue("@ID", itemId.ToString());
                cmd.ExecuteNonQuery();
                SqlCommand cmd2 = new SqlCommand("UPDATE CONTACTLOG SET Action = @Action, ActionDate = @ActionDate WHERE Action IS NULL");
                cmd2.CommandType = CommandType.Text;
                cmd2.Connection = connection;
                cmd2.Parameters.AddWithValue("@Action", action);
                cmd2.Parameters.AddWithValue("@ActionDate", DateTime.Now);
                cmd2.ExecuteNonQuery();
            }

        }
    }
}

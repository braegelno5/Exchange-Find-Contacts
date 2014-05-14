using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Threading;
using System.IO;

namespace FindContact
{
    class Program
    {
        String password = ConfigurationManager.AppSettings["Password"];
        String username = ConfigurationManager.AppSettings["Username"];
        String ewsAddress = ConfigurationManager.AppSettings["ExchangeURL"];
        String sSyncState;
        ContactAdpator contact = new ContactAdpator();

        private void ConnectToEWS()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
            service.UseDefaultCredentials = true;
            service.Credentials = new WebCredentials(username,password);
            service.Url = new Uri(ewsAddress);
            //Check if there is a sSyncState before running Sync. 
            ReadSyncState();
            //If sSyncState is null then the application has lost it sync.
            //Will need to empty database, and repopulate it with the contacts.
            SyncContacts(service);
        }

        private void ReadSyncState() 
        {
            if(File.Exists("data.dat")){
                StreamReader SyncStateReader = new StreamReader("data.dat");
                sSyncState = SyncStateReader.ReadLine();
                SyncStateReader.Close();
            }
        }

        private void WriteSyncState()
        {
            StreamWriter SyncStateWriter = new StreamWriter("data.dat");
            SyncStateWriter.Write(sSyncState);
            SyncStateWriter.Close();
        }

        private void SyncContacts(ExchangeService service)
        {
            // Initialize the flag that will indicate when there are no more changes.
            bool isEndOfChanges = false;

            // Call SyncFolderItems repeatedly until no more changes are available.
            // sSyncState represents the sync state value that was returned in the prior synchronization response.
            do
            {
                ChangeCollection<ItemChange> icc  = service.SyncFolderItems(new FolderId(WellKnownFolderName.Contacts), PropertySet.FirstClassProperties, null, 512, SyncFolderItemsScope.NormalItems, sSyncState);
                if (icc.Count == 0)
                {
                    Console.WriteLine("There are no item changes to synchronize.");
                }
                else
                {
                    foreach (ItemChange ic in icc)
                    {
                        
                        if (ic.ChangeType == ChangeType.Create)
                        {
                            Contact contacts = Contact.Bind(service, ic.ItemId);
                            contact.AddContact(contacts);
                            contact.PushUpdatesToAuditTable("Added the record", contacts);
                        }
                        else if (ic.ChangeType == ChangeType.Update)
                        {
                            Contact contacts = Contact.Bind(service, ic.ItemId);
                            contact.UpdateContact(contacts);
                            contact.PushUpdatesToAuditTable("Updated the record", contacts);
                        }
                        else if (ic.ChangeType == ChangeType.Delete)
                        {
                           contact.PushUpdatesToAuditTable("Deleted the record", ic.ItemId);
                           contact.DeleteContact(ic.ItemId);
                        }
                        else if (ic.ChangeType == ChangeType.Delete)
                        {
                            //TODO: Update the item's read flag on the client.
                        }
                    }
                }

                // Save the sync state for use in future SyncFolderHierarchy calls.
                sSyncState = icc.SyncState;
                WriteSyncState();
                if (!icc.MoreChangesAvailable)
                {
                    isEndOfChanges = false;
                }
            } while (isEndOfChanges); 
        }

        static void Main(string[] args)
        {
            new Program().ConnectToEWS();
            Console.Write("Please press any key to continue");
            Console.ReadKey();
        }
    }
}

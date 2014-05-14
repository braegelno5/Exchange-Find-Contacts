using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FindContact
{
    interface IContactDB
    {
        void AddContact(Contact contact);
        void UpdateContact(Contact contact);
        void DeleteContact(ItemId id);
        SqlConnection DBConnection();
        void GetAllContacts(ExchangeService service);
        void DeleteAllContacts();
        void PushUpdatesToAuditTable(String action, Contact contacts);
        void PushUpdatesToAuditTable(String action, ItemId itemId);
    }
}

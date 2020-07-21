using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Text;

namespace ServiceJournalEntryApDll
{
    public class Result
    {
        public string CreatedDocumentEntry { get; set; }
        public bool IsSuccessCode { get; set; }
        public string StatusDescription { get; set; }
        public BoObjectTypes ObjectType { get; set; }
    }
}

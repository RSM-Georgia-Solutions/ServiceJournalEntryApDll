using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;

namespace ServiceJournalEntryApDll
{
    public interface IDocumentHelper
    {
        IEnumerable<Result> PostIncomeTaxFromCreditMemo(string invDocEnttry, Company Company);
        IEnumerable<Result> PostIncomeTaxFromInvoice(string invDocEnttry, Company Company);
        IEnumerable<Result> PostIncomeTaxFromOutgoing(string invDocEnttry, Company Company);
        IEnumerable<Result> PostPension(string invoiceDocentry, Company Company);
    }
}

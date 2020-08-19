using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;

namespace ServiceJournalEntryApDll
{
    public interface IDocumentHelper
    {

        IEnumerable<Result> PostIncomeTaxFromCreditMemo(string invDocEnttry);
        IEnumerable<Result> PostIncomeTaxFromInvoice(string invDocEnttry);
        IEnumerable<Result> PostIncomeTaxFromOutgoing(string invDocEnttry);
        IEnumerable<Result> PostPension(string invoiceDocentry);
    }
}

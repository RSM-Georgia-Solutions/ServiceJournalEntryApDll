using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace ServiceJournalEntryApDll
{
    public class DocumentHelper : IDocumentHelper
    {
        public IEnumerable<Result> PostIncomeTaxFromCreditMemo(string invDocEnttry, Company Company)
        {
            List<Result> results = new List<Result>();
            Documents invoiceDi = (Documents)Company.GetBusinessObject(BoObjectTypes.oPurchaseCreditNotes);
            invoiceDi.GetByKey(int.Parse(invDocEnttry, CultureInfo.InvariantCulture));
            string bpCode = invoiceDi.CardCode;
            Recordset recSet = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery(
                $"SELECT U_IncomeTaxPayer, U_PensionPayer FROM OCRD WHERE OCRD.CardCode = N'{bpCode}'");
            bool isIncomeTaxPayer = recSet.Fields.Item("U_IncomeTaxPayer").Value.ToString() == "01";
            bool isPensionPayer = recSet.Fields.Item("U_PensionPayer").Value.ToString() == "01";
            recSet.DoQuery($"Select * From [@RSM_SERVICE_PARAMS]");
            string incomeTaxAccDr = recSet.Fields.Item("U_IncomeTaxAccDr").Value.ToString();
            string incomeTaxAccCr = recSet.Fields.Item("U_IncomeTaxAccCr").Value.ToString();
            string incomeTaxControlAccCr = recSet.Fields.Item("U_IncomeTaxAccCr").Value.ToString();
            string incomeTaxOnInvoice = recSet.Fields.Item("U_IncomeTaxOnInvoice").Value.ToString();

            if (!Convert.ToBoolean(incomeTaxOnInvoice))
            {
                results.Add(new Result { IsSuccessCode = false, StatusDescription = "არ არის საშემოსავლოს გადამხდელი" });
                return results;
            }

            BusinessPartners bp =
                (BusinessPartners)Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            bp.GetByKey(invoiceDi.CardCode);

            var incomeTaxPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_IncomeTaxPayerPercent").Value.ToString(),
                CultureInfo.InstalledUICulture);

            var pensionPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_PensionPayerPercent").Value.ToString());

            for (int i = 0; i < invoiceDi.Lines.Count; i++)
            {
                invoiceDi.Lines.SetCurrentLine(i);
                recSet.DoQuery(
                    $"SELECT U_PensionLiable FROM OITM WHERE OITM.ItemCode = N'{invoiceDi.Lines.ItemCode}'");
                bool isPensionLiable = recSet.Fields.Item("U_PensionLiable").Value.ToString() == "01";
                bool isFc = invoiceDi.DocCurrency != "GEL";

                double incomeTaxAmount;

                if (!isPensionLiable)
                {
                    double lineTotal = isFc ? invoiceDi.Lines.RowTotalFC : invoiceDi.Lines.LineTotal;
                    incomeTaxAmount = Math.Round(lineTotal * incomeTaxPayerPercent / 100, 6);
                }
                else
                {
                    if (isPensionPayer)
                    {
                        double lineTotal = isFc ? invoiceDi.Lines.RowTotalFC : invoiceDi.Lines.LineTotal;
                        double pensionAmount = Math.Round(lineTotal * pensionPayerPercent / 100, 6);
                        incomeTaxAmount = Math.Round((lineTotal - pensionAmount) * incomeTaxPayerPercent / 100, 6);
                    }
                    else
                    {
                        double lineTotal = isFc ? invoiceDi.Lines.RowTotalFC : invoiceDi.Lines.LineTotal;
                        incomeTaxAmount = Math.Round((lineTotal) * incomeTaxPayerPercent / 100, 6);
                    }

                }

                if (isIncomeTaxPayer && invoiceDi.CancelStatus == CancelStatusEnum.csNo)
                {
                    try
                    {
                        string incomeTaxPayerTransId = DiManager.AddJournalEntry(Company, incomeTaxAccCr,
                            incomeTaxAccDr, incomeTaxControlAccCr, invoiceDi.CardCode, -incomeTaxAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);
                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საშემოსავლოს გატარება კრედიტ მემო",
                            CreatedDocumentEntry = incomeTaxPayerTransId,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }
                if (isIncomeTaxPayer && invoiceDi.CancelStatus == CancelStatusEnum.csCancellation)
                {
                    try
                    {
                        string incomeTaxPayerTransId = DiManager.AddJournalEntry(Company, incomeTaxAccCr,
                            incomeTaxAccDr, incomeTaxControlAccCr, invoiceDi.CardCode, incomeTaxAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);

                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საშემოსავლოს გატარება (მაქენსელებელი) კრედიტ მემო",
                            CreatedDocumentEntry = incomeTaxPayerTransId,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }
            }
            return results;
        }
        public IEnumerable<Result> PostIncomeTaxFromInvoice(string invDocEnttry, Company Company)
        {
            List<Result> results = new List<Result>();
            Documents invoiceDi = (Documents)Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            invoiceDi.GetByKey(int.Parse(invDocEnttry, CultureInfo.InvariantCulture));
            string bpCode = invoiceDi.CardCode;
            bool isFc = invoiceDi.DocCurrency != "GEL";

            Recordset recSet = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            recSet.DoQuery(
                $"SELECT U_IncomeTaxPayer, U_PensionPayer FROM OCRD WHERE OCRD.CardCode = N'{bpCode}'");

            bool isIncomeTaxPayer = recSet.Fields.Item("U_IncomeTaxPayer").Value.ToString() == "01";
            bool isPensionPayer = recSet.Fields.Item("U_PensionPayer").Value.ToString() == "01";
            recSet.DoQuery($"Select * From [@RSM_SERVICE_PARAMS]");
            string incomeTaxAccDr = recSet.Fields.Item("U_IncomeTaxAccDr").Value.ToString();
            string incomeTaxAccCr = recSet.Fields.Item("U_IncomeTaxAccCr").Value.ToString();
            string incomeTaxControlAccCr = recSet.Fields.Item("U_IncomeTaxAccCr").Value.ToString();
            string incomeTaxOnInvoice = recSet.Fields.Item("U_IncomeTaxOnInvoice").Value.ToString();

            if (!Convert.ToBoolean(incomeTaxOnInvoice))
            {
                results.Add(new Result { IsSuccessCode = false, StatusDescription = "არ არის საშემოსავლოს გადამხდელი" });
                return results;
            }

            SAPbobsCOM.BusinessPartners bp =
                (SAPbobsCOM.BusinessPartners)Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            bp.GetByKey(invoiceDi.CardCode);

            var incomeTaxPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_IncomeTaxPayerPercent").Value.ToString(),
                CultureInfo.InstalledUICulture);

            var pensionPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_PensionPayerPercent").Value.ToString());

            for (int i = 0; i < invoiceDi.Lines.Count; i++)
            {
                invoiceDi.Lines.SetCurrentLine(i);
                recSet.DoQuery(
                    $"SELECT U_PensionLiable FROM OITM WHERE OITM.ItemCode = N'{invoiceDi.Lines.ItemCode}'");
                bool isPensionLiable = recSet.Fields.Item("U_PensionLiable").Value.ToString() == "01";


                double incomeTaxAmount;

                if (!isPensionLiable)
                {
                    double lineTotal = isFc ? invoiceDi.Lines.RowTotalFC : invoiceDi.Lines.LineTotal;
                    incomeTaxAmount = Math.Round(lineTotal * incomeTaxPayerPercent / 100, 6);
                }
                else
                {
                    if (isPensionPayer)
                    {
                        double lineTotal = isFc ? invoiceDi.Lines.RowTotalFC : invoiceDi.Lines.LineTotal;
                        double pensionAmount = Math.Round(lineTotal * pensionPayerPercent / 100, 6);
                        incomeTaxAmount = Math.Round((lineTotal - pensionAmount) * incomeTaxPayerPercent / 100, 6);
                    }
                    else
                    {
                        double lineTotal = isFc ? invoiceDi.Lines.RowTotalFC : invoiceDi.Lines.LineTotal;
                        incomeTaxAmount = Math.Round((lineTotal) * incomeTaxPayerPercent / 100, 6);
                    }

                }



                if (isIncomeTaxPayer && invoiceDi.CancelStatus == CancelStatusEnum.csNo)
                {
                    try
                    {
                        string incomeTaxPayerTransId = DiManager.AddJournalEntry(Company, incomeTaxAccCr,
                            incomeTaxAccDr, incomeTaxControlAccCr, invoiceDi.CardCode, incomeTaxAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);
                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საშემოსავლოს გატარება ინვოისი",
                            CreatedDocumentEntry = incomeTaxPayerTransId,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }
                if (isIncomeTaxPayer && invoiceDi.CancelStatus == CancelStatusEnum.csCancellation)
                {
                    try
                    {
                        string incomeTaxPayerTransId = DiManager.AddJournalEntry(Company, incomeTaxAccCr,
                            incomeTaxAccDr, incomeTaxControlAccCr, invoiceDi.CardCode, -incomeTaxAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);
                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საშემოსავლოს გატარება (მაქენსელებელი) ინვოისი",
                            CreatedDocumentEntry = incomeTaxPayerTransId,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }

            }
            return results;
        }
        public IEnumerable<Result> PostIncomeTaxFromOutgoing(string invDocEnttry, Company Company)
        {
            List<Result> results = new List<Result>();
            Documents invoiceDi = (Documents)Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            invoiceDi.GetByKey(int.Parse(invDocEnttry, CultureInfo.InvariantCulture));
            string bpCode = invoiceDi.CardCode;

            Recordset recSet = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);


            recSet.DoQuery(
                $"SELECT U_IncomeTaxPayer, U_PensionPayer FROM OCRD WHERE OCRD.CardCode = N'{bpCode}'");

            bool isIncomeTaxPayer = recSet.Fields.Item("U_IncomeTaxPayer").Value.ToString() == "01";
            bool isPensionPayer = recSet.Fields.Item("U_PensionPayer").Value.ToString() == "01";
            recSet.DoQuery($"Select * From [@RSM_SERVICE_PARAMS]");
            string incomeTaxAccDr = recSet.Fields.Item("U_IncomeTaxAccDr").Value.ToString();
            string incomeTaxAccCr = recSet.Fields.Item("U_IncomeTaxAccCr").Value.ToString();
            string incomeTaxControlAccCr = recSet.Fields.Item("U_IncomeTaxAccCr").Value.ToString();
            string incomeTaxOnInvoice = recSet.Fields.Item("U_IncomeTaxOnInvoice").Value.ToString();

            if (!Convert.ToBoolean(incomeTaxOnInvoice))
            {
                results.Add(new Result { IsSuccessCode = false, StatusDescription = "საშემოსავლოს გატარება ადახდაზე ალამი არ არის მონიშნული" });
                return results;
            }

            SAPbobsCOM.BusinessPartners bp =
                (SAPbobsCOM.BusinessPartners)Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            bp.GetByKey(invoiceDi.CardCode);

            var incomeTaxPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_IncomeTaxPayerPercent").Value.ToString(),
                CultureInfo.InstalledUICulture);

            var pensionPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_PensionPayerPercent").Value.ToString());

            for (int i = 0; i < invoiceDi.Lines.Count; i++)
            {
                invoiceDi.Lines.SetCurrentLine(i);
                recSet.DoQuery(
                    $"SELECT U_PensionLiable FROM OITM WHERE OITM.ItemCode = N'{invoiceDi.Lines.ItemCode}'");
                bool isPensionLiable = recSet.Fields.Item("U_PensionLiable").Value.ToString() == "01";


                double incomeTaxAmount;

                if (!isPensionLiable)
                {
                    double lineTotal = invoiceDi.Lines.LineTotal;
                    incomeTaxAmount = Math.Round(lineTotal * incomeTaxPayerPercent / 100, 6);
                }
                else
                {
                    if (isPensionPayer)
                    {
                        double lineTotal = invoiceDi.Lines.LineTotal;
                        double pensionAmount = Math.Round(lineTotal * pensionPayerPercent / 100, 6);
                        incomeTaxAmount = Math.Round((lineTotal - pensionAmount) * incomeTaxPayerPercent / 100, 6);
                    }
                    else
                    {
                        double lineTotal = invoiceDi.Lines.LineTotal;
                        incomeTaxAmount = Math.Round((lineTotal) * incomeTaxPayerPercent / 100, 6);
                    }

                }



                if (isIncomeTaxPayer && invoiceDi.CancelStatus == CancelStatusEnum.csNo)
                {
                    try
                    {
                        string incomeTaxPayerTransId = DiManager.AddJournalEntry(Company, incomeTaxAccCr,
                            incomeTaxAccDr, incomeTaxControlAccCr, invoiceDi.CardCode, incomeTaxAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);
                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საშემოსავლოს გატარება გამავალი გადახდა",
                            CreatedDocumentEntry = incomeTaxPayerTransId,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }
                if (isIncomeTaxPayer && invoiceDi.CancelStatus == CancelStatusEnum.csCancellation)
                {
                    try
                    {
                        string incomeTaxPayerTransId = DiManager.AddJournalEntry(Company, incomeTaxAccCr,
                            incomeTaxAccDr, incomeTaxControlAccCr, invoiceDi.CardCode, -incomeTaxAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);

                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საშემოსავლოს გატარება გამავალი გადახდა დაქენესელებული",
                            CreatedDocumentEntry = incomeTaxPayerTransId,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }

            }
            return results;
        }
        public IEnumerable<Result> PostPension(string invoiceDocentry, Company Company)
        {
            List<Result> results = new List<Result>();
            Documents invoiceDi = (Documents)Company.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
            invoiceDi.GetByKey(int.Parse(invoiceDocentry, CultureInfo.InvariantCulture));
            string bpCode = invoiceDi.CardCode;

            Recordset recSet = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);


            recSet.DoQuery(
                $"SELECT U_IncomeTaxPayer, U_PensionPayer FROM OCRD WHERE OCRD.CardCode = N'{bpCode}'");

            bool isPensionPayer = recSet.Fields.Item("U_PensionPayer").Value.ToString() == "01";

            recSet.DoQuery($"Select * From [@RSM_SERVICE_PARAMS]");
            string pensionAccDr = recSet.Fields.Item("U_PensionAccDr").Value.ToString();
            string pensionAccCr = recSet.Fields.Item("U_PensionAccCr").Value.ToString();
            string pensionControlAccDr = recSet.Fields.Item("U_PensionControlAccDr").Value.ToString();
            string pensionControlAccCr = recSet.Fields.Item("U_PensionControlAccCr").Value.ToString();

            SAPbobsCOM.BusinessPartners bp =
                (SAPbobsCOM.BusinessPartners)Company.GetBusinessObject(BoObjectTypes.oBusinessPartners);
            bp.GetByKey(invoiceDi.CardCode);

            var incomeTaxPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_IncomeTaxPayerPercent").Value.ToString(),
                CultureInfo.InstalledUICulture);

            var pensionPayerPercent = double.Parse(bp.UserFields.Fields.Item("U_PensionPayerPercent").Value.ToString());

            for (int i = 0; i < invoiceDi.Lines.Count; i++)
            {
                invoiceDi.Lines.SetCurrentLine(i);
                recSet.DoQuery(
                    $"SELECT U_PensionLiable FROM OITM WHERE OITM.ItemCode = N'{invoiceDi.Lines.ItemCode}'");
                bool isPensionLiable = recSet.Fields.Item("U_PensionLiable").Value.ToString() == "01";
                if (!isPensionLiable)
                {
                    continue;
                }

                double lineTotal = invoiceDi.Lines.LineTotal;
                double pensionAmount = Math.Round(lineTotal * pensionPayerPercent / 100, 6);

                if (isPensionPayer)
                {
                    //invoiceDi.CancelStatus == CancelStatusEnum.csNo
                    try
                    {
                        string incometaxpayertransidcomp = DiManager.AddJournalEntry(Company,
                            pensionAccCr, pensionAccDr, pensionControlAccCr, pensionControlAccDr, pensionAmount,
                            invoiceDi.Series, invoiceDi.Comments, invoiceDi.DocDate,
                            invoiceDi.BPL_IDAssignedToInvoice, invoiceDi.DocCurrency);
                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საპენსიოს გატარება ინვოისი",
                            CreatedDocumentEntry = incometaxpayertransidcomp,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {
                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }

                    try
                    {
                        string incometaxpayertransid = DiManager.AddJournalEntry(Company, pensionAccCr,
                            "", pensionControlAccCr, invoiceDi.CardCode, pensionAmount, invoiceDi.Series,
                            invoiceDi.Comments, invoiceDi.DocDate, invoiceDi.BPL_IDAssignedToInvoice,
                            invoiceDi.DocCurrency);
                        results.Add(new Result
                        {
                            IsSuccessCode = true,
                            StatusDescription = "საპენსიოს გატარება ინვოისი",
                            CreatedDocumentEntry = incometaxpayertransid,
                            ObjectType = BoObjectTypes.oJournalEntries
                        });
                    }
                    catch (Exception e)
                    {

                        results.Add(new Result { IsSuccessCode = false, StatusDescription = e.Message });
                        return results;
                    }
                }
                else
                {
                    results.Add(new Result { IsSuccessCode = false, StatusDescription = "არ არის საპენსიოს გადამხდელი" });
                    return results;
                }
            }
            return results;
        }

    }
}

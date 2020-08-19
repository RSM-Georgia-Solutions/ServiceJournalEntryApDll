using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SAPbobsCOM;
using ServiceJournalEntryApDll;

namespace Tests
{
    [TestClass]
    public class DocumentHelperTests
    {
        private IDocumentHelper _documentHelper;
        private Company _company;
        public DocumentHelperTests()
        {
            _documentHelper = new DocumentHelper();
            _company = new CompanyClass
            {
                Server = "SRV-TAMARI",
                DbServerType = BoDataServerTypes.dst_MSSQL2016,
                UserName = "manager",
                Password = "123456",
                CompanyDB = "KTWRealAnalog",
                language = BoSuppLangs.ln_English
            };
            _company.Connect();
            if (!_company.Connected)
            {
                throw new Exception($"Cannot Connect To the Server : {_company.GetLastErrorDescription()} : " +
                                    $"Server : {_company.Server}, " +
                                    $"DbServerType : {_company.DbServerType}," +
                                    $"UserName : {_company.UserName}," +
                                    $"CompanyDB : {_company.CompanyDB}");
            }

        }

        [TestMethod]
        public void PostIncomeTaxFromCreditMemo_TakesId_ReturnsNotEmptyResult()
        {
            _company.StartTransaction();
           var res =  _documentHelper.PostIncomeTaxFromCreditMemo("2", _company);
           var message = res.FirstOrDefault()?.StatusDescription;
           Assert.AreNotEqual(0,res.Count());
           _company.EndTransaction(BoWfTransOpt.wf_RollBack);

        }
        [TestMethod]
        public void PostIncomeTaxFromInvoice_TakesId_ReturnsNotEmptyResult()
        {
            _company.StartTransaction();
            var res = _documentHelper.PostIncomeTaxFromInvoice("14097", _company);
            var message = res.FirstOrDefault()?.StatusDescription;
            Assert.AreNotEqual(0, res.Count());
            _company.EndTransaction(BoWfTransOpt.wf_RollBack);

        }
        [TestMethod]
        public void PostIncomeTaxFromOutgoing_TakesId_ReturnsNotEmptyResult()
        {
            _company.StartTransaction();
            var res = _documentHelper.PostIncomeTaxFromOutgoing("2", _company);
            var message = res.FirstOrDefault()?.StatusDescription;
            Assert.AreNotEqual(0, res.Count());
            _company.EndTransaction(BoWfTransOpt.wf_RollBack);

        }
        [TestMethod]
        public void PostPension_TakesId_ReturnsNotEmptyResult()
        {
            _company.StartTransaction();
            var res = _documentHelper.PostPension("2", _company);
            var message = res.FirstOrDefault()?.StatusDescription;
           var aa = res.Count();
            Assert.AreNotEqual(0, aa);
            _company.EndTransaction(BoWfTransOpt.wf_RollBack);

        }
    }
}

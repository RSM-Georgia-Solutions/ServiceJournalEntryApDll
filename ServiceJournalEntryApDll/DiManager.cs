using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Translator;

namespace ServiceJournalEntryApDll
{
    class DiManager
    {
        public static Recordset Recordset => recSet.Value;
        public static Company Company => xCompany.Value;
        public static bool IsHana => IsHanax.Value;

        private static readonly Lazy<bool> IsHanax =
            new Lazy<bool>(() => Company.DbServerType.ToString() == "dst_HANADB" ? true : false);

        private static readonly Lazy<Company> xCompany =
            new Lazy<Company>(() => (Company)SAPbouiCOM.Framework
                .Application
                .SBO_Application
                .Company.GetDICompany());

        private static readonly Lazy<Recordset> recSet =
            new Lazy<SAPbobsCOM.Recordset>(() => (Recordset)
                Company
                    .GetBusinessObject(BoObjectTypes.BoRecordset));
        public static string QueryHanaTransalte(string query)
        {
            if (IsHana)
            {
                int numOfStatements;
                int numOfErrors;
                TranslatorTool TranslateTool = new TranslatorTool();
                query = TranslateTool.TranslateQuery(query, out numOfStatements, out numOfErrors);
                return query;
            }
            else
            {
                return query;
            }
        }

        public static string AddJournalEntry(Company _comp, string creditCode, string debitCode, string creditControlCode, string debitControlCode, double amount, int series, string comment, DateTime DocDate, int BPLID = 235,
            string currency = "GEL")
        {

            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = comment.Length < 50 ? comment : comment.Substring(0, 49);

            vJE.Lines.BPLID = BPLID; //branch

            if (currency == "GEL")
            {
                vJE.Lines.Debit = amount;
                vJE.Lines.FCDebit = 0;
            }
            else
            {
                vJE.Lines.FCCurrency = currency;
                vJE.Lines.FCDebit = amount;
            }

            vJE.Lines.Credit = 0;
            vJE.Lines.FCCredit = 0;

            if (string.IsNullOrWhiteSpace(debitCode))
            {
                vJE.Lines.ShortName = debitControlCode;
            }
            else
            {
                vJE.Lines.AccountCode = debitCode;
            }
            vJE.Lines.Add();


            vJE.Lines.BPLID = BPLID;
            if (string.IsNullOrWhiteSpace(creditCode))
            {
                vJE.Lines.ShortName = creditControlCode;
            }
            else
            {
                vJE.Lines.AccountCode = creditCode;
            }
            vJE.Lines.Debit = 0;
            vJE.Lines.FCDebit = 0;
            if (currency == "GEL")
            {
                vJE.Lines.Credit = amount;
                vJE.Lines.FCCredit = 0;
            }
            else
            {
                vJE.Lines.FCCurrency = currency;
                vJE.Lines.FCCredit = amount;
            }

            vJE.Lines.Add();

            int i = vJE.Add();
            if (i == 0)
            {
                string transId = _comp.GetNewObjectKey();
                return transId;
            }
            else
            {
                throw new Exception(_comp.GetLastErrorDescription());
            }
        }

        public bool CreateTable(string tableName, BoUTBTableType TableType)
        {
            GC.Collect();
            UserTablesMD oUTables;
            try
            {
                oUTables = (UserTablesMD)Company.GetBusinessObject(BoObjectTypes.oUserTables);

                if (oUTables.GetByKey(tableName) == false)
                {
                    GC.Collect();
                    oUTables.TableName = tableName;
                    oUTables.TableDescription = tableName;
                    oUTables.TableType = TableType;
                    int ret = oUTables.Add();

                    if (ret == 0)
                    {
                        //_application.StatusBar.SetText("UDT created:" + oUTables.TableName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUTables);
                        GC.Collect();
                        return true;
                    }
                    else
                    {
                        Application.SBO_Application.StatusBar.SetText("UDT failed: " + Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUTables);
                        GC.Collect();
                        return false;
                    }
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUTables);
                    GC.Collect();
                    return true;
                }
            }
            catch (Exception e)
            {
                //  _application.StatusBar.SetText("exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                GC.Collect();
            }

            return false;
        }

        public bool AddField(string tablename, string fieldname, string description, BoFieldTypes type, int size, IDictionary<string, string> validValues,
            bool isMandatory, bool isSapTable = false, string likedToTAble = "")
        {
            UserFieldsMD oUfield = (UserFieldsMD)Company.GetBusinessObject(BoObjectTypes.oUserFields);
            var recordset = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(QueryHanaTransalte($"SELECT * FROM CUFD WHERE AliasID = '{fieldname}' AND TableID = '@{tablename}'"));
            if (!recordset.EoF)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                GC.Collect();
                return true;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
            GC.Collect();
            try
            {
                if (isSapTable)
                {
                    oUfield.TableName = tablename;
                }
                else
                {
                    oUfield.TableName = "@" + tablename;
                }



                foreach (var validValue in validValues)
                {
                    oUfield.ValidValues.Value = validValue.Key;
                    oUfield.ValidValues.Description = validValue.Value;
                    oUfield.ValidValues.Add();
                }

                oUfield.DefaultValue = validValues.First().Key;



                oUfield.Name = fieldname;
                oUfield.Description = description;
                oUfield.Type = type;
                oUfield.Mandatory = isMandatory ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;

                if (type == BoFieldTypes.db_Float)
                {
                    oUfield.SubType = BoFldSubTypes.st_Price;
                }

                if (type == BoFieldTypes.db_Alpha || type == BoFieldTypes.db_Numeric)
                {
                    oUfield.EditSize = size;
                }
                oUfield.LinkedTable = likedToTAble;

                int ret = oUfield.Add();
                if (ret == 0 || ret == -2035)
                {
                    var x = Company.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);

                    GC.Collect();
                    return true;
                }
                else
                {
                    var x = Company.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);

                    GC.Collect();
                    return false;
                }

            }
            catch (Exception)
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);
                GC.Collect();
                return false;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);
                GC.Collect();
            }

        }

        public bool AddField(string tablename, string fieldname, string description, BoFieldTypes type, int size, bool isMandatory, bool isSapTable = false, string likedToTAble = "")
        {
            UserFieldsMD oUfield = (UserFieldsMD)Company.GetBusinessObject(BoObjectTypes.oUserFields);
            var recordset = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recordset.DoQuery(QueryHanaTransalte($"SELECT * FROM CUFD WHERE AliasID = '{fieldname}' AND TableID = '@{tablename}'"));
            if (!recordset.EoF)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
                GC.Collect();
                return true;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(recordset);
            GC.Collect();
            try
            {
                if (isSapTable)
                {
                    oUfield.TableName = tablename;
                }
                else
                {
                    oUfield.TableName = "@" + tablename;
                }

                oUfield.Name = fieldname;
                oUfield.Description = description;
                oUfield.Type = type;
                oUfield.Mandatory = isMandatory ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;

                if (type == BoFieldTypes.db_Float)
                {
                    oUfield.SubType = BoFldSubTypes.st_Price;
                }

                if (type == BoFieldTypes.db_Alpha || type == BoFieldTypes.db_Numeric)
                {
                    oUfield.EditSize = size;
                }
                oUfield.LinkedTable = likedToTAble;

                int ret = oUfield.Add();
                if (ret == 0 || ret == -2035)
                {
                    var x = Company.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);

                    GC.Collect();
                    return true;
                }
                else
                {
                    var x = Company.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);

                    GC.Collect();
                    return false;
                }

            }
            catch (Exception)
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);
                GC.Collect();
                return false;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUfield);
                GC.Collect();
            }

        }

        public bool AddKey(string tablename, string keyname, string fieldAlias, BoYesNoEnum IsUnique, string secondKeyAlias = "", string thirdKeyAlias = "")
        {
            int result;
            UserKeysMD oUkey = (UserKeysMD)Company.GetBusinessObject(BoObjectTypes.oUserKeys);
            try
            {

                oUkey.TableName = "@" + tablename;
                oUkey.KeyName = keyname;
                oUkey.Elements.ColumnAlias = fieldAlias;
                oUkey.Unique = IsUnique;
                oUkey.Elements.Add();
                if (secondKeyAlias != "")
                {
                    oUkey.Elements.ColumnAlias = secondKeyAlias;
                }
                if (thirdKeyAlias != "")
                {
                    oUkey.Elements.Add();
                    oUkey.Elements.ColumnAlias = thirdKeyAlias;
                }
                result = oUkey.Add();

                if (result == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUkey);
                    GC.Collect();
                    return true;
                }
                if (result == -1)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUkey);
                    GC.Collect();
                    return true;
                }
                else
                {
                    string str = Company.GetLastErrorDescription();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUkey);
                    GC.Collect();
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUkey);
                GC.Collect();
                return false;
            }
        }

        public static double GetCurrencyRate(string curCode, DateTime date, Company xCompany)
        {
            try
            {

                if (GetLocalCurrencyCode(xCompany) == GetSystemCurrencyCode(xCompany) && curCode == GetSystemCurrencyCode(xCompany) || curCode == GetLocalCurrencyCode(xCompany))
                {
                    return 1.0;
                }
                Company oCompany = xCompany;
                SBObob oSbObob = (SBObob)oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
                Recordset oRecordSet = oSbObob.GetCurrencyRate(curCode, date.Date);
                return double.Parse(oRecordSet.Fields.Item(0).Value.ToString());
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message, 1, "OK");
            }
            return 0;
        }

        public static string GetLocalCurrencyCode(SAPbobsCOM.Company xCompany)
        {
            try
            {
                Company oCompany = xCompany;

                SBObob oSbObob = (SBObob)oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
                Recordset oRecordSet = oSbObob.GetLocalCurrency();
                return oRecordSet.Fields.Item(0).Value.ToString();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message, 1, "OK");
            }
            return "0";
        }

        public static string GetSystemCurrencyCode(Company xCompany)
        {
            try
            {
                Company oCompany = xCompany;

                SBObob oSbObob = (SBObob)oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
                Recordset oRecordSet = oSbObob.GetSystemCurrency();
                return oRecordSet.Fields.Item(0).Value.ToString();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message, 1, "OK");
            }
            return "0";
        }
    }
}


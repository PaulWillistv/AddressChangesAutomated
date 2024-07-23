using System;

namespace AddressChangesAutomated
{
    internal class Util
    {
        public static System.Data.DataSet LoadExcelIntoDataTable(string fn, bool hdr7, bool JustGetTop5Records, string SheetName)
        {
            System.Data.DataSet ReturnDataSet = new System.Data.DataSet();
            try
            {

                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                string strHeader7 = "";
                strHeader7 = (hdr7) ? "Yes" : "No";
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fn + ";Extended Properties=\"Excel 12.0;HDR=" + strHeader7 + ";IMEX=1\"");
                if (JustGetTop5Records)
                {
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select top 5 * from [" + SheetName + "$]", MyConnection);
                }
                else
                {
                    MyCommand = new System.Data.OleDb.OleDbDataAdapter("select top 1000 * from [" + SheetName + "$]", MyConnection);
                }
                MyCommand.TableMappings.Add("Table", "MainTable");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                ReturnDataSet = DtSet;
                MyConnection.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return ReturnDataSet;

        }

    }
}

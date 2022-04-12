using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            var dt = ReadExcelFile("Sheet1", @"c:\temp\ClientLimt_20220412.xlsx");

            var ls = new List<CreditLimitExcelData>();

            foreach (DataRow item in dt.Rows)
            {
                var clientAccount = Convert.ToString(item["Counterparty Ref _(INT)"].ToString());

                if (clientAccount != "")
                {
                    ls.Add(new CreditLimitExcelData
                    {

                        Counterparty_Ref = clientAccount
                        ,
                        Client_Buy_Limit_Ccy = Convert.ToString(item["NEW Client BUY Limit _- Ccy_1564 (PI01)"].ToString())
                        ,
                        Client_Buy_Limit_Amount = Convert.ToDouble(item["NEW Client BUY Limit _- Amount_1261 (PAM1)"].ToString())
                        ,
                        Client_Sell_Limit_Amount = Convert.ToDouble(item["NEW Client SELL Limit _- Amount_1261 (PAM2)"].ToString())
                        ,
                        Client_Sell_Limit_Ccy = Convert.ToString(item["NEW Client SELL Limit _- Ccy_1564 (PI02)"].ToString())
                        ,
                        Client_Credit_Limit_Amount = Convert.ToDouble(item["NEW Client Credit Limit _- Amount_1261 (PAM3)"].ToString())
                        ,
                        Client_Credit_Limit_Ccy = Convert.ToString(item["NEW Client Credit Limit _- Ccy_1564 (PI03)"].ToString())

                    });
                }

                 
            }


        }


        private static DataTable ReadExcelFile(string sheetName, string path)
        {

            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                string fileExtension = Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";

                    comm.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }

                }
            }
        }
    }


    public class CreditLimitExcelData
    {
        public String Counterparty_Ref { get; set; }
        public String Client_Buy_Limit_Ccy { get; set; }
        public double Client_Buy_Limit_Amount { get; set; }
        public String Client_Sell_Limit_Ccy { get; set; }
        public double Client_Sell_Limit_Amount { get; set; }
        public String Client_Credit_Limit_Ccy { get; set; }
        public double Client_Credit_Limit_Amount { get; set; }

        public String Record_Update_Result { get; set; }

        public override string ToString()
        {
            return $"{{{nameof(Counterparty_Ref)}={Counterparty_Ref}, {nameof(Client_Buy_Limit_Ccy)}={Client_Buy_Limit_Ccy}, {nameof(Client_Buy_Limit_Amount)}={Client_Buy_Limit_Amount.ToString()}, {nameof(Client_Sell_Limit_Ccy)}={Client_Sell_Limit_Ccy}, {nameof(Client_Sell_Limit_Amount)}={Client_Sell_Limit_Amount.ToString()}, {nameof(Client_Credit_Limit_Ccy)}={Client_Credit_Limit_Ccy}, {nameof(Client_Credit_Limit_Amount)}={Client_Credit_Limit_Amount.ToString()}, {nameof(Record_Update_Result)}={Record_Update_Result}}}";
        }
    }
}

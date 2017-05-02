using Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace PopulateTickerTable
{
    class Program
    {
        private static string connectionString = @"Data Source=ALEX-PC;Initial Catalog=Numeraxial; Integrated Security=True;Connect Timeout=15;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";

        static void Main(string[] args)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string createTableCommand = "CREATE TABLE Stock_List1 (Stock_ID int IDENTITY (1000001, 1) Primary Key, Real_Ticker nvarchar(100), Storage_Ticker nvarchar(100), Comp_Name nvarchar(255), Exchange nvarchar(100), Country nvarchar(100), Category_Name nvarchar(100), Category_Number smallint, Initialized bit, Error bit);";
                using (SqlCommand command = new SqlCommand(createTableCommand, connection))
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }

            string currentDir = Environment.CurrentDirectory;
            DirectoryInfo directory = new DirectoryInfo(currentDir);

            string excelFilePath = directory.FullName;
            excelFilePath = excelFilePath.Replace(@"bin\Debug", "Yahoo Ticker Symbols - Jan 2016.xlsx");

            Console.WriteLine(excelFilePath);
            //Console.ReadLine(); 

            using (FileStream stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))

            {
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                
                char[] separatingChar = { '\t' };

                for (int i = 0;  i < 27724; i++)
                {
                    excelReader.Read();
                    if (i >= 4)
                    {
                        string queryString = String.Empty;
                        string realTicker = String.Empty;
                        string storageTicker = String.Empty;
                        string name = String.Empty;
                        string exchange = String.Empty;
                        string country = String.Empty;
                        string cat = String.Empty;
                        string num = String.Empty;

                        IDataRecord data = excelReader;

                        if (data[0] != null)
                        {
                            realTicker = data[0].ToString();
                            storageTicker = "TN_" + (realTicker.Replace("-", "dd").Replace("_", "uu").Replace(".", "pp").Replace("@", "aa").Replace("^", "cc"));
                        }
                        if (data[1] != null) name = data[1].ToString();
                        if (data[2] != null) exchange = data[2].ToString();
                        if (data[3] != null) country = data[3].ToString();
                        if (data[4] != null) cat = data[4].ToString();
                        if (data[5] != null) num = data[5].ToString();

                        queryString = "INSERT INTO Stock_List1 (Real_Ticker, Storage_Ticker, Comp_Name, Exchange, Country, Category_Name, Category_Number,Initialized, Error) values (@realTicker, @storageTicker, @name, @ex, @country, @cat_name, @cat_num, 0, 0);";
                        //Console.WriteLine(queryString);

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            using (SqlCommand cmd = new SqlCommand(queryString, connection))
                            {

                                cmd.Parameters.Add("@realTicker", SqlDbType.NVarChar);
                                cmd.Parameters["@realTicker"].Value = realTicker;

                                cmd.Parameters.Add("@storageTicker", SqlDbType.NVarChar);
                                cmd.Parameters["@storageTicker"].Value = storageTicker;

                                cmd.Parameters.Add("@name", SqlDbType.NVarChar);
                                cmd.Parameters["@name"].Value = name;

                                cmd.Parameters.Add("@ex", SqlDbType.NVarChar);
                                cmd.Parameters["@ex"].Value = exchange;

                                cmd.Parameters.Add("@country", SqlDbType.NVarChar);
                                cmd.Parameters["@country"].Value = country;

                                cmd.Parameters.Add("@cat_name", SqlDbType.NVarChar);
                                cmd.Parameters["@cat_name"].Value = cat;

                                cmd.Parameters.Add("@cat_num", SqlDbType.SmallInt);
                                cmd.Parameters["@cat_num"].Value = num;

                                connection.Open();
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                Console.WriteLine("Success!");
                Console.ReadLine();
            }
        }
    }
}

       
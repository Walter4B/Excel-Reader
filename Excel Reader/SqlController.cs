using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Data.SqlClient;


namespace Excel_Reader
{
    internal class SqlController
    {
        string connectionStringDatabase = ConfigurationManager.ConnectionStrings["ConnectionStringDatabase"].ConnectionString;

        ExcelManager excelManager = new ExcelManager();
        TableVisualisationEngine tableVisualisationEngine = new TableVisualisationEngine();
        OutputController outputController = new OutputController();

        internal void DatabaseToDisplayEngine()
        {
            outputController.DisplayMessage("SqlToConsole");
            using (SqlConnection connection = new SqlConnection(connectionStringDatabase))
            {
                using (SqlCommand command = connection.CreateCommand())
                {
                    string commandText = "SELECT * FROM ExcelTable";
                    command.CommandText = commandText;
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        var tableData = new List<List<Object>> { };

                        while (reader.Read())
                        {
                            tableData.Add(new List<object> {reader.GetInt32(0), reader.GetString(1), reader.GetString(2), reader.GetString(3),
                                                            reader.GetString(4), reader.GetInt32(5), reader.GetDecimal(6), reader.GetDecimal(7)});
                        }
                        tableVisualisationEngine.ConsoleDisplayData(tableData);
                    }
                }
            }
        }

        internal void ExcelToDatabase()
        {
            outputController.DisplayMessage("ImporttoDB");

             IList<Models.Acquisition> excelDataList = excelManager.GetExcelData();

            using (SqlConnection connection = new SqlConnection(connectionStringDatabase))
            {
                using (SqlCommand command = connection.CreateCommand())
                {
                    connection.Open();
                    for (int i = 0; i < excelDataList.Count; i++)
                    {
                        string commandText = $@"INSERT INTO ExcelTable (OrderDate, Region, Rep, Item, Units, UnitCost, Total) 
                                            VALUES ('{excelDataList[i].OrderDate.ToString()}','{excelDataList[i].Region}','{excelDataList[i].Rep}','{excelDataList[i].Item}',
                                                    '{excelDataList[i].Units}','{excelDataList[i].UnitCost}','{excelDataList[i].Total}')";
                        command.CommandText = commandText;
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}


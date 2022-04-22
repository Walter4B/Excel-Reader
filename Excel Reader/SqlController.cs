using System.Configuration;
using System.Data;
using System.Data.SqlClient;


namespace Excel_Reader
{
    internal class SqlController
    {
        string connectionStringDatabase = ConfigurationManager.ConnectionStrings["ConnectionStringDatabase"].ConnectionString;

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
    }
}

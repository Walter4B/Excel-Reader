using System.Configuration;
using System.Data.SqlClient;

namespace Excel_Reader
{
    internal class DatabaseManager
    {
        string connectionStringServer = ConfigurationManager.ConnectionStrings["ConnectionStringServer"].ConnectionString;
        string connectionStringDatabase = ConfigurationManager.ConnectionStrings["ConnectionStringDatabase"].ConnectionString;
        OutputController outputController = new OutputController();

        internal void CreateDatabase()
        {
            DeleteDatabase();
            outputController.DisplayMessage("CreateDatabase");
            using (SqlConnection connection = new SqlConnection(connectionStringServer))
            {
                using (SqlCommand command = connection.CreateCommand())
                {
                    string commandText = $@"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'ExcelDatabase')
                                            BEGIN
                                                CREATE DATABASE ExcelDatabase;
                                            END;
                                         ";
                      
                    command.CommandText = commandText;
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
            CreateTable();
        }

        internal void CreateTable()
        {
            outputController.DisplayMessage("CeateTable");
            using (SqlConnection connection = new SqlConnection(connectionStringDatabase))
            {
                using (SqlCommand command = connection.CreateCommand())
                {
                    string commandText =
                        $@"IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'ExcelTable')
                            CREATE TABLE ExcelTable (
                            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
                            OrderDate VARCHAR (30) NOT NULL,
                            Region VARCHAR (10) NOT NULL,
                            Rep	VARCHAR (30) NOT NULL,
                            Item VARCHAR (30) NOT NULL,
                            Units INT NOT NULL,
                            UnitCost DECIMAL NOT NULL,
                            Total DECIMAL NOT NULL
                            );
                        ";

                    command.CommandText = commandText;
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }

        internal void DeleteDatabase() //TODO fix so it works when there is no database
        {
            outputController.DisplayMessage("DeleteDatabase");
            using (SqlConnection connection = new SqlConnection(connectionStringServer))
            {
                using (SqlCommand command = connection.CreateCommand())
                {
                    string commandText = $@"IF EXISTS (SELECT * FROM sys.databases WHERE name = 'ExcelDatabase') 
                                            BEGIN
                                            USE MASTER 
                                            ALTER DATABASE [ExcelDatabase] SET single_user WITH ROLLBACK IMMEDIATE 
                                            DROP DATABASE [ExcelDatabase] 
                                            END";
                                          
                    command.CommandText = commandText;
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
//USE MASTER ALTER DATABASE [ExcelDatabase] SET single_user WITH ROLLBACK IMMEDIATE
//DROP DATABASE[ExcelDatabase]
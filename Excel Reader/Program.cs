
namespace Excel_Reader
{
    class ExcelReaderMain
    {
        public static void Main()
        {
            DatabaseManager databaseManager = new DatabaseManager();
            SqlController sqlController = new SqlController();
            ExcelManager excelManager = new ExcelManager();

            databaseManager.CreateDatabase();
            excelManager.GetExcelData(); //TODO
            //sqlController.DatabaseToDisplayEngine(); //needs excelManager to work
        }
    }
}
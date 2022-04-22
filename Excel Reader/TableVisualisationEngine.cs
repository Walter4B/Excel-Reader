using ConsoleTableExt;

namespace Excel_Reader
{
    internal class TableVisualisationEngine
    {
        internal void ConsoleDisplayData(List<List<Object>> ListOfTableRows)
        {
            ConsoleTableBuilder
                .From(ListOfTableRows)
                .WithColumn("Id", "Order Date", "Region", "Rep", "Item", "Units", "Unit Cost", "Total")
                .WithFormat(ConsoleTableBuilderFormat.Alternative)
                .ExportAndWriteLine();
        }
    }
}

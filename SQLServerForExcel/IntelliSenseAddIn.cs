using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace SQLServerForExcel
{
    public class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
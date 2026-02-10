using System.Windows;
using OfficeOpenXml;

namespace ProductionCalc.Desktop
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Set the EPPlus license for noncommercial use
            ExcelPackage.License.SetNonCommercialPersonal("PREYE");

            base.OnStartup(e);
        }
    }
}

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace ChangeLens
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // ここで DiffDialog や DiffProcessor を呼び出す
            DiffDialog dialog = new DiffDialog
            {
                RevitApp = commandData.Application
            };
            dialog.Show();


            return Result.Succeeded;
        }
    }
}

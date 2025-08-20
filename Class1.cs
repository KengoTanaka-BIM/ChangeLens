using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MepQuantifier
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            TaskDialog.Show("MepQuantifier", "MepQuantifier が起動しました");

            // 出力対象を選択
            DialogResult dr = MessageBox.Show(
                "選択要素のみを出力しますか？\n「いいえ」でモデル全体を出力",
                "出力範囲の選択",
                MessageBoxButtons.YesNoCancel
            );

            if (dr == DialogResult.Cancel) return Result.Cancelled;

            List<Element> targetElements = new List<Element>();

            if (dr == DialogResult.Yes)
            {
                var selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("MepQuantifier", "要素が選択されていません。");
                    return Result.Cancelled;
                }
                foreach (var id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (IsMepFamily(elem))
                        targetElements.Add(elem);
                }
            }
            else
            {
                var collector = new FilteredElementCollector(doc)
                    .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_MechanicalEquipment))
                    .UnionWith(new FilteredElementCollector(doc)
                        .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves)))
                    .UnionWith(new FilteredElementCollector(doc)
                        .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_PipeCurves)))
                    .UnionWith(new FilteredElementCollector(doc)
                        .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_PlumbingFixtures)))
                    .UnionWith(new FilteredElementCollector(doc)
                        .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_ElectricalEquipment)))
                    .ToElements();

                foreach (var e in collector)
                {
                    if (IsMepFamily(e))
                        targetElements.Add(e);
                }
            }

            if (targetElements.Count == 0)
            {
                TaskDialog.Show("MepQuantifier", "対象となるMEP要素が存在しません。");
                return Result.Cancelled;
            }

            // Excel 出力準備
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktop, "MepQuantifier_Export.csv");

            var sb = new StringBuilder();
            sb.AppendLine("Level,Category,ElementId,Name,Length/Size,SystemName");

            foreach (var elem in targetElements)
            {
                string levelName = GetLevelName(elem);
                string categoryName = elem.Category != null ? elem.Category.Name : "";
                string elemId = elem.Id.ToString();
                string name = elem.Name;
                string lengthOrSize = GetLengthOrSize(elem);
                string systemName = GetParameterValue(elem, BuiltInParameter.RBS_SYSTEM_NAME_PARAM);

                sb.AppendLine($"{levelName},{categoryName},{elemId},{name},{lengthOrSize},{systemName}");
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            TaskDialog.Show("MepQuantifier", $"Excel (CSV) 出力完了:\n{filePath}");

            return Result.Succeeded;
        }

        private bool IsMepFamily(Element e)
        {
            if (e == null || e.Category == null) return false;

            BuiltInCategory[] mepCategories = new BuiltInCategory[]
            {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_ElectricalEquipment
            };

            // ElementId と BuiltInCategory の安全比較
            foreach (var cat in mepCategories)
            {
                ElementId catId = new ElementId((int)cat);
                if (e.Category.Id == catId)
                    return true;
            }

            return false;
        }

        private string GetLevelName(Element elem)
        {
            Parameter param = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM);
            if (param != null && param.AsElementId() != ElementId.InvalidElementId)
            {
                Element lvl = elem.Document.GetElement(param.AsElementId());
                return lvl != null ? lvl.Name : "";
            }
            return "";
        }

        private string GetParameterValue(Element e, BuiltInParameter paramId)
        {
            Parameter p = e.get_Parameter(paramId);
            if (p != null)
            {
                switch (p.StorageType)
                {
                    case StorageType.String: return p.AsString();
                    case StorageType.Double: return p.AsDouble().ToString("F2");
                    case StorageType.Integer: return p.AsInteger().ToString();
                    case StorageType.ElementId: return p.AsElementId().ToString();
                }
            }
            return "";
        }

        private string GetLengthOrSize(Element e)
        {
            if (e is Pipe pipe)
                return pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble().ToString("F2") ?? "";
            else if (e is Duct duct)
                return duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble().ToString("F2") ?? "";
            else if (e is FamilyInstance fi)
            {
                var s = fi.Symbol;
                var volParam = s.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                if (volParam != null)
                    return volParam.AsDouble().ToString("F2");
            }
            return "";
        }
    }
}

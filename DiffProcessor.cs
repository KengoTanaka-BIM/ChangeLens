using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ChangeLens
{

    

    public static class DiffProcessor
    {
       
        // UI対応・進捗あり差分処理
        public static List<DiffResult> RunDiffWithProgress(
            Autodesk.Revit.DB.Document newDoc,
            string oldModelPath,
            List<BuiltInCategory> targetCategories,
            Autodesk.Revit.DB.Color addColor,
            Autodesk.Revit.DB.Color modColor,
            Autodesk.Revit.DB.Color paramColor,
            string excelPath,
            IProgress<int> progress)
        {
            var excelList = new List<DiffResult>();

            // ModelPath を Revit 内部形式に変換
            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(oldModelPath);

            // OpenOptions 作成
            OpenOptions openOptions = new OpenOptions();

            // oldDoc を開く
            Autodesk.Revit.DB.Document oldDoc = newDoc.Application.OpenDocumentFile(modelPath, openOptions);

            // 新旧要素取得
            var newElems = new FilteredElementCollector(newDoc)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && targetCategories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                .ToList();

            var oldElems = new FilteredElementCollector(oldDoc)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && targetCategories.Contains((BuiltInCategory)e.Category.Id.IntegerValue))
                .ToList();

            int total = newElems.Count;
            int count = 0;

            using (Transaction tx = new Transaction(newDoc, "Diff Highlight"))
            {
                tx.Start();

                foreach (var e in newElems)
                {
                    string status = "";
                    OverrideGraphicSettings ogs = new OverrideGraphicSettings();

                    // 位置とタイプ両方が一致する要素
                    var foundByPosition = oldElems.FirstOrDefault(oldEl =>
                        IsSameLocation(e, oldEl) &&
                        oldEl.GetTypeId().Value == e.GetTypeId().Value
                    );

                    // タイプだけ一致する要素（位置違い）
                    var foundByType = oldElems.FirstOrDefault(oldEl =>
                    oldEl.GetTypeId().Value == e.GetTypeId().Value &&
                    !IsSameLocation(e, oldEl)  // 位置が違う場合のみ
                    );


                    if (foundByPosition != null)
                    {
                        if (IsParamChangedLimited(e, foundByPosition))
                        {
                            // 重要パラメータ変更 → オレンジ
                            SetElementColor(e, ogs, paramColor, newDoc);
                            status = "ParamModified";
                        }
                        else
                        {
                            // 同じ位置・同じパラメータ → 何もしない
                            continue;
                        }
                    }
                    else
                    {
                        // 追加 → 赤
                        SetElementColor(e, ogs, addColor, newDoc);
                        status = "Added";
                    }

                    newDoc.ActiveView.SetElementOverrides(e.Id, ogs);

                    excelList.Add(new DiffResult
                    {
                        Id = (int)e.Id.Value,
                        Category = e.Category?.Name ?? "None",
                        Name = e.Name,
                        Status = status
                    });

                    count++;
                    progress?.Report((int)(count * 100.0 / total));
                }


                tx.Commit();
            }

            // 削除要素はExcelに記録のみ
            foreach (var oldEl in oldElems)
            {
                bool existsInNew = newElems.Any(newEl =>
                    newEl.GetTypeId().Value == oldEl.GetTypeId().Value &&
                    IsSameLocation(newEl, oldEl)
                );

                if (!existsInNew)
                {
                    excelList.Add(new DiffResult
                    {
                        Id = (int)oldEl.Id.Value,
                        Category = oldEl.Category?.Name ?? "None",
                        Name = oldEl.Name,
                        Status = "Deleted"
                    });
                }
            }

            oldDoc.Close(false);

            ExportToExcel(excelList, excelPath);

            return excelList;
        }

        private static void SetElementColor(Autodesk.Revit.DB.Element e, OverrideGraphicSettings ogs, Autodesk.Revit.DB.Color color, Autodesk.Revit.DB.Document doc)
        {
            ogs.SetProjectionLineColor(color);
            var fillId = GetSolidFillPatternId(doc);
            if (fillId != ElementId.InvalidElementId)
            {
                ogs.SetSurfaceForegroundPatternColor(color);
                ogs.SetSurfaceForegroundPatternId(fillId);
            }

        }

        private static bool IsParamChangedLimited(Element eNew, Element eOld)
        {
            string[] importantParams = { "Diameter", "直径", "Height", "高さ", "Width", "幅", "System Type", "システムタイプ" };

            bool hasImportant = false;

            foreach (Autodesk.Revit.DB.Parameter pNew in eNew.Parameters)
            {
                string name = pNew.Definition.Name;
                if (!importantParams.Contains(name)) continue;
                hasImportant = true;
                Autodesk.Revit.DB.Parameter pOld = eOld.LookupParameter(name);
                if (pOld != null && !ParameterEquals(pNew, pOld))
                    return true;
            }

            if (!hasImportant) return false; // 重要パラメータがなければ false にする
            return false;
        }

        private static bool ParameterEquals(Autodesk.Revit.DB.Parameter p1, Autodesk.Revit.DB.Parameter p2)
        {
            switch (p1.StorageType)
            {
                case StorageType.Double:
                    double val1_mm = p1.AsDouble() * 304.8; // フィート → mm
                    double val2_mm = p2.AsDouble() * 304.8;
                    return Math.Abs(val1_mm - val2_mm) < 1.0; // 1mm以内は同じ
                case StorageType.Integer:
                    return p1.AsInteger() == p2.AsInteger();
                case StorageType.String:
                    return (p1.AsString() ?? "") == (p2.AsString() ?? "");
                case StorageType.ElementId:
                    return p1.AsElementId().IntegerValue == p2.AsElementId().IntegerValue;
                default:
                    return true;
            }
        }


        private static ElementId GetSolidFillPatternId(Autodesk.Revit.DB.Document doc)
        {
            var solidFill = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .FirstOrDefault(x => x.GetFillPattern().IsSolidFill);

            return solidFill != null ? solidFill.Id : ElementId.InvalidElementId;
        }


        private static bool IsSameLocation(Autodesk.Revit.DB.Element e1, Autodesk.Revit.DB.Element e2)
        {
            if (e1.Location is LocationPoint lp1 && e2.Location is LocationPoint lp2)
                return lp1.Point.IsAlmostEqualTo(lp2.Point, 0.0328);
            else if (e1.Location is LocationCurve lc1 && e2.Location is LocationCurve lc2)
                return lc1.Curve.GetEndPoint(0).IsAlmostEqualTo(lc2.Curve.GetEndPoint(0), 0.0328) &&
　　　　　　　　　　　 lc1.Curve.GetEndPoint(1).IsAlmostEqualTo(lc2.Curve.GetEndPoint(1), 0.0328);
            return false; // これで比較できない要素は「別物」と判定される

        }

        public class DiffResult
        {
            public int Id;
            public string Category;
            public string Name;
            public string Status;
        }

        // 従来の簡易呼び出し
        public static void RunDiff(Autodesk.Revit.DB.Document newDoc, string oldModelPath)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelPath = System.IO.Path.Combine(desktopPath, "DiffReport.xlsx");

            RunDiffWithProgress(
                newDoc,
                oldModelPath,
                new List<BuiltInCategory>
                {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_CableTray
                },
                new Autodesk.Revit.DB.Color(255, 0, 0),     // 赤
                new Autodesk.Revit.DB.Color(0, 0, 255),     // 青
                new Autodesk.Revit.DB.Color(255, 165, 0),   // オレンジ
                excelPath,
                null
            );
        }


        // OpenXML で Excel 書き出し
        private static void ExportToExcel(List<DiffResult> list, string path)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "DiffReport" };
                sheets.Append(sheet);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // ヘッダ行
                Row headerRow = new Row();
                headerRow.Append(
                    new Cell() { CellValue = new CellValue("Id"), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue("Category"), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue("Name"), DataType = CellValues.String },
                    new Cell() { CellValue = new CellValue("Status"), DataType = CellValues.String }
                );
                sheetData.Append(headerRow);

                // データ行
                foreach (var r in list)
                {
                    Row row = new Row();
                    row.Append(
                        new Cell() { CellValue = new CellValue(r.Id.ToString()), DataType = CellValues.Number },
                        new Cell() { CellValue = new CellValue(r.Category), DataType = CellValues.String },
                        new Cell() { CellValue = new CellValue(r.Name), DataType = CellValues.String },
                        new Cell() { CellValue = new CellValue(r.Status), DataType = CellValues.String }
                    );
                    sheetData.Append(row);
                }

                workbookPart.Workbook.Save();
            }
        }
    }

    public static class XYZExtensions
    {
        public static bool IsAlmostEqualTo(this XYZ p1, XYZ p2, double tol)
        {
            return p1.DistanceTo(p2) < tol;
        }
    }
}

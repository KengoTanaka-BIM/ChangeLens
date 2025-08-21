using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ChangeLens
{
    public class DiffHandler : IExternalEventHandler
    {
        public Document Doc { get; set; }
        public string OldModelPath { get; set; }
        public IProgress<int> Progress { get; set; }
        public bool ResetColors { get; set; } = false;


        public void Execute(UIApplication app)
        {
            // 色リセットフラグが立っていたら、差分処理は飛ばしてリセットだけ
            if (ResetColors)
            {
                if (Doc == null) return;

                using (Transaction tx = new Transaction(Doc, "リセット:色付け解除"))
                {
                    tx.Start();

                    var allElems = new FilteredElementCollector(Doc)
                        .WhereElementIsNotElementType()
                        .ToElements();

                    foreach (var elem in allElems)
                    {
                        Doc.ActiveView.SetElementOverrides(elem.Id, new OverrideGraphicSettings());
                    }

                    tx.Commit();
                }

                ResetColors = false; // フラグを元に戻す
                TaskDialog.Show("ChangeLens", "ビュー上の色をリセットしました。");
                return; // ここで処理終了
            }

            // ↓既存の差分処理
            if (Doc == null || string.IsNullOrEmpty(OldModelPath)) return;

            DiffProcessor.RunDiffWithProgress(
                Doc,
                OldModelPath,
                new List<BuiltInCategory>
                {
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_CableTray
                },
                new Autodesk.Revit.DB.Color((byte)255, 0, 0),
                new Autodesk.Revit.DB.Color((byte)0, 0, 255),
                new Autodesk.Revit.DB.Color((byte)255, 165, 0),
                System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "DiffReport.xlsx"),
                Progress
            );
        }


        public string GetName() => "DiffHandler";
    }
}

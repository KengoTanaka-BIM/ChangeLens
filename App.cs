using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

namespace ChangeLens
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            // タブ作成（すでにあればスキップ）
            string tabName = "BIMCraftWorks";
            try
            {
                app.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // タブはすでにあるので無視
            }

            // パネル作成
            RibbonPanel panel = app.CreateRibbonPanel(tabName, "Difference detection");

            // ボタン作成
            PushButtonData buttonData = new PushButtonData(
                "DiffCommand",                     // 内部名
                "ChangeLens",                       // ボタン表示名
                Assembly.GetExecutingAssembly().Location, // このDLL
                "ChangeLens.Command"               // Command.csのフルクラス名
            );

            PushButton pushButton = panel.AddItem(buttonData) as PushButton;

            // アイコン設定（埋め込みリソース）
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string resourceName = "ChangeLens.Resources.icon.png"; // 名前空間.フォルダ.ファイル名
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        BitmapImage bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = stream;
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        pushButton.LargeImage = bitmap;
                    }
                    else
                    {
                        TaskDialog.Show("ChangeLens", "The embedded resource for the icon could not be found.");
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ChangeLens", "Failed to load icon: " + ex.Message);
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}

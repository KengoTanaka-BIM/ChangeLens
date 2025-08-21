using Autodesk.Revit.UI;
using System;
using System.Reflection;

namespace ChangeLens
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            // タブ作成（すでにあればスキップ）
            string tabName = "MyTools";
            try
            {
                app.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // タブはすでにあるので無視
            }

            // パネル作成
            RibbonPanel panel = app.CreateRibbonPanel(tabName, "Sheet Tools");

            // ボタン作成
            PushButtonData buttonData = new PushButtonData(
                "DiffCommand",                     // 内部名
                "Diff起動",                        // ボタン表示名
                Assembly.GetExecutingAssembly().Location, // このDLL
                "ChangeLens.Command"               // Command.csのフルクラス名
            );

            PushButton pushButton = panel.AddItem(buttonData) as PushButton;

            // アイコン設定（任意）
            // Uri uri = new Uri("pack://application:,,,/ChangeLens;component/Resources/icon.png");
            // pushButton.LargeImage = new System.Windows.Media.Imaging.BitmapImage(uri);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}

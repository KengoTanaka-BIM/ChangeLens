using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Windows;

namespace ChangeLens
{
    public partial class DiffDialog : Window
    {
        public UIApplication RevitApp { get; set; }

        private string oldModelPath;
        private DiffHandler handler;
        private ExternalEvent externalEvent;

        public DiffDialog()
        {
            InitializeComponent();

            handler = new DiffHandler();
            externalEvent = ExternalEvent.Create(handler);
        }

        // 古いモデル選択
        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Revit Model (*.rvt)|*.rvt",
                Title = "Please select an older model"
            };
            if (dlg.ShowDialog() == true)
            {
                oldModelPath = dlg.FileName;
                TxtExcelPath.Text = oldModelPath; // XAML 側の TextBox 名に合わせた
            }
        }

        // Diff開始
        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(oldModelPath))
            {
                MessageBox.Show("Please select an older model");
                return;
            }
            if (RevitApp == null)
            {
                MessageBox.Show("The Revit application is not set.");
                return;
            }

            handler.Doc = RevitApp.ActiveUIDocument.Document;
            handler.OldModelPath = oldModelPath;

            externalEvent.Raise();

        }



        // 色リセット
        private void BtnResetColor_Click(object sender, RoutedEventArgs e)
        {
            handler.ResetColors = true;
            externalEvent.Raise();
        }
    }
}

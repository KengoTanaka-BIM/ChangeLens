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

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Revit Model (*.rvt)|*.rvt",
                Title = "古いモデルを選択してください"
            };
            if (dlg.ShowDialog() == true)
            {
                oldModelPath = dlg.FileName;
                TxtExcelPath.Text = oldModelPath;
            }
        }

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(oldModelPath))
            {
                MessageBox.Show("古いモデルを選択してください。");
                return;
            }
            if (RevitApp == null)
            {
                MessageBox.Show("Revit アプリケーションがセットされていません。");
                return;
            }

            handler.Doc = RevitApp.ActiveUIDocument.Document;
            handler.OldModelPath = oldModelPath;
            handler.Progress = new System.Progress<int>(v => ProgressBar.Value = v);

            externalEvent.Raise();
        }
        private void BtnResetColor_Click(object sender, RoutedEventArgs e)
        {
            handler.ResetColors = true;
            externalEvent.Raise();

            
        }
        
        

    }
}

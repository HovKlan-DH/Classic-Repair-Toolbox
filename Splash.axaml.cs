using System;
using Avalonia.Controls;
using DataHandling;

namespace CRT
{
    public partial class Splash : Window
    {
        public Splash()
        {
            InitializeComponent();

            this.VersionNumberRun.Text = AppConfig.AppVersionString;

            DataManager.StatusChanged += this.OnStatusChanged;
            DataManager.FileDownloadChanged += this.OnFileDownloadChanged;
        }

        // ###########################################################################################
        // Updates the status label whenever DataManager reports a general progress change.
        // ###########################################################################################
        private void OnStatusChanged(string status)
        {
            this.StatusLabel.Text = status;
        }

        // ###########################################################################################
        // Updates the file label with the path of the file currently being downloaded.
        // Hides the label automatically when the download batch finishes (empty string).
        // ###########################################################################################
        private void OnFileDownloadChanged(string filePath)
        {
            this.FileLabel.Text = filePath;
            this.FileLabel.IsVisible = !string.IsNullOrEmpty(filePath);
        }

        protected override void OnClosed(EventArgs e)
        {
            DataManager.StatusChanged -= this.OnStatusChanged;
            DataManager.FileDownloadChanged -= this.OnFileDownloadChanged;
            base.OnClosed(e);
        }
    }
}
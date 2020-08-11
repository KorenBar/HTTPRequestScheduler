using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http;
using Utility;
using OfficeOpenXml;
using System.Reflection;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using KB.Processes;
using KB.Configuration;
using KB.Utility;

namespace HTTPRequestScheduler
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<string, object> WeekDays { get; set; }
            = new Dictionary<string, object>()
                { { DayOfWeek.Sunday.ToString(), DayOfWeek.Sunday },
                { DayOfWeek.Monday.ToString(), DayOfWeek.Monday },
                { DayOfWeek.Tuesday.ToString(), DayOfWeek.Tuesday },
                { DayOfWeek.Wednesday.ToString(), DayOfWeek.Wednesday },
                { DayOfWeek.Thursday.ToString(), DayOfWeek.Thursday },
                { DayOfWeek.Friday.ToString(), DayOfWeek.Friday },
                { DayOfWeek.Saturday.ToString(), DayOfWeek.Saturday } };

        public OpenFileDialog OpenFileDialog { get; private set; } 
            = new OpenFileDialog() { DefaultExt = ".xlsx", Filter = "Excel Workbook (*.xlsx)|*.xlsx" };

        public System.Windows.Forms.FolderBrowserDialog FolderBrowserDialog { get; private set; }
            = new System.Windows.Forms.FolderBrowserDialog() { ShowNewFolderButton = true };

        public Worker Worker { get; }

        #region Constructors
        bool loading = true;
        public MainWindow(Worker worker)
        {
            // First set the worker for binding to work.
            this.Worker = worker;
            InitializeComponent();

            Binding();

            // Worker Events
            this.Worker.SendStarted += Worker_SendStarted;
            this.Worker.SendFinished += Worker_SendFinished;
            this.Worker.Message += Worker_Message;
            this.Worker.Progress += Worker_Progress;
            this.Worker.ExcelPackageChanged += Worker_ExcelPackageChanged;

            this.Title = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location);
            loading = false;
        }

        private void Binding()
        {
            // RequestType
            if (this.requestTypeComboBox.Items.Contains(this.Worker.RequestType))
                this.requestTypeComboBox.SelectedItem = this.Worker.RequestType as object;
            else
            {
                this.customRequestTypeTextBox.Text = this.Worker.RequestType;
                this.requestTypeComboBox.SelectedItem = this.customRequestTypeTextBox;
            }
            this.requestTypeComboBox.SetBinding(ComboBox.TextProperty, "Worker.RequestType");

            // Days
            this.daysComboBox.ItemsSource = WeekDays;
            this.daysComboBox.Text = this.Worker.Days;

            // Time
            this.timePicker.Value = new DateTime() + this.Worker.Time;

            // Sheet
            Worker_ExcelPackageChanged(this, null);
            this.sheetComboBox.SetBinding(ComboBox.SelectedItemProperty, "Worker.SheetName");

            // RenameDownload
            RenameDownloadCheckBoxHeader_Click(this, null);
        }
        #endregion

        private void Save() => Ini.Default.SaveProperties(this.Worker, string.Empty);

        #region Worker Events
        private void Worker_ExcelPackageChanged(object sender, ExcelPackage e)
        {
            // Get sheets to ComboBox
            try
            {
                this.sheetComboBox.Items.Clear();

                if (this.Worker.ExcelPackage != null)
                    foreach (var s in this.Worker.ExcelPackage.Workbook.Worksheets.Select(s => s.Name))
                        this.sheetComboBox.Items.Add(s);

                if (this.sheetComboBox.Items.Count > 0)
                    this.sheetComboBox.SelectedIndex = 0;

                this.OpenFileDialog.InitialDirectory = Path.GetDirectoryName(Path.GetFullPath(this.excelFileTextBox.Text));
            }
            catch { }
        }

        private void Worker_Message(object sender, MessageEventArgs e)
        {
            Brush b;
            switch (e.Staus)
            {
                case MessageStaus.Success:
                    b = Brushes.Green;
                    break;
                case MessageStaus.Warning:
                    b = Brushes.Orange;
                    break;
                case MessageStaus.Error:
                    b = Brushes.Red;
                    break;
                case MessageStaus.Status:
                    b = Brushes.White;
                    break;
                default:
                    return; // Don't show anything else on UI
            }
            Dispatcher.Invoke(() =>
            {
                this.statusBar.Foreground = b;
                this.statusBar.Items[0] = e.Message + " | " + DateTime.Now.ToString();
            });
        }

        private void Worker_Progress(object sender, ProgressEventArgs e) => Dispatcher.Invoke(() => progressBar1.Value++);

        private void Worker_SendStarted(object sender, int e)
        {
            Dispatcher.Invoke(() =>
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = e;
                sendButton.Visibility = Visibility.Collapsed;
                progressBar1.Visibility = Visibility.Visible;
            });
        }

        private void Worker_SendFinished(object sender, int e)
        {
            Dispatcher.Invoke(() =>
            {
                progressBar1.Visibility = Visibility.Collapsed;
                sendButton.Visibility = Visibility.Visible;
            });
        }
        #endregion

        #region UI Events
        private async void sendButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show(this, "Do you want to send requests now?", "Request will be sent", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                Console.WriteLine("Sending required manually (by a button)");
                await this.Worker.Send();
            }
        }

        private void ValueChanged(object sender, object e)
        {
            if (!loading)
                saveButton.Visibility = Visibility.Visible;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Save();
            saveButton.Visibility = Visibility.Collapsed;
        }

        private void timePicker_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            DateTime? v = timePicker.Value;
            if (v != null) this.Worker.Time = ((DateTime)v).TimeOfDay;
            this.ValueChanged(sender, e);
        }

        private void daysComboBox_SelectedItemsChanged(object sender, EventArgs e)
        {
            this.Worker.Days = daysComboBox.Text;
            this.ValueChanged(sender, e);
        }

        private void triggerCheckBoxHeader_Click(object sender, RoutedEventArgs e) => this.ValueChanged(sender, e);

        private void excelFileTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
        }

        private void ExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.OpenFileDialog.ShowDialog(this) ?? false)
                this.excelFileTextBox.Text = this.OpenFileDialog.FileName;
        }

        private void StatusBar_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            contentPanel.Visibility = contentPanel.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        }

        private void DownloadFolderButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.FolderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                this.downloadFolderTextBox.Text = this.FolderBrowserDialog.SelectedPath;
        }

        private void DownloadFolderTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Directory.Exists(downloadFolderTextBox.Text))
                this.FolderBrowserDialog.SelectedPath = downloadFolderTextBox.Text;
            ValueChanged(sender, e);
        }

        private void RenameDownloadCheckBoxHeader_Click(object sender, RoutedEventArgs e)
        {
            renameDownloadTextBox.IsEnabled = renameDownloadCheckBoxHeader.IsChecked ?? false;
            ValueChanged(sender, e);
        }
        #endregion
    }
}

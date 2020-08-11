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
using System.Net.Mime;
using Utility;
using OfficeOpenXml;
using System.Reflection;
using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using KB.Processes;
using KB.Configuration;
using KB.Utility;
using KB.Threading;
using System.Threading;
using KB.Data;
using System.ComponentModel;
using System.Net;

namespace HTTPRequestScheduler
{
    public class Worker
    {

        #region Public Properties
        public bool LogToFile { get; set; } = true;

        internal ExcelPackage ExcelPackage => File.Exists(excelFile) ? new ExcelPackage(new FileInfo(excelFile)) : null;

        public InfoText HeadersInfo => (InfoText)Headers;

        public string Headers { get; set; } = string.Empty;

        public string RequestType { get; set; } = "GET";

        public string RequestUrl { get; set; } = string.Empty;

        public string Content { get; set; } = string.Empty;

        public string ContentType { get; set; } = "application/json";

        public int Delay { get; set; } = 0;

        private string excelFile = string.Empty;
        public string ExcelFile
        {
            get => excelFile;
            set
            {
                if (excelFile == value) return;
                excelFile = value;
                OnExcelPackageChanged(this.ExcelPackage);
            }
        }

        public string SheetName { get; set; } = string.Empty;

        private int firstRow = 1;
        public int FirstRow
        {
            get => firstRow;
            set => firstRow = Math.Max(value, 1);
        }

        public bool RecursiveInsert { get; set; } = true;

        private string days = "All";
        public string Days
        {
            get => days;
            set
            {
                days = value;
                this.SendScheduler.Revoke(); // Will be re-invoked
            }
        }

        internal DayOfWeek[] DaysOfWeek
        {
            get
            {
                var strDays = Days.Split(',').Select(s => s.Trim().ToLower());
                var ds = (DayOfWeek[])Enum.GetValues(typeof(DayOfWeek));
                if (strDays.Contains("all")) return ds;
                return ds.Where(d => strDays.Contains(d.ToString().ToLower())).ToArray();
            }
        }

        private TimeSpan time = new TimeSpan(4, 0, 0);
        public TimeSpan Time
        {
            get => time;
            set
            {
                time = value;
                this.SendScheduler.Revoke(); // Will be re-invoked
            }
        }

        public bool Trigger
        {
            get => SendScheduler.Enabled;
            set => SendScheduler.Enabled = value;
        }

        public bool Download { get; set; } = false;

        public string DownloadDirectory { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";

        public bool DownloadRename { get; set; } = false;

        public string DownloadFileName { get; set; } = string.Empty;
        #endregion

        private bool scheduling = false;
        private object schedulingLock = new object();
        private object sendLock = new object();

        private ActionScheduler SendScheduler { get; }

        private string assemblyName = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location);

        #region EventHandlers
        // Arg will be the number of requests we gonna send.
        public event EventHandler<int> SendStarted;
        public event EventHandler<int> SendFinished;
        public event EventHandler<ExcelPackage> ExcelPackageChanged;
        public event EventHandler<MessageEventArgs> Message;
        public event EventHandler<ProgressEventArgs> Progress;
        #endregion

        public Worker()
        {
            SendScheduler = new ActionScheduler(async () => await Send());
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        #region Actions
        /// <summary>
        /// Start endless scheduling
        /// </summary>
        public async void StartScheduling()
        {
            scheduling = true;
            // Only one scheduler can run at the same time
            if (!Monitor.TryEnter(schedulingLock)) return;

            try
            {
                var dow = DaysOfWeek;
                if (dow.Length > 0)
                {
                    DateTime nextTrigger = ActionScheduler.GetNextDateTime(dow, Time);
                    Console.WriteLine("Next Trigger at: " + nextTrigger.ToString());
                    await SendScheduler.Invoke(nextTrigger);
                }
                else await Task.Delay(1000);
            }
            catch (Exception ex)
            {
                OnMeassge(ex.Message, MessageStaus.Error);
            }
            finally
            {
                Monitor.Exit(schedulingLock);
                if (scheduling) this.StartScheduling();
            }
        }

        public void StopScheduling(bool revokeCurrentTriggers) => scheduling = false;

        public async Task Send()
        {
            if (!Monitor.TryEnter(sendLock))
            {
                ConsoleHelper.WriteLine("Send is locked and still in process, will not be invoked in parallel.", ConsoleColor.DarkYellow);
                return;
            }

            try
            {
                var ep = this.ExcelPackage;
                if (ep != null)
                {
                    ExcelWorkbook wb = ep.Workbook;
                    wb.Calculate();

                    ExcelWorksheet sheet = wb.Worksheets[this.SheetName];
                    if (sheet != null)
                    {
                        ExcelAddressBase dimension = sheet.Dimension;
                        if (dimension != null)
                        {
                            int totalRows = dimension.End.Row;
                            int firstRow = FirstRow;
                            int rowsToWorkWith = (totalRows - (firstRow - 1));
                            if (rowsToWorkWith > 0)
                            {
                                Console.WriteLine(rowsToWorkWith + " request rows.");
                                OnSendStarted(rowsToWorkWith);

                                int errors = 0;
                                bool recursiveInsert = this.RecursiveInsert;
                                HttpClient httpClient = new HttpClient();
                                for (int r = firstRow; r <= totalRows; r++)
                                {
                                    string url = RequestUrl.InsertValues(sheet, r, recursiveInsert);
                                    string method = RequestType.InsertValues(sheet, r, recursiveInsert);
                                    try
                                    {
                                        Uri uri = new Uri(url, UriKind.Absolute);
                                        string content = Content.InsertValues(sheet, r, recursiveInsert);

                                        Console.WriteLine("Sending " + method + " request " + (r - firstRow) + " to " + uri.AbsoluteUri + " with" + (string.IsNullOrEmpty(content) ? " no" : "") + " content");
                                        
                                        // Create request
                                        HttpRequestMessage msg = new HttpRequestMessage(new HttpMethod(method), uri);

                                        // Headers
                                        foreach(var h in ((InfoText)Headers.InsertValues(sheet, r, recursiveInsert)).ToDictionary())
                                            if (!string.IsNullOrEmpty(h.Key))
                                                msg.Headers.Add(h.Key, h.Value);

                                        // Content
                                        string cType = ContentType.InsertValues(sheet, r, recursiveInsert);
                                        if (msg.Method.Method != "GET")
                                            if (!string.IsNullOrEmpty(content))
                                                msg.Content = new StringContent(content, Encoding.UTF8, cType);

                                        // Send
                                        HttpResponseMessage httpResponse = await httpClient.SendAsync(msg);
                                        OnMeassge($"[{r}] {method} { url } => [{(int)httpResponse.StatusCode} {httpResponse.StatusCode.ToString()}] {httpResponse.Content.Headers.ContentType.MediaType}", MessageStaus.Information);

                                        // Download
                                        var resContent = httpResponse.Content;
                                        var resContentStream = await resContent.ReadAsStreamAsync();
                                        if (Download && resContentStream.Length > 0)
                                        {
                                            var dir = DownloadDirectory.InsertValues(sheet, r, recursiveInsert);
                                            Directory.CreateDirectory(dir);
                                            var resContentDisposition = resContent.Headers.ContentDisposition;
                                            // ?TODO: Convert response content type to file extension and add it to the renamed file.
                                            string filename = DownloadRename || string.IsNullOrEmpty(resContentDisposition.FileName)
                                                ? DownloadFileName.InsertValues(sheet, r, recursiveInsert)
                                                : resContentDisposition.FileName;

                                            // TODO: If it's a file, download as binary file just like now, otherwise decode to text file. (using utf-8 encoding)
                                            using (var fs = new FileStream(Path.Combine(dir, filename), FileMode.Create))
                                                await resContent.CopyToAsync(fs);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        errors++;
                                        string emsg = ex.Message + (ex.InnerException != null ? ("    " + ex.InnerException.Message) : "");
                                        emsg = Regex.Replace(emsg, @"\t|\n|\r", "    ");
                                        OnMeassge(DateTime.Now.ToString() + $" : [{r}] {method} {url} => Error: {emsg}", MessageStaus.Information);
                                    }

                                    int rowsSent = r - (firstRow - 1);
                                    int rowsLeft = rowsToWorkWith - rowsSent;
                                    OnMeassge($"success : {rowsSent - errors} | failure : {errors} | left : {rowsLeft} | rows : {rowsToWorkWith}", MessageStaus.Status);
                                    OnProgress(rowsSent, rowsToWorkWith);

                                    // Delay
                                    int delay = this.Delay;
                                    if (delay > 0) await Task.Delay(delay * 1000);
                                }
                                if (errors == 0)
                                    OnMeassge("Succeeded!", MessageStaus.Success);
                                else OnMeassge(errors + "/" + rowsToWorkWith + " failed.", MessageStaus.Error);
                            }
                            else OnMeassge("No rows", MessageStaus.Warning);
                        }
                        else OnMeassge("Sheet is empty", MessageStaus.Warning);
                    }
                    else OnMeassge("Sheet not found", MessageStaus.Warning);
                }
                else OnMeassge("No Workbook", MessageStaus.Warning);
            }
            catch (Exception ex)
            {
                OnMeassge("Error: " + ex.Message, MessageStaus.Error);
            }
            finally
            {
                Monitor.Exit(sendLock);
                OnSendFinished(0);
            }
        }
        #endregion

        #region Events
        protected virtual void OnSendStarted(int requests) => SendStarted?.Invoke(this, requests);
        protected virtual void OnSendFinished(int i) => SendFinished?.Invoke(this, i);
        protected virtual void OnExcelPackageChanged(ExcelPackage package) => ExcelPackageChanged?.Invoke(this, package);
        protected void OnMeassge(string msg, MessageStaus staus) => OnMeassge(new MessageEventArgs(msg, staus));
        protected virtual void OnMeassge(MessageEventArgs args)
        {
            if (LogToFile)
                switch (args.Staus)
                {
                    case MessageStaus.Error:
                    case MessageStaus.Warning:
                    case MessageStaus.Success:
                    case MessageStaus.Information:
                        File.AppendAllText($@".\{assemblyName}.log", 
                            DateTime.Now.ToString() + " : " + args.Message
                            + Environment.NewLine);
                        break;
                }

            ConsoleColor cc = Console.ForegroundColor;
            switch (args.Staus)
            {
                case MessageStaus.Progress:
                    cc = ConsoleColor.White;
                    break;
                case MessageStaus.Success:
                    cc = ConsoleColor.Green;
                    break;
                case MessageStaus.Warning:
                    cc = ConsoleColor.DarkYellow;
                    break;
                case MessageStaus.Error:
                    cc = ConsoleColor.Red;
                    break;
            }
            ConsoleHelper.WriteLine(args.Message, cc);
            Message?.Invoke(this, args);
        }
        protected void OnProgress(int rowsSent, int totalRows) => OnProgress(new ProgressEventArgs(rowsSent, totalRows));
        protected virtual void OnProgress(ProgressEventArgs args) => Progress?.Invoke(this, args);
        #endregion
    }
}

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;

using DevExpress.XtraRichEdit;

namespace DocumentServer_PrintToPDF {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window {
        PerformanceCounter counter;
        bool executing;

        public MainWindow() {
            InitializeComponent();

            tbPath.Text = Application.Current.StartupUri.AbsolutePath;

            string procName = Process.GetCurrentProcess().ProcessName;
            counter = new PerformanceCounter("Process", "Working Set - Private", procName);
            ShowMemoryUsage();
        }

        string PrintToPDF(RichEditDocumentServer server, string filePath) {
            try {
                server.LoadDocument(filePath);
            }
            catch (Exception ex) {
                server.CreateNewDocument();
                return String.Format("{0:T} Error:{1} -> {2}", DateTime.Now, ex.Message, filePath) + Environment.NewLine;
            }
            string outFileName = Path.ChangeExtension(filePath, "pdf");
            FileStream fsOut = File.Open(outFileName, FileMode.Create);
            server.ExportToPdf(fsOut);
            fsOut.Close();
            return String.Format("{0:T} Done-> {1}", DateTime.Now, outFileName) + Environment.NewLine;
        }

        void ConvertFiles(string path) {
            if (!Directory.Exists(path))
                return;

            string[] files = System.IO.Directory.GetFiles(path, "*.doc?", System.IO.SearchOption.AllDirectories);
            InitProgress(files.Length);
            using (RichEditDocumentServer server = new RichEditDocumentServer()) {
                int count = files.Length;
                for (int i = 0; i < count; i++) {
                    string progress = PrintToPDF(server, files[i]);
                    AppendProgress(progress, i + 1);
                    if (!executing)
                        return;
                }
            }
        }

        void btnConvert_Click(object sender, RoutedEventArgs e) {
            if (!executing)
                StartConversion();
            else
                FinishConversion();
        }

        void StartConversion() {
            executing = true;
            btnConvert.Content = "Stop!";
            tbPath.IsReadOnly = true;
            pnlProgress.Visibility = System.Windows.Visibility.Visible;
            Thread worker = new Thread(BackgroundWorker);
            worker.Start(tbPath.Text);
        }
        void FinishConversion() {
            pnlProgress.Visibility = System.Windows.Visibility.Collapsed;
            executing = false;
            tbPath.IsReadOnly = false;
            btnConvert.Content = "Start!";
        }
        void BackgroundWorker(object parameter) {
            string path = parameter as string;
            if (String.IsNullOrEmpty(path))
                return;
            ConvertFiles(path);
            ShowMemoryUsage();

            Action action = delegate() { FinishConversion(); };
            Dispatcher.BeginInvoke(action);
        }

        void InitProgress(int fileCount) {
            Action action = delegate() {
                edtProgress.Minimum = 0;
                edtProgress.Maximum = fileCount;
                edtProgress.EditValue = 0;
            };
            this.Dispatcher.Invoke(action);
        }
        void AppendProgress(string displayText, int fileIndex) {
            Action action = delegate() {
                tbLog.Text += displayText;
                LogScrollViewer.ScrollToVerticalOffset(LogScrollViewer.Height);
                edtProgress.EditValue = fileIndex;
            };
            this.Dispatcher.Invoke(action);
        }
        void ShowMemoryUsage() {
            Action action = delegate() { lblMemoryUsage.Text = String.Format("Memory usage: {0:N0} K", counter.RawValue / 1024); };
            this.Dispatcher.Invoke(action);
        }
        void Window_Closed(object sender, EventArgs e) {
            executing = false;
        }
    }
}

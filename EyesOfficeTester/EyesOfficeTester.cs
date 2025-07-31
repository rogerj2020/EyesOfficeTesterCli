using Applitools;
using Applitools.Images;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using IDataObject = System.Windows.Forms.IDataObject;

namespace EyesOfficeTesterLib
{
    public class EyesOfficeTester
    {
        readonly public static List<string> WordFileExtensions = new List<string> { ".docx", ".doc", ".docm", ".dotx", ".dotm" };
        readonly public static List<string> ExcelFileExtensions = new List<string> { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm" };
        public static EyesRunner EyesRunner = new ClassicRunner();
        public static BatchInfo BatchInfo;
        static string _batchId = Guid.NewGuid().ToString();

        private Eyes? Eyes;
        private string? _apiKey = "";
        private string? _appName = "";
        private string? _testName = "";
        private string? _batchName = "Microsoft Office Tests";
        private string? _serverUrl = "https://eyesapi.applitools.com";
        private bool? _notifyOnCompletion = false;
        private static bool _isFailOnDiff = true;
       
        
        private EyesOfficeProgressUpdate? _eyesOfficeProgressUpdate;
        private bool _reportImages = false;

        readonly IProgress<EyesOfficeProgressUpdate> _progress = new Progress<EyesOfficeProgressUpdate>(update =>
        {
            Console.WriteLine(update.progressMessage + " (" + update.progressValue + "%)");
        });

        public EyesOfficeTester()
        {
        }

        public EyesOfficeTester(string apiKey)
        {
            this._apiKey = apiKey;
        }

        public EyesOfficeTester(IProgress<EyesOfficeProgressUpdate> progress)
        {
            this._progress = progress;
        }

        public EyesOfficeTester(string apiKey, IProgress<EyesOfficeProgressUpdate> progress)
        {
            this._apiKey = apiKey;
            this._progress = progress;
        }

        public void ServerUrl(string serverUrl)
        {
            this._serverUrl = serverUrl;
        }

        public void TestName(string testName)
        {
            this._testName = testName;
        }

        public void AppName(string appName)
        {
            this._appName = appName;
        }

        public void ReportImages(bool reportImages)
        {
            this._reportImages = reportImages;
        }

        public void SetBatch(BatchInfo batchInfo)
        {
            BatchInfo = batchInfo;
        }

        public void NotifyOnCompletion(bool notifyOnCompletion) { 
            this._notifyOnCompletion = notifyOnCompletion;
        }

        public void FailOnDiff(bool isFailOnDiff)
        {
            _isFailOnDiff = isFailOnDiff;
        }

        private void ReportProgress(string message, int value, Bitmap bmp, bool isError)
        {
            if (_progress != null && _eyesOfficeProgressUpdate != null)
            {
                _eyesOfficeProgressUpdate.progressValue = value;
                _eyesOfficeProgressUpdate.progressMessage = message;

                // Don't reset error state if true
                if (!_eyesOfficeProgressUpdate.hasError)
                {
                    _eyesOfficeProgressUpdate.hasError = isError;
                }
                

                if (_reportImages && bmp != null)
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        bmp.Save(ms, ImageFormat.Png);
                        _eyesOfficeProgressUpdate.pngBytes = ms.ToArray();
                    }
                    _eyesOfficeProgressUpdate.bitmap = bmp;
                } 
                else
                {
                    _eyesOfficeProgressUpdate.pngBytes = null;
                    _eyesOfficeProgressUpdate.bitmap = null;
                }

                _progress.Report(_eyesOfficeProgressUpdate);
            }
        }

        void ResetProgress()
        {
            if(_eyesOfficeProgressUpdate != null)
            {
                _eyesOfficeProgressUpdate.progressValue = 0;
                _eyesOfficeProgressUpdate.progressMessage = "";
                _eyesOfficeProgressUpdate.hasError = false;
                _eyesOfficeProgressUpdate.pngBytes = null;
                _eyesOfficeProgressUpdate.bitmap = null;
            }
        }

        private void SetupEyes()
        {
            // Initialize the eyes SDK and set your private API key.
            Eyes = new Eyes(EyesRunner);
            if (BatchInfo == null)
            {
                BatchInfo = new BatchInfo(_batchName);
                BatchInfo.Id = _batchId;
            }
            BatchInfo.NotifyOnCompletion = _notifyOnCompletion;
            Eyes.Batch = BatchInfo;
            if(this._apiKey?.Length > 0)
            {
                Eyes.ApiKey = this._apiKey;
            }
        }

        private void OpenEyes(String appName, String testName)
        {
            // Start the session and set app name and test name.
            Eyes?.Open(appName, testName);
        }

        private void AbortEyes()
        {
            Eyes.Abort();
        }

        private void CloseEyes()
        {
            Eyes?.Close(false);
        }

        private static TestResultsSummary TearDownEyes()
        {
            return EyesRunner.GetAllTestResults(_isFailOnDiff);
        }

        private IEnumerable<FileInfo> GetFilesByExtensions(DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("extensions");
            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension));
        }

        [STAThread]
        public void CheckOfficeFiles(string directory, bool hasProgressBar)
        {
            if (_progress != null && _eyesOfficeProgressUpdate == null)
            {
                _eyesOfficeProgressUpdate = new EyesOfficeProgressUpdate();
            }
            
            SetupEyes();
            DirectoryInfo dInfo = new DirectoryInfo(directory);
            List<FileInfo> files = new List<FileInfo>();
            files.AddRange(GetFilesByExtensions(dInfo, [.. WordFileExtensions]));
            files.AddRange(GetFilesByExtensions(dInfo, [.. ExcelFileExtensions]));
            foreach (FileInfo file in files)
            {
                Thread thread = new Thread(() => CheckFile(file));
                thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
                thread.Start();
                thread.Join(); //Wait for the thread to end
            }

            TestResultsSummary summary = TearDownEyes();
            Console.WriteLine("[Test Results]==============\n" + summary.ToString());
            Console.WriteLine("Done checking Office Documents with Eyes.");
        }

        private void CheckFile(FileInfo file)
        {
            ResetProgress();
            string fileExt = Path.GetExtension(file.FullName);
            if (WordFileExtensions.Contains(fileExt))
            {
                CheckWordPages(file.FullName);
            }
            else if (ExcelFileExtensions.Contains(fileExt))
            {
                CheckSheets(file.FullName);
            }
        }

        [STAThread]
        private void CheckWordPages(string filePath)
        {
            string progressMessage = "Checking Word Document: " + filePath + "...";
            int progressValue = 0;
            ReportProgress(progressMessage, progressValue, null, false);
            Microsoft.Office.Interop.Word.Application myWordApp = 
                new Microsoft.Office.Interop.Word.Application();

            Document myWordDoc = new Document();
            object missing = System.Type.Missing;
            myWordDoc = myWordApp.Documents.Add(filePath, missing, missing, missing);

            // Start the session and set app name and test name.
            String testName = String.IsNullOrEmpty(_testName) ?
                new FileInfo(filePath).Name : _testName;

            String appName = String.IsNullOrEmpty(_appName) ?
                "Microsoft Word" : _appName;

            OpenEyes(appName, testName);

            foreach (Microsoft.Office.Interop.Word.Window window in myWordDoc.Windows)
            {
                // Select Pane 1
                Microsoft.Office.Interop.Word.Pane pane = window.Panes[1];

                // Capture all word pages with Applitools Eyes
                for (var i = 1; i <= pane.Pages.Count; i++)
                {
                    var bits = pane.Pages[i].EnhMetaFileBits;
                    try
                    {
                        Bitmap bmp = EyesCheckWordPageBits(bits,i);
                        progressMessage = "Checked Word Document: " +
                            filePath + " - Page " + i + "...";

                        progressValue = 100 * i / pane.Pages.Count;
                        ReportProgress(progressMessage, progressValue, bmp, false);
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine("EyesOfficeTester: " + ex.Message);
                        Console.WriteLine(ex.StackTrace);
                        progressMessage = "ERROR: " + ex.Message;
                        progressValue = 100 * i / pane.Pages.Count;
                        ReportProgress(progressMessage, progressValue, null, true);
                    }
                }
            }
            if (Eyes.IsOpen)
            {
                if (_eyesOfficeProgressUpdate != null &&
                _eyesOfficeProgressUpdate.hasError)
                {
                    AbortEyes();
                } else
                {
                    CloseEyes();
                }
            }

            myWordDoc.Close(Type.Missing, Type.Missing, Type.Missing);
            myWordApp.Quit(Type.Missing, Type.Missing, Type.Missing);

            progressMessage = "Done checking Word Document: " + filePath + ".";
            progressValue = 100;
            ReportProgress(progressMessage, progressValue, null, false);
        }

        [STAThread]
        private Bitmap EyesCheckWordPageBits(object pageBits, int pageIndex)
        {
            Bitmap bmp = null;
            using (var ms = new MemoryStream((byte[])(pageBits)))
            {
                var image = System.Drawing.Image.FromStream(ms);

                if (Eyes.IsOpen)
                {
                    bmp = new Bitmap(image.Width, image.Height, 
                        PixelFormat.Format32bppArgb);

                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(image,
                            new System.Drawing.Rectangle(
                                new System.Drawing.Point(), image.Size),
                            new System.Drawing.Rectangle(
                                new System.Drawing.Point(), image.Size),
                            GraphicsUnit.Pixel);
                    }

                    // Visual checkpoint.
                    Eyes.CheckImage(bmp, _testName + " - Page " + pageIndex);
                }
            }
            return bmp;
        }

        [STAThread]
        private void CheckSheets(string xlsFile)
        {
            if (BatchInfo == null)
            {
                BatchInfo = new BatchInfo(_batchName);
                BatchInfo.Id = _batchId;
                BatchInfo.NotifyOnCompletion = true;
            }
            string progressMessage = "Checking Excel Document: " + xlsFile + "...";
            int progressValue = 0;
            ReportProgress(progressMessage, progressValue, null, false);

            Excel.Application xl = new Excel.Application();
            xl.CutCopyMode = Excel.XlCutCopyMode.xlCopy;

            if (xl == null)
            {
                progressMessage = "ERROR: No Excel!";
                progressValue = 0;
                ReportProgress(progressMessage, progressValue, null, true);
                return;
            }

            Excel.Workbook wb = xl.Workbooks.Open(xlsFile);
            try
            {
                String testName = String.IsNullOrEmpty(_testName) ?
                new FileInfo(xlsFile).Name : _testName;

                String appName = String.IsNullOrEmpty(_appName) ?
                    "Microsoft Excel" : _appName;

                OpenEyes(appName, testName);

                foreach (Excel.Worksheet sheet in wb.Worksheets)
                {
                    progressValue = 100 * sheet.Index / wb.Worksheets.Count;
                    Bitmap bmp = EyesCheckExcelSheet(sheet, progressValue);
                    if (bmp == null)
                    {
                        progressMessage = "ERROR Checking: " + sheet.Name + 
                            " (" + sheet.Index + "/" + wb.Worksheets.Count + ")";

                        ReportProgress(progressMessage, progressValue, null, true);
                    } else
                    {
                        progressMessage = "Checking: " + sheet.Name +
                            " (" + sheet.Index + "/" + wb.Worksheets.Count + ") ";

                        ReportProgress(progressMessage, progressValue, bmp, false);
                    }
                }

                // End the test.
                CloseEyes();
            }
            catch (Exception ex)
            {
                if (Eyes.IsOpen)
                {
                    Eyes.Abort();
                }
                progressMessage = "ERROR: " + ex.Message + "\n" + ex.StackTrace;
                ReportProgress(progressMessage, progressValue, null, true);
            }
            finally
            {
                wb.Close(0);
                xl.Quit();
                progressMessage = "Done checking Excel Workbook: " + xlsFile + ".";
                progressValue = 100;
                ReportProgress(progressMessage, progressValue, null, false);
            }
        }

        private Bitmap EyesCheckExcelSheet(Excel.Worksheet sheet, int progressValue)
        {
            string progressMessage = "";
            string sheetTag = sheet.Name + "-" + sheet.Index;
            
            Excel.Range r = (Excel.Range)sheet.UsedRange;

            Thread.Sleep(800);
            r.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                           Excel.XlCopyPictureFormat.xlBitmap);

            Thread.Sleep(800);

            Bitmap bmp = null;

            if (Clipboard.GetDataObject() != null)
            {
                IDataObject data = Clipboard.GetDataObject();

                if (data.GetDataPresent(DataFormats.Bitmap))
                {
                    Image image = (Image)data.GetData(DataFormats.Bitmap, true);
                    bmp = new Bitmap(image);

                    if (Eyes.IsOpen)
                    {
                        // Visual checkpoint.
                        Eyes.CheckImage(bmp, sheetTag);
                    }

                }
                else
                {
                    progressMessage = "No image in Clipboard !!";
                    ReportProgress(progressMessage, progressValue, null, true);
                }
            }
            else
            {
                progressMessage = "Clipboard Empty !!";
                ReportProgress(progressMessage, progressValue, null, true);
            }

            return bmp;
        }
    }
}

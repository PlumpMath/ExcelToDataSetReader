namespace ExcelReader
{
   using System;
   using System.Collections.Generic;
   using System.Linq;
   using System.Text;
   using Microsoft.Office.Interop.Excel;
   using System.IO;
   using System.Diagnostics;

   /// <summary>
   /// Class that contains logic for support reading and writing operations using Excel API.
   /// </summary>
   /// <remarks>Handles just one Excel application.</remarks>
   public abstract class ExcelReaderBase : IDisposable 
   {
      #region "Member constants"
      
      private const string ExcelComLabel = "EXCEL";
      /// <summary>
      /// Maximum number of rows of excel 2007.
      /// </summary>
      public const int MaxExcel2007Rows = 65536;
      /// <summary>
      /// Maximum number of columns of excel 2007.
      /// </summary>
      public const int MaxExcel2007Columns = 256;

      // error codes
      private const string ErrorCreateComObj1 = "0x80080005";
      private const string ErrorCreateComObj2 = "0x80070005";
      private const string ErrorCreateComObj3 = "429";
      private const string ActiveXKeyword = "ActiveX";
      private const string ErrorDComAuth1 = "0x800A175D";
      private const string ErrorDComAuth2 = "1004";
      private const string ErrorPrinterNotInstalled = "0x800A03EC";

      #endregion

      private bool _disposed = false;
      private System.Globalization.CultureInfo _XlCulture = new System.Globalization.CultureInfo("en-US");
      private System.Globalization.CultureInfo _originalCulture;
      private bool _hasApplication;
      private readonly List<string> _paths = new List<string>();
      
      /// <summary>
      /// Excel Application.
      /// </summary>
      /// <returns>Excel Application</returns>
      protected Application Application { get; private set; }

      /// <summary>
      /// Keep files.
      /// </summary>
      /// <value>If <v>True</v> then keep the generated files on the hard drive</value>
      protected bool KeepFiles { get; set; }

      /// <summary>
      /// Classifier of the possible groups of exceptions caused by COM components.
      /// </summary>
      private enum ComExceptionClassifier
      {
         CreateCOMObjectExc = 1,
         UserDCOMAuthorizationExc = 2,
         PrinterNotInstalledExc = 3,	
         SystemIOException = 9,	
         OtherExc = 10
      }
  
      /// <summary>
      /// Get the current file path being used by the application.
      /// </summary>
      /// <returns>Path</returns>
      protected string CurrentTempPath {
         get { 
            return _paths.Last(); 
         }
      }

      /// <summary>
      /// Create a new object, without opening an Excel Application.
      /// </summary>
      /// <param name="keepFile">If <v>True</v> the files used by Excel Application will be saved</param>
      protected ExcelReaderBase(bool keepFile)
      {
         Application = null;
         this.KeepFiles = keepFile;
         _hasApplication = false;
      }

      #region "Excel Application methods"
      
      /// <summary>
      /// Creates the Excel application.
      /// /// </summary>
      /// /// <returns>Excel Application</returns>
      protected Application CreateExcelApplication()
      {
         InitializeUsCulture();
         if (_hasApplication)
            CloseApplication();
         try {
            Application = new Application();
            Application.DisplayAlerts = false;
            _hasApplication = true;
            return Application;
         }
         catch (Exception ex) {
            ReleaseComObject(Application);
            throw new InvalidProgramException("Excel not installed in the machine", ex);
         } 
         finally {
            ResetOriginCulture();
         }
      }

      /// <summary>
      /// Closes the Excel application.
      /// </summary>
      /// <remarks>Release all COM components of application and workbooks</remarks>
      private void CloseApplication()
      {
         InitializeUsCulture();
         try {
            if (Application.Workbooks != null && Application.Workbooks.Count > 0) {
               foreach (Workbook workbook in Application.Workbooks) {
                  workbook.Close(KeepFiles == true ? true : false);
                  ReleaseComObject(workbook);
               }
            }
            Application.Quit();
         }
         finally {
            Application.DisplayAlerts = false;
            ResetOriginCulture();
            ReleaseComObject(Application);
            if (!KeepFiles)
               DeleteTempFile();
         }
      }

      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="path">Excel file path</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(string path)
      {
         Workbook workbook = null;
         try {
            string tmpPath = GetTemporaryExcelFile(path);
            workbook = Application.Workbooks.Open(tmpPath);
            return workbook;
         } 
         catch (Exception ex) {
            throw new InvalidProgramException(GetExceptionMessage(ex));
         }
      }
      
      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="file">Binary file</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(byte[] file)
      {
         Workbook workbook = null;
         try {
            string tmpPath = GetTemporaryExcelFile(file);
            workbook = Application.Workbooks.Open(tmpPath);
            return workbook;
         } 
         catch (Exception ex) {
            throw new InvalidProgramException(GetExceptionMessage(ex));
         }
      }

      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="stream">Excel stream</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(Stream stream)
      {
         Workbook workbook = null;
         try {
            string tmpPath = GetTemporaryExcelFile(stream);
            workbook = Application.Workbooks.Open(tmpPath);
            return workbook;
         } 
         catch (Exception ex) {
            throw new InvalidProgramException(GetExceptionMessage(ex));
         }
      }

      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="stream">Excel stream</param>
      /// <param name="extension">File extension (e.g. XLS, XLSX)</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(Stream stream, string extension)
      {
         Workbook workbook = null;
         try {
            string tmpPath = GetTemporaryExcelFile(stream, extension);
            workbook = Application.Workbooks.Open(tmpPath);
            return workbook;
         } 
         catch (Exception ex) {
            throw new InvalidProgramException(GetExceptionMessage(ex));
         }
      }

      #endregion

      #region "Utility Methods"
      /// <summary>
      /// Classify exception based on the message contents.
      /// </summary>
      /// <param name="ex">Generic exception</param>
      /// <returns>Exception classifier</returns>
      private static ComExceptionClassifier ClassifyException(Exception ex)
      {
         if (ex is System.IO.IOException)
            return ComExceptionClassifier.SystemIOException;
         if (ex.Message.Contains(ErrorCreateComObj1) || ex.Message.Contains(ErrorCreateComObj2) || ex.Message.Contains(ActiveXKeyword) || ex.Message.Contains(ErrorCreateComObj3))
            return ComExceptionClassifier.CreateCOMObjectExc;
         if (ex.Message.Contains(ErrorDComAuth1) || ex.Message.Contains(ErrorDComAuth2))
            return ComExceptionClassifier.UserDCOMAuthorizationExc;
         if (ex.Message.Contains(ErrorPrinterNotInstalled))
            return ComExceptionClassifier.PrinterNotInstalledExc;
         return ComExceptionClassifier.OtherExc;
      }

      /// <summary>
      /// Gets the message bounded to a specific COM exception classification
      /// </summary>
      /// <param name="ex">Exception for which the message should be retrieved</param>
      /// <returns>Message to throw</returns>
      private static string GetExceptionMessage(Exception ex)
      {
         ComExceptionClassifier classifier = ClassifyException(ex);
         string msg = string.Empty;
         switch (classifier) {
            case ComExceptionClassifier.SystemIOException:
               msg = "Error IO File";
               break;
            case ComExceptionClassifier.CreateCOMObjectExc:
               msg = "Error while creating COM+ object";
               break;
            case ComExceptionClassifier.PrinterNotInstalledExc:
               msg = "Printer not found";
               break;
            case ComExceptionClassifier.UserDCOMAuthorizationExc:
               msg = "User DCOM+ authorization error";
               break;
            case ComExceptionClassifier.OtherExc:
               msg = "Generic Excel application error";
               break;
            default:
               msg = "Generic Excel application error";
               break;
         }
         return msg + Environment.NewLine + ex.Message;
      }

      /// <summary>
      /// Releases resources uses by the COM object.
      /// </summary>
      /// <param name="com">Instance of the COM object to be released</param>
      /// <remarks>Code provided by MSDN support</remarks>
      protected static void ReleaseComObject(object com)
      {
         try
         {
            while ((System.Runtime.InteropServices.Marshal.ReleaseComObject(com) > 0))
            {
            }
         }
         catch
         {
         }
         finally
         {
            com = null;
         }
      }

      /// <summary>
      /// Set the US culture to the thread.
      /// </summary>
      private void InitializeUsCulture()
      {
         _originalCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
         System.Threading.Thread.CurrentThread.CurrentCulture = _XlCulture;
      }

      /// <summary>
      /// Reset the original culture of the thread.
      /// </summary>
      private void ResetOriginCulture()
      {
         System.Threading.Thread.CurrentThread.CurrentCulture = _originalCulture;
      }

      /// <summary>
      /// Deletes temporary files use in creating the sheet from binary or copying an original Excel sheet.
      /// </summary>
      /// <remarks>Do nothing if an exception is created while deleting the file</remarks>
      private void DeleteTempFile()
      {
         try {
            foreach (string path in _paths) {
               if (!string.IsNullOrEmpty(path) && File.Exists(path)) {
                  File.SetAttributes(path, FileAttributes.Archive | FileAttributes.Normal);
                  File.Delete(path);
               }
            }
         } 
         finally {
            _paths.Clear();
         }
      }

      /// <summary>
      /// Creates a temporary Path and save the Excel stream into a file.
      /// </summary>
      /// <param name="excelFileStream">MS Excel stream</param>
      /// <returns>Path</returns>
      private string GetTemporaryExcelFile(Stream excelFileStream)
      {
         return GetTemporaryExcelFile(excelFileStream, "tmp");
      }

      /// <summary>
      /// Creates a temporary Path and save the Excel stream into a file.
      /// </summary>
      /// <param name="excelFileStream">MS Excel stream</param>
      /// <param name="extension">File extension (e.g. XLS, XLSX)</param>
      /// <returns>Path</returns>
      private string GetTemporaryExcelFile(Stream excelFileStream, string extension)
      {
         string path = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetRandomFileName())) + "." + extension;
         MemoryStream ExcelFileMemoryStream = CopyStreamToMemoryStream(excelFileStream);
         using (FileStream fStrem = new FileStream(path, FileMode.Create, FileAccess.Write)) {
            ExcelFileMemoryStream.WriteTo(fStrem);
         }
         _paths.Add(path);
         return path;
      }

      /// <summary>
      /// Create a temporary Path and save the excel file into a temporary file.
      /// </summary>
      /// <param name="excelPath">Excel file path</param>
      /// <returns>Path</returns>
      protected string GetTemporaryExcelFile(string excelPath)
      {
         string tmpPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetRandomFileName()));
         if (File.Exists(excelPath)) {
            File.Copy(excelPath, tmpPath, true);
         }
         _paths.Add(tmpPath);
         return tmpPath;
      }

      /// <summary>
      /// Create a temporary Path and save the Excel file into a temporary file.
      /// </summary>
      /// <param name="file">Excel binary</param>
      /// <returns>Path</returns>
      private string GetTemporaryExcelFile(byte[] file)
      {
         string path = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetRandomFileName()));
         File.WriteAllBytes(path, file);
         _paths.Add(path);
         return path;
      }

      /// <summary>
      /// Copy a Stream into a new Memory Stream.
      /// </summary>
      /// <param name="stream">Input stream.</param>
      protected static MemoryStream CopyStreamToMemoryStream(Stream stream)
      {
         MemoryStream res = new MemoryStream();
         //TODO: .NET >= 4.0? Input.CopyTo(Output, 8 * 1024);
         byte[] buffer = new byte[8 * 1024];
         int Length = 0;
         Length = stream.Read(buffer, 0, buffer.Length);
         while (Length > 0) {
            res.Write(buffer, 0, Length);
            Length = stream.Read(buffer, 0, buffer.Length);
         }
         return res;
      }

      /// <summary>
      /// Get the sheet count for a given workbook.
      /// </summary>
      /// <param name="workbook">Workbook.</param>
      /// <returns>Number of sheets, <c>0</c> if any exception is caught.</returns>
      /// <remarks>If the workbook is not a workbook, is logically correct to return 0, since a workbook must have at least 1 sheet.</remarks>
      protected static int GetSheetCount(Workbook workbook)
      {
         try {
            return workbook.Worksheets.Count;
         } 
         catch {
            return 0;
         }
      }

      /// <summary>
      /// Get the sheet names for a given workbook.
      /// </summary>
      /// <param name="workbook">Workbook.</param>
      /// <returns>Array containing the sheet names. If no sheet found, array containing one empty string.</returns>
      protected static string[] GetSheetNames(Workbook workbook)
      {
         int count = GetSheetCount(workbook);
         if (count != 0) {
            string[] sheetNames = new string[count];
            int i = 0;
            foreach (Worksheet Sheet in workbook.Worksheets) {
               sheetNames[i] = Sheet.Name;
               i++;
            }
            return sheetNames;
         }
         return new string[] { string.Empty };
      }

      /// <summary>
      /// Number of Excel COM processes running in the System.
      /// </summary>
      /// <returns>Excel COM processes number.</returns>
      protected static int ExcelComProcesses()
      {
         Process[] procList = Process.GetProcesses();
         int counter = 0;
         for (int i = 0; i <= procList.GetUpperBound(0); i++)
         {
            if (procList[i].ProcessName.Contains(ExcelComLabel))
               counter += 1;
         }
         return counter;
      }

      #endregion

      #region "Dispose and Finalize"
      /// <summary>
      /// Dispose the base object resources.
      /// </summary>
      /// <param name="disposing">Called by Dispose (<c>True</c>) or Finalize (<c>False</c>) methods</param>
      protected virtual void Dispose(bool disposing)
      {
         if (!_disposed) {
            if (disposing && _hasApplication) {
               CloseApplication();
               _hasApplication = false;
            }
            DeleteTempFile();
            KeepFiles = false;
         }
         _disposed = true;
      }

      /// <summary>
      /// Dispose the object deleting application and files, releasing the COM components and setting parameters to default.
      /// </summary>
      public void Dispose()
      {        
         Dispose(true);
         GC.SuppressFinalize(this);
      }
      
      #endregion
   }
}
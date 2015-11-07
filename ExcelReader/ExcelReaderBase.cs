using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;

namespace ExcelReader
{
   /// <summary>
   /// Class that contains logic for support reading and writing operations using Excel API.
   /// </summary>
   /// <remarks>Handles just one Excel application.</remarks>
   public abstract class ExcelReaderBase : IDisposable 
   {
      #region "Member constants"
      
      private const string ExcelCOMLabel = "EXCEL";
      public const int MaxExcel2007Rows = 65536;
      public const int MaxExcel2007Columns = 256;

      // error codes
      private const string ErrorCreateCOMObj1 = "0x80080005";
      private const string ErrorCreateCOMObj2 = "0x80070005";
      private const string ErrorCreateCOMObj3 = "429";
      private const string ActiveXKeyword = "ActiveX";
      private const string ErrorDCOMAuth1 = "0x800A175D";
      private const string ErrorDCOMAuth2 = "1004";
      private const string ErrorPrinterNotInstalled = "0x800A03EC";

      #endregion

      protected bool _disposed = false;
      private System.Globalization.CultureInfo _XLCulture = new System.Globalization.CultureInfo("en-US");
      private System.Globalization.CultureInfo _OrigCulture;
      private bool _HasApplication;
      private readonly List<string> _FilePaths = new List<string>();
      
      /// <summary>
      /// Excel Application.
      /// </summary>
      /// <returns>Excel Application</returns>
      public Application ExcelApplication {get; private set;}

      /// <summary>
      /// Keep files.
      /// </summary>
      /// <value>If <v>True</v> then keep the generated files on the hard drive</value>
      public bool KeepFiles {get; set;}

      /// <summary>
      /// Classifier of the possible groups of exceptions caused by COM components.
      /// </summary>
      private enum COMExceptionClassifier
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
            return _FilePaths.Last(); 
         }
      }

      /// <summary>
      /// Create a new object, without opening an Excel Application.
      /// </summary>
      /// <param name="KeepFiles">If <v>True</v> the files used by Excel Application will be saved</param>
      public ExcelReaderBase(bool KeepFiles)
      {
         ExcelApplication = null;
         this.KeepFiles = KeepFiles;
         _HasApplication = false;
      }

      #region "Excel Application methods"
      
      /// <summary>
      /// Creates the Excel application.
      /// /// </summary>
      /// /// <returns>Excel Application</returns>
      protected Application CreateExcelApplication()
      {
         InitializeUSCulture();
         if (_HasApplication)
            CloseApplication();
         try {
            ExcelApplication = new Application();
            ExcelApplication.DisplayAlerts = false;
            _HasApplication = true;
            return ExcelApplication;
         }
         catch (Exception ex) {
            ReleaseCOMObject(ExcelApplication);
            throw new ApplicationException("Excel not installed in the machine", ex);
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
         InitializeUSCulture();
         try {
            if (ExcelApplication.Workbooks != null && ExcelApplication.Workbooks.Count > 0) {
               foreach (Workbook Workbook in ExcelApplication.Workbooks) {
                  Workbook.Close(KeepFiles == true ? true : false);
                  ReleaseCOMObject(Workbook);
               }
            }
            ExcelApplication.Quit();
         }
         finally {
            ExcelApplication.DisplayAlerts = false;
            ResetOriginCulture();
            ReleaseCOMObject(ExcelApplication);
            if (!KeepFiles)
               DeleteTempFile();
         }
      }

      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="Path">Excel file path</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(string Path)
      {
         Workbook Workbook = null;
         try {
            string TempPath = GetTemporaryExcelFile(Path);
            Workbook = ExcelApplication.Workbooks.Open(TempPath);
            return Workbook;
         } 
         catch (Exception ex) {
            throw new Exception(GetExceptionMessage(ex));
         }
      }
      
      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="BinaryFile">Binary stream</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(byte[] BinaryFile)
      {
         Workbook Workbook = null;
         try {
            string TempPath = GetTemporaryExcelFile(BinaryFile);
            Workbook = ExcelApplication.Workbooks.Open(TempPath);
            return Workbook;
         } 
         catch (Exception ex) {
            throw new Exception(GetExceptionMessage(ex));
         }
      }

      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="Stream">Excel stream</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(Stream Stream)
      {
         Workbook Workbook = null;
         try {
            string TempPath = GetTemporaryExcelFile(Stream);
            Workbook = ExcelApplication.Workbooks.Open(TempPath);
            return Workbook;
         } 
         catch (Exception ex) {
            throw new Exception(GetExceptionMessage(ex));
         }
      }

      /// <summary>
      /// Open an Excel workbook in the application.
      /// </summary>
      /// <param name="Stream">Excel stream</param>
      /// <param name="Extension">File extension (e.g. XLS, XLSX)</param>
      /// <returns>Instance of the opened workbook</returns>
      protected Workbook OpenWorkbook(Stream Stream, string Extension)
      {
         Workbook Workbook = null;
         try {
            string TempPath = GetTemporaryExcelFile(Stream, Extension);
            Workbook = ExcelApplication.Workbooks.Open(TempPath);
            return Workbook;
         } 
         catch (Exception ex) {
            throw new Exception(GetExceptionMessage(ex));
         }
      }

      #endregion

      #region "Utility Methods"
      /// <summary>
      /// Classify exception based on the message contents.
      /// </summary>
      /// <param name="ex">Generic exception</param>
      /// <returns>Exception classifier</returns>
      private COMExceptionClassifier ClassifyException(Exception ex)
      {
         if (ex is System.IO.IOException)
            return COMExceptionClassifier.SystemIOException;
         if (ex.Message.Contains(ErrorCreateCOMObj1) || ex.Message.Contains(ErrorCreateCOMObj2) || ex.Message.Contains(ActiveXKeyword) || ex.Message.Contains(ErrorCreateCOMObj3))
            return COMExceptionClassifier.CreateCOMObjectExc;
         if (ex.Message.Contains(ErrorDCOMAuth1) || ex.Message.Contains(ErrorDCOMAuth2))
            return COMExceptionClassifier.UserDCOMAuthorizationExc;
         if (ex.Message.Contains(ErrorPrinterNotInstalled))
            return COMExceptionClassifier.PrinterNotInstalledExc;
         return COMExceptionClassifier.OtherExc;
      }

      /// <summary>
      /// Gets the message bounded to a specific COM exception classification
      /// </summary>
      /// <param name="ex">Exception for which the message should be retrieved</param>
      /// <returns>Message to throw</returns>
      private string GetExceptionMessage(Exception ex)
      {
         COMExceptionClassifier uCOMExceptionClassifier = ClassifyException(ex);
         string sMessage = string.Empty;
         switch (uCOMExceptionClassifier) {
            case COMExceptionClassifier.SystemIOException:
               sMessage = "Error IO File";
               break;
            case COMExceptionClassifier.CreateCOMObjectExc:
               sMessage = "Error while creating COM+ object";
               break;
            case COMExceptionClassifier.PrinterNotInstalledExc:
               sMessage = "Printer not found";
               break;
            case COMExceptionClassifier.UserDCOMAuthorizationExc:
               sMessage = "User DCOM+ authorization error";
               break;
            case COMExceptionClassifier.OtherExc:
               sMessage = "Generic Excel application error";
               break;
            default:
               sMessage = "Generic Excel application error";
               break;
         }
         return sMessage + Environment.NewLine + ex.Message;
      }

      /// <summary>
      /// Releases resources uses by the COM object.
      /// </summary>
      /// <param name="objCOM">Instance of the COM object to be released</param>
      /// <remarks>Code provided by MSDN support</remarks>
      public static void ReleaseCOMObject(object objCOM)
      {
         try
         {
            while ((System.Runtime.InteropServices.Marshal.ReleaseComObject(objCOM) > 0))
            {
            }
         }
         catch
         {
         }
         finally
         {
            objCOM = null;
         }
      }

      /// <summary>
      /// Set the US culture to the thread.
      /// </summary>
      private void InitializeUSCulture()
      {
         _OrigCulture = System.Threading.Thread.CurrentThread.CurrentCulture;
         System.Threading.Thread.CurrentThread.CurrentCulture = _XLCulture;
      }

      /// <summary>
      /// Reset the original culture of the thread.
      /// </summary>
      private void ResetOriginCulture()
      {
         System.Threading.Thread.CurrentThread.CurrentCulture = _OrigCulture;
      }

      /// <summary>
      /// Deletes temporary files use in creating the sheet from binary or copying an original Excel sheet.
      /// </summary>
      /// <remarks>Do nothing if an exception is created while deleting the file</remarks>
      private void DeleteTempFile()
      {
         try {
            foreach (string path in _FilePaths) {
               if (!string.IsNullOrEmpty(path) && File.Exists(path)) {
                  File.SetAttributes(path, FileAttributes.Archive | FileAttributes.Normal);
                  File.Delete(path);
               }
            }
         } 
         finally {
            _FilePaths.Clear();
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
         string StringPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetRandomFileName())) + "." + extension;
         MemoryStream ExcelFileMemoryStream = CopyStreamToMemoryStream(excelFileStream);
         using (FileStream FileStream = new FileStream(StringPath, FileMode.Create, FileAccess.Write)) {
            ExcelFileMemoryStream.WriteTo(FileStream);
         }
         _FilePaths.Add(StringPath);
         return StringPath;
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
         _FilePaths.Add(tmpPath);
         return tmpPath;
      }

      /// <summary>
      /// Create a temporary Path and save the Excel file into a temporary file.
      /// </summary>
      /// <param name="BinaryFile">Excel binary</param>
      /// <returns>Path</returns>
      private string GetTemporaryExcelFile(byte[] BinaryFile)
      {
         string StringPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetRandomFileName()));
         File.WriteAllBytes(StringPath, BinaryFile);
         _FilePaths.Add(StringPath);
         return StringPath;
      }

      /// <summary>
      /// Copy a Stream into a new Memory Stream.
      /// </summary>
      /// <param name="Input">Input stream.</param>
      public static MemoryStream CopyStreamToMemoryStream(Stream Input)
      {
         MemoryStream Output = new MemoryStream();
         // .NET >= 4.0? Input.CopyTo(Output, 8 * 1024);
         byte[] cBuffer = new byte[8 * 1024];
         int Length = 0;
         Length = Input.Read(cBuffer, 0, cBuffer.Length);
         while (Length > 0) {
            Output.Write(cBuffer, 0, Length);
            Length = Input.Read(cBuffer, 0, cBuffer.Length);
         }
         return Output;
      }

      /// <summary>
      /// Get the sheet count for a given workbook.
      /// </summary>
      /// <param name="Workbook">Workbook.</param>
      /// <returns>Number of sheets, <c>0</c> if any exception is caught.</returns>
      /// <remarks>If the workbook is not a workbook, is logically correct to return 0, since a workbook must have at least 1 sheet.</remarks>
      protected int GetSheetCount(Workbook Workbook)
      {
         try {
            return Workbook.Worksheets.Count;
         } 
         catch {
            return 0;
         }
      }

      /// <summary>
      /// Get the sheet names for a given workbook.
      /// </summary>
      /// <param name="Workbook">Workbook.</param>
      /// <returns>Array containing the sheet names. If no sheet found, array containing one empty string.</returns>
      protected string[] GetSheetNames(Workbook Workbook)
      {
         int count = GetSheetCount(Workbook);
         if (count != 0) {
            string[] SheetNames = new string[count];
            int i = 0;
            foreach (Worksheet Sheet in Workbook.Worksheets) {
               SheetNames[i] = Sheet.Name;
               i++;
            }
            return SheetNames;
         }
         return new string[] { string.Empty };
      }

      /// <summary>
      /// Number of Excel COM processes running in the System.
      /// </summary>
      /// <returns>Excel COM processes number.</returns>
      public int ExcelCOMProcesses()
      {
         Process[] procList = Process.GetProcesses();
         int ExcelCOMCounter = 0;
         for (int i = 0; i <= procList.GetUpperBound(0); i++)
         {
            if (procList[i].ProcessName.Contains(ExcelCOMLabel))
               ExcelCOMCounter += 1;
         }
         return ExcelCOMCounter;
      }

      #endregion

      #region "Dispose and Finalize"
      /// <summary>
      /// Dispose the base object resources.
      /// </summary>
      /// <param name="Disposing">Called by Dispose (<c>True</c>) or Finalize (<c>False</c>) methods</param>
      protected virtual void Dispose(bool Disposing)
      {
         if (!_disposed) {
            if (Disposing && _HasApplication) {
               CloseApplication();
               _HasApplication = false;
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
      }
      
      #endregion
   }
}
using System;
using System.Data;

namespace ExcelReader
{
   /// <summary>
   /// Generic interface handling the transformation from generic files to datasets.
   /// </summary>
   public interface IDatasetReader
   {  
      /// <summary>
      /// Get a dataset out of a file.
      /// </summary>
      /// <param name="FileData">Generic file containing data that can be translated to a dataset</param>
      /// <returns>Translated dataset</returns>
      DataSet GetDataSet(byte[] FileData);
   }
}

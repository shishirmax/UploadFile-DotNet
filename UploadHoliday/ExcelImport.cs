using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Excel;
using System.Data;
using System.IO;

namespace UploadHoliday
{
    public class ExcelData
    {
        readonly string _path;
        private static readonly object mlock = new object();
        public ExcelData(string path)
        {
            _path = path;
        }

        /// <summary>
        /// Sets Binary(.xls)/OpenXml(.xlsx) Reader as per file type.
        /// </summary>
        /// <returns></returns>
        private IExcelDataReader GetExcelReader()
        {
            FileStream stream = null;
            IExcelDataReader reader = null;
            try
            {

                //to avoid concurrency error:"An item with the same key has already been added" 
                //lock below code.Otherwise it will give avobe error when different files from different
                //report type folders will execute.
                lock (mlock)
                {
                    if (_path.EndsWith(".xls"))  //will not parse if file is corrupted.
                    {
                        stream = File.Open(_path, FileMode.Open, FileAccess.Read);
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);

                    }
                    if (_path.EndsWith(".xlsx") || (reader != null && !reader.IsValid))  //Extra check if .xls is not parsed correctly.
                    {
                        stream = File.Open(_path, FileMode.Open, FileAccess.Read);
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    }
                }

            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
            return reader;
        }


        /// <summary>
        /// Fetches all the rows for first sheet.
        /// </summary>
        /// <param name="firstRowIsColumnNames"></param>
        /// <returns></returns>
        public IEnumerable<DataRow> GetData(bool firstRowIsColumnNames = false)
        {
            IEnumerable<DataRow> rows = null;
            DataTable workSheet;
            try
            {
                var reader = this.GetExcelReader();

                reader.IsFirstRowAsColumnNames = firstRowIsColumnNames;

                workSheet = reader.AsDataSet().Tables[0];

                rows = (from DataRow a in workSheet.Rows select a);
            }
            catch (Exception e)
            {
                throw;
            }

            return rows;
        }


        /// <summary>
        /// Converts excel to datatable.
        /// </summary>
        /// <param name="firstRowIsColumnNames">Set true if first row contains columns</param>
        /// <returns></returns>
        public DataTable GetExcelDataRows(int sheet = 0, bool firstRowIsColumnNames = false)
        {
            DataTable dt = null;
            try
            {
                using (IExcelDataReader reader = this.GetExcelReader())
                {
                    if (reader.IsValid)
                    {
                        reader.IsFirstRowAsColumnNames = firstRowIsColumnNames;
                        dt = reader.AsDataSet(true).Tables[sheet];
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }

            return dt;
        }

    }
}
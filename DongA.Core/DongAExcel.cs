using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using AsposeCells = Aspose.Cells;

namespace DongA.Core
{

    public class DongAExcel
    {
        /// <summary>
        /// Content Type - Excel 2007 trở lên
        /// application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        /// </summary>
        public const string XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        /// <summary>
        /// application/pdf
        /// </summary>
        public const string PDF = "application/pdf";

        /// <summary>
        /// [Aspose] Save to stream
        /// </summary>
        /// <param name="workBook"></param>
        /// <returns></returns>
        public Stream SaveToStream(AsposeCells.Workbook workBook)
        {
            Stream stream = new MemoryStream();
            //Save theo định dạng của excel (office97-2003/2007/2010)
            AsposeCells.SaveFormat saveFormat = AsposeCells.SaveFormat.Xlsx;
            // Save
            saveFormat = AsposeCells.SaveFormat.Xlsx;
            //Save the Excel file.
            workBook.Save(stream, saveFormat);
            stream.Seek(0, SeekOrigin.Begin);
            return stream;
        }

        /// <summary>
        /// [Aspose] Open excel
        /// </summary>
        /// <param name="excelTemplatePath"></param>
        /// <returns></returns>
        public AsposeCells.Workbook OpenExcelFile(string excelPath)
        {
            //Load theo định dạng của excel (office97-2003/2007/2010)
            AsposeCells.LoadFormat loadFormat = AsposeCells.LoadFormat.Xlsx;
            loadFormat = AsposeCells.LoadFormat.Xlsx;

            var workBook = new AsposeCells.Workbook(excelPath,
                new AsposeCells.LoadOptions(loadFormat));
            return workBook;
        }

        /// <summary>
        /// Convert to listObject to datatable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;

        }

        /// <summary>
        /// convert từ list object to dataset
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public DataSet ConvertListObjectToDataSet<T>(IList<T> data)
        {
            DataTable dataTable = ConvertToDataTable(data);

            DataSet ds = new DataSet();
            if (dataTable.Rows.Count > 0)
            {
                ds.Tables.Add(dataTable);
            }
            return ds;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using WB = Microsoft.Office.Interop.Excel.Workbook;
using WS = Microsoft.Office.Interop.Excel.Worksheet;
using RG = Microsoft.Office.Interop.Excel.Range;
using xl = Microsoft.Office.Interop.Excel;
using xlApp = Microsoft.Office.Interop.Excel.Application;
using Microsoft.Office.Tools.Ribbon;
using DKARibbon;
using System.ComponentModel;
using SD = System.Data;

namespace DKAExcelStuff
{
    class DKAWrite : KAXL
    {        
        public void WriteDataFromTable(SD.DataTable dt,Worksheet ws)
        {        
            int rows = dt.Rows.Count;

            // Create new worksheet, and write datatable
            ws.Activate();
            WS newWorksheet = (WS)Globals.ThisAddIn.Application.Worksheets.Add();

            // column headings
            for (var i = 0; i < dt.Columns.Count; i++)
            {
                newWorksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
            }

            int writeRow = 1;

            foreach (SD.DataRow datarow in dt.Rows)
            {
                // first element of list is all null...start with element[1]
                if (writeRow != 1)
                {
                    for (int i = 1; i <= dt.Columns.Count; i++)
                    {
                        newWorksheet.Cells[writeRow, i] = datarow[i - 1].ToString();
                    }
                }                
                writeRow++;
            }
        }
        public SD.DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
            TypeDescriptor.GetProperties(typeof(T));
            SD.DataTable table = new SD.DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                SD.DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }        
    }
}

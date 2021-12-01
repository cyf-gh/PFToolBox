using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcel {
    public class EUtil {
        FileStream stream = null;
        public IExcelDataReader ExcelDataReader { get; set; }
        public DataTableCollection Tables { get; set; }
        public DataTable GetTableByName( string tableName )
        {
            foreach ( DataTable t in Tables ) {
                if ( t.TableName == tableName ) {
                    return t;
                }
            }
            return null;
        }
        public DataTableCollection OpenExcel()
        {
            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;
                string filePath = string.Empty;

                if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
                    filePath = openFileDialog.FileName;
                    filePath += ".copy.xlsx";
                    if ( File.Exists( filePath ) ) {
                        File.Delete( filePath );
                    }
                    File.Copy( openFileDialog.FileName, filePath );
                    try {
                        stream = File.Open( filePath, FileMode.Open, FileAccess.Read );
                    } catch ( Exception ex ) {
                        MessageBox.Show( $"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试" );
                        return null;
                    }
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    ExcelDataReader = ExcelReaderFactory.CreateReader( stream );
                    var result = ExcelDataReader.AsDataSet();
                    Tables = result.Tables;
                }
            }
            return Tables;
        }
        Microsoft.Office.Interop.Excel.Application xlApp;
        public EUtil()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
        }
        public Microsoft.Office.Interop.Excel.Workbook CreateNewWorkBook()
        {
            return xlApp.Workbooks.Add();
        }
        public Microsoft.Office.Interop.Excel.Worksheet GetWorksheet( Microsoft.Office.Interop.Excel.Workbook workbook, int i = 1 )
        {
            return workbook.Worksheets.get_Item( i ) as Microsoft.Office.Interop.Excel.Worksheet;
        }
    }
}

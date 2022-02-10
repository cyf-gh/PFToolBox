using ExcelDataReader;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcel {
    public class NPOIUtil {
        string strTmpExcelPath = string.Empty;
        public IWorkbook WB { get; set; } = null;

        public ISheet GetSheet( string name )
        {
            return WB.GetSheet( name );
        }
        public ISheet GetSheetAt( int i )
        {
            return WB.GetSheetAt( i );
        }

        public void OpenFileDialog()
        {
            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;
                string filePath = string.Empty;

                if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
                    filePath = openFileDialog.FileName;
                    var newFilePath = openFileDialog.FileName;
                    filePath += ".npio.copy.xlsx";
                    strTmpExcelPath = filePath;
                    if ( File.Exists( filePath ) ) {
                        File.Delete( filePath );
                    }
                    File.Copy( openFileDialog.FileName, filePath );

                    try {
                        var stream = File.Open( filePath, FileMode.Open, FileAccess.Read );
                        WB = new XSSFWorkbook( stream );
                        stream.Close();
                    } catch ( Exception ex ) {
                        MessageBox.Show( $"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试" );
                    }
                }
            }
        }
        public void OpenFile( string path )
        {
            var filePath = string.Empty;
            var newFilePath = path;
            filePath += ".copy.xlsx";
            strTmpExcelPath = filePath;
            if ( File.Exists( filePath ) ) {
                File.Delete( filePath );
            }
            File.Copy( path, filePath );

            try {
                var stream = File.Open( filePath, FileMode.Open, FileAccess.Read );
                WB = new XSSFWorkbook( stream );
                stream.Close();
            } catch ( Exception ex ) {
                MessageBox.Show( $"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试" );
                return;
            }
        }
        static void InsertRows( ref HSSFSheet s1, int fromRowIndex, int rowCount )
        {
            s1.ShiftRows( fromRowIndex, s1.LastRowNum, rowCount, true, false, true );

            for ( int rowIndex = fromRowIndex; rowIndex < fromRowIndex + rowCount; rowIndex++ ) {
                HSSFRow rowSource = s1.GetRow( rowIndex + rowCount ) as HSSFRow;
                HSSFRow rowInsert = s1.CreateRow( rowIndex ) as HSSFRow;
                rowInsert.Height = rowSource.Height;
                for ( int colIndex = 0; colIndex < rowSource.LastCellNum; colIndex++ ) {
                    HSSFCell cellSource = rowSource.GetCell( colIndex ) as HSSFCell;
                    HSSFCell cellInsert = rowInsert.CreateCell( colIndex ) as HSSFCell;
                    if ( cellSource != null ) {
                        cellInsert.CellStyle = cellSource.CellStyle;
                    }
                }
            }
        }
    }
    public class EUtil {
        FileStream stream = null;
        string strTmpExcelPath = string.Empty;
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
        public DataTable GetTableByIndex( int i )
        {
            return Tables[i];
        }
        public void ClearTempExcelFile()
        {
            if ( File.Exists( strTmpExcelPath ) ) {
                File.Delete( strTmpExcelPath );
            }
        }
        public Microsoft.Office.Interop.Excel.Workbooks wkbks = null;
        public Microsoft.Office.Interop.Excel.Workbook wkbk = null;
        public DataTableCollection OpenExcel()
        {
            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;
                string filePath = string.Empty;

                if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
                    filePath = openFileDialog.FileName;
                    var newFilePath = openFileDialog.FileName;
                    filePath += ".copy.xlsx";
                    strTmpExcelPath = filePath;
                    if ( File.Exists( filePath ) ) {
                        File.Delete( filePath );
                    }
                    File.Copy( openFileDialog.FileName, filePath );
                    wkbks = xlApp.Workbooks;
                    wkbk = wkbks.Open( newFilePath );

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
        public DataTableCollection OpenExcelByPath( string path )
        {

            var filePath = path;
            filePath += ".copy.xlsx";
            strTmpExcelPath = filePath;
            if ( File.Exists( filePath ) ) {
                File.Delete( filePath );
            }
            File.Copy( path, filePath );
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
            return Tables;
        }
        Microsoft.Office.Interop.Excel.Application xlApp;
        public EUtil()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
        }
        public Microsoft.Office.Interop.Excel.Workbook CreateNewWorkbook()
        {
            return xlApp.Workbooks.Add();
        }
        public Microsoft.Office.Interop.Excel.Worksheet GetWorksheet( Microsoft.Office.Interop.Excel.Workbook workbook, int i = 1 )
        {
            return workbook.Worksheets.get_Item( i ) as Microsoft.Office.Interop.Excel.Worksheet;
        }
    }
}

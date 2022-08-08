using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Merge_FinaAnaly : Form {
        public Form_Merge_FinaAnaly() {
            InitializeComponent();
        }
        public string filePath { get; set; }
        FileStream stream;
        IExcelDataReader reader;
        private void Form_Merge_FinaAnaly_Load( object sender, EventArgs e ) {
            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;

                filePath = @"C:\Users\cyf-thinkpad\Desktop\授信财务报表\2021 (3).xlsx";
                string rawfp = filePath;
                filePath += ".copy.xlsx";
                if ( File.Exists( filePath ) ) {
                    File.Delete( filePath );
                }
                File.Copy( rawfp, filePath );
                try {
                    stream = File.Open( filePath, FileMode.Open, FileAccess.Read );
                } catch ( Exception ex ) {
                    MessageBox.Show( $"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试" );
                    return;
                }
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                reader = ExcelReaderFactory.CreateReader( stream );
                // 2. Use the AsDataSet extension method
                // var result = excelDataReader.AsDataSet();
                // The result of each spreadsheet is in result.Tables
                var result = reader.AsDataSet();
                var t = result.Tables;
                var defaultT = t[0];
                var a = defaultT.Rows[0].ItemArray;
            }
        }
    }
}

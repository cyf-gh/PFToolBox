using ExcelDataReader;

using Microsoft.Office.Interop.Excel;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Merge_FinaAnalyPlus : Form {
        public Form_Merge_FinaAnalyPlus() {
            InitializeComponent();
        }
        public class Sheet {
            public static List<Sheet> Sheets = new List<Sheet>();
            public static Sheet GetSheetByYearName( string year, string name ) { return Sheets.Find( m => { return m.Year == year && m.Name == name; } ); }
            public Sheet( string year, string index ) {
                Year = year;
                Index = index;

                dataReader = getReader();
                Rows = dataReader.AsDataSet().Tables[0].Rows;
                var r1 = Rows[0].ItemArray;

                Name = r1[0].ToString();
                loadSections();
                if ( Name == "资产负债表" ) {
                    loadSections( 4 );
                }
                Sheets.Add( this );
            }
            public class Section {
                private readonly Sheet sheet;
                string removeWhitespace( string str ) {
                    if ( str == null ) {
                        return String.Empty;
                    }
                    return Regex.Replace( str, @"\s+", "" );
                }
                public Section( Sheet parent ) {
                    this.sheet = parent;
                }
                public string Name { get; set; }
                public int NameWhiteSpaceCount {
                    get {
                        var spaceCount = Name.Split( ' ' ).Length - 1;
                        var tabCount = Name.Split( '　' ).Length - 1;
                        return spaceCount == 0 ? tabCount : spaceCount;
                    }
                }
                private double v;
                public double Value { get { Read = true; return v; } set { v = value; } }
                public double Value10SByYear( string year ) {
                    var s = GetSheetByYearName( year, sheet.Name );
                    var ss = s.Sections.Find( m => removeWhitespace( m.Name ) == removeWhitespace( Name ) ); // 匹配section name
                    return ss == null ? 0 : ss.Value10S;
                }
                public double ValueIncreasmentRate( string year ) {
                    var diff = Value10SByYear( year ) - Value10S;
                    diff = int.Parse( year ) > int.Parse( sheet.Year ) ? diff : -diff;
                    var basement = int.Parse( year ) > int.Parse( sheet.Year ) ? Value10S : Value10SByYear( year );
                    return diff / basement;
                }
                public double Value10S { get { return Value / 10000; } }
                public bool IsSum { get { return Name.Contains( "合计" ); } }
                /// <summary>
                /// 是否被读取过数据
                /// </summary>
                public bool Read { get; set; } = false;
            }
            public string Name { get; set; }
            public readonly string Year;
            public readonly string Index;
            public List<Section> Sections { get; set; } = new List<Section>();

            void loadSections( int nameIndex = 0 ) {
                var valueIndex = Name == "现金流量表" ? nameIndex + 2 : nameIndex + 3;
                for ( int j = 2; j < Rows.Count; j++ ) {
                    var r = Rows[j].ItemArray;
                    Sections.Add( new Section( this ) {
                        Name = r[nameIndex].ToString(),
                        Value = r[valueIndex] is DBNull ? 0 : double.Parse( r[valueIndex].ToString() ),
                    } );
                }
            }
            IExcelDataReader dataReader;
            DataRowCollection Rows;
            string filePath;
            string fileTmpPath { get { return filePath + ".copy.xlsx"; } }
            FileStream stream;
            IExcelDataReader getReader() {
                filePath = $@"C:\Users\cyf-thinkpad\Desktop\1\{Year}\{Year} ({Index}).xlsx";
                if ( File.Exists( fileTmpPath ) ) {
                    File.Delete( fileTmpPath );
                }
                File.Copy( filePath, fileTmpPath );
                try {
                    stream = File.Open( fileTmpPath, FileMode.Open, FileAccess.Read );
                } catch ( Exception ex ) {
                    MessageBox.Show( $"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试" );
                    return null;
                }
                return ExcelReaderFactory.CreateReader( stream );
            }
            ~Sheet() {
                if ( File.Exists( fileTmpPath ) ) {
                    File.Delete( fileTmpPath );
                }
            }
        }

        private void Form_Merge_FinaAnalyPlus_Load( object sender, EventArgs e ) {
            for ( int ii = 0; ii < 4; ii++ ) {
                for ( int i = 1; i < 4; i++ ) {
                    var sheet = new Sheet( ( 2019 + ii ).ToString(), i.ToString() ) { };
                }
            }
            var a = JsonConvert.SerializeObject( Sheet.Sheets );
            var s2019s = Sheet.Sheets.FindAll( m => m.Year == "2019" );
            string res = "";
            foreach ( var s in s2019s ) {
                res += $"\n{s.Name}\n";
                foreach ( var ss in s.Sections ) {
                    res += $"{ss.Name}\t{ss.Value10S}\t{ss.Value10SByYear( "2020" )}\t{ss.Value10SByYear( "2021" )}\t{ss.ValueIncreasmentRate("2020")}\t{ss.ValueIncreasmentRate( "2021" )}\n";
                }
            }

            #region SAVE_ANALY
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            var xlWorkBook = xlApp.Workbooks.Add();
            var ws = xlWorkBook.Worksheets.get_Item( 1 ) as Worksheet;
            // Cells[Row, Col]
            // 绘制表头
            setCellValue( ws.Cells[1, 1], "近三年主要财务数据列表：", true );
            setCellValue( ws.Cells[2, 1], "财务分析表", true, XlHAlign.xlHAlignCenter );
            mergeCell( ws, 2, 1, 2, 5 );
            setCellValue( ws.Cells[3, 1], "项目（单位：人民币万元）" );
            setCellValue( ws.Cells[4, 1], "报表类型（单一/合并）" );
            setCellValue( ws.Cells[5, 1], "是否审计" );
            setCellValue( ws.Cells[6, 1], "审计单位" );
            setCellValue( ws.Cells[7, 1], "审计意见类型" );
            int row = 8;
            foreach ( var s in s2019s ) {
                foreach ( var ss in s.Sections ) {
                    setCellValue( ws.Cells[row, 1], $"{ss.Name}" );
                    setCellValue( ws.Cells[row, 2], $"{ss.Value10S}" );
                    setCellValue( ws.Cells[row, 3], $"{ss.Value10SByYear( "2020" )}" );
                    setCellValue( ws.Cells[row, 4], $"{ss.Value10SByYear( "2021" )}" );
                    setCellValue( ws.Cells[row, 6], $"{ss.ValueIncreasmentRate( "2020" )}" );
                    setCellValue( ws.Cells[row, 7], $"{ss.ValueIncreasmentRate( "2021" )}" );
                    row++;
                }
            }
            // eWSheet.Range[eWSheet.Cells[1, 1], eWSheet.Cells[4, 1]].Merge(); 合并单元格
            // sheet.Cells[rowCount, column].Formula = string.Format("=SUM(G1:G{0})", rowCount); 公式
            // 自适应
            ws.Cells.AutoFit();
            xlWorkBook.SaveCopyAs( $@"C:\Users\cyf-thinkpad\Documents\test.xlsx" );
            Process.Start( $@"C:\Users\cyf-thinkpad\Documents\test.xlsx" );
            #endregion

            File.WriteAllText( "./fal_final_res.txt", res );
            File.WriteAllText( "./fal_final.json", a );
        }
        static public void mergeCell( Worksheet ws, int r1, int c1, int r2, int c2 ) {
            ws.Range[ws.Cells[r1, c1], ws.Cells[r2, c2]].Merge();
        }
        static public void setCellValue( Range c, string v, bool isBold = false, XlHAlign align = XlHAlign.xlHAlignLeft, bool isBordered = false ) {
            c.Value = v;
            c.Font.Bold = isBold;
            c.HorizontalAlignment = align;
            // c.Style = "Comma";
        }
    }
}

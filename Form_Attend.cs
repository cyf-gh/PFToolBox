using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Attend : Form {
        public Form_Attend()
        {
            InitializeComponent();
        }
        IExcelDataReader excelDataReader = null;
        FileStream stream = null;

        class Row {
            public double[] D { get; set; } = new double[9];
            public double[] AVG { get; set; } = new double[4];
        }
        int ColI( char col )
        {
            return col - 'A';
        }
        int ToInt( string s )
        {
            int res = 0;
            int.TryParse( s, out res );
            return res;
        }
        double min( double a )
        {
            return a == 0 ? 1 : a;
        } 
        int HHMM2H( string a ) {
            if ( a == "0" ) {
                return 0;
            }
            var p = new Regex( @"(?<h>[\S]+)小时(?<m>[\S])+分钟" );
            Match m = p.Match( a );
            if ( m.Success ) {
                return Convert.ToInt32( m.Groups["h"].Value ) * 60 + Convert.ToInt32( m.Groups["m"].Value );
            }
            return 0;
        }
        private void button1_Click( Object sender, EventArgs e )
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;

                if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
                    filePath = openFileDialog.FileName;
                    if ( checkBox1.Checked ) {
                        filePath += ".copy.xlsx";
                        if ( File.Exists( filePath ) ) {
                            File.Delete( filePath );
                        }
                        File.Copy( openFileDialog.FileName, filePath );
                    }
                    try {
                        stream = File.Open( filePath, FileMode.Open, FileAccess.Read );
                    } catch ( Exception ex ) {
                        MessageBox.Show( $"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试" );
                        return;
                    }
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    excelDataReader = ExcelReaderFactory.CreateReader( stream );
                    // 2. Use the AsDataSet extension method
                    // var result = excelDataReader.AsDataSet();
                    // The result of each spreadsheet is in result.Tables
                    var result = excelDataReader.AsDataSet();
                    var ts = result.Tables;

                    var dic = new Dictionary<string, Row>();
                    for ( int ii = 0; ii < ts[0].Rows.Count; ii++ ) {
                        var d = ts[0].Rows[ii];
                        var r = d.ItemArray;
                        // group
                        var key = r[ColI( 'B' )].ToString();
                        if ( !dic.ContainsKey( key ) ) {
                            dic[key] = new Row();
                        }
                        dic[key].D[0]++; // 人数

                        var d1 = ToInt( r[ColI( 'F' )].ToString() );
                        dic[key].D[1] += d1; // 平均值项:出勤天数
                        dic[key].AVG[1] += d1 == 0 ? 0 : 1;


                        var d3 = HHMM2H( r[ColI( 'H' )].ToString() );
                        dic[key].D[2] += d1 == 0 ? 0 : ( d3 / d1 / 60 ); // 平均值项:工作时长(小时)
                        dic[key].AVG[2] += d3 == 0 ? 0 : 1;

                        dic[key].D[3] += ToInt( r[ColI( 'T' )].ToString() ); // 求和项:外出时长
                        dic[key].D[4] += ToInt( r[ColI( 'I' )].ToString() ); // 求和项:迟到次数
                        dic[key].D[5] += ToInt( r[ColI( 'N' )].ToString() ); // 求和项:早退次数
                        dic[key].D[6] += ToInt( r[ColI( 'P' )].ToString() ); // 求和项:上班缺卡次数
                        dic[key].D[7] += ToInt( r[ColI( 'Q' )].ToString() ); // 求和项:下班缺卡次数
                        dic[key].D[8] += ToInt( r[ColI( 'R' )].ToString() ); // 求和项: 旷工天数
                    }
                    dic.Remove( "" );
                    dic.Remove( "考勤组" );
                    dic.Remove( "未加入考勤组" );
                    foreach ( var d in dic ) {
                        d.Value.D[1] /= min( d.Value.AVG[1] );
                        d.Value.D[2] /= min( d.Value.AVG[2] );
                    }
                    /*
                     列      列数据

                     G      出勤天数
                     I      工作时长(小时)
                     U      外出时长
                     J      迟到次数
                     O      早退次数
                     Q      上班缺卡次数
                     R      下班缺卡次数
                     S      旷工天数

                     */
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    var xlWorkBook = xlApp.Workbooks.Add();
                    var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 1 );

                    xlWorkSheet.Cells[1, 1] = "行标签";
                    xlWorkSheet.Cells[1, 2] = "计数项:姓名";
                    xlWorkSheet.Cells[1, 3] = "平均值项:出勤天数";
                    xlWorkSheet.Cells[1, 4] = "平均值项:工作时长(小时)";
                    xlWorkSheet.Cells[1, 5] = "求和项:外出时长";
                    xlWorkSheet.Cells[1, 6] = "求和项:迟到次数";
                    xlWorkSheet.Cells[1, 7] = "求和项:早退次数";
                    xlWorkSheet.Cells[1, 8] = "求和项:上班缺卡次数";
                    xlWorkSheet.Cells[1, 9] = "求和项:下班缺卡次数";
                    xlWorkSheet.Cells[1, 10] = "求和项:旷工天数";

                    
                    var sums = new double[11];

                    int i = 2;
                    foreach ( var d in dic ) {
                        xlWorkSheet.Cells[i, 1] = d.Key;
                        for ( int j = 0; j < d.Value.D.Length; j++ ) {
                            xlWorkSheet.Cells[i, j + 2] = d.Value.D[j];
                            sums[j] += d.Value.D[j];
                        }
                        i++;
                    }
                    int @is = 2;

                    sums[1] = sums[1] / dic.Count;
                    sums[2] = sums[2] / dic.Count;

                    xlWorkSheet.Cells[i, 1] = "合计";
                    foreach ( var s in sums ) {
                        xlWorkSheet.Cells[i, @is] = s;
                        @is++;
                    }
                    var now = DateTime.Now.ToString( "MM-dd-yyyy-HH-mm-ss" );
                    var ff = Path.Combine( Environment.CurrentDirectory, $"最终结果{now}.xlsx" );
                    SAVE:
                    try {
                        xlWorkBook.SaveCopyAs( ff );
                    } catch ( Exception ex ) {
                        MessageBox.Show($"{ex.Message}\n please retry");
                        goto SAVE;
                        throw;
                    }
                    Process.Start( ff );
                }
            }
        }

        private void FormAttend_Load( Object sender, EventArgs e )
        {
            label1.Text = @"
                    为保证最终数据的正确，请保证列与列数据的正确性

                     列      列数据

                     F      出勤天数
                     T      外出时长
                     H      工作时长(小时)
                     I      迟到次数
                     N      早退次数
                     P      上班缺卡次数
                     Q      下班缺卡次数
                     R      旷工天数";
        }
    }
}

using ExcelDataReader;

using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Attend : Form {
        public Form_Attend() {
            InitializeComponent();
        }
        IExcelDataReader excelDataReader = null;
        FileStream stream = null;

        class Row {
            public double[] D { get; set; } = new double[9];
            public double[] AVG { get; set; } = new double[4];
        }
        int ColI( char col ) {
            return col - 'A';
        }
        int ColI2( int col ) {
            return col - 'A';
        }
        int ToInt( string s ) {
            int res = 0;
            int.TryParse( s, out res );
            return res;
        }
        double min( double a ) {
            return a == 0 ? 1 : a;
        }
        int HHMM2H( string a ) {
            if ( string.IsNullOrEmpty( a ) ) {
                return 0;
            }
            if ( a == "0" ) {
                return 0;
            }
            var p = new Regex( @"(?<h>[\S]+)小时(?<m>[\S]+)分钟" );
            Match m = p.Match( a );
            if ( m.Success ) {
                return Convert.ToInt32( m.Groups["h"].Value ) * 60 + Convert.ToInt32( m.Groups["m"].Value );
            } else {
                try {
                    return Convert.ToInt32( a );
                } catch ( Exception ) {
                    return 0;
                }
            }
            return 0;
        }
        struct OT {
            public int org, target;
        }
        private void button1_Click( Object sender, EventArgs e ) {
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


                    var r34Dict = new Dictionary<string, int>();
                    // 读取第一第二行，判断列的正确性并提示用户
                    var r3 = ts[0].Rows[2].ItemArray;
                    for ( int iii = 0; iii < r3.Length; iii++ ) {
                        var r = r3[iii];
                        r34Dict[r.ToString()] = iii + 'A';
                    }
                    var r4 = ts[0].Rows[3].ItemArray;
                    for ( int iii = 0; iii < r4.Length; iii++ ) {
                        var r = r4[iii];
                        r34Dict[r.ToString()] = iii + 'A';
                    }

                    var diffDict = new Dictionary<string, OT>();

                    foreach ( var od in orgDict ) {
                        foreach ( var r34 in r34Dict ) {
                            if ( r34.Key == od.Key ) {
                                if ( od.Value != r34.Value ) {
                                    var ot = new OT() { org = od.Value, target = r34.Value };

                                    diffDict.Add( od.Key, ot );
                                }
                            }
                        }
                    }
                    string t = $"发现 {diffDict.Count} 个与默认列号不一致的列\n";
                    foreach ( var dd in diffDict ) {
                        t += $"{dd.Key} -> 原本是 {(char)dd.Value.org} 但目标 Excel表格 的列可能是 {(char)dd.Value.target}\n";

                    }
                    t += $"按确定使用检测到的列进行生成，点击取消使用原始列";
                    if ( DialogResult.OK == MessageBox.Show( t, "请注意", MessageBoxButtons.OKCancel ) ) {
                        foreach ( var dd in diffDict ) {
                            orgDict[dd.Key] = (char)dd.Value.target;
                        }
                    } else {
                        return;
                    }

                    for ( int ii = 0; ii < ts[0].Rows.Count; ii++ ) {
                        var d = ts[0].Rows[ii];
                        var r = d.ItemArray;
                        // group
                        var key = r[ColI( 'B' )].ToString();
                        if ( !dic.ContainsKey( key ) ) {
                            dic[key] = new Row();
                        }
                        dic[key].D[0]++; // 人数

                        var attendDay = ToInt( r[ColI2( orgDict["出勤天数"] )].ToString() );
                        dic[key].D[1] += attendDay; // 平均值项:出勤天数
                        dic[key].AVG[1] += attendDay == 0 ? 0 : 1;

                        var outT = ToInt( r[ColI2( orgDict["外出"] )].ToString() ); // 求和项:外出时长
                        // var outDay = outT / 8; // 外出固定为8小时

                        var workD = HHMM2H( r[ColI2( orgDict["工作时长"] )].ToString() );
                        Console.WriteLine( $"{r[ColI2( orgDict["工作时长"] )]} == {workD}" );
                        var workDay = attendDay;
                        dic[key].D[2] += workDay == 0 ? 0 : ( workD / workDay / 60 ); // 平均值项:工作时长(小时) 工作时长包括出勤时长
                        dic[key].AVG[2] += workD == 0 ? 0 : 1;


                        dic[key].D[3] += outT;
                        dic[key].D[4] += ToInt( r[ColI2( orgDict["迟到次数"] )].ToString() ); // 求和项:迟到次数
                        dic[key].D[5] += ToInt( r[ColI2( orgDict["早退次数"] )].ToString() ); // 求和项:早退次数
                        dic[key].D[6] += ToInt( r[ColI2( orgDict["上班缺卡次数"] )].ToString() ); // 求和项:上班缺卡次数
                        dic[key].D[7] += ToInt( r[ColI2( orgDict["下班缺卡次数"] )].ToString() ); // 求和项:下班缺卡次数
                        dic[key].D[8] += ToInt( r[ColI2( orgDict["旷工天数"] )].ToString() ); // 求和项: 旷工天数
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
                        MessageBox.Show( $"{ex.Message}\n please retry" );
                        goto SAVE;
                        throw;
                    }
                    Process.Start( ff );
                }
            }
        }
        Dictionary<string, char> orgDict = new Dictionary<string, char>();
        private void FormAttend_Load( Object sender, EventArgs e ) {
            orgDict = JsonConvert.DeserializeObject<Dictionary<string, char>>( File.ReadAllText( "./c.json" ) );

            label1.Text = @"
        为保证最终数据的正确，请保证列与列数据的正确性

        列      列数据

        " + Environment.NewLine;
            foreach ( var or in orgDict ) {
                label1.Text += $"        {or.Value}        {or.Key}{Environment.NewLine}";
            }
        }

        private void button2_Click( object sender, EventArgs e ) {
            Process.Start( Path.Combine( Application.StartupPath, "c.json" ) );
        }
    }
}

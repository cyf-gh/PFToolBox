using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_TeenScore : Form {
        public Form_TeenScore() {
            InitializeComponent();
        }
        int ColI( char col ) {
            return col - 'A';
        }

        class Worker {
            public string g; // 条线
            public List<string> Bs = new List<string>(); // 
        }

        class Recommend {
            public string BeRWorkerName;
            public string Description;
        }
        string co = "（2022年1季度）";
        string filePath;
        FileStream stream;
        IExcelDataReader excelDataReader = null;
        private void Form_TeenScore_Load( object sender, EventArgs e ) {
            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;

                filePath = @"C:\Users\cyf-desktop\Desktop\158059950_2_青年员工积分上报（2022一季度）_42_38.xlsx";
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
                excelDataReader = ExcelReaderFactory.CreateReader( stream );
                // 2. Use the AsDataSet extension method
                // var result = excelDataReader.AsDataSet();
                // The result of each spreadsheet is in result.Tables
                var result = excelDataReader.AsDataSet();
                var t = result.Tables;
                var defaultT = t[0];

                var Heads = new List<string>();

                {
                    var head = defaultT.Rows[0];
                    var r = head.ItemArray;
                    for ( int i = ColI( 'I' ); i < ColI( 'Y' ); i++ ) {
                        Heads.Add( r[i].ToString() );
                    }
                }

                var rec = new Dictionary<string, List<Recommend>>();
                var dic = new Dictionary<string, Worker>();
                for ( int j = 1; j < defaultT.Rows.Count; j++ ) {
                    var d = defaultT.Rows[j];
                    var r = d.ItemArray;
                    var name = ( r[ColI( 'G' )].ToString() ); // 姓名
                    var w = dic.ContainsKey( name ) ? dic[name] : new Worker();
                    if ( !dic.ContainsKey( name ) ) {
                        dic[name] = w;
                    }
                    w.g = r[ColI( 'H' )].ToString(); // 条线
                                                     // 自我打分
                    for ( int i = ColI( 'I' ), ii = 0; i < ColI( 'Y' ); i++ ) {
                        var content = r[i].ToString();
                        var contentHead = Heads[i - ColI( 'I' )].ToString();
                        if ( contentHead.Contains( "(1)" ) ) {
                            if ( content == "(空)" ) {
                                ++i;
                                continue;
                            } else {
                                w.Bs.Add( $"{contentHead.Replace( "___", $"[{content}]\t" ).Replace( "\n", "" )}" );
                                ++ii;
                            }
                        } else if ( contentHead.Contains( "(2)" ) ) {
                            var p = new Regex( @"（(?<min>[\S]+)分-(?<max>[\S]+)分）" );
                            Match m = p.Match( contentHead );
                            string min = "0", max = "100";
                            if ( m.Success ) {
                                min = m.Groups["min"].Value;
                                max = m.Groups["max"].Value;
                            }
                            w.Bs[w.Bs.Count - 1] += $"自评描述：[{content.Replace( "\n", "" )}] ~{name}~";
                            w.Bs[w.Bs.Count - 1] = w.Bs[w.Bs.Count - 1].Replace( "分值填写：", "\t自评分：" );
                            w.Bs[w.Bs.Count - 1] = w.Bs[w.Bs.Count - 1].Replace( "(1)", "" );
                            w.Bs[w.Bs.Count - 1] += $"[单行文本题](分值范围：{min}~{max})\n\n";
                        }
                    }
                    var recrr = rec.ContainsKey( name ) ? rec[name] : new List<Recommend>();
                    if ( !rec.ContainsKey( name ) ) {
                        rec[name] = recrr;
                    }
                    var recName = r[ColI( 'Y' )].ToString();
                    var recDesc = r[ColI( 'Z' )].ToString();
                    if ( recName == "(空)" ) {
                        continue;
                    } else {
                        recrr.Add( new Recommend { BeRWorkerName = recName, Description = recDesc } );
                    }
                }

                var ggg = new Dictionary<string, string>();
                ggg.Add( "1", "运营条线条线负责人打分统计" + co );
                ggg.Add( "2", "零售条线条线负责人打分统计" + co );
                ggg.Add( "3", "公司条线条线负责人打分统计" + co );
                ggg.Add( "4", "风险条线条线负责人打分统计" + co );
                ggg.Add( "5", "办公室条线负责人打分统计" + co );

                var gggstr = new Dictionary<string, string>();
                foreach ( var d in dic ) {

                    string str = "";
                    if ( gggstr.ContainsKey( ggg[d.Value.g] ) ) {
                        str = gggstr[ggg[d.Value.g]];
                    } else {
                        gggstr[ggg[d.Value.g]] = str;
                    }
                    str += $"\n\n{d.Key}[段落说明]\n";
                    foreach ( var des in d.Value.Bs ) {
                        str += $"{des}";
                    }
                    var rrr = rec[d.Key];
                    string fuckrrr = "";
                    foreach ( var r in rrr ) {
                        fuckrrr += $"{r.Description}[单行文本题]({r.BeRWorkerName})\n";
                    }
                    str += $"条线打分理由:[多行文本题]\n推荐人打分[段落说明]\n如有多位员工姓名，则每个人都会获得相同的打分值[段落说明]\n{fuckrrr}\n===分页===";
                    gggstr[ggg[d.Value.g]] = str;
                }
                foreach ( var ggggggg in gggstr ) {
                    string sum = "";
                    sum += $"{ggggggg.Key} \n\n{ggggggg.Value}";
                    File.WriteAllText( $"./{ggggggg.Key}.txt", sum );
                }
                Console.WriteLine( dic );
                Application.Exit();
            }
        }
    }
}

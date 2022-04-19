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
            public string name;
            public string g; // 条线
            public string Content;
            public string ZPScore;
            public string ZPContent;
            public string EMContent;
        }

        class Recommend {
            public string RName;
            public string g;
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

                var rec = new List<Recommend>();
                var dic = new List<Worker>();
                for ( int j = 1; j < defaultT.Rows.Count; j++ ) {
                    var d = defaultT.Rows[j];
                    var r = d.ItemArray;
                    var name = ( r[ColI( 'G' )].ToString() ); // 姓名
                    var w = new Worker();
                    string g = r[ColI( 'H' )].ToString(); // 条线
                                                          // 自我打分
                    for ( int i = ColI( 'I' ), ii = 0; i < ColI( 'Y' ); i++ ) {
                        var content = r[i].ToString();
                        var contentHead = Heads[i - ColI( 'I' )].ToString();
                        if ( contentHead.Contains( "(1)" ) ) {
                            if ( content == "(空)" ) {
                                ++i;
                                continue;
                            } else {
                                w = new Worker();
                                w.g = g;
                                w.name = name;
                                w.ZPScore = content;
                                w.Content = contentHead.Replace( "分值填写：___", "" ).Replace( "(1)", "" );
                                // w.Bs.Add( $"{contentHead.Replace( "___", $"[{content}]\t" ).Replace( "\n", "" )}" );
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
                            w.ZPContent = content;
                            w.Content += $"({min}~{max})";
                            dic.Add( new Worker {
                                name = w.name,
                                Content = w.Content,
                                g = g,
                                ZPScore = w.ZPScore,
                                ZPContent = w.ZPContent,
                                EMContent = w.EMContent
                            } );
                        }
                    }
                    var recName = r[ColI( 'Y' )].ToString();
                    var recDesc = r[ColI( 'Z' )].ToString();
                    if ( recName == "(空)" ) {
                        continue;
                    }
                    rec.Add( new Recommend { BeRWorkerName = recName, Description = recDesc, RName = name, g=g } );
                }

                var ggg = new Dictionary<string, string>();
                ggg.Add( "1", "运营条线条线负责人打分统计" + co );
                ggg.Add( "2", "零售条线条线负责人打分统计" + co );
                ggg.Add( "3", "公司条线条线负责人打分统计" + co );
                ggg.Add( "4", "风险条线条线负责人打分统计" + co );
                ggg.Add( "5", "办公室条线负责人打分统计" + co );

                foreach ( var gg in ggg ) {
                    var allw = dic.FindAll( a => { return a.g == gg.Key; } );
                    if ( allw.Count == 0 ) {
                        continue;
                    }
                    SaveFuck( ggg[allw[0].g], allw, rec.FindAll( a => { return a.g == gg.Key; } ) );
                }


                Application.Exit();
            }

            void SaveFuck( string name, List<Worker> fff, List<Recommend> rrs ) {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                var xlWorkBook = xlApp.Workbooks.Add();
                xlWorkBook.Sheets.Add( After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count] );
                var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 1 );
                var xlWorkSheetReco = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 2 );

                xlWorkSheet.Cells[1, 1] = "姓名";
                xlWorkSheet.Cells[1, 2] = "项目";
                xlWorkSheet.Cells[1, 3] = "自评分";
                xlWorkSheet.Cells[1, 4] = "项目内容";
                xlWorkSheet.Cells[1, 5] = "条线打分";
                xlWorkSheet.Cells[1, 6] = "打分理由";
                xlWorkSheetReco.Cells[1, 1] = "推荐人";
                xlWorkSheetReco.Cells[1, 2] = "被推荐人";
                xlWorkSheetReco.Cells[1, 3] = "推荐理由";
                xlWorkSheetReco.Cells[1, 4] = "条线打分";
                xlWorkSheetReco.Cells[1, 5] = "打分理由";

                for ( int i = 0; i < fff.Count; i++ ) {
                    var f = fff[i];
                    xlWorkSheet.Cells[i + 2, 1] = f.name;
                    xlWorkSheet.Cells[i + 2, 2] = f.Content;
                    xlWorkSheet.Cells[i + 2, 3] = f.ZPScore;
                    xlWorkSheet.Cells[i + 2, 4] = f.ZPContent;
                }
                for ( int i = 0; i < rrs.Count; i++ ) {
                    var r = rrs[i];
                    xlWorkSheetReco.Cells[i + 2, 1] = r.RName;
                    xlWorkSheetReco.Cells[i + 2, 2] = r.BeRWorkerName;
                    xlWorkSheetReco.Cells[i + 2, 3] = r.Description;
                }
                File.Delete( "./{name}.xlsx" );
                xlWorkBook.SaveCopyAs( $"./{name}.xlsx" );
            }
        }
    }
}

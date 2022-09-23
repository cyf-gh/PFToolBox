using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace MergeExcel {
    public partial class Form_Merge_FinaAnaly : Form {
        static public Dictionary<string, string> Aliases = new Dictionary<string, string>();
        public Form_Merge_FinaAnaly() {
            Aliases[" 一、营业总收入"] = "营业收入";
            Aliases["减：营业成本"] = "营业成本";
            Aliases["负债和所有者权益（或股东权益）总计"] = "负债和所有者权益总计";
            Aliases["所有者权益（或股东权益）合计"] = "所有者权益合计";
            InitializeComponent();
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
        public class Report {
            /// <summary>
            /// 资产负债表采用
            /// </summary>
            public TreeNode BalanceSheet;
            public TreeNode IncomeSheet;
            // public FinanState.StateSection 
            public Report() {
                BalanceSheet = new TreeNode( "资产负债表" );
                var ass = BalanceSheet.Nodes.Add( "资产总计" );
                ass.Nodes.Add( "流动资产合计" );
                ass.Nodes.Add( "非流动资产合计" );

                var liaProp = BalanceSheet.Nodes.Add( "负债和所有者权益（或股东权益）总计" );
                var lia = liaProp.Nodes.Add( "负债合计" );
                lia.Nodes.Add( "流动负债合计" );
                lia.Nodes.Add( "非流动负债合计" );
                liaProp.Nodes.Add( "所有者权益（或股东权益）合计" );
            }
            /// <summary>
            /// 年报年份
            /// </summary>
            public string Year { get; set; }
            /// <summary>
            /// 非即期月报此项为空
            /// </summary>
            public string Month { get; set; }
            public bool IsAnnualReport { get { return string.IsNullOrEmpty( Month ); } }
            /// <summary>
            /// 报表列表
            /// </summary>
            public List<FinanState> FinanStates { get; set; } = new List<FinanState>();

            public FinanState GetBS() {
                return FinanStates.Find( m => m.Name == "资产负债表" );
            }

            public FinanState GetIC() {
                return FinanStates.Find( m => m.Name == "损益表" );
            }

            public FinanState GetCA() {
                return FinanStates.Find( m => m.Name == "现金流量表" );
            }

            /// <summary>
            /// 报表 分为 资产负债 利润 现金流量
            /// </summary>
            public class FinanState {
                public class ViewModel {
                    public string Name { get; set; } = string.Empty;
                    public double Value { get; set; }
                    public bool IsBold { get; set; } = false;
                    public string Fomula { get; set; } = string.Empty;
                    public bool Visiable { get; set; } = true;
                    public string Alias { get; set; } = string.Empty;
                }
                string removeWhitespace( string str ) {
                    if ( str == null ) {
                        return String.Empty;
                    }
                    return Regex.Replace( str, @"\s+", "" );
                }
                public Dictionary<string, ViewModel> BSS = new Dictionary<string, ViewModel>();
                public List<int> Fl_RigidLia { get; set; } = new List<int>();
                public void PrintBS( Worksheet ws, TreeNode root, ref int row, int col = 2 ) {
                    foreach ( TreeNode node in root.Nodes ) {
                        var section = Sections.Find( m => removeWhitespace( m.SumName ) == node.Text );
                        section = section == null ? Sections.Find( m => removeWhitespace( m.SumName ) == Aliases[node.Text] ) : section;
                        // BSS.Add( node.Text, new ViewModel { Name = node.Text, IsBold = true, Value = ( section.SumValue / 10000 ) } ); 
                        Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, 1], node.Text, true );
                        Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, col], ( section.SumValue / 10000 ).ToString(), true );
                        row++;
                        foreach ( var kv in section.States ) {
                            // BSS.Add( kv.Key, new ViewModel { Name = removeWhitespace( kv.Key ), IsBold = false, Value = ( kv.Value / 10000 ) } );
                            Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, 1], procRigidLia( removeWhitespace( kv.Key ), row ) );
                            Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, col], ( kv.Value / 10000 ).ToString(), false, XlHAlign.xlHAlignCenter );
                            row++;
                        }
                        if ( node.Text == "非流动负债合计" ) {
                            setCellValue( ws.Cells[row, 1], "刚性负债合计＝①＋②＋③＋④＋⑤" );
                            ws.Cells[row, col].Formula = Formula_RigidLia( col );
                            row++;
                        }
                        PrintBS( ws, node, ref row, col );
                    }
                }
                public void PrintIC( Worksheet ws, TreeNode root, ref int row, int col = 2 ) {
                    //foreach ( TreeNode node in root.Nodes ) {
                    //    var section = Sections.Find( m => removeWhitespace( m.SumName ) == node.Text );
                    //    Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, 1], node.Text, true );
                    //    Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, col], ( section.SumValue / 10000 ).ToString(), true );
                    //    row++;
                    //    foreach ( var kv in section.States ) {
                    //        Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, 1], removeWhitespace( kv.Key ) );
                    //        Form_Merge_FinaAnaly.setCellValue( ws.Cells[row, col], ( kv.Value / 10000 ).ToString(), false, XlHAlign.xlHAlignCenter );
                    //        row++;
                    //    }
                    //    PrintIC( ws, node, ref row, col );
                    //}
                }
                public string Formula_RigidLia( int col ) {
                    string r = "=";
                    string C = Convert.ToString( System.Text.Encoding.ASCII.GetString( new byte[1] { (byte)Convert.ToInt32( col + 64 ) } ) );
                    foreach ( var c in Fl_RigidLia ) {
                        r += $"{C}{c}+";
                    }
                    r += "0";
                    return r;
                }
                public string procRigidLia( string name, int col ) {
                    switch ( name ) {
                        case "短期借款":
                            name += "①";
                            break;
                        case "应付票据":
                            name += "②";
                            break;
                        case "一年内到期的非流动负债":
                            name += "③";
                            break;
                        case "长期借款":
                            name += "④";
                            break;
                        case "应付债券":
                            name += "⑤";
                            break;
                        default:
                            return name;
                    }
                    Fl_RigidLia.Add( col );
                    return name;
                }
                public readonly string MaskName;
                public StateSection GetSection( string name ) {
                    return Sections.Find( m => m.Name.Contains( name ) );
                }
                public FinanState ProcIC( DataRowCollection rows, int sNameIndex = 0 ) {
                    var s = new StateSection();
                    int sValueIndex = sNameIndex + 3;

                    for ( int i = 2; i < rows.Count; i++ ) {
                        var r = rows[i].ItemArray;
                        bool isValueNull = r[sValueIndex] is DBNull;
                        var Value = r[sValueIndex] is DBNull ? 0 : double.Parse( r[sValueIndex] as string );
                        var Name = r[sNameIndex] as string;
                        Name = Regex.Replace( Name, @"\s+", "" );
                        if ( s.TrySetName( Name, Value, ref Sections ) ) { s = Sections.Last(); continue; }
                        s.States[Name] = Value;
                    }
                    return this;
                }
                /// <summary>
                /// 处理资产负债表
                /// </summary>
                /// <param name="rows">Rows数据</param>
                /// <param name="sNameIndex">科目名称列，数据列汇由这列偏移3</param>
                /// <returns></returns>
                public FinanState ProcAL( DataRowCollection rows, int sNameIndex = 0 ) {
                    // var fs = new Report.FinanState();
                    var s = new StateSection();
                    int sValueIndex = sNameIndex + 3;
                    for ( int i = 2; i < rows.Count; i++ ) {
                        var r = rows[i].ItemArray;
                        bool isValueNull = r[sValueIndex] is DBNull;

                        // 资产：
                        var Name = r[sNameIndex] as string;
                        if ( string.IsNullOrEmpty( Name ) ) {
                            continue;
                        }
                        // 添加子科目名,
                        if ( isValueNull ) {
                            if ( s.TrySetName( Name, 0, ref Sections ) ) { s = Sections.Last(); continue; }
                        } else {
                            // 添加总和
                            var Value = double.Parse( r[sValueIndex] as string );
                            if ( Name.Contains( "总计" ) && Name.Contains( "资产" ) ) {
                                SumAss = Value; // 资产总计
                            } else {
                                SumLia = Value; // 负债及所有者权益总计
                            }
                            var res = s.TryEndSection( Name, Value, ref Sections );
                            if ( res == null ) {
                                s.States[Name] = Value;
                            } else {
                                s = res;
                            }
                        }
                    }
                    return this;
                }
                /// <summary>
                /// 处理现金流量表
                /// </summary>
                /// <param name="rows"></param>
                /// <returns></returns>
                public FinanState ProcCa( DataRowCollection rows ) {
                    var s = new StateSection();
                    const int sValueIndex = 2, sNameIndex = 0;
                    for ( int i = 2; i < rows.Count; i++ ) {
                        var r = rows[i];
                        var Name = r[sNameIndex] as string;
                        var Value = r[sValueIndex] is DBNull ? 0 : double.Parse( r[sValueIndex] as string );
                        if ( s.TrySetName( Name, Value, ref Sections ) ) { s = Sections.Last(); continue; }
                        var res = s.TryEndSection( Name, Value, ref Sections );
                        if ( res == null ) {
                            s.States[Name] = Value;
                        } else {
                            s = res;
                        }
                    }
                    return this;
                }
                /// <summary>
                ///  资产总计
                /// </summary>
                public double SumAss { get; set; }
                /// <summary>
                /// 负债和所有者权益（或股东权益）总计
                /// </summary>
                public double SumLia { get; set; }
                /// <summary>
                /// 资产是否等于负债+所有者权益
                /// </summary>
                public bool IsLRBalanced { get { return SumAss - SumLia < 0.01; } }
                public string Name { get; set; } = String.Empty;
                public List<StateSection> Sections = new List<StateSection>();
                /// <summary>
                /// 科目块
                /// </summary>
                public class StateSection {
                    public StateSection TryEndSection( string name, double v, ref List<StateSection> sections ) {
                        var sectionEndMarks = new string[] { "小计", "合计", "流量净额", "总计" };
                        foreach ( var sem in sectionEndMarks ) {
                            if ( name.Contains( sem ) ) {
                                SumName = name;
                                SumValue = v;
                                if ( !sections.Contains( this ) ) {
                                    sections.Add( this );
                                }
                                return new StateSection() { Name = this.Name };
                            }
                        }
                        return null;
                    }
                    /// <summary>
                    /// 尝试设置科目块的名字
                    /// </summary>
                    /// <param name="name"></param>
                    /// <param name="s"></param>
                    /// <returns>true表示是科目块标题</returns>
                    public bool TrySetName( string name, double Value, ref List<StateSection> sections ) {
                        // Console.WriteLine( Name );
                        var sectionMarks = new char[] { '：', ':', '、' };
                        foreach ( var sm in sectionMarks ) {
                            if ( name.Contains( sm ) ) {
                                switch ( sm ) {
                                    case '、':
                                        if ( name.IndexOf( sm ) >= 3 ) {
                                            return false;
                                        }
                                        break;
                                }
                                SumValue = Value;
                                // 单行科目单独添加
                                if ( Name != String.Empty ) {
                                    sections.Add( new StateSection() { Name = name } );
                                    return true;
                                }
                                Name = name;
                                return true;
                            }
                        }
                        return false;
                    }
                    /// <summary>
                    /// XXX 合计
                    /// </summary>
                    public string Name { get; set; }
                    public double SumValue { get; set; }
                    public string SumName { get; set; }
                    public Dictionary<string, double> States { get; set; } = new Dictionary<string, double>();
                    private double statesSum() {
                        double fSum = 0;
                        foreach ( var s in States ) {
                            fSum = s.Key.Contains( "减" ) ? fSum - s.Value : fSum + s.Value;
                        }
                        return fSum;
                    }
                    public double StatesSum { get { return statesSum(); } }
                    /// <summary>
                    /// 该科目是否轧平
                    /// </summary>
                    public bool IsAccountFinished { get { return Math.Abs( StatesSum - SumValue ) < 0.01; } } // double误差精确到分以下
                }
            }
        }

        public string filePath { get; set; }
        FileStream stream;
        IExcelDataReader reader;
        private void Form_Merge_FinaAnaly_Load( object sender, EventArgs e ) {
            var ys = new List<Report>();
            for ( int ii = 0; ii < 4; ii++ ) {
                var y = new Report() { Year = ( 2019 + ii ).ToString() };
                for ( int i = 1; i < 4; i++ ) {
                    using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                        #region
                        openFileDialog.InitialDirectory = "c:\\";
                        openFileDialog.RestoreDirectory = true;

                        filePath = $@"C:\Users\cyf-thinkpad\Desktop\1\{2019 + ii}\{2019 + ii} ({i}).xlsx";
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
                        var Rows = defaultT.Rows;
                        #endregion
                        var r1 = Rows[0].ItemArray;
                        var fs = new Report.FinanState();
                        fs.Name = r1[0] as string;
                        switch ( fs.Name ) {
                            case "资产负债表":
                                fs.ProcAL( Rows )     // 处理资产部分 
                                  .ProcAL( Rows, 4 ); // 处理负债部分
                                break;
                            case "现金流量表":
                                fs.ProcCa( Rows );
                                break;
                            case "损益表":
                                fs.ProcIC( Rows );
                                break;
                            default:
                                break;
                        }
                        y.FinanStates.Add( fs );
                    }
                }
                ys.Add( y );
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
            int j = 2;
            foreach ( var y in ys ) {
                int row = 8;
                setCellValue( ws.Cells[3, j], $"{y.Year}年12月" );
                var bs = y.FinanStates.Find( m => m.Name == "资产负债表" );
                bs.PrintBS( ws, y.BalanceSheet, ref row, j );
                // bs.PrintIC( ws, y.BalanceSheet, ref row, j );
                var a = bs.BSS.Count;
                ++j;
            }
            // tv_zcfz.Nodes.Add( y.BalanceSheet );
            // 

            // eWSheet.Range[eWSheet.Cells[1, 1], eWSheet.Cells[4, 1]].Merge(); 合并单元格
            // sheet.Cells[rowCount, column].Formula = string.Format("=SUM(G1:G{0})", rowCount); 公式
            // 自适应
            ws.Cells.AutoFit();
            xlWorkBook.SaveCopyAs( $"./test.xlsx" );
            Process.Start( $@"C:\Users\cyf-thinkpad\Documents\test.xlsx" );
            #endregion
        }
    }
}

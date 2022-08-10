using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Merge_FinaAnaly : Form {
        public Form_Merge_FinaAnaly() {
            InitializeComponent();
        }

        public class Report {
            /// <summary>
            /// 资产负债表采用
            /// </summary>
            public TreeNode BalanceSheet { get; set; }
            // public FinanState.StateSection 
            public Report() {
                BalanceSheet = new TreeNode( "资产负债表" );
                var ass = BalanceSheet.Nodes.Add( "资产总计" );
                ass.Nodes.Add( "流动资产合计" );
                ass.Nodes.Add( "非流动资产合计" );

                var liaProp = BalanceSheet.Nodes.Add( "负债和所有者权益（或股东权益）总计" );
                liaProp.Nodes.Add( "所有者权益（或股东权益）合计" );
                var lia = liaProp.Nodes.Add( "负债合计" );
                lia.Nodes.Add( "非流动负债合计" );
                lia.Nodes.Add( "流动负债合计" );
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
            /// <summary>
            /// 报表 分为 资产负债 利润 现金流量
            /// </summary>
            public class FinanState {
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
                        Console.WriteLine( Name );
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
            var y = new Report();
            for ( int i = 1; i < 4; i++ ) {
                using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                    #region
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.RestoreDirectory = true;

                    filePath = $@"C:\Users\cyf-thinkpad\Desktop\授信财务报表\2021 ({i}).xlsx";
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
                    var ys = new List<Report>();
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
        }
    }
}

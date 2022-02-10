using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Test : Form {
        public class DailySheetModel {
            public string Date { get; set; }
            public string ComCodeComp { get; set; }
            public string ComCode { get; set; }
            public double MonthlyTotal { get; set; }
        }
        public class ComManModel {
            public string ComCode { get; set; }
            public string Manager { get; set; }
            public double MonthlyHistory { get; set; }
            public double MonthlyTotal { get; set; }
        }
        public class PersonModel {
            public string Manager { get; set; }
            public double Target { get; set; }
            public double Total { get; set; }
            public double CompleteRate { get; set; }

            public void CalculateRate()
            {
                CompleteRate = Total / ( Target * 10000 ) * 100;
            }
        }
        public class DepModel {
            public string Name { get; set; }
            public double Total { get; set; }
            public double Target { get; set; }
            public double CompleteRate { get; set; }
            public List<PersonModel> ps { get; set; } = new List<PersonModel>();

            public void CalculateRate()
            {
                CompleteRate = Total / ( Target * 10000 ) * 100;
            }
        }
        public Form_Test()
        {
            InitializeComponent();
        }
        int ColI( char col )
        {
            return col - 'A';
        }
        int GetMonthFromString( string str )
        {
            return Convert.ToInt32( str.Substring( 5, 2 ) );
        }
        int GetYearFromString( string str )
        {
            return Convert.ToInt32( str.Substring( 0, 5 ) );
        }
        NPOIUtil nu_daily = new NPOIUtil();
        NPOIUtil nu = new NPOIUtil();

        private void Form_Test_Load( Object sender, EventArgs e )
        {


        }
        List<DailySheetModel> ds;
        int Month;
        int Year;
        int MonthDay;
        private void button1_Click( Object sender, EventArgs e )
        {
            // var a = eu_daily.OpenExcel();
            nu_daily.OpenFileDialog();
            // ==================== 打开每天下发的代发金额表 ====================
            // eu_daily.OpenExcelByPath( @"C:\Users\cyf-m\Documents\1.xlsx" );
            // var t = eu_daily.GetTableByIndex( 0 );

            var t = nu_daily.GetSheetAt( 0 );
            ds = new List<DailySheetModel>();

            // 应当保证
            // 统计日期           A
            // 代发单位编号       D
            // 当月累计代发金额   L
            for ( int i = 1; i < t.LastRowNum; i++ ) {
                var r = t.GetRow( i );
                ds.Add( new DailySheetModel() {
                    Date = r.GetCell( ColI( 'A' ) ).ToString(),
                    ComCodeComp = r.GetCell( ColI( 'D' ) ).ToString(),
                    MonthlyTotal = Convert.ToDouble( r.GetCell( ColI( 'L' ) ).ToString() )
                } );
            }
            foreach ( var d in ds ) {
                if ( d.ComCodeComp == "" ) {
                    continue;
                }
                d.ComCode = d.ComCodeComp.Split( '*' )[1];
            }
            Month = GetMonthFromString( ds[0].Date );
            Year = GetYearFromString( ds[0].Date );
            MonthDay = Convert.ToInt32( ds[0].Date.Substring( 5 ) ); ;
            button2.Enabled = true;
        }

        private void button2_Click( Object sender, EventArgs e )
        {
            // ==================== 打开含有“代发单位”表的文件 ====================
            nu.OpenFileDialog();
            var tcl = nu.GetSheet( "代发单位" );
            // 代发编号     B
            // 开拓人       G
            var cms = new List<ComManModel>();
            // 计算历史总和
            var r0 = tcl.GetRow( 0 );
            var StartMonth = Convert.ToInt32( r0.GetCell( ColI( 'I' ) ).ToString().Replace( '月', '\r' ) );
            var SumColCount = Month - StartMonth;

            for ( 
                int i = 1; i < tcl.LastRowNum; i++ ) {
                var r = tcl.GetRow( i );
                var cmmm = new ComManModel() {
                    ComCode = r.GetCell( ColI( 'B' ) ).ToString(),
                    Manager = r.GetCell( ColI( 'G' ) ).ToString(),
                    MonthlyHistory = 0
                };
                // 计算之前总数
                for ( int j = 0; j < SumColCount; j++ ) {
                    cmmm.MonthlyHistory += Convert.ToDouble( r.GetCell( ColI( 'I' ) + j ).ToString() );
                }
                cms.Add( cmmm );
            }
            // 将代发金额搬至代发单位表
            
            foreach ( var cm in cms ) {
                var dd = ds.Find( d => { return d.ComCode == cm.ComCode; } );
                if ( dd == null ) {
                    // Console.WriteLine($"{cm.ComCode}");
                }
                cm.MonthlyTotal = dd == null ? 0 : dd.MonthlyTotal;
            }
            var ccs = new List<string>();
            foreach ( var d in ds ) {
                var dd = cms.Find( cm => { return cm.ComCode == d.ComCode; } );
                if ( dd == null ) {
                    ccs.Add( d.ComCode );
                    Console.WriteLine( $"{d.ComCode}" );
                }
            }
            // 到个人
            var cmm = new Dictionary<string, double>();
            foreach ( var cm in cms ) {
                if ( !cmm.ContainsKey( cm.Manager ) ) {
                    cmm[cm.Manager] = 0;
                }
                cmm[cm.Manager] += ( cm.MonthlyTotal + cm.MonthlyHistory );
            }
            var cmmD = new Dictionary<string, double>( cmm );
            foreach ( var cm in cmmD ) {
                // 处理比例问题
                if ( cm.Key.Contains( '/' ) ) {
                    var ms = cm.Key.Split( '/' );
                    foreach ( var mmm in ms ) { // 遍历分割的字符串 xxx2 ttt3
                        foreach ( var cm2 in cmm ) { // 遍历名字 寻找 Manager
                            if ( !cm2.Key.Contains( '/' ) ) {
                                if ( mmm.Contains( cm2.Key ) ) {
                                    if ( cm2.Key.Length < mmm.Length ) {
                                        var mr = Convert.ToDouble( mmm.Substring( cm2.Key.Length ) );
                                        cmm[cm2.Key] += mr * cm.Value / 10;
                                    } else {
                                        cmm[cm2.Key] += 0.5 * cm.Value;
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            var dds = new List<DepModel>();
            var ps = new List<PersonModel>();

            var tp = nu.GetSheet( "到个人" );
            DepModel ddd = null;
            for ( int i = 4; i < tp.LastRowNum; i++ ) 
                {
                var r = tp.GetRow( i );
                if ( ddd == null ) {
                    ddd = new DepModel() {
                        Name = r.GetCell( ColI( 'A' )).ToString()
                    };
                }
                Console.WriteLine( r.GetCell( ColI( 'A' ) ).ToString() );
                var p = new PersonModel() {
                    Manager = r.GetCell( ColI( 'B' ) ).ToString(),
                    Target = Convert.ToDouble( r.GetCell( ColI( 'J' ) ).ToString() ),
                };
                if ( p.Manager == "小计" ) {
                    var total = new PersonModel() {
                        Manager = "小计",
                        Total = 0,
                        Target = Convert.ToDouble( r.GetCell( ColI( 'J' ) ).ToString() )
                    };
                    foreach ( var ppp in ps ) {
                        total.Total += ppp.Total;
                    }
                    p.CalculateRate();
                    total.CalculateRate();
                    ps.Add( total );
                    ddd.ps = ps;
                    ddd.Total = total.Total;
                    dds.Add( ddd );

                    ps = new List<PersonModel>();
                    ddd = null;
                } else {
                    if ( !cmm.ContainsKey( p.Manager ) ) {
                        continue;
                    }
                    p.Total = cmm[p.Manager];
                    p.CalculateRate();
                    ps.Add( p );
                }
            }
            var team = nu.GetSheet( "团队" );
            for ( int i = 0; i < team.LastRowNum; i++ ) {
                var r = team.GetRow(i);
                var teamName = r.GetCell( ColI( 'A' ) )?.ToString();
                var dddd = dds.Find( dd => { return dd.Name == teamName; } );
                if ( dddd != null ) {
                    dddd.Target = Convert.ToDouble( r.GetCell( ColI( 'C' ) ).ToString() );
                    dddd.CalculateRate();
                }
            }
            /*
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            var xlWorkBook = xlApp.Workbooks.Add();
            xlWorkBook.Sheets.Add( After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count] );
            xlWorkBook.Sheets.Add( After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count] );
            xlWorkBook.Sheets.Add( After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count] );

            var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 1 );
            var xlWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 2 );
            var xlWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 3 );
            var xlWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 4 );

            var Q = (int)( Month / 3 );
            Q += ( Month % 3 ) == 0 ? 0 : 1;
            xlWorkSheet.Cells[1, 1] = "团队";
            xlWorkSheet.Cells[1, 2] = $"{Year}年{Q}季度代发量目标";
            xlWorkSheet.Cells[1, 3] = $"{Q}季度完成量（截至{MonthDay}）";
            xlWorkSheet.Cells[1, 4] = $"{Q}季度完成率（截至{MonthDay}）";
            for ( int i = 0; i < dds.Count; i++ ) {
                var s = dds[i];
                xlWorkSheet.Cells[i + 2, 1] = s.Name;
                xlWorkSheet.Cells[i + 2, 2] = s.Target;
                xlWorkSheet.Cells[i + 2, 3] = s.Total / 10000;
                xlWorkSheet.Cells[i + 2, 4] = s.CompleteRate.ToString() + "%";
            }

            xlWorkSheet2.Cells[1, 1] = "团队";
            xlWorkSheet2.Cells[1, 2] = "客户经理";
            xlWorkSheet2.Cells[1, 3] = $"{Year}年{Q}季度代发量目标";
            xlWorkSheet2.Cells[1, 4] = $"{Year}年{Q}季度完成量（截至{MonthDay}）";
            xlWorkSheet2.Cells[1, 5] = $"{Year}年{Q}季度完成率（截至{MonthDay}）";

            int tttttt = 0;
            for ( int i = 0; i < dds.Count; i++ ) {
                var s = dds[i];
                for ( int j = 0; j < s.ps.Count; j++ ) {
                    var p = s.ps[j];
                    xlWorkSheet2.Cells[tttttt + 2, 1] = s.Name;
                    xlWorkSheet2.Cells[tttttt + 2, 2] = p.Manager;
                    xlWorkSheet2.Cells[tttttt + 2, 3] = p.Target;
                    xlWorkSheet2.Cells[tttttt + 2, 4] = p.Total / 10000;
                    xlWorkSheet2.Cells[tttttt + 2, 5] = p.CompleteRate.ToString() + "%";
                    tttttt++;
                }
            }
            xlWorkSheet3.Cells[1, 1] = $"{Month}月";
            for ( int i = 0; i < cms.Count; i++ ) {
                xlWorkSheet3.Cells[i + 2, 1] = cms[i].MonthlyTotal;
            }
            xlWorkSheet4.Cells[1, 1] = $"任务完成情况表-代发单位表 中 缺失的单位名单";
            for ( int i = 0; i < ccs.Count; i++ ) {
                xlWorkSheet4.Cells[i + 2, 1] = ccs[i];
            }
            var ff = Path.Combine( Environment.CurrentDirectory, $"最终结果{Year}{MonthDay}.xlsx" );
            xlWorkBook.SaveCopyAs( Path.Combine( Environment.CurrentDirectory, $"最终结果{Year}{MonthDay}.xlsx" ) );
            Process.Start( ff );
            */
        }
    }
}

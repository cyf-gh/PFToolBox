using ExcelDataReader;

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
    public partial class Form_MergeExcel : Form {
        public Form_MergeExcel()
        {
            InitializeComponent();
        }
        IExcelDataReader excelDataReader = null;
        FileStream stream = null;
        private void button1_Click( Object sender, EventArgs e )
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using ( OpenFileDialog openFileDialog = new OpenFileDialog() ) {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.RestoreDirectory = true;

                if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
                    filePath = openFileDialog.FileName;
                    try {
                        stream = File.Open( filePath, FileMode.Open, FileAccess.Read );
                    } catch ( Exception ex) {
                        MessageBox.Show($"{ex.Message}\n请检查excel表格是否在打开状态，或excel表格文件是否正确再重试");
                        return;
                    }
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    excelDataReader = ExcelReaderFactory.CreateReader( stream );
                    // 2. Use the AsDataSet extension method
                    var result = excelDataReader.AsDataSet();

                    var ts = result.Tables;
                    checkedListBox1.Items.Clear();
                    foreach ( DataTable t in ts ) {
                        checkedListBox1.Items.Add( t.TableName );
                    }
                    checkedListBox2.Items.Clear();
                    foreach ( DataTable t in ts ) {
                        checkedListBox2.Items.Add( t.TableName );
                    }
                    // The result of each spreadsheet is in result.Tables
                }
            }
        }

        public class Row {
            public string A;
            public string ARaw;
            public List<string> Data = new List<string>();
            public Row Parent;
            public List<string> Desc = new List<string>();
        }

        class RowDiff {
            public string A;
            public int Col;
            public string ARaw;
            public string Data1, Data2;
            public Row Parent;
            public string IsSame;
        }

        private void button2_Click( Object sender, EventArgs e )
        {
            var ci = checkedListBox1.CheckedItems;
            if ( ci.Count != 2 ) {
                MessageBox.Show( "必须选取2个表格进行合并，当前勾选的表格数量不为2" );
                return;
            }
            var diff = checkedListBox2.CheckedItems;
            if ( diff.Count != 1 ) {
                MessageBox.Show( "必须选取1个表格进行最终校对，当前勾选的表格数量不为1" );
                return;
            }
            var result = excelDataReader.AsDataSet();
            var ts = result.Tables;
            List<DataTable> dt = new List<DataTable>();
            DataTable diffT = new DataTable();
            foreach ( var c in ci ) {
                foreach ( DataTable t in ts ) {
                    if ( t.TableName == (string)c ) {
                        dt.Add( t );
                        break;
                    }
                }
            }
            foreach ( DataTable t in ts ) {
                if ( t.TableName == (string)diff[0] ) {
                    diffT = t;
                    break;
                }
            }
            var rows1 = DataRow2MERow( dt[0].Rows );
            var rows2 = DataRow2MERow( dt[1].Rows );
            Dictionary<string, Row> rows2_cp = new Dictionary<string, Row>();
            var res1 = MergeSheet( rows1, rows2, out rows2_cp );
            foreach ( var row2_cp in rows2_cp ) {
                for ( Int32 i1 = 0; i1 < res1.Count; i1++ ) {
                    Row i = res1[i1];
                    if ( i.Parent?.ARaw == row2_cp.Value.Parent?.ARaw ) {
                        res1.Insert( i1, row2_cp.Value );
                        break;
                    }
                }
            }
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if ( xlApp == null ) {
                MessageBox.Show( "excel未正确安装，请将输出的文本文件手动转化为excel文件\n不用担心，这不会影响最终的校对结果。" );
                string output = string.Empty;
                for ( Int32 i = 0; i < res1.Count; i++ ) {
                    Row re = res1[i];
                    output += re.A + ",";
                    foreach ( var d in re.Data ) {
                        output += d + ",";
                    }
                    foreach ( var d in re.Desc ) {
                        output += d + ",";
                    }
                    output += '\n';
                }
                textBox1.Text = output;
            } else {
                var xlWorkBook = xlApp.Workbooks.Add();
                var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 1 );
                for ( int i = 0; i < res1.Count; i++ ) {
                    int j = 1;
                    Row re = res1[i];
                    xlWorkSheet.Cells[i + 1, j] = re.A;
                    foreach ( var d in re.Data ) {
                        j += 1;
                        xlWorkSheet.Cells[i + 1, j] = d;
                    }
                    foreach ( var d in re.Desc ) {
                        j += 1;
                        xlWorkSheet.Cells[i + 1, j] = d;
                    }
                }
                xlWorkBook.SaveCopyAs( Path.Combine( Environment.CurrentDirectory, "合并报表.xlsx" ) );
                Process.Start( Environment.CurrentDirectory );
            }
            // 核对报表
            var rowDiff = DataRow2MERow( diffT.Rows );
            var merged = List2Dict( res1 );
            var deficiency = new List<RowDiff>(); // 缺少项目
            var diffs = new List<RowDiff>();
            var sames = new List<RowDiff>();
            var nonData = new List<RowDiff>();
            foreach ( var m in merged ) {
                if ( !rowDiff.ContainsKey( m.Key ) ) {
                    deficiency.Add(
                        new RowDiff {
                            A = m.Value.A,

                            Col = 1,
                            ARaw = m.Value.ARaw,
                            IsSame = "目标表格中不存在该科目",
                            Parent = m.Value.Parent,
                        }
                    );
                    continue;
                } else {
                    var di = rowDiff[m.Key];
                    for ( Int32 i = 0; i < m.Value.Data.Count; i++ ) {
                        String ddd = m.Value.Data[i];
                        String ddi = di.Data[i];
                        double d1, d2;
                        var isR1d = double.TryParse( ddd, out d1 );
                        var isR2d = double.TryParse( ddi, out d2 );
                        if ( isR1d && isR2d ) {
                            if ( d1 == d2 ) {
                                sames.Add( new RowDiff {
                                    A = m.Value.A,
                                    ARaw = m.Value.ARaw,
                                    Col = i + 1,
                                    IsSame = "相同",
                                    Data1 = ddd,
                                    Data2 = ddi,
                                    Parent = m.Value.Parent,
                                } );
                            } else {
                                diffs.Add( new RowDiff {
                                    A = m.Value.A,
                                    ARaw = m.Value.ARaw,
                                    Col = i + 1,
                                    IsSame = "不同",
                                    Data1 = ddd,
                                    Data2 = ddi,
                                    Parent = m.Value.Parent,
                                } );
                            }
                        } else {
                            if ( ddd == ddi ) {
                                nonData.Add( new RowDiff {
                                    A = m.Value.A,
                                    ARaw = m.Value.ARaw,
                                    Col = i + 1,
                                    IsSame = "非数据",
                                    Data1 = ddd,
                                    Data2 = ddi,
                                    Parent = m.Value.Parent,
                                } );
                            }
                        }
                    }

                }
                rowDiff.Remove( m.Key );
            }
            var excresent = new List<RowDiff>();
            foreach ( var d in rowDiff ) {
                excresent.Add( new RowDiff {
                    A = d.Value.A,

                    Col = 1,
                    ARaw = d.Value.ARaw,
                    IsSame = "合并表格中不存在该科目",
                    Parent = d.Value.Parent,
                } );
            }
            if ( xlApp == null ) {
                MessageBox.Show( "excel未正确安装，请将输出的文本文件手动转化为excel文件\n不用担心，这不会影响最终的校对结果。" );
                string output = string.Empty;
                for ( Int32 i = 0; i < res1.Count; i++ ) {
                    Row re = res1[i];
                    output += re.A + ",";
                    foreach ( var d in re.Data ) {
                        output += d + ",";
                    }
                    foreach ( var d in re.Desc ) {
                        output += d + ",";
                    }
                    output += '\n';
                }
                textBox1.Text = output;
            } else {
                var xlWorkBook = xlApp.Workbooks.Add();
                xlWorkBook.Sheets.Add( After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count] );
                var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 1 );
                var xlWorkSheetInfo = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 2 );
                
                xlWorkSheetInfo.Cells[1, 1] = "缺失科目个数";
                xlWorkSheetInfo.Cells[2, 1] = deficiency.Count;
                xlWorkSheetInfo.Cells[1, 2] = "相同科目个数";
                xlWorkSheetInfo.Cells[2, 2] = sames.Count;
                xlWorkSheetInfo.Cells[1, 3] = "不同科目个数";
                xlWorkSheetInfo.Cells[2, 3] = diffs.Count;
                xlWorkSheetInfo.Cells[1, 4] = "非数据个数";
                xlWorkSheetInfo.Cells[2, 4] = nonData.Count;
                xlWorkSheetInfo.Cells[1, 5] = "目标表格中多出的科目个数";
                xlWorkSheetInfo.Cells[2, 5] = excresent.Count;

                var sum = new List<RowDiff>();
                sum.AddRange( deficiency );
                sum.AddRange( sames );
                sum.AddRange( diffs );
                sum.AddRange( nonData );
                sum.AddRange( excresent );

                xlWorkSheet.Cells[1, 1] = "科目（仅数字）";
                xlWorkSheet.Cells[1, 2] = "科目";
                xlWorkSheet.Cells[1, 3] = "所在列数";
                xlWorkSheet.Cells[1, 4] = "合并表中数据";
                xlWorkSheet.Cells[1, 5] = "目标表中数据";
                xlWorkSheet.Cells[1, 6] = "描述";
                xlWorkSheet.Cells[1, 7] = "父级类型名";
                for ( int i = 0; i < sum.Count; i++ ) {
                    var s = sum[i];
                    xlWorkSheet.Cells[i + 2, 1] = s.A;
                    xlWorkSheet.Cells[i + 2, 2] = s.ARaw;
                    xlWorkSheet.Cells[i + 2, 3] = ((char)(s.Col + 'A')).ToString();
                    xlWorkSheet.Cells[i + 2, 4] = s.Data1;
                    xlWorkSheet.Cells[i + 2, 5] = s.Data2;
                    xlWorkSheet.Cells[i + 2, 6] = s.IsSame;
                    xlWorkSheet.Cells[i + 2, 7] = s.Parent?.ARaw;
                }
                var ff = Path.Combine( Environment.CurrentDirectory, "最终结果.xlsx" );
                xlWorkBook.SaveCopyAs( Path.Combine( Environment.CurrentDirectory, "最终结果.xlsx" ) );
                Process.Start( ff );
            }
        }

        static Dictionary<string, Row> DataRow2MERow( DataRowCollection rows )
        {
            var rows2 = new Dictionary<string, Row>();
            Row p = null;
            foreach ( DataRow row in rows ) {
                var araw = row.ItemArray[0].ToString();
                var a = row.ItemArray[0].ToString().Replace( " ", "" );
                string xxxx = a;
                if ( a.Length >= 4 ) {
                    xxxx = a.Substring( 0, 4 );
                }
                rows2[xxxx] = new Row();
                rows2[xxxx].ARaw = araw;
                rows2[xxxx].A = xxxx;
                for ( int i = 1; i < row.ItemArray.Length; i++ ) {
                    rows2[xxxx].Data.Add( row.ItemArray[i].ToString() );
                }
                if ( !double.TryParse( xxxx, out _ ) ) {
                    p = rows2[xxxx];
                } else {
                    rows2[xxxx].Parent = p;
                }
            }
            return rows2;
        }

        static Dictionary<string, Row> List2Dict( List<Row> rows )
        {
            var dic = new Dictionary<string, Row>();
            foreach ( var r in rows ) {
                dic[r.A] = r;
            }
            return dic;
        }

        List<Row> MergeSheet( Dictionary<string, Row> rows1, Dictionary<string, Row> rows2, out Dictionary<string, Row> rows2_cp )
        {
            rows2_cp = new Dictionary<string, Row>( rows2 );
            var repeats = new List<Row>();
            var res = new List<Row>();
            foreach ( var r in rows1 ) {
                // 重复项相加
                if ( rows2.ContainsKey( r.Key ) ) {
                    var r1 = r.Value;
                    rows2_cp.Remove( r.Key );
                    repeats.Add( r1 );
                    var r2 = rows2[r.Key];
                    var newData = new List<string>();
                    var descData = new List<string>();

                    double d1 = 0, d2 = 0;
                    for ( Int32 i = 0; i < r1.Data.Count; i++ ) {
                        if ( i >= r2.Data.Count ) {
                            break;
                        }
                        String r1d = r1.Data[i];
                        String r2d = r2.Data[i];

                        var isR1d = double.TryParse( r1d, out d1 );
                        var isR2d = double.TryParse( r2d, out d2 );
                        if ( isR1d && isR2d ) {
                            newData.Add( ( d1 + d2 ).ToString() );
                            descData.Add( ( d1 == 0 ) && ( d2 == 0 ) ? "" : $"{d1}+{d2}" );
                        } else {
                            newData = ( r1.Data );
                        }
                    }
                    res.Add( new Row {
                        A = r1.A,
                        ARaw = r1.ARaw, /// TODO
                        Data = newData,
                        Parent = r1.Parent,
                        Desc = descData
                    } );
                } else {
                    // 不重复项目直接添加
                    res.Add( r.Value );
                }
            }
            return res;
        }

        private void Form1_FormClosed( Object sender, FormClosedEventArgs e )
        {
            if ( stream != null ) {
                stream.Close();
            }
        }

        private void Form1_Load( Object sender, EventArgs e )
        {

        }
    }
}

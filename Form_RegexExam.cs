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
    public partial class Form_RegexExam : Form {
        public Form_RegexExam()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 转全角的函数(SBC case)
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string ToSBC( string input )
        {
            //半角转全角：
            char[] c = input.ToCharArray();
            for ( int i = 0; i < c.Length; i++ ) {
                if ( c[i] == 32 ) {
                    c[i] = (char)12288;
                    continue;
                }
                if ( c[i] < 127 )
                    c[i] = (char)( c[i] + 65248 );
            }
            return new string( c );
        }
        public static string ToDBC( string input )
        {
            char[] c = input.ToCharArray();
            for ( int i = 0; i < c.Length; i++ ) {
                if ( c[i] == 12288 ) {
                    c[i] = (char)32;
                    continue;
                }
                if ( c[i] > 65280 && c[i] < 65375 )
                    c[i] = (char)( c[i] - 65248 );
            }
            return new string( c );
        }
        public class Question {
            public string Id { get; set; }
            public string Stem { get; set; }
            public string Ans { get; set; }
            public Dictionary<string, string> Choices { get; set; } = new Dictionary<string, string>();
        }
        private void FormRegexExam_Load( Object sender, EventArgs e )
        {
//            string exam = @"21、下列哪个业务不属于贵金属业务范围（D）
//A、账户外汇
//B、外汇宝
//C、双向外汇
//D、结售汇";


//            var lines = exam.Split( '\n' );

            var lines = File.ReadAllLines("./input.txt");

            //var lines = new List<string>();
            //foreach ( var l in flines ) {
            //    lines.AddRange( l.Split(' ').Where( x => !string.IsNullOrWhiteSpace( x ) ).ToList() );
            //}

            var qs = new List<Question>();
            int i = 0;
            qs.Add( new Question { } );
            foreach ( var l in lines ) {
                var ll = ToDBC( l );
                var questionPattern = new Regex( @"(?<id>\d+)(、|\.)(?<stem>[\s\S]+)\(\s*(?<ans>[A-G|a-d]+)\s*\)(?<stem2>[\s\S]*)" );
                Match m = questionPattern.Match( ll );
                if ( m.Success ) {
                    qs.Add( new Question {
                        Id = m.Groups["id"].Value,
                        Stem = m.Groups["stem"].Value + "()" + m.Groups["stem2"].Value,
                        Ans = m.Groups["ans"].Value,
                    } );
                    ++i;
                    // 匹配题干
                    continue;
                } else {
                    // 匹配选项
                    var choicePattern = new Regex( @"(?<ag>[A-G])(、|\.)[\s]*(?<desc>[\S]+)[\s]*" );
                    MatchCollection mc = choicePattern.Matches( ll );
                    foreach ( Match mcc in mc ) {
                        if ( mcc.Success ) {
                            qs[i].Choices[mcc.Groups["ag"].Value] = mcc.Groups["desc"].Value;
                        }
                    }
                }
            }
            // 完成匹配
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            var xlWorkBook = xlApp.Workbooks.Add();
            var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item( 1 );

            xlWorkSheet.Cells[1, 1] = "序号";
            xlWorkSheet.Cells[1, 2] = "题干";
            xlWorkSheet.Cells[1, 3] = "答案";
            xlWorkSheet.Cells[1, 4] = "A";
            xlWorkSheet.Cells[1, 5] = "B";
            xlWorkSheet.Cells[1, 6] = "C";
            xlWorkSheet.Cells[1, 7] = "D";
            xlWorkSheet.Cells[1, 8] = "E";
            xlWorkSheet.Cells[1, 9] = "F";
            xlWorkSheet.Cells[1, 10] = "G";
            xlWorkSheet.Cells[1, 11] = "类型";
            i = 1;
            qs.RemoveAt(0);
            foreach ( var d in qs ) {
                ++i;
                xlWorkSheet.Cells[i, 1] = d.Id;
                xlWorkSheet.Cells[i, 2] = d.Stem;

                if ( d.Choices.Count == 2 ) {
                    xlWorkSheet.Cells[i, 3] = d.Ans; //== "A" ? "对" : "错";
                    xlWorkSheet.Cells[i, 4] = "对";
                    xlWorkSheet.Cells[i, 5] = "错";
                    continue;
                }
                int j = 4;
                xlWorkSheet.Cells[i, 3] = d.Ans;
                foreach ( var c in d.Choices ) {
                    xlWorkSheet.Cells[i, j] = $"{c.Key}、"+ c.Value;
                    j++;
                }
            }
            var ff = Path.Combine( Environment.CurrentDirectory, "题目.xlsx" );
        SAVE:
            try {
                xlWorkBook.SaveCopyAs( ff );
            } catch ( Exception ex ) {
                MessageBox.Show( $"{ex.Message}\n请重试" );
                goto SAVE;
                throw;
            }
        }
    }
}

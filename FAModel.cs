using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace MergeExcel.FA {
    public class SheetBalance {
        private static Application App = null;
        public static Application InitExcelApp() {
            return App == null ? new Application() : App;
        }
        public readonly string FullName;
        public SheetBalance( string filePath ) {
            var xlWB = App.Workbooks.Open( filePath );
            var xlWS = xlWB.Worksheets["Sheet1"];
            int color_n = Convert.ToInt32( ( xlWS.Cells[1, "D"] ).Interior.Color );
            // Color color = ColorTranslator.FromOle( color_n );
        }
        public Dictionary<string, List<string>> NameAlias { get; set; } = new Dictionary<string, List<string>>();
    }

    public class ItemModel {

        public string Name { get; set; }
        public string Index { get; set; }
        public double ValuePrev { get; set; }
        public double ValueCur { get; set; }
        public bool IsSum { get; set; }
    }
    public class SectionModel {

    }
}
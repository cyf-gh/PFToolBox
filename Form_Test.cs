using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcel {
    public partial class Form_Test : Form {
        public Form_Test()
        {
            InitializeComponent();
        }
        EUtil eu = new EUtil();

        private void Form_Test_Load( Object sender, EventArgs e )
        {
            var a = eu.OpenExcel();
            Console.Write( a );
        }
    }
}

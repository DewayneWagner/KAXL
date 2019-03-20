using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DKAExcelStuff;

namespace DKARibbon
{
    public partial class frmCurrencyConvert : Form
    {
        KAXLApp K;

        public frmCurrencyConvert(KAXLApp k)
        {

            InitializeComponent();
            ControlBox = false;

            K = k;
        }

        private void btn_ConvertData_Click(object sender, EventArgs e)
        {
            KAXL.CADtoUSDConverter(K, Convert.ToDouble(txt_ExRateCADtoUSD.Text));
        }

        private void btn_Exit_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}

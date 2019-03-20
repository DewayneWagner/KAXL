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

namespace EXPREP_V2
{
    public partial class frmEXPREP_V2_WINDOW : Form
    {
        KAXLApp k;
        Master M;
        public frmEXPREP_V2_WINDOW(KAXLApp kaxlApp)
        {
            InitializeComponent();
            k = kaxlApp;
        }

        private void frmEXPREP_V2_WINDOW_Load(object sender, EventArgs e)
        {

        }

        private void btnUpdateExpRep_Click(object sender, EventArgs e)
        {
            M = new Master(k, this);
            try
            {
                txtDone.Text = "Working...";
                

                txtDone.Text = "DONE!!!";
                txtQDeliveryDatesUpdated.Text = Convert.ToString(M.updateMetrics.QUpdatedRevisedDeliveryDates);
                txtQLinesUpdated.Text = Convert.ToString(M.updateMetrics.QTotalUpdatedLines);
                txtQReceivedDatesUpdated.Text = Convert.ToString(M.updateMetrics.QUpdatedReceivedDates);
                txtTimeElapsed.Text = Convert.ToString(M.stopWatch.ElapsedTime);
            }
            catch
            {
                txtDone.Text = "F'd Up..." + "\n" + M.errorTracker.GetErrorMessage();
            }
        }

        public void KillProgramLeaveWindowOpen(string message)
        {
            txtDone.Text = message;
            
        }
        
        private void btnExit_Click(object sender, EventArgs e)
        {
            //KAXLApp.CloseSheet(k);
            Close();
        }
    }
}

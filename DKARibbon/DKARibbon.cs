using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using DKAExcelStuff;
using EXPREP_V2;
using System.Windows.Forms;
using XL = Microsoft.Office.Interop.Excel;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new DKARibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace DKARibbon
{
    [ComVisible(true)]
    public class DKARibbon : Office.IRibbonExtensibility
    {
        private Form currencyConversionForm;
        private Office.IRibbonUI ribbon;

        public DKARibbon()
        {

        }
        public void TestButton(Office.IRibbonControl control)
        {
            KAXLApp k = new KAXLApp();
            DKAExcelStuff.TestButton.TestM(k);
        }

        public void OnCurrencyConversion(Office.IRibbonControl control)
        {
            KAXLApp k = new KAXLApp();
            currencyConversionForm = new frmCurrencyConvert(k);
            currencyConversionForm.ShowDialog();
        }
        public void OnScrubItemNumbers(Office.IRibbonControl control)
        {            
            KAXL.ScrubItemNumbers(new KAXLApp());
        }

        public void OnWarRoomButton(Office.IRibbonControl control)
        {
            DKAExcelStuff.WarRoomFormat.ReqData();
        }
        
        public void OnWarRoomPresentButton(Office.IRibbonControl control)
        {
            DKAExcelStuff.WarRoomPresent.PresentData();
        }
        
        public void OnFirstRowFormulaButton(Office.IRibbonControl control)
        {
            KAXLApp k = new KAXLApp();
            KAXL.TopRowFormulas(k);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(k.XL);
        }

        public void btnIfError(Office.IRibbonControl control)
        {
            KAXLApp k = new KAXLApp();
            KAXL.IfError(k);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(k.XL);
        }
        public void btnOverWrite(Office.IRibbonControl control)
        {
            KAXLApp kaxlApp = new KAXLApp();
            KAXL.OverWriteFormulas(kaxlApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(kaxlApp.XL);
        }
        public void btnEXPREP_V2(Office.IRibbonControl control)
        {
            frmEXPREP_V2_WINDOW expRepForm = new frmEXPREP_V2_WINDOW(new KAXLApp());
            expRepForm.ShowDialog();

            //frmEXPREP_V2_WINDOW expRepForm = new frmEXPREP_V2_WINDOW(new KAXLApp());
            //expRepForm.ShowDialog();

            //KAXLApp kaxlApp = new KAXLApp();
            //System.Windows.Forms.Application.EnableVisualStyles();
            //System.Windows.Forms.Application.Run(new frmEXPREP_V2_WINDOW(kaxlApp));
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("DKARibbon.DKARibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

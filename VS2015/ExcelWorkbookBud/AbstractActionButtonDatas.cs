using Excel = Microsoft.Office.Interop.Excel;
using ExcelWorkbookBud;
using Microsoft.Office.Tools.Excel;
using SageWSSelData;


namespace ExcelWorkbook.Actions
{
    abstract class AbstractActionButtonDatas : IActionButtonDatas
    {
        protected static readonly string START_FORMULA = "~~S";
        protected static readonly string END_FORMULA = "~~E";
        // le tableau TABLE_BUD contient les données des datas.
        // Ces données sont dans un rectangle contenu dans TABLE_BUD
        // A---------------------B
        // |                   |
        // D---------------------C
        protected Excel.Range rangeTableBudA;

        public Excel.Range RangeTableBudA
        {
            get { return rangeTableBudA; }

        }
        protected Excel.Range rangeTableBudB;

        public Excel.Range RangeTableBudB
        {
            get { return rangeTableBudB; }

        }
        protected Excel.Range rangeTableBudC;

        public Excel.Range RangeTableBudC
        {
            get { return rangeTableBudC; }

        }
        protected Excel.Range rangeTableBudD;

        public Excel.Range RangeTableBudD
        {
            get { return rangeTableBudD; }

        }

        private static string PASSWD_PROTECT_EXCEL = "sage";

        //protected Excel._Worksheet sheetForBud = null;

        /*public Excel._Worksheet SheetForBud
        {
            get { return sheetForBud; }

        }
         */

        //protected Excel._Worksheet sheetGlobalParams = null;
        //protected Excel._Worksheet sheetHome = null;
        //protected Excel._Worksheet sheetLogWebService = null;
        //protected Feuil5 feuilLogWebService=null;

        protected int staret = 0;
        public int Staret
        {
            get { return staret; }

        }
        protected string mesret = "";

        public string Mesret
        {
            get { return mesret; }

        }

        public string DataType
        {
            get
            {
                return dataType;
            }

            set
            {
                dataType = value;
            }
        }

        private string dataType = "";  // Backing store


        protected int tabBudNbRow = 0;
        protected int tabBudNbCol = 0;

        public abstract void call();

        static public void protectFeuilForBud()
        {

            //Globals.FeuilDataForm.Protect(PASSWD_PROTECT_EXCEL, true, true, true, true, false, false, false, false, false, false, false, false, true, true, true);
            Globals.FeuilDataForm.Protect(PASSWD_PROTECT_EXCEL,true,true,true,true,false,false,false,false,true,false,false,true,true,true,true);

            Globals.FeuilDataForm.EnableOutlining = true;
        }
        static public void deProtectFeuilForBud()
        {

            //Globals.FeuilDataForm.Protect(PASSWD_PROTECT_EXCEL, false, false);
        }

        static public void protectFeuilHiddenDatas()
        {

            Globals.FeuilHiddenDatas.Protect(PASSWD_PROTECT_EXCEL, true, true, true, true, true, true, true, false, false, false, false, false, true, true, true);

            Globals.FeuilHiddenDatas.EnableOutlining = true;
        }
        static public void deProtectFeuilHiddenDatas()
        {

            Globals.FeuilHiddenDatas.Protect(PASSWD_PROTECT_EXCEL, false, false);
        }
        protected void initObjectsreference()
        {
            //sheetForBud = Globals.ThisWorkbook.Worksheets["For Bud"];
            //sheetGlobalParams = Globals.ThisWorkbook.Worksheets["Global Params"];
            //sheetLogWebService = Globals.ThisWorkbook.Worksheets["Log web service"];
            //feuilLogWebService = (Feuil5)Globals.Feuil5;

            //rangeTableBudA = sheetForBud.get_Range("TABLE_BUD").Cells[1, 1];
            rangeTableBudA = Globals.FeuilDataForm.TABLE_BUD.Cells[1, 1];

            // le tableau TABLE_BUD contient les données du des datas.
            // Ces données sont dans un rectangle contenu dans TABLE_BUD
            // A---------------------B
            // |                   |
            // D---------------------C
        }

        static protected void activeLog(CallWebServiceX3 callWebServiceX3)
        {
            callWebServiceX3.CAdxCallContext.requestConfig = "";
            //if (feuilLogWebService.checkBoxLogWebServeur.Checked == true)
            if (Globals.FeuilLogWebService.checkBoxLogWebServeur.Checked == true)
            {
                callWebServiceX3.CAdxCallContext.requestConfig = "adxwss.trace.on=on";
                if (Globals.FeuilLogWebService.checkBoxLogX3.Checked == true)
                {
                    callWebServiceX3.CAdxCallContext.requestConfig += "&adonix.trace.on=on";
                }
            }
            else
            {
                if (Globals.FeuilLogWebService.checkBoxLogX3.Checked == true)
                {
                    callWebServiceX3.CAdxCallContext.requestConfig = "adonix.trace.on=on";
                }
            }
        }

        static protected void setFeuilLog(CallWebServiceX3 callWebServiceX3)
        {
            if (Globals.FeuilLogWebService.checkBoxLogWebServeur.Checked == true || Globals.FeuilLogWebService.checkBoxLogX3.Checked == true)
            {
                Globals.FeuilLogWebService.textBoxInputXML.Text = callWebServiceX3.InputXmlString;
                Globals.FeuilLogWebService.textBoxOutputXML.Text = callWebServiceX3.CAdxResultXml.resultXml;
                Globals.FeuilLogWebService.textBoxLogWebService.Text = callWebServiceX3.CAdxResultXml.technicalInfos.traceRequest;
            }
        }

     
    }
}

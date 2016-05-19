using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelWorkbook.Model;
using ExcelWorkbook.Actions;

namespace ExcelWorkbookBud
{
    public partial class ThisWorkbook
    {
        internal ConnectionWSX3             connectionWSX3;
        internal ActionButtonLoadDatasForm     actionButtonLoadDatasForm;
        internal ActionButtonLoadDatasRep     actionButtonLoadDatasRep;
        internal ActionButtonUpdateDatas   actionButtonUpdateDatas;
        //internal ExcelFormulas              excelFormulas;

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            actionButtonLoadDatasForm = new ActionButtonLoadDatasForm();
            actionButtonLoadDatasRep  = new ActionButtonLoadDatasRep();
            actionButtonUpdateDatas = new ActionButtonUpdateDatas();
            //excelFormulas = new ExcelFormulas(Globals.FeuilForBud.TABLE_BUD.Rows.Count, Globals.FeuilForBud.TABLE_BUD.Columns.Count);
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par le Concepteur VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du Concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}

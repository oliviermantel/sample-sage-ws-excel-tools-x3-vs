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

namespace ExcelWorkbookBud
{
    public partial class FeuilHiddenDatas
    {
        private void Feuil10_Startup(object sender, System.EventArgs e)
        {
        }

        private void Feuil10_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par le Concepteur VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du Concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.Feuil10_Startup);
            this.Shutdown += new System.EventHandler(this.Feuil10_Shutdown);

        }

        #endregion

        private void TABLE_BUD_HIDDEN_Change(Excel.Range Target)
        {

        }

    }
}

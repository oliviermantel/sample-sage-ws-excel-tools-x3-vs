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
using ExcelWorkbook.Actions;
using ExcelWorkbook.Model;


namespace ExcelWorkbookBud
{
    public partial class FeuilDataForm
    {
        //ActionButtonLoadDatas actionButtonLoadDatas;
       
        DescriptionColumn descriptionColumn = new DescriptionColumn();

        private void Feuil3_Startup(object sender, System.EventArgs e)
        {   
        }

        private void Feuil3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par le Concepteur VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du Concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.TABLE_BUD.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.TABLE_BUD_Change);
            this.Startup += new System.EventHandler(this.Feuil3_Startup);
            this.Shutdown += new System.EventHandler(this.Feuil3_Shutdown);

        }

        #endregion



        private void TABLE_BUD_Change(Excel.Range Target) {

            Excel._Application app;
            app = Globals.ThisWorkbook.Application;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            TABLE_BUD_Sage_AM(Target);
            app.ScreenUpdating = true;
            app.EnableEvents = true;
            
        }

        private void TABLE_BUD_Sage_AM(Excel.Range Target)
        {
            
            int lig, col;
            int ligA, colA;
            // M(lig,col) est dans le referentiel avec le point A comme origine

            ligA = Globals.ThisWorkbook.actionButtonLoadDatasForm.RangeTableBudA.Row;
            colA = Globals.ThisWorkbook.actionButtonLoadDatasForm.RangeTableBudA.Column;
            lig = Target.Row - ligA + 1;
            col = Target.Column - colA + 1;
            int[] ligcol;
            Excel.Range r1, r2;
            //r1 = Globals.ThisWorkbook.actionButtonLoadDatas.RangeTableBudA.Cells[2, col];
            r1 = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[2, col];
            descriptionColumn.setDescription(r1.Value);
            if (descriptionColumn.TypDisplay == "Enter" && descriptionColumn.ActionsAM!="Empty" &&    descriptionColumn.TabActionsAM != null
                && descriptionColumn.TabActionsAM[0] != "" && lig != 2 && lig != 1)
            {
                for (int i = 0; i < descriptionColumn.TabActionsAM.Length; i++)
                {
                    if (descriptionColumn.TabActionsAM[i] != "")
                    {
                        ligcol = DescriptionColumn.getRowColumn(descriptionColumn.TabActionsAMLeft[i]);
                        if (ligcol[0] == 0)
                        {
                            r2 = Globals.ThisWorkbook.actionButtonLoadDatasForm.RangeTableBudA.Cells[lig, ligcol[1]];
                            r2.FormulaR1C1 = "=" + DescriptionColumn.translateRangeColumn("",Globals.ThisWorkbook.actionButtonLoadDatasForm.RangeTableBudA.Column, descriptionColumn.TabActionsAMRight[i]);
                            //r2.FormulaR1C1 = r2.Value;
                        }
                    }
                }
            }
        }

        private void buttonAddLigne_Click(object sender, EventArgs e)
        {
            Excel._Application app;
            app = Globals.ThisWorkbook.Application;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;
            AbstractActionButtonDatas.deProtectFeuilForBud();
            AbstractActionButtonDatas.deProtectFeuilHiddenDatas();
            Globals.FeuilDataForm.TABLE_BUD.AutoFilter();
            Excel.Range c = Globals.FeuilDataForm.TABLE_BUD.Cells[4,1];
            c.Value = "toto";
            c.Copy();
            c.Insert();
            c.PasteSpecial();
            app.ScreenUpdating = true;
            app.EnableEvents = true;
            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            AbstractActionButtonDatas.protectFeuilForBud();
            AbstractActionButtonDatas.protectFeuilHiddenDatas();
            Globals.ThisWorkbook.RefreshAll();
        }
    }
}

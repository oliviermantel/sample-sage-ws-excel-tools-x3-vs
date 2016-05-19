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
    public partial class FeuilCalculation
    {
        DescriptionColumn descriptionColumn = new DescriptionColumn();
        FormConnectionX3 formConnectionX3;

        private void Feuil9_Startup(object sender, System.EventArgs e)
        {
            
            AbstractActionButtonDatas.protectFeuilForBud();
            AbstractActionButtonDatas.protectFeuilHiddenDatas();
            //comboBoxSite.DataSource = Globals.FeuilDataForm.dataTableSites;
            //comboBoxSite.ValueMember = "CodeSite";
            //comboBoxSite.DisplayMember = "Designation";
           
            comboBoxSite.DataSource = Globals.FeuilDataForm.bindingSourceREP;
            //comboBoxSite.DataSource = Globals.FeuilDataForm.dataTableREP;
            comboBoxSite.ValueMember = "Code";
            comboBoxSite.DisplayMember = "Designation";
            Globals.ThisWorkbook.actionButtonLoadDatasRep.majDataSet();
            //comboBoxSite.SelectedIndex =-2;
        }
    

        private void Feuil9_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par le Concepteur VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du Concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.buttonLoadDatasForm.Click += new System.EventHandler(this.buttonLoadDatasForm_Click);
            this.buttonConnection.Click += new System.EventHandler(this.buttonConnection_Click);
            this.buttonUpdateDatas.Click += new System.EventHandler(this.buttonUpdateDatas_Click);
            this.comboBoxSite.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            this.comboBoxSite.SelectionChangeCommitted += new System.EventHandler(this.comboBoxSite_SelectionChangeCommitted);
            this.buttonLoadDep.Click += new System.EventHandler(this.buttonLoadDep_Click);
            this.X3CRIT1.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.X3CRIT1_Change);
            this.X3CRIT3.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.X3CRIT3_Change);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.FeuilCalculation_ActivateEvent);
            this.Startup += new System.EventHandler(this.Feuil9_Startup);
            this.Shutdown += new System.EventHandler(this.Feuil9_Shutdown);

        }

        #endregion

        private void buttonConnection_Click(object sender, EventArgs e)
        {
            formConnectionX3 = new FormConnectionX3();
            formConnectionX3.ShowDialog();

            Globals.ThisWorkbook.connectionWSX3 = new ConnectionWSX3(formConnectionX3.textBoxLoginX3.Text, formConnectionX3.textBoxPasswordX3.Text, formConnectionX3.comboBoxLanguageX3.Text);
            Globals.FeuilCalculation.X3LOGIN.Value = Globals.ThisWorkbook.connectionWSX3.Login;
           
        }

        private void buttonUpdateDatas_Click(object sender, EventArgs e)
        {
            if (Globals.ThisWorkbook.connectionWSX3 == null)
            {
                System.Windows.Forms.MessageBox.Show("Please, you must connect to X3");
                return;
            }
            if (System.Windows.Forms.MessageBox.Show("Update forecast X3 ?", "Confirmation", System.Windows.Forms.MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {

                Excel._Application app;
                app = Globals.ThisWorkbook.Application;
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                //AbstractActionButtonDatas.deProtectFeuilHiddenDatas();
                Globals.ThisWorkbook.actionButtonUpdateDatas.call();

                if (Globals.ThisWorkbook.actionButtonUpdateDatas.Staret > 0)
                {
                    System.Windows.Forms.MessageBox.Show(Globals.ThisWorkbook.actionButtonUpdateDatas.Mesret);
                }
                else
                {
                    //System.Windows.Forms.MessageBox.Show("End Loading budget");
                    System.Windows.Forms.MessageBox.Show("End updating forecast");
                }
                app.ScreenUpdating = true;
                app.EnableEvents = true;
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

            }
        }


        private void buttonLoadDatasForm_Click(object sender, EventArgs e)
        {
            
            buttonLoadDatasAll(Globals.ThisWorkbook.actionButtonLoadDatasForm,"Load forecast X3 ?", "End Loading forecast");

        }

        private void buttonLoadDatasAll(ActionButtonLoadDatas actionButtonLoadDatas, string mess1, string mess2)
        {
            if (Globals.ThisWorkbook.connectionWSX3 == null)
            {
                System.Windows.Forms.MessageBox.Show("Please, you must connect to X3");
                return;
            }
            //if (System.Windows.Forms.MessageBox.Show("Load budget X3 ?", "Confirmation", System.Windows.Forms.MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            if (System.Windows.Forms.MessageBox.Show(mess1, "Confirmation", System.Windows.Forms.MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)

                {
                Excel._Application app;
                app = Globals.ThisWorkbook.Application;
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                AbstractActionButtonDatas.deProtectFeuilForBud();
                AbstractActionButtonDatas.deProtectFeuilHiddenDatas();

                if (actionButtonLoadDatas.DataType == "DATAFORM")
                {
                    Globals.ThisWorkbook.actionButtonLoadDatasForm.call();
                }
                else
                {
                    Globals.ThisWorkbook.actionButtonLoadDatasRep.call();
                }

                app.ScreenUpdating = true;
                app.EnableEvents = true;
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                AbstractActionButtonDatas.protectFeuilForBud();
                AbstractActionButtonDatas.protectFeuilHiddenDatas();
                Globals.ThisWorkbook.RefreshAll();

                


                if (actionButtonLoadDatas.Staret > 0)
                {
                    System.Windows.Forms.MessageBox.Show(actionButtonLoadDatas.Mesret);
                }
                else
                {
                    //System.Windows.Forms.MessageBox.Show("End Loading budget");
                    System.Windows.Forms.MessageBox.Show(mess2);
                }


            }

        }

        private void X3CRIT1_Change(Excel.Range Target)
        {

        }

        private void X3CRIT3_Change(Excel.Range Target)
        {

        }

        private void FeuilDataFormBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void FeuilDataFormBindingSource_CurrentChanged_1(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBoxSite.SelectedIndex!=0)
            Globals.FeuilCalculation.X3CRIT1.Value= comboBoxSite.SelectedValue;
        }

        private void buttonLoadDep_Click(object sender, EventArgs e)
        {
            Globals.FeuilDataRep1.TableauDataRep1.DataBodyRange.Clear();
            Globals.Feuil4.TableauDataRep2.DataBodyRange.Clear();
            buttonLoadDatasAll(Globals.ThisWorkbook.actionButtonLoadDatasRep, "Load site and Products sales X3 ?", "End Loading site and Products sales X3");
        }

        private void comboBoxSite_SelectionChangeCommitted(object sender, EventArgs e)
        {

        }

        private void FeuilCalculation_ActivateEvent()
        {
            Globals.ThisWorkbook.actionButtonLoadDatasRep.majDataSet();
        }
    }
}
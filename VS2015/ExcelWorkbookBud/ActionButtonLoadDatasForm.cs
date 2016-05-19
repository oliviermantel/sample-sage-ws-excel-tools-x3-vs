using System;
using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;
using SageWSSelData;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelWorkbookBud;
using ExcelWorkbook.Model;
using Microsoft.Office.Interop.Excel;


namespace ExcelWorkbook.Actions
{
    class ActionButtonLoadDatasForm : ActionButtonLoadDatas
    {

        public ActionButtonLoadDatasForm()
        {
            this.DataType = "DATAFORM";

        }


        public override void call()
        {
            tabBudNbRow = 0;
            CAdxCallContext cAdxCallContext;
            CallWebServiceX3 callWebServiceX3;
            int allsel, nextt, nbsel;
            int tabnb = 0;
            int nbtabval = 0;
            String[] tabcrit;
            String[,] tabvalstr;
            int staret = 0;
            String mesret = "";
            int noReq = 0;
            Excel.Range range, range2;
            String url;
            string formule = "";
            //int nbLigDataRep1=1, nbLigDataRep2=1;
            //Excel.Range r1;
            //DescriptionColumn descriptionColumn = new DescriptionColumn();
            //int colDataRep = 0;
            tabcrit = new String[4];
            if (Globals.FeuilCalculation.X3CRIT1.Value!=null)
             tabcrit[0] = (string)Globals.FeuilCalculation.X3CRIT1.Value;
            if (Globals.FeuilCalculation.X3CRIT2.Value != null)
                tabcrit[1] = Globals.FeuilCalculation.X3CRIT2.Value.ToString();
            if (Globals.FeuilCalculation.X3CRIT3.Value != null)
                tabcrit[2] = Globals.FeuilCalculation.X3CRIT3.Value.ToString();

            if (tabcrit[0] == null || tabcrit[0] == "")
            {
                this.staret = 1000;
                this.mesret = "Site X3 required";
                return;
            }
            if (tabcrit[1] == null || tabcrit[1] == "")
            {
                this.staret = 1001;
                this.mesret = "Starting period required";
                return;
            }

            if (tabcrit[2] == null || tabcrit[2] == "")
            {
                this.staret = 1002;
                this.mesret = "End period required";
                return;
            }
            tabcrit[3] = DataType;

            cAdxCallContext = new CAdxCallContext();
            cAdxCallContext.poolAlias = (string)Globals.FeuilGlobalParams.POOLWS.Value; //"DEMO";
            url = (string)Globals.FeuilGlobalParams.X3WSENDPOINT.Value;
            Globals.ThisWorkbook.connectionWSX3.setConnect(cAdxCallContext);
            cAdxCallContext.requestConfig = "";
            callWebServiceX3 = new CallWebServiceX3(url, cAdxCallContext,Globals.ThisWorkbook.connectionWSX3.Login, Globals.ThisWorkbook.connectionWSX3.Password);
            int lig0 = 1;
            Excel.Range rows = Globals.FeuilDataForm.TABLE_BUD.Rows;
            
            for (int i = 2; i <= rows.Count; i++) // on n'efface pas la première ligne
            {
                    rows[i].ClearContents();
            }
            //Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_FORMULA.Clear();
            Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Clear();
            //Globals.FeuilHiddenDatas.nbRowsTAB_BUD.Clear();
            //Globals.FeuilHiddenDatas.nbColumnsTAB_BUD.Clear();
            
            rows = null;
            allsel = 0;
            nextt = 1;


            nbsel = 0;
            nbtabval = 0;
            tabnb = 0;
            activeLog(callWebServiceX3);

            while (allsel != 2)
            {
                callWebServiceX3.SelData("YSELDATA", (String)Globals.FeuilGlobalParams.X3TYPWS.Value, nextt, tabcrit, ref allsel,
                    ref nbsel, ref tabnb, ref nbtabval, out tabvalstr, ref staret, ref mesret, ref noReq);

                setFeuilLog(callWebServiceX3);

                this.staret = staret;
                this.mesret = mesret;
                if (staret > 0)
                {
                    break;
                }

                tabBudNbRow += tabnb;
                tabBudNbCol = nbtabval;

                nextt = 2;
                for (int lig = 0; lig < tabnb; lig++)
                {
                    for (int col = 0; col < nbtabval; col++)
                    {
                        if (tabvalstr[lig, col] != null && !tabvalstr[lig, col].Equals(""))
                        {
                            
                            if (tabvalstr[lig, col].StartsWith("="))
                            {
                               
                                    range = rangeTableBudA.Cells[lig + lig0, col + 1];
                                    //range.FormulaR1C1 = DescriptionColumn.translateRangeColumn(Globals.FeuilForBud.Name,rangeTableBudA.Column, tabvalstr[lig, col]);

                                    formule = DescriptionColumn.translateRangeColumn(Globals.FeuilDataForm.Name, rangeTableBudA.Column, tabvalstr[lig, col]);
                                    range.FormulaR1C1 = formule;
                                    //Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_FORMULA.Cells[lig + lig0, col + 1] = AbstractActionButtonDatas.START_FORMULA + range.FormulaR1C1 + AbstractActionButtonDatas.END_FORMULA; // S=Start, E=End 
                                
                            }
                            else
                            {
                                if (lig == 1 && lig0 == 1)
                                {
                                    range2 = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[lig + lig0, col + 1];
                                    range2.Value = tabvalstr[lig, col];
                                    
                                    
                                }
                                else if (lig == 0 && lig0 == 1)
                                { // On ne fait rien car on ne modifie pas les entetes de colonne
                                }
                                else
                                {   
                                    rangeTableBudA.Cells[lig + lig0, col + 1] = tabvalstr[lig, col];
                                }


                            }
                        }
                    }

                }


                lig0 = lig0 + tabnb;

            }

            rangeTableBudB = Globals.FeuilDataForm.TABLE_BUD.Cells[1, tabBudNbCol];
            rangeTableBudC = Globals.FeuilDataForm.TABLE_BUD.Cells[tabBudNbRow, tabBudNbCol];
            rangeTableBudD = Globals.FeuilDataForm.TABLE_BUD.Cells[tabBudNbRow, 1];
            //Globals.FeuilHiddenDatas.nbRowsTAB_BUD.Value = tabBudNbRow;
            //Globals.FeuilHiddenDatas.nbColumnsTAB_BUD.Value = tabBudNbCol;
            updatePresentation();

           
        }


        private void updatePresentation()
        {
            DescriptionColumn descriptionColumn = new DescriptionColumn();
            Excel.Range r1, r2;
            r1 = rangeTableBudA.Cells[2, 1];
            //r1 = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[2, 1];
            r2 = r1.EntireRow;
            r2.Hidden = true;

            for (int col = 0; col < tabBudNbCol; col++)
            {
                r1 = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[2, col + 1];
                descriptionColumn.setDescription(r1.Value);
                r2 = rangeTableBudA.Cells[2, col + 1];
                r2 = r2.EntireColumn;
                if (descriptionColumn.TypDisplay == "Hidden")
                {
                    r2.Hidden = true;
                }
                else
                {
                    r2.Hidden = false;
                }

                if (descriptionColumn.TypDisplay == "Hidden" || descriptionColumn.TypDisplay == "Display")
                {
                    r2.Locked = true;
                }
                else
                {
                    r2.Locked = false;
                }
            }

        }

    }



}

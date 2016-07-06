using System;
using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;
using SageWSSelData;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelWorkbookBud;
using ExcelWorkbook.Model;
using Microsoft.Office.Interop.Excel;


namespace ExcelWorkbook.Actions
{
    sealed class ActionButtonLoadDatasRep : ActionButtonLoadDatas
    {

        public ActionButtonLoadDatasRep()
        {
            initObjectsreference();
            this.DataType = "DATAREP";

        }


        /// <summary>
        /// 
        /// </summary>
        public override void call()
        {
            //tabBudNbRow = 0;
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
            String url;
            ///
            int nbLigDataRep1 = 1, nbLigDataRep2 = 1;
            DescriptionColumn descriptionColumn = new DescriptionColumn();
            int colDataRep = 0;
            ///
            tabcrit = new String[4];
            tabcrit[0] = "Empty";
            tabcrit[1] = "Empty";
            tabcrit[2] = "Empty";
            tabcrit[3] = DataType;

            cAdxCallContext = new CAdxCallContext();
            cAdxCallContext.poolAlias = (string)Globals.FeuilGlobalParams.POOLWS.Value; //"DEMO";
            url = (string)Globals.FeuilGlobalParams.X3WSENDPOINT.Value;

            Globals.ThisWorkbook.connectionWSX3.setConnect(cAdxCallContext);
            cAdxCallContext.requestConfig = "";
            //tmpoma ajout
            //url = "http://52.17.90.248:8124/soap-generic/syracuse/collaboration/syracuse/CAdxWebServiceXmlCC";
            callWebServiceX3 = new CallWebServiceX3(url, cAdxCallContext,Globals.ThisWorkbook.connectionWSX3.Login, Globals.ThisWorkbook.connectionWSX3.Password);
            int lig0 = 1;
            //Excel.Range rows = Globals.FeuilDataForm.TABLE_BUD.Rows;

            allsel = 0;
            nextt = 1;


            nbsel = 0;
            nbtabval = 0;
            tabnb = 0;
            activeLog(callWebServiceX3);

            //Globals.FeuilCalculation.listBoxSite.Items.Clear();
            //Globals.FeuilDataForm.dataTableSites.Clear();
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

                //tabBudNbRow += tabnb;
                //tabBudNbCol = nbtabval;

                nextt = 2;
                for (int lig = 0; lig < tabnb; lig++)
                {
                    for (int col = 0; col < nbtabval; col++)
                    {
                        if (tabvalstr[lig, col] != null && !tabvalstr[lig, col].Equals(""))
                        {
                            if (tabvalstr[lig, 0] == "DATAREP1" && col == colDataRep && col != 0)
                            {
                                //tabBudNbRow -= 1;
                                nbLigDataRep1 += 1;
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 1] = tabvalstr[lig, colDataRep];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 2] = tabvalstr[lig, colDataRep + 1];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 3] = tabvalstr[lig, colDataRep + 2];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 4] = tabvalstr[lig, colDataRep + 3];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 5] = tabvalstr[lig, colDataRep + 4];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 6] = tabvalstr[lig, colDataRep + 5];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 7] = tabvalstr[lig, colDataRep + 6];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 8] = tabvalstr[lig, colDataRep + 7];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 9] = tabvalstr[lig, colDataRep + 8];
                                Globals.FeuilDataRep1.TableauDataRep1.Range.Cells[nbLigDataRep1, 10] = tabvalstr[lig, colDataRep + 9];

                                //Globals.FeuilDataForm.dataTableSites.Rows.Add(tabvalstr[lig, colDataRep], tabvalstr[lig, colDataRep + 1]);

                            }
                            else if (tabvalstr[lig, 0] == "DATAREP2" && col == colDataRep && col != 0)
                            {
                                //tabBudNbRow -= 1;
                                nbLigDataRep2 += 1;
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 1] = tabvalstr[lig, colDataRep];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 2] = tabvalstr[lig, colDataRep + 1];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 3] = tabvalstr[lig, colDataRep + 2];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 4] = tabvalstr[lig, colDataRep + 3];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 5] = tabvalstr[lig, colDataRep + 4];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 6] = tabvalstr[lig, colDataRep + 5];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 7] = tabvalstr[lig, colDataRep + 6];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 8] = tabvalstr[lig, colDataRep + 7];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 9] = tabvalstr[lig, colDataRep + 8];
                                Globals.Feuil4.TableauDataRep2.Range.Cells[nbLigDataRep2, 10] = tabvalstr[lig, colDataRep + 9];

                            }
                            if (tabvalstr[lig, col].StartsWith("="))
                            {
                                // ne rien faire
                            }
                            else
                            {
                                if (lig == 1 && lig0 == 1)
                                {
                                    // pas fini à epurer cette methode   je dois voir si j'ai besoin d'une cellule particuliere
                                    //range2 = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[lig + lig0, col + 1];
                                    //range2.Value = tabvalstr[lig, col];
                                    //descriptionColumn.setDescription(range2.Value);
                                    descriptionColumn.setDescription(tabvalstr[lig, col]);
                                    if (descriptionColumn.NameColumn == "REP")
                                    {
                                        colDataRep = col;
                                    }

                                }
                                else if (lig == 0 && lig0 == 1)
                                { // On ne fait rien car on ne modifie pas les entetes de colonne
                                }
                                else
                                {
                                    //rangeTableBudA.Cells[lig + lig0, col + 1] = tabvalstr[lig, col];
                                }


                            }
                        }
                    }

                }


                lig0 = lig0 + tabnb;

            }



            majDataSet();



        }

        public void majDataSet()
        {
            Globals.FeuilDataForm.dataSetDatas.Clear();
            Excel.Range r, r1, r2;
            Globals.FeuilDataForm.dataTableREP.Rows.Add("DATAREP1", "", "","");
            foreach (ListRow lr in Globals.FeuilDataRep1.TableauDataRep1.ListRows)
            {
                r = lr.Range;
                r1 = r.Cells[1, 1];
                r2 = r.Cells[1, 2];
                if (r1.Value==null || r1.Value == "")
                    break;
                Globals.FeuilDataForm.dataTableREP.Rows.Add("DATAREP1", r1.Value, r2.Value,"");

            }
            /* Pas besoin pour l'instant
            Globals.FeuilDataForm.dataTableREP.Rows.Add("DATAREP2", "", "","");
            foreach (ListRow lr in Globals.Feuil4.TableauDataRep2.ListRows)
            {
                r = lr.Range;
                r1 = r.Cells[1, 1];
                r2 = r.Cells[1, 2];
                r3 = r.Cells[1, 3];
                if (r1.Value == "")
                    break;
                Globals.FeuilDataForm.dataTableREP.Rows.Add("DATAREP2", r1.Value, r2.Value,r3.Value);
                

            }

            */
        }



    }
}

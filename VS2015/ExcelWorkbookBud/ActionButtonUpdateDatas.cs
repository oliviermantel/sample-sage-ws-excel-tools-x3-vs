using System;
using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;
using SageWSSelData;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelWorkbookBud;
using ExcelWorkbook.Model;
using System.Collections.Generic;
using System.Text;
using System.Globalization;


namespace ExcelWorkbook.Actions
{
    sealed class ActionButtonUpdateDatas : AbstractActionButtonDatas
    {


        private static readonly int NB_ROWS_UPDATE_BUD = 100;
        public ActionButtonUpdateDatas()
        {
            initObjectsreference();

        }

        public override void call()
        {
            CAdxCallContext cAdxCallContext;
            CallWebServiceX3 callWebServiceX3;
            DescriptionColumn descriptionColumn = new DescriptionColumn();
            string filt = "";
            int nblig = 0;
            int nbupd = 0;
            string transac = "Begin";
            int nbtabclb = 0;
            int nbkey = 0;

            int nbval = 0;
            int nbtransac;
            int nbcolumns = Globals.FeuilDataForm.TABLE_BUD.Columns.Count;

            filt = (string)Globals.FeuilCalculation.X3LOGIN.Value;
            if (filt == null || filt == "")
            {
                this.staret = 1000;
                this.mesret = "user code X3 required";
                return;
            }
            StringBuilder sbkey = new StringBuilder();
            StringBuilder sbkeyVide = new StringBuilder();
            StringBuilder sbval = new StringBuilder();
            String[] tabkeyclb = null, tabvalclb = null;
            List<String> listkeyclb = new List<string>(), listvalclb = new List<string>();
            //String[,] tabvalstr;
            int staret = 0;
            String mesret = "";
            Excel.Range range, range2;

            String url;
            string s, sVide, sValue;
            cAdxCallContext = new CAdxCallContext();

            cAdxCallContext.poolAlias = (string)Globals.FeuilGlobalParams.POOLWS.Value; //"DEMO";
            url = (string)Globals.FeuilGlobalParams.X3WSENDPOINT.Value;
            //cAdxCallContext.codeUser = sheetHome.get_Range("USERX3WS").Value; //"ADMIN";
            //cAdxCallContext.password = sheetHome.get_Range("PASSWDX3WS").Value; //"";
            //cAdxCallContext.codeLang = sheetHome.get_Range("LANGX3WS").Value; // FRA
            Globals.ThisWorkbook.connectionWSX3.setConnect(cAdxCallContext);
            //sheetForBud.get_Range("REP").Value = Globals.ThisWorkbook.connectionWSX3.Login;
            cAdxCallContext.requestConfig = "";

            callWebServiceX3 = new CallWebServiceX3(url, cAdxCallContext,Globals.ThisWorkbook.connectionWSX3.Login, Globals.ThisWorkbook.connectionWSX3.Password);

            filt = (string)Globals.FeuilCalculation.X3LOGIN.Value;
            //range = this.rangeTableBudA;
            listkeyclb.Clear();
            listvalclb.Clear();

            int lig = 0, col = 0, colk = 0;
            nblig = 0;
            range2 = null;
            // je compte les lignes à mettre à jour, je considère que la premiere cle est non null
            while (lig == 0 || range2.Value != null)
            {
                lig += 1;
                col = 0;
                while (colk == 0 && col <= Globals.FeuilDataForm.TABLE_BUD.Columns.Count)
                {
                    col += 1;
                    if (colk == 0)
                    {
                        //range = RangeTableBudA.Cells[2, col];
                        range = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[2, col];
                        descriptionColumn.setDescription(range.Value);
                        if (descriptionColumn.TypUPD == "K")
                        {
                            range2 = rangeTableBudA.Cells[lig + 2, col];
                            if (range2.Value != null)
                            {
                                colk = col;
                                //range2.Value = null;
                                //nbupd += 1;
                                break;
                            }
                        }
                    }
                }

                if (colk > 0)
                {
                    range2 = rangeTableBudA.Cells[lig + 2, colk];
                    if (range2.Value != null)
                    {
                        nblig += 1;

                    }
                }
                else
                {
                    break;
                }
            }

            if (colk == 0)
            {
                this.staret = 300;
                this.mesret = "No Update, please load forecast first";
                return;
            }

            if (
                (int)(nblig / NB_ROWS_UPDATE_BUD)
                ==
                NB_ROWS_UPDATE_BUD
                )
            {
                nbtransac = (int)(nblig / NB_ROWS_UPDATE_BUD);
            }
            else
            {
                nbtransac = (int)(nblig / NB_ROWS_UPDATE_BUD) + 1;
            }


            listkeyclb.Clear();
            listvalclb.Clear();
            lig = 2;
            activeLog(callWebServiceX3);
            for (int i = 0; i < nbtransac; i++) // nb transactions
            {
                if (i > 0 && i <= nbtransac - 2)
                    transac = "Current";
                else if (i == nbtransac - 1)
                    transac = "End";

                listkeyclb.Clear();
                listvalclb.Clear();
                for (int j = 0; j < NB_ROWS_UPDATE_BUD; j++) // nb lignes
                {
                    lig += 1;
                    sbkey.Clear();
                    sbkeyVide.Clear();
                    sbkey.Append("{");
                    sbkeyVide.Append("{");
                    nbkey = 0;

                    sbval.Clear();
                    sbval.Append("{");
                    nbval = 0;
                    col = 0;
                    for (int k = 0; k < nbcolumns; k++) // nb colonnes
                    {

                        //range = RangeTableBudA.Cells[2, k + 1];
                        range = Globals.FeuilHiddenDatas.TABLE_BUD_HIDDEN_VALUE.Cells[2, k + 1];
                        descriptionColumn.setDescription(range.Value);
                        if (descriptionColumn.TypUPD.Equals("K") || descriptionColumn.TypUPD.Equals("V"))
                        {
                            range2 = rangeTableBudA.Cells[lig, k + 1];
                            if (range2.Value != null)
                            {
                                sValue = range2.Value.ToString();
                                if (range2.Value is double)
                                {
                                    sValue = sValue.Replace(",", ".");
                                    sValue = sValue.Replace(" ", "");
                                }
                            }
                            else
                            {
                                sValue = "";
                            }
                            if (descriptionColumn.TypUPD == "K")
                            {
                                sbkey.Append('"');
                                sbkeyVide.Append('"');
                            }
                            else
                                sbval.Append('"');

                            //if (range2.Value != null)
                            //{
                            if (descriptionColumn.TypUPD == "K")
                            {
                                //sbkey.Append(range2.Value);
                                sbkey.Append(sValue);
                                sbkeyVide.Append("");
                            }
                            else
                                //sbval.Append(range2.Value);
                                sbval.Append(sValue);
                            //}
                            //else
                            //{
                            //sb.Append("");
                            //break;
                            //}
                            if (descriptionColumn.TypUPD == "K")
                            {
                                sbkey.Append('"');
                                sbkey.Append(",");
                                sbkeyVide.Append('"');
                                sbkeyVide.Append(",");
                                nbkey += 1;

                            }
                            else
                            {
                                sbval.Append('"');
                                sbval.Append(",");
                                nbval += 1;
                            }

                        }

                    }
                    ///

                    s = sbkey.ToString();
                    s = s.Remove(s.Length - 1);
                    s += "}";
                    sVide = sbkeyVide.ToString();
                    sVide = sVide.Remove(sVide.Length - 1);
                    sVide += "}";
                    if (s != sVide)
                    {
                        listkeyclb.Add(s);
                    }
                    s = sbval.ToString();
                    s = s.Remove(s.Length - 1);
                    s += "}";
                    if (s != "{}")
                    {
                        listvalclb.Add(s);
                    }




                }
                tabkeyclb = listkeyclb.ToArray();
                tabvalclb = listvalclb.ToArray();
                nbtabclb = tabkeyclb.Length;
                callWebServiceX3.UpdData("YUPDDATA", (String)Globals.FeuilGlobalParams.X3TYPWS2.Value, filt, ref nbupd, transac, nbtabclb, nbkey, nbval, tabkeyclb, tabvalclb,
                        ref staret, ref mesret);
                this.staret = staret;
                this.mesret = mesret;
                setFeuilLog(callWebServiceX3);

                if (staret > 0)
                    break;

            }

        }
        

        private void updatePresentation()
        {

        }


    }
}

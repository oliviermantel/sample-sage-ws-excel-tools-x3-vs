using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelWorkbook.Model
{
    class DescriptionColumn
    {
        private string typDisplay;

        public string TypDisplay
        {
            get { return typDisplay; }

        }
        private string nameColumn;

        public string NameColumn
        {
            get { return nameColumn; }

        }
        // actionsAM --> action1;action2;action3; ou action1;action2;action3;
        private String actionsAM;

        public String ActionsAM
        {
            get { return actionsAM; }
        }


        // tabActionsAM --> {action1,action2,action3}
        private String[] tabActionsAM;

        public String[] TabActionsAM
        {
            get { return tabActionsAM; }

        }

        // tabActionsAMLeft --> {action1Left,action2Left,action3Left}
        private String[] tabActionsAMLeft;

        public String[] TabActionsAMLeft
        {
            get { return tabActionsAMLeft; }

        }

        // tabActionsAMRight --> {action1Right,action2Right,action3Right}
        private String[] tabActionsAMRight;

        public String[] TabActionsAMRight
        {
            get { return tabActionsAMRight; }

        }

        private string typUPD;

        public string TypUPD // K or V --> Key ou valeur maj
        {
            get { return typUPD; }
        }

        private string format;

        public DescriptionColumn()
        {
            init();
            transform();
        }
        // object car le format peut être string ou bien entier,... , type d'une cellule
        public DescriptionColumn(Object format)
        {
            init();
            this.format = format.ToString();
            transform();
        }

        public void setDescription(Object format)
        {
            init();
            if (format == null)
                format = "";
            if ((String)format != "")
            {
                this.format = format.ToString();
                transform();
            }
        }

        private void init()
        {
            this.typDisplay = "";
            this.nameColumn = "";
            this.actionsAM = "";
            this.tabActionsAM = null;
            this.tabActionsAMLeft = null;
            this.tabActionsAMRight = null;
            this.typUPD = "";
            this.format = "";

        }
        private void transform()
        // typDisplay::nameColumn::actionsAM
        {
            String[] tabString, tabString2;
            Regex regex1, regex2;
            if (format != null && format != "")
            {
                regex1 = new Regex("::");
                tabString = regex1.Split(format);
                typDisplay = tabString[0];
                if (tabString.Length >= 2)
                    nameColumn = tabString[1];
                if (tabString.Length >= 3)
                {
                    actionsAM = tabString[2];
                    tabActionsAM = getTabActions();
                    regex2 = new Regex("=");
                    tabActionsAMLeft = new String[tabActionsAM.Length];
                    tabActionsAMRight = new String[tabActionsAM.Length];
                    for (int i = 0; i < tabActionsAM.Length; i++)
                    {

                        tabString2 = regex2.Split(tabActionsAM[i]);
                        if (tabString2.Length == 2)
                        {
                            tabActionsAMLeft[i] = tabString2[0];
                            tabActionsAMRight[i] = tabString2[1];
                        }
                    }


                }
                if (tabString.Length >= 4)
                {
                    typUPD = tabString[3];
                }
            }
        }

        private String[] getTabActions()
        {
            // actionsAM --> action1;action2;action3; ou action1;action2;action3;
            String[] ret = null;

            if (actionsAM != null && actionsAM != "")
            {
                Regex regex = new Regex(";");
                ret = regex.Split(actionsAM);

            }

            return ret;
        }
        // RC120 --> {0,120}
        // R12C45 --> {12,45} non traité pour l'instant
        public static int[] getRowColumn(String formatRC)
        {
            int[] ret = { 0, 0 };
            Regex regex = new Regex("RC");
            String[] tabstring;
            tabstring = regex.Split(formatRC);
            if (tabstring.Length >= 1)
            {
                ret = new int[2];
                ret[0] = 0;
                ret[1] = Int16.Parse(tabstring[1]);
            }

            return ret;
        }

        // RC120+RC130 --> tranlastion de col 5 --> RC125 + RC135

        public static String translateRangeColumn(string nameFeuil, int col, String formatRC)
        {
            String ret = "";
            String s_nb;
            char[] tabCar;
            int j, nb;
            tabCar = formatRC.ToCharArray();
            int i = 0;
            while (i < tabCar.Length)
            {
                if (tabCar[i] == 'R')
                {
                    if (tabCar[i + 1] == 'C')
                    {
                        if (nameFeuil != "")
                        {
                            ret += "'" + nameFeuil + "'!" + tabCar[i];
                        }
                        else
                        {
                            ret += tabCar[i];
                        }

                    }
                    else
                    {
                        ret += tabCar[i];
                    }
                    i = i + 1;
                    if (i < tabCar.Length && tabCar[i] == 'C')
                    {
                        ret += tabCar[i];
                        j = i + 1;
                        i = i + 1;
                        s_nb = "";
                        while (j < tabCar.Length && Char.IsDigit(tabCar[j]))
                        {
                            s_nb += tabCar[j];
                            j += 1;
                            i += 1;

                        }

                        if (s_nb != "")
                        {
                            nb = Int32.Parse(s_nb) + col - 1;
                            ret += nb.ToString();
                        }
                    }
                    else { ret += tabCar[i]; i += 1; }
                }
                else
                {
                    ret += tabCar[i];
                    i = i + 1;
                }

            }
            return ret;
        }
    }
}

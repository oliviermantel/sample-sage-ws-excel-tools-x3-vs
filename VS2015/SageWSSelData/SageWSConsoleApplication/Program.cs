using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;
using System.Xml;
using SageWSSelData;


namespace SageWSConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {

            CAdxCallContext cAdxCallContext;
            CallWebServiceX3 callWebServiceX3;
            String url;
            int tabnb = 0, allsel = 0, nextt = 1, nbsel = 0, nbtabval = 0;
            String[] tabcrit = { "TEST", "2013", "", "", "" };


            //List<List<String>> list2valstr = new List<List<String>>();
            String[,] tabvalstr;
            int staret = 0;
            String mesret = "";
            int noReq = 0;
            cAdxCallContext = new CAdxCallContext();
            cAdxCallContext.poolAlias = "CAPBUDXLS";
            //cAdxCallContext.codeUser = "OMA";
            //cAdxCallContext.password = "";
            cAdxCallContext.codeLang = "FRA";
            cAdxCallContext.requestConfig = "";

            url = "http://d01-x3v6:28880/adxwsvc/services/CAdxWebServiceXmlCC?wsdl";

            callWebServiceX3 = new CallWebServiceX3(url, cAdxCallContext,"admin","admin");




            while (allsel != 2)
            {



                Console.Out.WriteLine("nextt -->" + nextt);
                Console.Out.WriteLine("Appel web service");
                callWebServiceX3.SelData("YSELDATA", "TRT(YS06WSSELBUD):SUBPROG(SELBUD)", nextt, tabcrit, ref allsel,
                    ref nbsel, ref tabnb, ref nbtabval, out tabvalstr, ref staret, ref mesret, ref noReq);
                //callWebServiceX3.SelData("YSELDATA", "TRT(YWSSDWM):SUBPROG(SELBUD)", nextt, tabcrit, ref allsel,
                //    ref nbsel, ref tabnb, ref nbtabval, ref list2valstr);
                Console.Out.WriteLine("allsel -->" + allsel);
                Console.Out.WriteLine("nbsel -->" + nbsel);
                Console.Out.WriteLine("tabnb -->" + tabnb);
                Console.Out.WriteLine("staret -->" + staret);
                Console.Out.WriteLine("mesret -->" + mesret);
                Console.Out.WriteLine("noReq -->" + noReq);
                nextt = 2;

                for (int lig = 0; lig < tabnb; lig++)
                {
                    //Console.Out.WriteLine("Ligne -->" + lig);
                    for (int col = 0; col < nbtabval; col++)
                    {
                        //Console.Out.WriteLine("Ligne --> Colonne -->"+lig+"--:"+col+"-->" + tabvalstr[lig,col]);
                    }
                }


                /* foreach (var list in list2valstr)
                 {

                     Console.Out.WriteLine(list);
                     foreach (var cellule in list)
                      {
                     //Console.Out.WriteLine("Cellule -->"+cellule);
                    
                      }
                 }
                 * */
            }

        }


    }
}


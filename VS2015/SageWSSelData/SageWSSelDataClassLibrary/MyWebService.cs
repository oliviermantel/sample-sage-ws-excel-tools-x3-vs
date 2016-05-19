using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace SageWSSelDataClassLibrary
{
    class MyWebService : CAdxWebServiceXmlCCService
    {


        
        /*
        protected override WebRequest GetWebRequest(Uri address)
        {

            
            WebRequest request = (WebRequest)base.GetWebRequest(address);

            
            request.ContentType = "text/xml;charset=UTF-8";
            NetworkCredential netCredential = new NetworkCredential("admin","admin");
            //Uri uri = new Uri("http://52.17.90.248:8124/soap-wsdl/syracuse/collaboration/syracuse/CAdxWebServiceXmlCC?wsdl");
            //Uri uri = new Uri(address);
            ICredentials credentials = netCredential.GetCredential(address, "Basic");
            request.Credentials = credentials;
            request.PreAuthenticate =true;






            
            return request;
        }
        */

        protected override WebRequest GetWebRequest(Uri uri)
        {
            HttpWebRequest request;
            request = (HttpWebRequest)base.GetWebRequest(uri);
            NetworkCredential networkCredentials =
            Credentials.GetCredential(uri, "Basic");
            if (networkCredentials != null)
            {
                byte[] credentialBuffer = new UTF8Encoding().GetBytes(
                networkCredentials.UserName + ":" +
                networkCredentials.Password);
                request.Headers["Authorization"] =
                "Basic " + Convert.ToBase64String(credentialBuffer);
                request.Headers["Cookie"] = "BCSI-CS-2rtyueru7546356=1";
                request.Headers["Cookie2"] = "$Version=1";
            }
            else
            {
                throw new ApplicationException("No network credentials");
            }
            return request;
        }



    }
}

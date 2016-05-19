using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using SageWSSelDataClassLibrary.SageWSWebReferenceSyracuse;

namespace ExcelWorkbook.Model
{
    public class ConnectionWSX3
    {
        private string login;

        public string Login
        {
            get { return login; }
            //set { login = value; }
        }

        public string Password
        {
            get
            {
                return password;
            }
        }

        private string password;

        private string language;


        public ConnectionWSX3(string login, string password, string language)
        {
            this.login = login.ToUpper();
            this.password = password;
            this.language = language;
        }

        public void setConnect(CAdxCallContext cAdxCallContext)
        {
            //cAdxCallContext.codeUser = login;
            //cAdxCallContext.password = password;
            cAdxCallContext.codeLang = language;
        }
    }
}

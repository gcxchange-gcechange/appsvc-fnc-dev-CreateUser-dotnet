using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace appsvc_fnc_dev_CreateUser_dotnet
{
    public class UserInfo
    {
        public string emailcloud { get; set; }
        public string emailwork { get; set; }
        public string firstname { get; set; }
        public string lastname { get; set; }
        public string department { get; set; }
    }

    public class UserEmail
    {
        public string emailUser { get; set; }
        public string firstname { get; set; }
        public string lastname { get; set; }
    }

}
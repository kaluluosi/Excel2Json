using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data
{
    public class CfgProvider
    {
        public static readonly string directoryPath = "";

        public static string Get(string cfgName)
        {
            return directoryPath+cfgName + ".json";
        }
    }
}

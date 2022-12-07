using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Comparator.OKVED.Services
{
    class AssemblyInfo
    {
        public string GetAssemblyInfo()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;

            string name = Assembly.GetExecutingAssembly().GetName().Name;

            string majourVersion = version.Major.ToString();
            string minorVersion = version.Minor.ToString();
            string build = version.Build.ToString();

            return $"{name} {majourVersion.Substring(1,1)}.{minorVersion.Substring(1,1)}";
        }
    }
}

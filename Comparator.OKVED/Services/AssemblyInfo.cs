using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using NLog;

namespace Comparator.OKVED.Services
{
    class AssemblyInfo
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public static string GetAssemblyInfo()
        {
            try
            {
                Version version = Assembly.GetExecutingAssembly().GetName().Version;

                string name = Assembly.GetExecutingAssembly().GetName().Name;

                string majourVersion = version.Major.ToString();
                string minorVersion = version.Minor.ToString();
                string build = version.Build.ToString();

                return $"{name}  ver. {majourVersion}.{minorVersion} build:{build}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                logger.Error(ex.Message + ex.StackTrace);
                return null;
            }
        }
    }
}

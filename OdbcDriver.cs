using System;
using System.Collections.Generic;

namespace Gewatec.Database.Odbc
{
    /// <summary>
    /// Summary description for ODBCDriver.
    /// </summary>
    public class OdbcDriver
    {
        internal OdbcDriver() { }

        public string ODBCDriverName { get; private set; }
        public string APILevel { get; private set; }
        public string ConnectFunctions { get; private set; }
        public string DriverDLL { get; private set; }
        public string DriverODBCVer { get; private set; }

        public string FileExtns { get; private set; }
        public string FileUsage { get; private set; }

        public string SetUp { get; private set; }

        public string SQLLevel { get; private set; }

        public string UsageCount { get; private set; }

        // Additional atttributes for INTERSOLV 3.00 32-BIT ParadoxFile (*.db).
        // Additional atttributes for INTERSOLV 3.11 32-BIT ParadoxFile (*.db).
        // Microsoft ODBC for Oracle has 'm_CPTimeOut' extra but doesn't have
        // 'm_FileExtns' atttribute.
        public string CPTimeOut { get; private set; }
        public string PdxUnInstall { get; private set; }

        public override string ToString()
        {
            return ODBCDriverName;
        }

        public static OdbcDriver ParseForDriver(string driverName, Dictionary<string, string> data)
        {
            OdbcDriver odbcdriver = new OdbcDriver()
            {
                ODBCDriverName = driverName
            };


            // For each element defined for a typical Driver get
            // its value.
            foreach (var driverElement in data)
            {
                string v = driverElement.Value;
                switch (driverElement.Key.ToLower())
                {
                    case "apilevel":
                        odbcdriver.APILevel = v;
                        break;
                    case "connectfunctions":
                        odbcdriver.ConnectFunctions = v;
                        break;
                    case "driver":
                        odbcdriver.DriverDLL = v;
                        break;
                    case "driverodbcver":
                        odbcdriver.DriverODBCVer = v;
                        break;
                    case "fileextns":
                        odbcdriver.FileExtns = v;
                        break;
                    case "fileusage":
                        odbcdriver.FileUsage = v;
                        break;
                    case "setup":
                        odbcdriver.SetUp = v;
                        break;
                    case "sqllevel":
                        odbcdriver.SQLLevel = v;
                        break;
                    case "usagecount":
                        odbcdriver.UsageCount = v;
                        break;
                    case "cptimeout":
                        odbcdriver.CPTimeOut = v;
                        break;
                    case "pdxuninstall":
                        odbcdriver.PdxUnInstall = v;
                        break;
                }
            }
            return odbcdriver;
        }
    }
}

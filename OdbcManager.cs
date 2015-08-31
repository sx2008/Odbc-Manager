using System.Collections;
using Microsoft.Win32;
using System.Collections.Generic;

namespace Gewatec.Database.Odbc
{

    /// <summary>
    /// ODBCManager is the static class which provide
    /// access to the various ODBC components such as Drivers List, DSNs List
    /// etc. 
    /// </summary>
    public static class OdbcManager
    {
        private const string ODBC_LOC_IN_REGISTRY = @"SOFTWARE\ODBC\";
        private const string ODBC_INI_LOC_IN_REGISTRY =
            ODBC_LOC_IN_REGISTRY + @"ODBC.INI\";

        private const string DSN_LOC_IN_REGISTRY =
            ODBC_INI_LOC_IN_REGISTRY + @"ODBC Data Sources\";

        private const string ODBCINST_INI_LOC_IN_REGISTRY =
            ODBC_LOC_IN_REGISTRY + @"ODBCINST.INI\";

        private const string ODBC_DRIVERS_LOC_IN_REGISTRY =
            ODBCINST_INI_LOC_IN_REGISTRY + @"ODBC Drivers\";


        // Set this Variable to RegistryView.Registry32 if you want to access 32-Bit ODBC Settings
        // from a 64-Bit Application
        // Set this Variable to RegistryView.Registry64 if you want to access 64-Bit ODBC Settings
        // from a 32-Bit Application
        public static RegistryView RegistryViewType = RegistryView.Default;


        private static RegistryKey GetLocalMachineKey()
        {
            return RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryViewType);
        }

        private static RegistryKey GetCurrentUserKey()
        {
            return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryViewType);
        }

        /// <summary>
        /// returns a list with all ODBC driver names
        /// </summary>
        public static IEnumerable<string> GetODBCDriverNames()
        {
            // Get the key for
            // "KHEY_LOCAL_MACHINE\\SOFTWARE\\ODBC\\ODBCINST.INI\\ODBC Drivers\\"
            // (ODBC_DRIVERS_LOC_IN_REGISTRY) that contains all the drivers
            // that are installed in the local machine.
            using (RegistryKey odbcDrvLocKey = GetLocalMachineKey().OpenSubKey(ODBC_DRIVERS_LOC_IN_REGISTRY, false))
            {
                if (odbcDrvLocKey != null)
                {
                    // Get all Driver entries defined in ODBC_DRIVERS_LOC_IN_REGISTRY.
                    string[] driverNames = odbcDrvLocKey.GetValueNames();
                    if (driverNames != null)
                    {
                        return driverNames;
                    }
                }
            }

            return null;
        }



        /// <summary>
        /// Method that gives the ODBC Drivers installed in the local machine.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<OdbcDriver> GetODBCDrivers()
        {
            List<OdbcDriver> result = new List<OdbcDriver>();
            var driverNames = GetODBCDriverNames();

            if (driverNames != null)
            {
                // Foreach Driver entry in the ODBC_DRIVERS_LOC_IN_REGISTRY,
                // goto the Key ODBCINST_INI_LOC_IN_REGISTRY+driver and get
                // elements of the DSN entry to create ODBCDSN objects.
                foreach (string driverName in driverNames)
                {
                    OdbcDriver odbcDriver = GetODBCDriver(driverName);
                    if (odbcDriver != null)
                        result.Add(odbcDriver);
                }
            }

            return result;
        }

        /// <summary>
        /// Method that returns driver object based on the driver name.
        /// </summary>
        /// <param name="driverName"></param>
        /// <returns>ODBCDriver object</returns>
        public static OdbcDriver GetODBCDriver(string driverName)
        {
            OdbcDriver odbcDriver = null;

            // Get the key for ODBCINST_INI_LOC_IN_REGISTRY+dsnName.
            using (RegistryKey driverNameKey = GetLocalMachineKey().OpenSubKey(ODBCINST_INI_LOC_IN_REGISTRY + driverName, false))
            {
                if (driverNameKey != null)
                {
                    // Get all elements defined in the above key
                    string[] driverElements = driverNameKey.GetValueNames();

                    var data = new Dictionary<string, string>(driverElements.Length);

                    // For each element defined for a typical Driver get
                    // its value.
                    foreach (string driverElement in driverElements)
                    {
                        data.Add(driverElement, driverNameKey.GetValue(driverElement).ToString());
                    }

                    // Create ODBCDriver Object.
                    odbcDriver = OdbcDriver.ParseForDriver(driverName, data);
                }
                return odbcDriver;
            }
        }
        /// <summary>
        /// Method that gives the System Data Source Name (DSN) entries as
        /// array of ODBCDSN objects.
        /// </summary>
        /// <returns>Array of System DSNs</returns>
        public static IList<OdbcDsn> GetSystemDSNList()
        {
            return GetDSNList(GetLocalMachineKey());
        }

        /// <summary>
        /// Method that returns one System ODBCDSN Object.
        /// </summary>
        /// <param name="dsnName"></param>
        /// <returns></returns>
        public static OdbcDsn GetSystemDSN(string dsnName)
        {
            return GetDSN(GetLocalMachineKey(), dsnName);
        }

        /// <summary>
        /// Method that gives the User Data Source Name (DSN) entries as
        /// array of ODBCDSN objects.
        /// </summary>
        /// <returns>Array of User DSNs</returns>
        public static IList<OdbcDsn> GetUserDSNList()
        {
            var list = GetDSNList(GetCurrentUserKey());
            foreach (var dsn in list)
                dsn.UserDSN = true;
            return list;
        }

        /// <summary>
        /// Method that returns one User ODBCDSN Object.
        /// </summary>
        /// <param name="dsnName"></param>
        /// <returns></returns>
        public static OdbcDsn GetUserDSN(string dsnName)
        {
            var dsn = GetDSN(GetCurrentUserKey(), dsnName);
            if (dsn != null)
                dsn.UserDSN = true;
            return dsn;
        }

        /// <summary>
        /// returns the UserDSN (if exists) else return the SystemDSN
        /// </summary>
        /// <param name="dsnName"></param>
        /// <returns></returns>
        public static OdbcDsn GetDSN(string dsnName)
        {
            var dsn = GetUserDSN(dsnName);
            if (dsn == null)
                dsn = GetSystemDSN(dsnName);
            return dsn;
        }


        /// <summary>
        /// returns a list of all DSNs (UserDNSs override SystemDSNs)
        /// </summary>
        /// <returns></returns>
        public static IList<OdbcDsn> GetDSNList()
        {
            IList<OdbcDsn> result = GetSystemDSNList();

            IList<OdbcDsn> user = GetUserDSNList();

            foreach (var dsn in user)
            {
                bool found = false;
                for (int i = 0; i < result.Count; i++)
                {
                    if (result[i].DSNName == dsn.DSNName)
                    {
                        result[i] = dsn;
                        found = true;
                        break;
                    }
                }
                if (!found)
                    result.Add(dsn);
            }
            return result;
        }



        /// <summary>
        /// Method that gives the Data Source Name (DSN) entries as list of
        /// ODBCDSN objects.
        /// </summary>
        /// <returns>list of DSNs based on the baseKey parameter</returns>
        private static List<OdbcDsn> GetDSNList(RegistryKey baseKey)
        {
            if (baseKey == null)
                return null;

            List<OdbcDsn> dsnList = new List<OdbcDsn>();

            // Get the key for (using the baseKey parmetre passed in)
            // "\\SOFTWARE\\ODBC\\ODBC.INI\\ODBC Data Sources\\" (DSN_LOC_IN_REGISTRY)
            // that contains all the configured Data Source Name (DSN) entries.
            using (RegistryKey dsnNamesKey = baseKey.OpenSubKey(DSN_LOC_IN_REGISTRY, false))
            {
                if (dsnNamesKey == null)
                    return null;

                // Get all DSN entries defined in DSN_LOC_IN_REGISTRY.
                string[] dsnNames = dsnNamesKey.GetValueNames();
                if (dsnNames != null)
                {
                    // Foreach DSN entry in the DSN_LOC_IN_REGISTRY, goto the
                    // Key ODBC_INI_LOC_IN_REGISTRY+dsnName and get elements of
                    // the DSN entry to create ODBCDSN objects.
                    foreach (string dsnName in dsnNames)
                    {
                        // Get ODBC DSN object.
                        OdbcDsn odbcDSN = GetDSN(baseKey, dsnName);
                        if (odbcDSN != null)
                            dsnList.Add(odbcDSN);
                    }
                }

                return dsnList;
            }
        }

        /// <summary>
        /// Method that gives one ODBC DSN object
        /// </summary>
        /// <param name="baseKey"></param>
        /// <param name="dsnName"></param>
        /// <returns>ODBC DSN object</returns>
        private static OdbcDsn GetDSN(RegistryKey baseKey, string dsnName)
        {
            string dsnDriverName;

            // Get the key for (using the baseKey parmetre passed in)
            // "\\SOFTWARE\\ODBC\\ODBC.INI\\ODBC Data Sources\\" (DSN_LOC_IN_REGISTRY)
            // that contains all the configured Data Source Name (DSN) entries.
            using (RegistryKey dsnNamesKey = baseKey.OpenSubKey(DSN_LOC_IN_REGISTRY, false))
            {
                if (dsnNamesKey == null)
                    return null;

                object value = dsnNamesKey.GetValue(dsnName);

                // Get the name of the driver for which the DSN is 
                // defined.
                dsnDriverName = (value != null) ? value.ToString() : null;
            }

            // Get the key for ODBC_INI_LOC_IN_REGISTRY+dsnName.
            using (RegistryKey dsnNameKey = baseKey.OpenSubKey(ODBC_INI_LOC_IN_REGISTRY + dsnName, false))
            {
                OdbcDsn odbcDSN = null;
                if (dsnNameKey != null)
                {
                    // Get all elements defined in the above key
                    string[] dsnElements = dsnNameKey.GetValueNames();

                    Dictionary<string, string> data = new Dictionary<string, string>(dsnElements.Length);

                    // For each element defined for a typical DSN get
                    // its value.
                    foreach (string dsnElement in dsnElements)
                    {
                        data.Add(dsnElement, dsnNameKey.GetValue(dsnElement).ToString());
                    }

                    // Create ODBCDSN Object.
                    odbcDSN = OdbcDsn.ParseForODBCDSN(dsnName, dsnDriverName, data);
                }
                return odbcDSN;
            }

        }

    }
}

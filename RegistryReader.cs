using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArchivarTIA
{
    /// <summary>
    /// Thanks to K_Ohjus of support.industry.siemens forum for this class!!
    /// </summary>
    class RegistryReader
    {

        private const string BASE_PATH = "SOFTWARE\\Siemens\\Automation\\Openness\\";

        public static List<string> GetVersions()
        {
            RegistryKey key = GetRegistryKey(BASE_PATH);

            if (key != null)
            {
                var names = key.GetSubKeyNames().OrderBy(x => x).ToList();
                key.Dispose();

                return names;
            }

            return new List<string>();
        }
        /// <summary>
        /// Searches registry for all assemblies under selected TIA version.
        /// </summary>
        /// <param name="version">Selected assembly version</param>
        /// <returns>List of assemblies</returns>
        public static List<string> GetAssemblies(string version)
        {
            RegistryKey key = GetRegistryKey(BASE_PATH + version);

            if (key != null)
            {
                try
                {
                    var subKey = key.OpenSubKey("PublicAPI");
                    var result = subKey.GetSubKeyNames().OrderBy(x => x).ToList();

                    subKey.Dispose();

                    return result;
                }
                finally
                {
                    key.Dispose();
                }
            }

            return new List<string>();
        }
        /// <summary>
        /// Searches registry for assemblies outputs paths to Siemens.Engineering and Siemens.Engineering.Hmi DLLs.
        /// </summary>
        /// <param name="version">Selected assembly version</param>
        /// <param name="assemblyPath">Path to Siemens.Engineering.DLL</param>
        /// <param name="assemblyPathHmi">Path to Siemens.Engineering.Hmi.DLL</param>
        public static void GetAssemblyPath(string version, string assembly, out string assemblyPath, out string assemblyPathHmi)
        {
            assemblyPath = "";
            assemblyPathHmi = "";
            RegistryKey key = GetRegistryKey(BASE_PATH + version + "\\PublicAPI\\" + assembly);

            if (key != null)
            {
                try
                {
                    assemblyPath = key.GetValue("Siemens.Engineering").ToString();
                    assemblyPathHmi = key.GetValue("Siemens.Engineering.Hmi").ToString();

                }
                finally
                {
                    key.Dispose();
                }
            }

        }
        /// <summary>
        /// Gets any key from Registry.
        /// </summary>
        /// <param name="keyname">Selected assembly version</param>
        /// /// <returns>key</returns>
        private static RegistryKey GetRegistryKey(string keyname)
        {
            //HKEY_LOCAL_MACHINE
            RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey key = baseKey.OpenSubKey(keyname);
            //If key is not found in 64-bit registry
            if (key == null)
            {
                baseKey.Dispose();
                baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
                key = baseKey.OpenSubKey(keyname);
            }
            if (key == null)
            {
                baseKey.Dispose();
                baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
                key = baseKey.OpenSubKey(keyname);
            }
            baseKey.Dispose();

            return key;

        }
    }
}

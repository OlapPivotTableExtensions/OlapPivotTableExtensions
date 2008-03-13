using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace CustomActions
{
    /// <summary>
    /// This class takes advantage of the User Settings registry in which Office will spread HKLM registry settings to the HKCU branch
    /// </summary>
    internal class ManageUserSettings
    {
        private static string REGISTRY_PATH = @"Software\Microsoft\Office\12.0\User Settings\OlapPivotTableExtensions";

        public static void Install(bool bAllUsers)
        {
            //whether installing for all users or not, go ahead and remove the delete key
            IncrementCount();
            RemoveDeleteInstruction();
        }

        public static void Uninstall(bool bAllUsers)
        {
            if (bAllUsers)
            {
                IncrementCount();
                RegisterDeleteInstruction();
            }
        }

        /// <summary>
        /// necessary to increment the counter with ever action that's taken because that's what signals to Office to run these updates (if the Count key under HKLM is greater than the Count key under HKCU)
        /// </summary>
        private static void IncrementCount()
        {
            RegistryKey appKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(REGISTRY_PATH);

            object oCount = appKey.GetValue("Count");
            if (oCount == null)
            {
                appKey.SetValue("Count", 1);
            }
            else
            {
                appKey.SetValue("Count", Convert.ToInt32(oCount) + 1);
            }

            appKey.Close();
        }


        /// <summary>
        /// If a previous uninstall set the Delete registry key, it this function will remove it
        /// </summary>
        private static void RemoveDeleteInstruction()
        {
            RegistryKey appKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(REGISTRY_PATH);

            RegistryKey deleteKey = appKey.OpenSubKey("Delete", false);
            if (deleteKey != null)
            {
                deleteKey.Close();
                appKey.DeleteSubKeyTree("Delete");
            }

            appKey.Close();
        }

        /// <summary>
        /// Create a Delete registry key so that future executions of Office will uninstall the add-in from the HKCU registry branch
        /// </summary>
        private static void RegisterDeleteInstruction()
        {
            RegistryKey appKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(REGISTRY_PATH);
            appKey.CreateSubKey(@"Delete\Software\Microsoft\Office\Excel\AddIns\OlapPivotTableExtensions");
            appKey.Close();
        }
        
    }
}

//-----------------------------------------------------------------------
// 
//  Copyright (C) Microsoft Corporation.  All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;

namespace CustomActions
{
    [RunInstaller(true)]
    [System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class CustomInstallActions : Installer
    {
        private static string _TargetDir;

        public CustomInstallActions()
        {
            InitializeComponent();
        }

        public override void Install(System.Collections.IDictionary stateSaver)
        {
            try
            {
                // Call the base implementation.
                base.Install(stateSaver);

                string allUsersString = this.Context.Parameters["allUsers"];
                string targetDir = this.Context.Parameters["targetDir"];

                _TargetDir = targetDir;
                if (String.IsNullOrEmpty(targetDir))
                    throw new InstallException("Cannot set the security policy. The specified target directory is not valid.");
                if (stateSaver == null)
                    throw new ArgumentNullException("stateSaver");

                bool allUsers = String.Equals(allUsersString, "1");

                ManageUserSettings.Install(allUsers);

                ClearDisabledItems.CheckDisabledItems(targetDir);
            }
            catch (Exception ex)
            {
                //doesn't appear to report this exception to the user, so we have to do it ourselves
                System.Windows.Forms.MessageBox.Show("Problem in CustomInstallActions.Install: " + ex.Message + "\r\n" + ex.StackTrace);
                throw;
            }
        }

        public override void Rollback(System.Collections.IDictionary savedState)
        {
            try
            {
                try
                {
                    string targetDir = this.Context.Parameters["targetDir"];
                    _TargetDir = targetDir;
                }
                catch { }

                // Call the base implementation.
                base.Rollback(savedState);

                // Check whether the "allUsers" property is saved.
                // If it is not set, the Install method did not set the security policy.
                if ((savedState == null) || (savedState["allUsers"] == null))
                    return;

                bool allUsers = (bool)savedState["allUsers"];

                ManageUserSettings.Uninstall(allUsers);
            }
            catch (Exception ex)
            {
                //doesn't appear to report this exception to the user, so we have to do it ourselves
                System.Windows.Forms.MessageBox.Show("Problem in CustomInstallActions.Rollback: " + ex.Message + "\r\n" + ex.StackTrace);
                throw;
            }
        }


        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            try
            {
                try
                {
                    string targetDir = this.Context.Parameters["targetDir"];
                    _TargetDir = targetDir;
                }
                catch { }

                // Call the base implementation.
                base.Uninstall(savedState);

                // Check whether the "allUsers" property is saved.
                // If it is not set, the Install method did not set the security policy.
                if ((savedState == null) || (savedState["allUsers"] == null))
                    return;

                bool allUsers = (bool)savedState["allUsers"];

                ManageUserSettings.Uninstall(allUsers);
            }
            catch (Exception ex)
            {
                //doesn't appear to report this exception to the user, so we have to do it ourselves
                System.Windows.Forms.MessageBox.Show("Problem in CustomInstallActions.Uninstall: " + ex.Message + "\r\n" + ex.StackTrace);
                throw;
            }
        }

    }
}
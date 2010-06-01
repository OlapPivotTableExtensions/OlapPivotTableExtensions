// NoImpersonate.js <msi-file>
// Performs a post-build fixup of an msi to change all deferred custom actions to NoImpersonate

// Constant values from Windows Installer
var msiOpenDatabaseModeTransact = 1;

var msiViewModifyInsert         = 1
var msiViewModifyUpdate         = 2
var msiViewModifyAssign         = 3
var msiViewModifyReplace        = 4
var msiViewModifyDelete         = 6

var msidbCustomActionTypeInScript       = 0x00000400;
var msidbCustomActionTypeNoImpersonate = 0x00000800;

var msidbLocatorType64bit = 0x010;
var msidbComponentAttributes64bit = 0x00000100;



if (WScript.Arguments.Length < 2) {
	WScript.StdErr.WriteLine(WScript.ScriptName + " file - too few arguments");
	WScript.Quit(1);
}

var filespec = WScript.Arguments(0);
var configuration = WScript.Arguments(1);

var installer = WScript.CreateObject("WindowsInstaller.Installer");
var database = installer.OpenDatabase(filespec, msiOpenDatabaseModeTransact);

if (configuration == "Release64") {
    WScript.StdOut.WriteLine("Configuration set to Release64 so changing MSI Platform in SummaryInformation to x64");
    var summaryInfo = database.SummaryInformation(1);
    summaryInfo.Property(7) = "x64; 1033";
    summaryInfo.Persist();
}




var sql
var view
var record

try
{
	sql = "SELECT `Action`, `Type`, `Source`, `Target` FROM `CustomAction`";
	view = database.OpenView(sql);
	view.Execute();
	record = view.Fetch();
	while (record)
	{
	    if (record.IntegerData(2) & msidbCustomActionTypeInScript)
	    {
	        record.IntegerData(2) = record.IntegerData(2) | msidbCustomActionTypeNoImpersonate;
        	view.Modify(msiViewModifyReplace, record);
        	WScript.StdOut.WriteLine("Changed CustomAction " + record.StringData(1) + " to NoImpersonate");
     }
        if (configuration == "Release64" && record.StringData(1) == "DIRCA_TARGETDIR") {
            WScript.StdOut.WriteLine("Configuration set to Release64 so changing MSI TARGETDIR to be 64-bit Program Files");
            record.StringData(4) = record.StringData(4).replace("[ProgramFilesFolder]", "[ProgramFiles64Folder]");
            view.Modify(msiViewModifyReplace, record);
        }
        record = view.Fetch();
	}

	view.Close();
	database.Commit();
}
catch(e)
{
	WScript.StdErr.WriteLine(e);
	WScript.Quit(1);
}

//fix the VersionMin property so that it will properly delete prior versions during install
try
{
	sql = "SELECT `VersionMin`, `VersionMax`, `ActionProperty` FROM `Upgrade`";
	view = database.OpenView(sql);
	view.Execute();
	record = view.Fetch();
	while (record)
	{
	    if (record.StringData(3) == "PREVIOUSVERSIONSINSTALLED")
	    {
	        record.StringData(1) = "0.1.0.0"; //this defaults to 1.0.0.0 which is greater than the first version we released
        	view.Modify(msiViewModifyReplace, record);
        }
        record = view.Fetch();
	}

	view.Close();
	database.Commit();
}
catch(e)
{
	WScript.StdErr.WriteLine(e);
	WScript.Quit(1);
}




try {
    if (configuration == "Release64") {
        sql = "SELECT `AppSearch`.`Property`, `RegLocator`.`Key`, `RegLocator`.`Name`, `RegLocator`.`Type` FROM `AppSearch`, `RegLocator` where `AppSearch`.`Signature_` = `RegLocator`.`Signature_`";
        view = database.OpenView(sql);
        view.Execute();
        record = view.Fetch();
        while (record) {
            if (record.StringData(1).substring(record.StringData(1).length - 3) == "X64" && (record.IntegerData(4) & msidbLocatorType64bit) == 0) {
                //if a 64-bit build, have certain registry searches use the 64-bit registry
                WScript.StdOut.WriteLine("Configuration set to Release64 so changing Registry Search " + record.StringData(1) + " (" + record.StringData(2) + ") to look in 64-bit hive");
                record.IntegerData(4) = record.IntegerData(4) + msidbLocatorType64bit;
                view.Modify(msiViewModifyUpdate, record);
            }
            record = view.Fetch();
        }

        view.Close();
        database.Commit();
    }
}
catch (e) {
    WScript.StdErr.WriteLine(e);
    WScript.Quit(1);
}



try {
    if (configuration == "Release64") {
        sql = "SELECT `Component`.`Attributes`, `Registry`.`Key`, `Registry`.`Name` FROM `Component`, `Registry` where `Component`.`Component` = `Registry`.`Component_`";
        view = database.OpenView(sql);
        view.Execute();
        record = view.Fetch();
        while (record) {
            if ((record.IntegerData(1) & msidbComponentAttributes64bit) == 0) {
                //if a 64-bit build, have certain registry searches use the 64-bit registry
                WScript.StdOut.WriteLine("Configuration set to Release64 so changing Registry " + record.StringData(2) + "\\" + record.StringData(3) + " to use 64-bit hive");
                record.IntegerData(1) = record.IntegerData(1) + msidbComponentAttributes64bit;
                view.Modify(msiViewModifyUpdate, record);
            }
            record = view.Fetch();
        }

        view.Close();
        database.Commit();
    }
}
catch (e) {
    WScript.StdErr.WriteLine(e);
    WScript.Quit(1);
}




try {
    if (configuration == "Release64") {
        sql = "INSERT INTO `Property` (`Property`, `Value`) VALUES ('INSTALLER64', 1)";
        view = database.OpenView(sql);
        view.Execute();
        view.Close();
        database.Commit();
    }
}
catch (e) {
    WScript.StdErr.WriteLine(e);
    WScript.Quit(1);
}

    


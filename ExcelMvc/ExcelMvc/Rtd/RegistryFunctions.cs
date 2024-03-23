using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelMvc.Rtd
{
    /*
    /// COM class:
    /// 
    /// namespace MyServer
    /// {
    ///     [ClassInterface(ClassInterfaceType.None), Guid("3346EEFA-D567-447A-92A9-B941D1BAB751"), ProgId("MyServer.MyObject")]
    ///     public class MyObject
    ///     {
    ///     }
    /// }
    /// 
    /// Registry:
    ///   .NET Framework
    ///   Windows Registry Editor Version 5.00 
    /// 
    ///   [HKEY_CLASSES_ROOT\CLSID\{3346EEFA-D567-447A-92A9-B941D1BAB751}]
    ///   @="MyServer.MyObject"
    /// 
    ///   [HKEY_CLASSES_ROOT\CLSID\{3346EEFA-D567-447A-92A9-B941D1BAB751}\Implemented Categories]
    /// 
    ///   [HKEY_CLASSES_ROOT\CLSID\{3346EEFA-D567-447A-92A9-B941D1BAB751}\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}]
    /// 
    ///   [HKEY_CLASSES_ROOT\CLSID\{3346EEFA-D567-447A-92A9-B941D1BAB751}\InprocServer32]
    ///   @="mscoree.dll"
    ///   "ThreadingModel"="Both"
    ///   "Class"="MyServer.MyObject"
    ///   "Assembly"="ClassLibrary3, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    ///   "RuntimeVersion"="v4.0.30319"
    ///   "CodeBase"="file:////D:/Temp/ClassLibrary3/ClassLibrary3/bin/Debug/ClassLibrary3.dll"
    /// 
    ///   [HKEY_CLASSES_ROOT\CLSID\{3346EEFA-D567-447A-92A9-B941D1BAB751}\InprocServer32\1.0.0.0]
    ///   "Class"="MyServer.MyObject"
    ///   "Assembly"="ClassLibrary3, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    ///   "RuntimeVersion"="v4.0.30319"
    ///   "CodeBase"="file:////D:/Temp/ClassLibrary3/ClassLibrary3/bin/Debug/ClassLibrary3.dll"
    /// 
    ///   [HKEY_CLASSES_ROOT\CLSID\{3346EEFA-D567-447A-92A9-B941D1BAB751}\ProgId]
    ///   @="MyServer.MyObject"
    ///   
    ///   .NET Core
    ///   Windows Registry Editor Version 5.00
    ///   [HKEY_CLASSES_ROOT\CLSID\{9F35B6F5-2C05-4E7F-93AA-EE087F6E7AB6}]
    ///    @= "CoreCLR COMHost Server"
    ///
    ///   [HKEY_CLASSES_ROOT\CLSID\{9F35B6F5-2C05-4E7F-93AA-EE087F6E7AB6}\InProcServer32]
    ///   @= "D:\\Temp\\classlib\\classlib\\bin\\Debug\\net6.0\\classlib.comhost.dll"
    ///   "ThreadingModel" = "Both"
    ///   [HKEY_CLASSES_ROOT\CLSID\{ 9F35B6F5 - 2C05 - 4E7F - 93AA - EE087F6E7AB6}\ProgID]
    ///   @= "classlib.server"
    ///   
    /// 
    ///   [HKEY_CLASSES_ROOT\classlib.server]
    ///   @="classlib.Server"
    ///   [HKEY_CLASSES_ROOT\classlib.server\CLSID]
    ///   @="{9F35B6F5-2C05-4E7F-93AA-EE087F6E7AB6}"
    ///
    ///
    /// Replace HKEY_CLASSES_ROOT with HKEY_CURRENT_USER\Software\Classes.
    /// 
    */
    public static class RegistryFunctions
    {
        public static (string progId, Guid clsId) Register()
        {
            var clsId = Guid.NewGuid();
            var guid = clsId.ToString("B").ToUpperInvariant();
            var progId = $"ExcelMvc.{clsId.ToString("N")}";

            foreach (var key in OpenClassesKeys())
            {
                using (key)
                {
                    ///[HKEY_CURRENT_USER\Software\Classes\Prog.ID]
                    using (var keyProgID = CreateSubKey(key, progId))
                    {
                        keyProgID.SetValue(null, progId);
                        ///[HKEY_CURRENT_USER\Software\Classes\Prog.ID\CLSID]
                        using (var x = CreateSubKey(keyProgID, "CLSID")) x.SetValue(null, guid);
                        using (var x = CreateSubKey(keyProgID, "Time")) x.SetValue(null, $"{DateTime.Now:O}");
                    }

                    ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}]
                    using (var clsKey = OpenSubKey(key, "CLSID") ?? CreateSubKey(key, "CLSID"))
                    {
                        using (var keyGuid = CreateSubKey(clsKey, guid))
                        {
                            keyGuid.SetValue(null, progId);
                            ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\ProgId]
                            using (var x = CreateSubKey(keyGuid, "ProgId")) x.SetValue(null, progId);

                            ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\InprocServer32]
                            using (var inprocServer32 = keyGuid.CreateSubKey("InProcServer32"))
                            {
                                inprocServer32.SetValue(null, AddIn.ModuleFileName);
                                inprocServer32.SetValue("ThreadingModel", "Both");
                            }
                        }
                    }
                }
            }

            return (progId, clsId);
        }

        public static void Unregister(string progId)
        {
            DeleteProgId(progId);
        }

        public const string ProgIdPattern = "ExcelMvc.(.)+";
        public static void PurgeProgIds()
        {
            var pattern = new Regex(ProgIdPattern);
            foreach (var key in OpenClassesKeys())
            {
                using (key)
                {
                    foreach (var progId in key.GetSubKeyNames().Where(x => pattern.IsMatch(x)))
                    {
                        using (var progKey = key.OpenSubKey(progId))
                        using (var timeKey = progKey.OpenSubKey("Time"))
                        {
                            var time = $"{timeKey.GetValue(null)}";
                            if (string.IsNullOrWhiteSpace(time) || !DateTime.TryParse(time, out var x) ||
                                (DateTime.Now - x).TotalSeconds > 60)
                                DeleteProgId(progId);
                        }
                    }
                }
            }
        }

        public static void DeleteProgId(string progId)
        {
            var guids = new List<string>();
            foreach (var key in OpenClassesKeys())
            {
                using (key)
                {
                    using (var progKey = key.OpenSubKey(progId))
                    {
                        if (progKey == null) continue;
                        using (var guidKey = progKey.OpenSubKey("CLSID"))
                        {
                            guids.Add($"{guidKey.GetValue(null)}");
                        }
                        key.DeleteSubKeyTree(progId, false);
                    }
                }
            }

            foreach (var key in OpenClassesKeys())
            {
                using (key)
                {
                    using (var clsKey = OpenSubKey(key, "CLSID"))
                    {
                        if (clsKey == null) continue;
                        foreach (var guid in guids)
                            clsKey.DeleteSubKeyTree(guid, false);
                    }
                }
            }
        }

        /*
        private static void SetKeyValues(RegistryKey key, Type type, bool versionNode)
        {
            if (!versionNode)
            {
                key.SetValue(null, "mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
            }
            key.SetValue("Class", type.FullName);
            key.SetValue("Assembly", type.Assembly.FullName);
            key.SetValue("RuntimeVersion", type.Assembly.ImageRuntimeVersion);
            key.SetValue("CodeBase", type.Assembly.Location);
        }*/

        private const string ClassesPath = @"Software\Classes\";
        public static RegistryKey[] OpenClassesKeys()
        {
            var x86 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            var x64 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
            return new[] { OpenSubKey(x86, ClassesPath), OpenSubKey(x64, ClassesPath) };
        }

        private static RegistryKey OpenSubKey(RegistryKey key, string path)
            => key.OpenSubKey(path, RegistryKeyPermissionCheck.ReadWriteSubTree);

        private static RegistryKey CreateSubKey(RegistryKey key, string path)
            => key.CreateSubKey(path, RegistryKeyPermissionCheck.ReadWriteSubTree);
    }
}

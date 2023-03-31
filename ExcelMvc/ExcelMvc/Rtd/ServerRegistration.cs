using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


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
    public static class ServerRegistration
    {
        private const string ClassesPath = @"Software\Classes\";

        public static string RegisterType(Type type)
        {
            var x86 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            var x64 = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);

            var keys = new[]
            {
                x86.OpenSubKey(ClassesPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl),
                x64.OpenSubKey(ClassesPath, RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl)
            };

            var progId = GetProgId(type);
            var guid = $"{{{GetGuid(type)}}}";


            foreach (var key in keys)
            {
                ///[HKEY_CURRENT_USER\Software\Classes\Prog.ID]
                var keyProgID = key.CreateSubKey(progId);
                keyProgID.SetValue(null, progId);

                ///[HKEY_CURRENT_USER\Software\Classes\Prog.ID\CLSID]
                keyProgID.CreateSubKey(@"CLSID").SetValue(null, guid);


                ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}]
                var keyCLSID = key.OpenSubKey(@"CLSID", RegistryKeyPermissionCheck.ReadWriteSubTree,
                    System.Security.AccessControl.RegistryRights.FullControl).CreateSubKey(guid);
                keyCLSID.SetValue(null, progId);

                ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\ProgId]
                keyCLSID.CreateSubKey("ProgId").SetValue(null, progId);

#if NET5_0_OR_GREATER
                ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\InprocServer32]
                var inprocServer32 = keyCLSID.CreateSubKey("InProcServer32");
                inprocServer32.SetValue(null, type.Assembly.Location.ToLower().Replace(".dll", ".comhost.dll"));
                inprocServer32.SetValue("ThreadingModel", "Both");

#else
                ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\InprocServer32]
                var inprocServer32 = keyCLSID.CreateSubKey("InprocServer32");

                SetKeyValues(inprocServer32, type, false);
                ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\InprocServer32\1.0.0.0]
                SetKeyValues(inprocServer32.CreateSubKey("Version"), type, true);
                ///[HKEY_CURRENT_USER\Software\Classes\CLSID\{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}]
                keyCLSID.CreateSubKey(@"Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
#endif
                keyCLSID.Close();
            }

            return progId;
        }

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
            key.SetValue("CodeBase", type.Assembly.CodeBase);
        }

        private static string GetProgId(Type type) => type.GetCustomAttributes(typeof(ProgIdAttribute), false)
            .Cast<ProgIdAttribute>().Single().Value;
        private static string GetGuid(Type type) => type.GetCustomAttributes(typeof(GuidAttribute), false)
            .Cast<GuidAttribute>().Single().Value;
    }
}

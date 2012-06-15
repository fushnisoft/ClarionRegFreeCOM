using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace UCTest
{
    /// <summary>
    /// Base class for COM interop user controls to derive from.
    /// This class handles all the details of exposing the public methods for interop.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ActiveXControl : System.Windows.Forms.UserControl
    {
        [ComRegisterFunction]
        static void ComRegister(Type t)
        {
            string keyName = @"CLSID\" + t.GUID.ToString("B");
            using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(keyName, true))
            {
                key.CreateSubKey("Control").Close();
                using (RegistryKey subkey = key.CreateSubKey("MiscStatus"))
                {
                    subkey.SetValue("", "131457");
                }
                using (RegistryKey subkey = key.CreateSubKey("TypeLib"))
                {
                    Guid libid = Marshal.GetTypeLibGuidForAssembly(t.Assembly);
                    subkey.SetValue("", libid.ToString("B"));
                }
                using (RegistryKey subkey = key.CreateSubKey("Version"))
                {
                    Version ver = t.Assembly.GetName().Version;
                    string version = string.Format("{0}.{1}", ver.Major, ver.Minor);
                    if (version == "0.0")
                        version = "1.0";
                    subkey.SetValue("", version);
                }
            }
        }

        [ComUnregisterFunction]
        static void ComUnregister(Type t)
        {
            Registry.ClassesRoot.DeleteSubKeyTree(@"CLSID\" + t.GUID.ToString("B"));
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // ActiveXControl
            // 
            this.Name = "ActiveXControl";
            this.Size = new System.Drawing.Size(128, 38);
            this.ResumeLayout(false);

        }

    }

}

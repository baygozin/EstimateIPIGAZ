using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace EstimatesName {
    class Ini {
        public string Path;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public Ini(string iniPath) {
            Path = iniPath;
        }

        public void IniWriteValue(string section, string key, string value) {
            var inipath = System.IO.Path.GetDirectoryName(Path);
            if (inipath != null && !Directory.Exists(inipath))
                Directory.CreateDirectory(inipath);
            if (!File.Exists(Path))
                using (File.Create(Path)) { }
            WritePrivateProfileString(section, key, value, Path);
        }

        public string IniReadValue(string section, string key) {
            var temp = new StringBuilder(255);
            var i = GetPrivateProfileString(section, key, "", temp, 255, Path);
            return temp.ToString();
        }
    }
}

using System;
using System.Windows.Forms;

namespace EstimatesAssembly {
    static class Program {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainFormAsm());
//            try {
//                Application.Run(new MainFormAsm());
//            }
//            catch (Exception excep) {
//                MessageBox.Show(excep.Message);
//            }
        }
    }
}

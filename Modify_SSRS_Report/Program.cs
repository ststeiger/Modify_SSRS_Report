
namespace Modify_SSRS_Report
{


    static class Program
    {


        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [System.STAThread]
        static void Main()
        {
            // Run();
            Portal_Reports.LeaseContractFormPostProcessing.ProcessWorkbook();


            if (false)
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
                System.Windows.Forms.Application.Run(new Form1());
            } // End if (false) 

        } // End Sub Main 



        public static void Run()
        {
            using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
            {

                for (int i = 1; i < 11; ++i)
                {
                    zip.AddEntry("LeaseContractForm_" + i.ToString() + ".xlsx", delegate(string filename, System.IO.Stream output)
                    {
                        // ByteArray from ExecuteReport - only ONE ByteArray at a time, because i might be > 100, and ba.size might be > 20 MB
                        byte[] ba = Portal_Reports.LeaseContractFormPostProcessing.ProcessWorkbook();
                        output.Write(ba, 0, ba.Length);
                    });
                } // Next i 

                using (System.IO.Stream someStream = new System.IO.FileStream(@"D:\test.zip", System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.None))
                {
                    zip.Save(someStream);
                }
            } // End Using zip 
        } // End Sub Run 


    } // End Class Program


} // End Namespace Modify_SSRS_Report 

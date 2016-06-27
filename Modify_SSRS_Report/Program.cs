
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
            Portal_Reports.LeaseContractFormPostProcessing.ProcessWorkbook();

            if (false)
            {
                System.Windows.Forms.Application.EnableVisualStyles();
                System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
                System.Windows.Forms.Application.Run(new Form1());
            } // End if (false) 

        } // End Sub Main 


    } // End Class Program


} // End Namespace Modify_SSRS_Report 

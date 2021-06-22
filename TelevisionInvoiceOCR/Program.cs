using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TelevisionInvoiceOCR
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            //Application.Run(new DebugForm());
        }
    }
}


/*


...
#a = [29.75,43.509,819.613,464.472]
#df = read_pdf("Final_est.pdf", lattice=True, pages = 1, multiple_tables=True)
#df.to_csv("output.csv")
#df.to_excel("out.xlsx")
def DataExtract(PdfFile, Result) : 
    tabula.convert_into(PdfFile, Result, lattice=True, pages = "all", multiple_tables=True)
   # tabula.convert_into("Final_est.pdf", "output.xlsx", lattice=True, pages = "all", multiple_tables=True)

#tabula.convert_into("‪Final_est.pdf", "output.csv", lattice=True, pages = "all", multiple_tables=False)
#print(df)




#import tabula

# Read pdf into DataFrame
#df = tabula.read_pdf("test.pdf", options)

# Read remote pdf into DataFrame
#df2 = tabula.read_pdf("https://github.com/tabulapdf/tabula-java/raw/master/src/test/resources/technology/tabula/arabic.pdf")

# convert PDF into CSV
#tabula.convert_into("test.pdf", "output.csv", output_format="csv")

# convert all PDFs in a directory
#tabula.convert_into_by_batch("input_directory", output_format= "pdf")
...
 


 */

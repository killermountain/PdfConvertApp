using System;
using Aspose.Words;
using System.IO;

namespace AsposePdfConvert
{
    class Program
    {
        static void Main(string[] args)
        {
            License license = new License();
            //Console.WriteLine(Directory.GetCurrentDirectory());

            try
            {
                license.SetLicense("Aspose.Words.NET.lic");
                Console.WriteLine("License set successfully.");
            }
            catch (Exception e)
            {
                // We do not ship any license with this example, visit the Aspose site to obtain either a temporary or permanent license. 
                Console.WriteLine("\nThere was an error setting the license: " + e.Message);
            }

            Console.WriteLine("\n********************** PDF to HTML Converter **********************\n");
            string dir_name = "MSE"; // Broomfield | Southend | MSE
            string curr_dir = Directory.GetCurrentDirectory();
            string input_pdfs_dir = curr_dir + "\\input.pdfs\\"+ dir_name;
            string output_pdf_dir = curr_dir + "\\processed.pdfs\\" + dir_name + "\\";
            string output_html_dir = curr_dir + "\\output.html\\" + dir_name + "\\";
            
            if (!Directory.Exists(output_pdf_dir))
                Directory.CreateDirectory(output_pdf_dir);
            
            string[] fileEntries = Directory.GetFiles(input_pdfs_dir, "*.pdf");
            //Console.WriteLine("Found " + fileEntries.Length + " PDF files in '"+dir_name+"' folder.");
            //Console.WriteLine("CWD: " + curr_dir);
            //Console.Read();

            if (!Directory.Exists(output_html_dir))
                Directory.CreateDirectory(output_html_dir);

            int n = 1;
            foreach (string file in fileEntries)
            {
                string filename = Path.GetFileName(file);
                string output_dir = output_html_dir + filename.Replace(".pdf", "");

                if (!Directory.Exists(output_dir))
                    Directory.CreateDirectory(output_dir);

                string output_file = output_dir +"\\"+ filename.Replace(".pdf", ".html");
                
                //Console.WriteLine("FileName: " + filename);
                
                //Console.WriteLine("\n\nMoved file path ");
                //Console.WriteLine(output_pdf_dir + filename);
                
                //Console.WriteLine("\nOutput file name:");
                //Console.WriteLine(output_file);
                
                //Console.Read();

                try 
                {
                    var doc = new Document(file);
                    doc.Save(output_file);
                    File.Move(file, output_pdf_dir + filename);
                    Console.WriteLine("\n'"+n+": " + filename + "' converted.");
                    n++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("----------------------------------------------------");
                    Console.WriteLine("Error occurred while processing: '" + filename + "'");
                    Console.WriteLine("Error: "+ ex.Message);
                    Console.WriteLine("----------------------------------------------------");
                }
                //break;
            }
            
            Console.Read();
        }
    }
}

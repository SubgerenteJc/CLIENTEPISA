using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtraerInfo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program archivo = new Program();
            archivo.Extraer();
        }
        public void Extraer()
        {
            DirectoryInfo di = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
            FileInfo[] files = di.GetFiles("*.xlsx");

            int cantidad = files.Length;
            if (cantidad > 0)
            {
                foreach (var item in files)
                {
                    string sourceFile = @"C:\Administración\Proyecto PISA\ArchivosExcel\" + item.Name;
                    Console.WriteLine("Archivo seleccionado: " + sourceFile);
                    

                    //Obtener extension del archivo
                    // get extension
                    //string nombrem = sourceFile;
                    //var workbook = new Aspose.Cells.Workbook(nombrem);
                    //// guardar como formatos XLSX, ODS, SXC y FODS
                    //workbook.Save(@"C:\Administración\Proyecto PISA\ArchivosExcel\output.xlsx", Aspose.Cells.SaveFormat.Xlsx);


                    string extension = Path.GetExtension(item.Name);
                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".XLS": //Excel 97-03.
                            conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFile + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;ImportMixedtypes=Text;TypeGuessRows=0'";
                            break;
                        case ".xlsx": //Excel 07 and above.
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFile + ";Extended Properties='Excel 8.0;HDR=YES'";
                            break;
                    }


                    

                    DataTable dt = new DataTable();
                    //dt.Columns.Add("Id", typeof(int));
                    conString = string.Format(conString, sourceFile);
                    using (OleDbConnection connExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = new OleDbCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = connExcel;

                                //Get the name of First Sheet.
                                connExcel.Open();
                                DataTable dtExcelSchema;
                                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                connExcel.Close();

                                //Read Data from First Sheet.
                                connExcel.Open();
                                cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                               

                                connExcel.Close();
                            }

                        }

                        string cadena = @"Data source=DESKTOP-CV57FOU\SQLEXPRESS; Initial Catalog=BDFarmacia; User ID=jdev; Password=tdr123;Trusted_Connection=false;MultipleActiveResultSets=true";


                        using (SqlConnection con = new SqlConnection(cadena))
                        {
                            using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                            {
                                //Set the database table name.
                                sqlBulkCopy.DestinationTableName = "EXCELPISA";
                                int conta = 1;
                                foreach (DataColumn col in dt.Columns)
                                {
                                    
                                    
                                        switch (conta)
                                        {
                                            case 1:
                                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "idenvio");
                                            conta++;
                                            break;
                                            case 2:
                                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcvendedora");
                                            break;
                                    }
                                }
                                
                                // Map the Excel columns with that of the database table, this is optional but good if you do
                                // 
                                //sqlBulkCopy.ColumnMappings.Add("idenvio", "{0}");
                                //sqlBulkCopy.ColumnMappings.Add("rfcvendedora", "RFC Sociedad Vendedora");
                                //sqlBulkCopy.ColumnMappings.Add("razonsocialremitente", "Av_cmd_description");
                                //sqlBulkCopy.ColumnMappings.Add("Af_count", "Af_count");
                                //sqlBulkCopy.ColumnMappings.Add("Av_countunit", "Av_countunit");
                                //sqlBulkCopy.ColumnMappings.Add("Av_description_parts", "Av_description_parts");
                                //sqlBulkCopy.ColumnMappings.Add("Af_weight", "Af_weight");
                                //sqlBulkCopy.ColumnMappings.Add("Av_weightunit", "Av_weightunit");
                                //sqlBulkCopy.ColumnMappings.Add("Av_description_units", "Av_description_units");
                                con.Open();
                                sqlBulkCopy.WriteToServer(dt);
                                con.Close();
                            }
                        }
                        //Fin extension
                    } //FOREACH END 

                    //var ultimo_archivo = (from f in di.GetFiles()
                    //                      orderby f.LastWriteTime descending
                    //                      select f).First();



                    //string datestring = DateTime.Now.ToString("yyyyMMddHHmmss");
                    //string aname = datestring + "-" + ultimo_archivo.Name;
                    //string farchivo = ultimo_archivo + datestring;
                    ////Console.WriteLine("Copia existosa: " + farchivo);


                    //string sourceFile = @"C:\Users\Administrator\Documents\SAYER\" + ultimo_archivo;


                    //string destinationFile = @"C:\inetpub\wwwroot\SWUpload\Uploads\" + datestring + "-" + ultimo_archivo;
                    //System.IO.File.Move(sourceFile, destinationFile);
                    //DirectoryInfo dis = new DirectoryInfo(@"C:\inetpub\wwwroot\SWUpload\Uploads");
                    //FileInfo[] filess = dis.GetFiles("*.xml");
                    //var lasts = filess.Last();
                    ////cargarEnSQL(aname);
                    //Console.WriteLine("Copia existosa: " + lasts);
                }

            }
        }
    }
}

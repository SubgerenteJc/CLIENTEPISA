using Aspose.Cells;
using IronXL;
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
            //VERIFICAR SI LOS ARCHIVOS SON XLSX DE ORIGEN

            DirectoryInfo typefiles = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
            FileInfo[] filestype = typefiles.GetFiles("*.xlsx");
            int cantidadfilestype = filestype.Length;

            if (cantidadfilestype > 0)
            {
                //PROCESAR Y EXTRAER LA INFORMACION DE ARCHIVOS XLSX 
                DirectoryInfo difile = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
                FileInfo[] filesxlsx = difile.GetFiles("*.xlsx");

                int cantidadxlsx = filesxlsx.Length;
                if (cantidadxlsx > 0)
                {
                    foreach (var item in filesxlsx)
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

                            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
                            //string cadena = @"Data source=DESKTOP-CV57FOU\SQLEXPRESS; Initial Catalog=BDFarmacia; User ID=jdev; Password=tdr123;Trusted_Connection=false;MultipleActiveResultSets=true";


                            using (SqlConnection con = new SqlConnection(cadena))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {
                                    //Set the database table name.
                                    sqlBulkCopy.DestinationTableName = "TESTPISAUPLOAD";
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
                                                conta++;
                                                break;
                                            case 3:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialremitente");
                                                conta++;
                                                break;
                                            case 4:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcoperador");
                                                conta++;
                                                break;
                                            case 5:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcontratante");
                                                conta++;
                                                break;
                                            case 6:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfccliente");
                                                conta++;
                                                break;
                                            case 7:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcliente");
                                                conta++;
                                                break;
                                            case 8:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuencia");
                                                conta++;
                                                break;
                                            case 9:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorallegada");
                                                conta++;
                                                break;
                                            case 10:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorasalida");
                                                conta++;
                                                break;
                                            case 11:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveprodservicio");
                                                conta++;
                                                break;
                                            case 12:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "descripcion");
                                                conta++;
                                                break;
                                            case 13:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveunidad");
                                                conta++;
                                                break;
                                            case 14:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "materialpeligroso");
                                                conta++;
                                                break;
                                            case 15:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pesoenkg");
                                                conta++;
                                                break;
                                            case 16:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "valormercancia");
                                                conta++;
                                                break;
                                            case 17:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "moneda");
                                                conta++;
                                                break;
                                            case 18:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "numpiezas");
                                                conta++;
                                                break;
                                            case 19:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "unidadpeso");
                                                conta++;
                                                break;
                                            case 20:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciaorigen");
                                                conta++;
                                                break;
                                            case 21:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio1");
                                                conta++;
                                                break;
                                            case 22:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle1");
                                                conta++;
                                                break;
                                            case 23:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado1");
                                                conta++;
                                                break;
                                            case 24:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais1");
                                                conta++;
                                                break;
                                            case 25:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia1");
                                                conta++;
                                                break;
                                            case 26:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal1");
                                                conta++;
                                                break;
                                            case 27:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciadestino");
                                                conta++;
                                                break;
                                            case 28:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio2");
                                                conta++;
                                                break;
                                            case 29:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle2");
                                                conta++;
                                                break;
                                            case 30:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado2");
                                                conta++;
                                                break;
                                            case 31:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais2");
                                                conta++;
                                                break;
                                            case 32:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia2");
                                                conta++;
                                                break;
                                            case 33:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal2");
                                                break;
                                        }
                                    }
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(dt);
                                    con.Close();
                                }
                            }

                        }

                        string destinationFile = @"C:\Administración\Proyecto PISA\Uploads\" + item.Name;
                        System.IO.File.Move(sourceFile, destinationFile);
                        Console.WriteLine("Carga exitosa del archivo: " + item.Name);
                    }

                }
            }
            
            else //AQUI LOS ARCHIVOS SON XLS DE ORIGEN
            {
                //CONVERTIR XLS A ARCHIVOS XLSX
                DirectoryInfo difiles = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
                FileInfo[] files2 = difiles.GetFiles("*.xls");
                int cantidadfiles = files2.Length;

                if (cantidadfiles > 0)
                {
                    foreach (var item2 in files2)
                    {
                        string sourceFile2 = @"C:\Administración\Proyecto PISA\ArchivosExcel\" + item2.Name;
                        string namefiles = item2.Name.Replace(".XLS", "");
                        // Load XLS file
                        var converter = new GroupDocs.Conversion.Converter(sourceFile2);
                        // Set conversion parameters for XLSX format
                        var convertOptions = converter.GetPossibleConversions()["xlsx"].ConvertOptions;
                        // Convert to XLSX format

                        converter.Convert(@"C:\Administración\Proyecto PISA\ArchivosExcel\" + namefiles + ".xlsx", convertOptions);
                        item2.Delete();
                    }
                }




                //PROCESAR Y EXTRAER LA INFORMACION DE ARCHIVOS XLSX 
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
                                    string sheetName = dtExcelSchema.Rows[1]["TABLE_NAME"].ToString();
                                    connExcel.Close();

                                    //Read Data from First Sheet.
                                    connExcel.Open();
                                    cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                    odaExcel.SelectCommand = cmdExcel;
                                    odaExcel.Fill(dt);


                                    connExcel.Close();
                                }

                            }

                            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
                            //string cadena = @"Data source=DESKTOP-CV57FOU\SQLEXPRESS; Initial Catalog=BDFarmacia; User ID=jdev; Password=tdr123;Trusted_Connection=false;MultipleActiveResultSets=true";


                            using (SqlConnection con = new SqlConnection(cadena))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {
                                    //Set the database table name.
                                    sqlBulkCopy.DestinationTableName = "TESTPISAUPLOAD";
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
                                                conta++;
                                                break;
                                            case 3:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialremitente");
                                                conta++;
                                                break;
                                            case 4:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcoperador");
                                                conta++;
                                                break;
                                            case 5:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcontratante");
                                                conta++;
                                                break;
                                            case 6:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfccliente");
                                                conta++;
                                                break;
                                            case 7:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcliente");
                                                conta++;
                                                break;
                                            case 8:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuencia");
                                                conta++;
                                                break;
                                            case 9:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorallegada");
                                                conta++;
                                                break;
                                            case 10:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorasalida");
                                                conta++;
                                                break;
                                            case 11:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveprodservicio");
                                                conta++;
                                                break;
                                            case 12:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "descripcion");
                                                conta++;
                                                break;
                                            case 13:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveunidad");
                                                conta++;
                                                break;
                                            case 14:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "materialpeligroso");
                                                conta++;
                                                break;
                                            case 15:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pesoenkg");
                                                conta++;
                                                break;
                                            case 16:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "valormercancia");
                                                conta++;
                                                break;
                                            case 17:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "moneda");
                                                conta++;
                                                break;
                                            case 18:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "numpiezas");
                                                conta++;
                                                break;
                                            case 19:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "unidadpeso");
                                                conta++;
                                                break;
                                            case 20:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciaorigen");
                                                conta++;
                                                break;
                                            case 21:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio1");
                                                conta++;
                                                break;
                                            case 22:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle1");
                                                conta++;
                                                break;
                                            case 23:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado1");
                                                conta++;
                                                break;
                                            case 24:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais1");
                                                conta++;
                                                break;
                                            case 25:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia1");
                                                conta++;
                                                break;
                                            case 26:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal1");
                                                conta++;
                                                break;
                                            case 27:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciadestino");
                                                conta++;
                                                break;
                                            case 28:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio2");
                                                conta++;
                                                break;
                                            case 29:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle2");
                                                conta++;
                                                break;
                                            case 30:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado2");
                                                conta++;
                                                break;
                                            case 31:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais2");
                                                conta++;
                                                break;
                                            case 32:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia2");
                                                conta++;
                                                break;
                                            case 33:
                                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal2");
                                                break;
                                        }
                                    }
                                    con.Open();
                                    sqlBulkCopy.WriteToServer(dt);
                                    con.Close();
                                }
                            }

                        }

                        string destinationFile = @"C:\Administración\Proyecto PISA\Uploads\" + item.Name;
                        System.IO.File.Move(sourceFile, destinationFile);
                        Console.WriteLine("Carga exitosa del archivo: " + item.Name);
                    }

                }
            }
            // FIN DE ARCHIVOS XLSX

            
        }
    }
}

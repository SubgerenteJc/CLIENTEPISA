using Aspose.Cells;
using ExcelDataReader;
using Ganss.Excel;
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
        public class Students
        {
            public string Name { get; set; }
        }
        public void Extraer()
        {
            string[] values;
            DataTable tbl = new DataTable();
            //DirectoryInfo di24 = new DirectoryInfo(@"\\10.223.208.41\Users\Administrator\Documents\DHLORDENES");
            DirectoryInfo dir = new DirectoryInfo(@"C:\Administración\Proyecto PISA\Ordenes");

            FileInfo[] files = dir.GetFiles("*.XLS");
            int count = files.Length;
            if (count > 0)
            {
                foreach (var item in files)
                {
                    string sourceFile = @"C:\Administración\Proyecto PISA\Ordenes\" + item.Name;
                    string[] strAllLines = File.ReadAllLines(sourceFile, Encoding.UTF8);
                    File.WriteAllLines(sourceFile, strAllLines.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray());
                    string lna = item.Name.ToLower();
                    string Ai_orden = lna.Replace(".xls", "");

                    
                    string[] lineas1 = File.ReadAllLines(sourceFile, Encoding.UTF8);
                    lineas1 = lineas1.Skip(1).ToArray();
                    foreach (string line in lineas1)
                    {
                        string renglones = line;
                        char delimitador = '\t';
                        string[] valores = renglones.Split(delimitador);
                        int coln = int.Parse(valores[0]);
                        string col1 = coln.ToString();
                        string col2 = valores[1].ToString();
                        string col3 = valores[2].ToString();
                        string col4 = valores[3].ToString();
                        string col5 = valores[4].ToString();
                        string col6 = valores[5].ToString();
                        string col7 = valores[6].ToString();
                        int cols = int.Parse(valores[7]);
                        string col8 =cols.ToString();
                        string col9 = valores[8].ToString();
                        string col10 = valores[9].ToString();
                        string clave = valores[10].ToString();
                        string Av_cmd_code = clave.Replace("'", "");
                        string descrip = valores[11].ToString();
                        string Av_cmd_description = descrip.Replace("\"", "");
                        string Av_countunit = valores[12].ToString();
                        string col14 = valores[13].ToString();
                        string Af_weight = valores[14].ToString();
                        string col16 = valores[15].ToString();
                        string col17 = valores[16].ToString();
                        string Af_count = Math.Floor(Convert.ToDecimal(valores[17])).ToString();
                        //string Af_count = valores[17].ToString();
                        string Av_weightunit = valores[18].ToString();
                        string col20 = valores[19].ToString();
                        string col21 = valores[20].ToString();
                        string col22 = valores[21].ToString();
                        string col23 = valores[22].ToString();
                        string col24 = valores[23].ToString();
                        int colt = int.Parse(valores[24]);
                        string col25 = colt.ToString();
                        string col26 = valores[25].ToString();
                        string col27 = valores[26].ToString();
                        string col28 = valores[27].ToString();
                        string col29 = valores[28].ToString();
                        string col30 = valores[29].ToString();
                        string col31 = valores[30].ToString();
                        string col32 = valores[31].ToString();
                        string col33 = valores[32].ToString();

                        if (Av_cmd_code != "")
                        {

                            InsertMerc(col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,Av_cmd_code, Av_cmd_description, Av_countunit,col14, Af_weight,col16,col17, Af_count, Av_weightunit,col20,col21,col22,col23,col24,col25,col26,col27,col28,col29,col30,col31,col32,col33);

                        }

                    }
                    string destinationFile = @"C:\Administración\Proyecto PISA\Procesadas\" + item.Name;
                    System.IO.File.Move(sourceFile, destinationFile);
                }
            }

            //string[] values;
            //DataTable tbl = new DataTable();
            //foreach (string line in File.ReadLines(@"C:\Administración\Proyecto PISA\ArchivosExcel\6415643.TXT"))
            //{

            //    values = line.Split(';');
            //    DataRow dr = tbl.NewRow();

            //    string colum1 = values[0];
            //    //string colum2 = values[1];
            //    //string colum3 = values[2];
            //    //string colum4 = values[3];
            //    //string colum5 = values[4];
            //    //string colum6 = values[5];
            //    //string colum7 = values[6];
            //    //string colum8 = values[7];
            //    //string colum9 = values[8];
            //    //string colum10 = values[9];
            //    //string colum11 = values[10];
            //    //string colum12 = values[11];
            //    //string colum13 = values[12];
            //    //string colum14 = values[13];
            //    //string colum15 = values[14];
            //    //string colum16 = values[15];
            //    //string colum17 = values[16];
            //    //string colum18 = values[17];
            //    //string colum19 = values[18];
            //    //string colum20 = values[19];
            //    //string colum21 = values[20];
            //    //string colum22 = values[21];
            //    //string colum23 = values[22];
            //    //string colum24 = values[23];
            //    //string colum25 = values[24];
            //    //string colum26 = values[25];
            //    //string colum27 = values[26];
            //    //string colum28 = values[27];
            //    //string colum29 = values[28];
            //    //string colum30 = values[29];
            //    //string colum31 = values[30];
            //    //string colum32 = values[31];
            //    //string colum33 = values[32];
            //}

            //PROCESAR Y EXTRAER LA INFORMACION DE ARCHIVOS XLSX 
            //DirectoryInfo di24 = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
            //FileInfo[] files24 = di24.GetFiles("*.xlsx");

            //int cantidad24 = files24.Length;
            //if (cantidad24 > 0)
            //{
            //    foreach (var item in files24)
            //    {
            //        string sourceFile = @"C:\Administración\Proyecto PISA\ArchivosExcel\" + item.Name;
            //        Console.WriteLine("Archivo seleccionado: " + sourceFile);

            //        string extension = Path.GetExtension(item.Name);
            //        string conString = string.Empty;
            //        switch (extension)
            //        {
            //            case ".XLS": //Excel 97-03.
            //                conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFile + ";Extended Properties='Excel 8.0;HDR=YES;CharacterSet=UTF8;IMEX=1;ImportMixedtypes=Text;TypeGuessRows=0'";
            //                break;
            //            case ".xlsx": //Excel 07 and above.
            //                conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFile + ";Extended Properties='Excel 8.0;CharacterSet=UTF8;HDR=YES'";
            //                break;
            //        }




            //        DataTable dt = new DataTable();
            //        //dt.Columns.Add("Id", typeof(int));
            //        conString = string.Format(conString, sourceFile);
            //        using (OleDbConnection connExcel = new OleDbConnection(conString))
            //        {
            //            using (OleDbCommand cmdExcel = new OleDbCommand())
            //            {
            //                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
            //                {
            //                    cmdExcel.Connection = connExcel;

            //                    //Get the name of First Sheet.
            //                    connExcel.Open();
            //                    DataTable dtExcelSchema;
            //                    dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //                    string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            //                    connExcel.Close();

            //                    //Read Data from First Sheet.
            //                    connExcel.Open();
            //                    cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
            //                    odaExcel.SelectCommand = cmdExcel;
            //                    odaExcel.Fill(dt);
                                


            //                    connExcel.Close();
            //                }

            //            }

            //            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //            //string cadena = @"Data source=DESKTOP-CV57FOU\SQLEXPRESS; Initial Catalog=BDFarmacia; User ID=jdev; Password=tdr123;Trusted_Connection=false;MultipleActiveResultSets=true";


            //            using (SqlConnection con = new SqlConnection(cadena))
            //            {
            //                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
            //                {
            //                    //Set the database table name.
            //                    sqlBulkCopy.DestinationTableName = "TESTPISAUPLOAD";
            //                    int conta = 1;
            //                    foreach (DataColumn col in dt.Columns)
            //                    {
            //                        switch (conta)
            //                        {
            //                            case 1:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "idenvio");
            //                                conta++;
            //                                break;
            //                            case 2:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcvendedora");
            //                                conta++;
            //                                break;
            //                            case 3:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialremitente");
            //                                conta++;
            //                                break;
            //                            case 4:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcoperador");
            //                                conta++;
            //                                break;
            //                            case 5:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcontratante");
            //                                conta++;
            //                                break;
            //                            case 6:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfccliente");
            //                                conta++;
            //                                break;
            //                            case 7:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcliente");
            //                                conta++;
            //                                break;
            //                            case 8:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuencia");
            //                                conta++;
            //                                break;
            //                            case 9:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorallegada");
            //                                conta++;
            //                                break;
            //                            case 10:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorasalida");
            //                                conta++;
            //                                break;
            //                            case 11:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveprodservicio");
            //                                conta++;
            //                                break;
            //                            case 12:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "descripcion");
            //                                conta++;
            //                                break;
            //                            case 13:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveunidad");
            //                                conta++;
            //                                break;
            //                            case 14:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "materialpeligroso");
            //                                conta++;
            //                                break;
            //                            case 15:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pesoenkg");
            //                                conta++;
            //                                break;
            //                            case 16:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "valormercancia");
            //                                conta++;
            //                                break;
            //                            case 17:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "moneda");
            //                                conta++;
            //                                break;
            //                            case 18:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "numpiezas");
            //                                conta++;
            //                                break;
            //                            case 19:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "unidadpeso");
            //                                conta++;
            //                                break;
            //                            case 20:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciaorigen");
            //                                conta++;
            //                                break;
            //                            case 21:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio1");
            //                                conta++;
            //                                break;
            //                            case 22:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle1");
            //                                conta++;
            //                                break;
            //                            case 23:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado1");
            //                                conta++;
            //                                break;
            //                            case 24:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais1");
            //                                conta++;
            //                                break;
            //                            case 25:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia1");
            //                                conta++;
            //                                break;
            //                            case 26:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal1");
            //                                conta++;
            //                                break;
            //                            case 27:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciadestino");
            //                                conta++;
            //                                break;
            //                            case 28:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio2");
            //                                conta++;
            //                                break;
            //                            case 29:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle2");
            //                                conta++;
            //                                break;
            //                            case 30:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado2");
            //                                conta++;
            //                                break;
            //                            case 31:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais2");
            //                                conta++;
            //                                break;
            //                            case 32:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia2");
            //                                conta++;
            //                                break;
            //                            case 33:
            //                                sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal2");
            //                                break;
            //                        }
            //                    }
                                
            //                    con.Open();
            //                    sqlBulkCopy.WriteToServer(dt);
            //                    con.Close();
            //                }
            //            }

            //        }

            //        string destinationFile = @"C:\Administración\Proyecto PISA\Uploads\" + item.Name;
            //        System.IO.File.Move(sourceFile, destinationFile);
            //        Console.WriteLine("Carga exitosa del archivo: " + item.Name);
            //    }

            //}
            //DirectoryInfo difiles = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
            //    FileInfo[] files2 = difiles.GetFiles("*.xls");
            //    int cantidadfiles = files2.Length;

            //    if (cantidadfiles > 0)
            //    {
            //        foreach (var item2 in files2)
            //        {
            //            string sourceFile2 = @"C:\Administración\Proyecto PISA\ArchivosExcel\" + item2.Name;
            //            string namefiles = item2.Name.Replace(".XLS", "");
                   
                   

                    
            //            // Load XLS file
                        
            //            var converter = new GroupDocs.Conversion.Converter(sourceFile2);
                    
                    
            //            // Set conversion parameters for XLSX format
            //            var convertOptions = converter.GetPossibleConversions()["xlsx"].ConvertOptions;
                          
                    
            //        // Convert to XLSX format

            //       converter.Convert(@"C:\Administración\Proyecto PISA\ArchivosExcel\" + namefiles + ".xlsx", convertOptions);
                    
                   
                    
            //            item2.Delete();
            //        }
            //    }




                ////PROCESAR Y EXTRAER LA INFORMACION DE ARCHIVOS XLSX 
                //DirectoryInfo di = new DirectoryInfo(@"C:\Administración\Proyecto PISA\ArchivosExcel");
                //FileInfo[] files = di.GetFiles("*.xlsx");

                //int cantidad = files.Length;
                //if (cantidad > 0)
                //{
                //    foreach (var item in files)
                //    {
                //        string sourceFile = @"C:\Administración\Proyecto PISA\ArchivosExcel\" + item.Name;
                //        Console.WriteLine("Archivo seleccionado: " + sourceFile);

                //        string extension = Path.GetExtension(item.Name);
                //        string conString = string.Empty;
                //        switch (extension)
                //        {
                //            case ".XLS": //Excel 97-03.
                //                conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sourceFile + ";Extended Properties='Excel 8.0;HDR=YES;CharacterSet=UTF8;IMEX=1;ImportMixedtypes=Text;TypeGuessRows=0'";
                //                break;
                //            case ".xlsx": //Excel 07 and above.
                //                conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFile + ";Extended Properties='Excel 8.0;CharacterSet=UTF8;IMEX=1;HDR=YES'";
                //                break;
                //        }




                //        DataTable dt = new DataTable();
                //        //dt.Columns.Add("Id", typeof(int));
                //        conString = string.Format(conString, sourceFile, Encoding.Unicode);
                //        using (OleDbConnection connExcel = new OleDbConnection(conString))
                //        {
                //            using (OleDbCommand cmdExcel = new OleDbCommand())
                //            {
                //                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                //                {
                //                    cmdExcel.Connection = connExcel;

                //                    //Get the name of First Sheet.
                //                    connExcel.Open();
                //                    DataTable dtExcelSchema;
                //                    dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                //                    string sheetName = dtExcelSchema.Rows[1]["TABLE_NAME"].ToString();
                //                    connExcel.Close();

                //                    //Read Data from First Sheet.
                //                    connExcel.Open();
                //                    cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                //                    odaExcel.SelectCommand = cmdExcel;
                //                    odaExcel.Fill(dt);


                //                    connExcel.Close();
                //                }

                //            }

                //            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
                //            //string cadena = @"Data source=DESKTOP-CV57FOU\SQLEXPRESS; Initial Catalog=BDFarmacia; User ID=jdev; Password=tdr123;Trusted_Connection=false;MultipleActiveResultSets=true";


                //            using (SqlConnection con = new SqlConnection(cadena))
                //            {
                //                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                //                {
                //                    //Set the database table name.
                //                    sqlBulkCopy.DestinationTableName = "TESTPISAUPLOAD";
                //                    int conta = 1;
                //                foreach (DataColumn col in dt.Columns)
                //                {
                //                    switch (conta)
                //                    {
                //                        case 1:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "idenvio");
                //                            conta++;
                //                            break;
                //                        case 2:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcvendedora");
                //                            conta++;
                //                            break;
                //                        case 3:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialremitente");
                //                            conta++;
                //                            break;
                //                        case 4:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfcoperador");
                //                            conta++;
                //                            break;
                //                        case 5:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcontratante");
                //                            conta++;
                //                            break;
                //                        case 6:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "rfccliente");
                //                            conta++;
                //                            break;
                //                        case 7:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "razonsocialcliente");
                //                            conta++;
                //                            break;
                //                        case 8:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuencia");
                //                            conta++;
                //                            break;
                //                        case 9:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorallegada");
                //                            conta++;
                //                            break;
                //                        case 10:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "fechahorasalida");
                //                            conta++;
                //                            break;
                //                        case 11:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveprodservicio");
                //                            conta++;
                //                            break;
                //                        case 12:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "descripcion");
                //                            conta++;
                //                            break;
                //                        case 13:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "claveunidad");
                //                            conta++;
                //                            break;
                //                        case 14:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "materialpeligroso");
                //                            conta++;
                //                            break;
                //                        case 15:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pesoenkg");
                //                            conta++;
                //                            break;
                //                        case 16:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "valormercancia");
                //                            conta++;
                //                            break;
                //                        case 17:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "moneda");
                //                            conta++;
                //                            break;
                //                        case 18:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "numpiezas");
                //                            conta++;
                //                            break;
                //                        case 19:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "unidadpeso");
                //                            conta++;
                //                            break;
                //                        case 20:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciaorigen");
                //                            conta++;
                //                            break;
                //                        case 21:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio1");
                //                            conta++;
                //                            break;
                //                        case 22:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle1");
                //                            conta++;
                //                            break;
                //                        case 23:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado1");
                //                            conta++;
                //                            break;
                //                        case 24:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais1");
                //                            conta++;
                //                            break;
                //                        case 25:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia1");
                //                            conta++;
                //                            break;
                //                        case 26:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal1");
                //                            conta++;
                //                            break;
                //                        case 27:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "secuenciadestino");
                //                            conta++;
                //                            break;
                //                        case 28:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "municipio2");
                //                            conta++;
                //                            break;
                //                        case 29:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "calle2");
                //                            conta++;
                //                            break;
                //                        case 30:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "estado2");
                //                            conta++;
                //                            break;
                //                        case 31:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "pais2");
                //                            conta++;
                //                            break;
                //                        case 32:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "colonia2");
                //                            conta++;
                //                            break;
                //                        case 33:
                //                            sqlBulkCopy.ColumnMappings.Add(col.ColumnName, "codigopostal2");
                //                            break;
                //                    }
                //                }
                //                //int crows = 1;
                //                //foreach (DataRow reng in dt.Rows)
                //                //{

                //                //    string row20 = reng[20].ToString();
                //                //    byte[] bytes = Encoding.Default.GetBytes(row20);
                //                //    row20 = Encoding.UTF8.GetString(bytes);
                //                //    System.Text.Encoding.Unicode.GetString(System.Text.Encoding.UTF8.GetBytes(row20));
                //                //    Console.WriteLine(row20);

                //                //}
                //                con.Open();
                //                    sqlBulkCopy.WriteToServer(dt);
                //                    con.Close();
                //                }
                //            }

                //        }

                //        string destinationFile = @"C:\Administración\Proyecto PISA\Uploads\" + item.Name;
                //        System.IO.File.Move(sourceFile, destinationFile);
                //        Console.WriteLine("Carga exitosa del archivo: " + item.Name);
                //    }

                //}
            
            // FIN DE ARCHIVOS XLSX

            
        }
        public void InsertMerc(string col1, string col2, string col3, string col4, string col5, string col6, string col7, string col8, string col9, string col10, string Av_cmd_code, string Av_cmd_description, string Av_countunit, string col14, string Af_weight, string col16, string col17, string Af_count, string Av_weightunit, string col20, string col21, string col22, string col23, string col24, string col25, string col26, string col27, string col28, string col29, string col30, string col31, string col32, string col33)
        {
            string cadena2 = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(cadena2))
            {

                using (SqlCommand selectCommand = new SqlCommand("sp_Insert_Merc_Pisa_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@col1", col1);
                    selectCommand.Parameters.AddWithValue("@col2", col2);
                    selectCommand.Parameters.AddWithValue("@col3", col3);
                    selectCommand.Parameters.AddWithValue("@col4", col4);
                    selectCommand.Parameters.AddWithValue("@col5", col5);
                    selectCommand.Parameters.AddWithValue("@col6", col6);
                    selectCommand.Parameters.AddWithValue("@col7", col7);
                    selectCommand.Parameters.AddWithValue("@col8", col8);
                    selectCommand.Parameters.AddWithValue("@col9", col9);
                    selectCommand.Parameters.AddWithValue("@col10", col10);
                    selectCommand.Parameters.AddWithValue("@Av_cmd_code", Av_cmd_code);
                    selectCommand.Parameters.AddWithValue("@Av_cmd_description", Av_cmd_description);
                    selectCommand.Parameters.AddWithValue("@Av_countunit", Av_countunit);
                    selectCommand.Parameters.AddWithValue("@col14", col14);
                    selectCommand.Parameters.AddWithValue("@Af_weight", Af_weight);
                    selectCommand.Parameters.AddWithValue("@col16", col16);
                    selectCommand.Parameters.AddWithValue("@col17", col17);
                    selectCommand.Parameters.AddWithValue("@Af_count", Af_count);
                    selectCommand.Parameters.AddWithValue("@Av_weightunit", Av_weightunit);
                    selectCommand.Parameters.AddWithValue("@Af_count", Af_count);
                    selectCommand.Parameters.AddWithValue("@col20", col20);
                    selectCommand.Parameters.AddWithValue("@col21", col21);
                    selectCommand.Parameters.AddWithValue("@col22", col22);
                    selectCommand.Parameters.AddWithValue("@col23", col23);
                    selectCommand.Parameters.AddWithValue("@col24", col24);
                    selectCommand.Parameters.AddWithValue("@col25", col25);
                    selectCommand.Parameters.AddWithValue("@col26", col26);
                    selectCommand.Parameters.AddWithValue("@col27", col27);
                    selectCommand.Parameters.AddWithValue("@col28", col28);
                    selectCommand.Parameters.AddWithValue("@col29", col29);
                    selectCommand.Parameters.AddWithValue("@col30", col30);
                    selectCommand.Parameters.AddWithValue("@col31", col31);
                    selectCommand.Parameters.AddWithValue("@col32", col32);
                    selectCommand.Parameters.AddWithValue("@col33", col33);



                    try
                    {
                        connection.Open();
                        selectCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }

        }
    }
}

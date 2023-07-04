using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Xml.Linq;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.Text;
using System.IO;
using System.Data.Odbc;
using System.Collections.Generic;

namespace POSender
{
    /// <summary>
    /// Summary description for Service1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    [System.Web.Script.Services.ScriptService]
    public class Service1 : System.Web.Services.WebService
    {

        [WebMethod]
        public string SendPO(string PONumber)
        {
            string responseMsg = string.Empty;
            string query = string.Empty;
            Dictionary<string, object> parameters = null;
            DataRowCollection dataRows = null;
            string cab = string.Empty;
            string gdg = string.Empty;
            string pemasok = string.Empty;
            string jenisPO = string.Empty;
            string kdTujuan = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(PONumber))
                {
                    throw new ArgumentNullException("PONumber");
                }

                query = "select * from get_po where C_PONO = ? limit 1";
                parameters = new Dictionary<string, object>();
                parameters.Add("@C_PONO", PONumber);
                dataRows = ExecuteQuery(query, parameters);
                if (dataRows.Count > 0)
                {
                    pemasok = dataRows[0].Field<string>("C_NOSUP");
                    jenisPO = dataRows[0].Field<string>("JENISPO");
                    kdTujuan = dataRows[0].Field<string>("KDTUJUAN");
                }
                if (jenisPO.Trim().ToUpper() == "HO2")
                {
                    cab = "X8";
                }
                else
                {
                    cab = "X9";
                }

                DataTable dt = new DataTable(),
                  dt1 = new DataTable();
                DataSet ds = new DataSet(),
                  ds1 = new DataSet();

                string path = null,
                  sNoPO = null,
                  sNosup = null;

                sNosup = pemasok;
                sNoPO = PONumber;

                query = "select C_PARAMETERNM from tbl_parameter_sales where C_PARAMETERCD = ?";
                parameters = new Dictionary<string, object>();
                parameters.Add("@C_PARAMETERCD", "SCMS_SEND_PO");
                dataRows = ExecuteQuery(query, parameters);
                string sSupplier = string.Empty;
                string[] sSupplierArray = new string[] { "" };
                if (dataRows.Count > 0)
                {
                    sSupplier = dataRows[0].Field<string>("C_PARAMETERNM");
                }

                sSupplierArray = sSupplier.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                if (!sSupplierArray.Contains(sNosup))
                {
                    return string.Format("Hanya melayani kode Principal berikut: {0}", sSupplier);
                }
                path = ConfigurationManager.AppSettings[sNosup];
                
                if (!string.IsNullOrEmpty(path))
                {
                    query = @"SELECT DISTINCT RIGHT(C_PONO,8) AS c_corno, a.C_ITENO AS c_iteno, b.C_ITENOPRI AS c_itenopri, a.C_ITNAM AS c_itnam, a.C_UNDES AS c_undes, SUM(a.N_QTYINPUT) n_qty, a.C_HNARBP AS n_salpri, '' AS c_nosp, 'D' AS c_via 
                            FROM get_po AS a INNER JOIN tbl_product AS b ON a.C_ITENO = b.C_ITENO WHERE a.C_PONO = ? GROUP BY a.C_PONO, a.C_ITENO, b.C_ITENOPRI, a.C_ITNAM, a.C_UNDES, a.C_HNARBP;";
                    parameters = new Dictionary<string, object>();
                    parameters.Add("@C_PONO", PONumber);
                    dataRows = ExecuteQuery(query, parameters);
                    if (dataRows.Count > 0)
                    {
                        ds.Tables.Add(dataRows[0].Table);
                    }

                    query = string.Format(@"SELECT RIGHT(C_PONO,8) AS c_corno, D_ORDATE AS d_corda, 'Head Office' AS c_komen1, '' AS c_komen2, CASE WHEN 1=1 THEN 1 ELSE 0 END AS L_load, '' AS c_kddepo, '{0}' AS c_kdcab, NULL AS d_posender 
                            FROM get_po WHERE c_pono = ? LIMIT 1", kdTujuan);
                    parameters = new Dictionary<string, object>();
                    parameters.Add("@c_pono", PONumber);
                    dataRows = ExecuteQuery(query, parameters);
                    if (dataRows.Count > 0)
                    {
                        ds1.Tables.Add(dataRows[0].Table);
                    }

                    bool isSukses = false;
                    string sNama = string.Empty;

                    if (ds.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)))
                    {
                        #region "Old code"
                        //query = "select C_PARAMETERNM from tbl_parameter_sales where C_PARAMETERCD = ?";
                        //parameters = new Dictionary<string, object>();
                        //parameters.Add("@C_PARAMETERCD", "SCMS_SEND_PO_NEW_PRINCIPAL");
                        //dataRows = ExecuteQuery(query, parameters);
                        //sSupplier = string.Empty;
                        //sSupplierArray = new string[] { "" };
                        //if (dataRows.Count > 0)
                        //{
                        //    sSupplier = dataRows[0].Field<string>("C_PARAMETERNM");
                        //}

                        //sSupplierArray = sSupplier.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                        ///*
                        // * jika Principal lama, penamaan file m'gunakan format 23010003Header.DBF,23010003Detail.DBF,23010003.txt
                        // * akan direname m'jadi .SP1 & .SP2 oleh program CutData.exe
                        // * selain itu (Principal baru) m'gunakan format 23010003.SP1,23010003.SP2,23010003.txt
                        // */
                        //if (!sSupplierArray.Contains(sNosup))
                        //{
                        //    isSukses = ExportDBF(ds1, path, sNoPO.Substring(2) + "Header", true, false);
                        //    if (ds1.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)) && isSukses)
                        //    {
                        //        isSukses = ExportDBF(ds, path, sNoPO.Substring(2) + "Detil", false, false);
                        //        if (ds1.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)) && isSukses)
                        //        {
                        //            isSukses = ExportDBF(ds, path, sNoPO.Substring(2) + ".txt", false, true);
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        //    isSukses = ExportDBF(ds1, path, sNoPO.Substring(2), true, false, sNosup);
                        //    if (ds1.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)) && isSukses)
                        //    {
                        //        isSukses = ExportDBF(ds, path, sNoPO.Substring(2), false, false, sNosup);
                        //        if (ds1.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)) && isSukses)
                        //        {
                        //            isSukses = ExportDBF(ds, path, sNoPO.Substring(2) + ".txt", false, true, sNosup);
                        //        }
                        //    }

                        //}
                        #endregion
                        
                        isSukses = ExportDBF(ds1, path, sNoPO.Substring(2) + "Header", true, false, sNosup);
                        if (ds1.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)) && isSukses)
                        {
                            isSukses = ExportDBF(ds, path, sNoPO.Substring(2) + "Detil", false, false, sNosup);
                            if (ds1.Tables.Count > 0 && (!string.IsNullOrEmpty(path)) && (!string.IsNullOrEmpty(sNoPO)) && isSukses)
                            {
                                isSukses = ExportDBF(ds, path, sNoPO.Substring(2) + ".txt", false, true, sNosup);
                            }
                        }
                    }

                }

                responseMsg = "Sukses";
            }
            catch (Exception e)
            {
                if (e == null)
                {
                    responseMsg = "Unknown error";
                }
                else
                {
                    if (string.IsNullOrEmpty(e.Message))
                    {
                        responseMsg = "Error is not specified";
                    }
                    else
                    {
                        responseMsg = e.Message;
                    }
                }
            }

            return responseMsg;
        }

        [Obsolete]
        private bool ExportDBF(System.Data.DataSet dsExport, string folderPath, string sNama, bool isHeader, bool isText)
        {
            if (!string.IsNullOrEmpty(folderPath) && !Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string tableName = sNama;

            try
            {
                string createStatement = "Create Table " + tableName + " ( ";
                string insertStatement = "Insert Into " + tableName + " Values ( ";
                string insertTemp = string.Empty;
                OleDbCommand cmd = new OleDbCommand();
                OleDbConnection conn = null;
                if (dsExport.Tables[0].Columns.Count <= 0) { throw new Exception(); }

                StringBuilder sb = new StringBuilder();
                int nLoop = 0,
                  nLen = 0,
                  nLenC = 0,
                  nLoopC = 0;
                System.Data.DataColumn col = null;
                System.Data.DataRow row = null;
                string reslt = null;

                string sFile = folderPath + sNama + ".dbf";

                bool bData = false;
                DateTime d_corda;

                if (!isText)
                {
                    #region Create Table

                    conn = new System.Data.OleDb.OleDbConnection(string.Format("Provider=vfpoledb;Data Source='{0}';Collating Sequence=general;", folderPath));
                    conn.Open();

                    cmd = conn.CreateCommand();

                    DataTable table = dsExport.Tables[0];

                    sb.AppendFormat("CREATE TABLE {0} (", tableName);

                    for (nLoop = 0, nLen = table.Columns.Count; nLoop < nLen; nLoop++)
                    {
                        if ((nLoop + 1) >= nLen)
                        {
                            sb.AppendFormat(" {0}", DbfColumnParser(table.Columns[nLoop], table.Columns[nLoop].Caption));
                        }
                        else
                        {
                            sb.AppendFormat(" {0},", DbfColumnParser(table.Columns[nLoop], table.Columns[nLoop].Caption));
                        }
                    }
                    sb.Append(" )");

                    cmd.CommandText = sb.ToString();
                    cmd.ExecuteNonQuery();

                    cmd.Dispose();

                    sb.Remove(0, sb.Length);

                    #endregion

                    #region Populate Data

                    cmd = conn.CreateCommand();

                    nLenC = table.Columns.Count;

                    for (nLoopC = 0; nLoopC < nLenC; nLoopC++)
                    {
                        col = table.Columns[nLoopC];

                        reslt = string.Concat(reslt, ",", col.ColumnName);
                    }

                    reslt = (reslt.StartsWith(",", StringComparison.OrdinalIgnoreCase) ?
                      reslt.Remove(0, 1) : reslt);

                    for (nLoop = 0, nLen = table.Rows.Count; nLoop < nLen; nLoop++)
                    {
                        row = table.Rows[nLoop];

                        sb.AppendFormat("Insert Into {0} ({1}) Values (", tableName, reslt);

                        for (nLoopC = 0; nLoopC < nLenC; nLoopC++)
                        {
                            col = table.Columns[nLoopC];

                            if (col.DataType.Equals(typeof(float)) ||
                               col.DataType.Equals(typeof(double)) ||
                               col.DataType.Equals(typeof(decimal)))
                            {
                                sb.AppendFormat("{0} ,", decimal.Parse(row[col].ToString()));
                            }
                            else if (col.DataType.Equals(typeof(ushort)) ||
                               col.DataType.Equals(typeof(short)) ||
                               col.DataType.Equals(typeof(uint)) ||
                               col.DataType.Equals(typeof(int)) ||
                               col.DataType.Equals(typeof(ulong)) ||
                               col.DataType.Equals(typeof(long)))
                            {
                                sb.AppendFormat("{0} ,", int.Parse(row[col].ToString()));
                            }
                            else if (col.DataType.Equals(typeof(DateTime)))
                            {
                                if (!string.IsNullOrEmpty(row[col].ToString()))
                                {
                                    d_corda = DateTime.Parse(row[col].ToString());
                                    sb.AppendFormat("Date({0},{1},{2}) ,", d_corda.Year, d_corda.Month, d_corda.Day);
                                }
                                else
                                {
                                    sb.AppendFormat("NULL ,");
                                }
                            }
                            else if (col.DataType.Equals(typeof(bool)))
                            {
                                bData = bool.Parse(row[col].ToString());
                                sb.AppendFormat("{0} ,", (bData ? ".t." : ".f."));
                            }
                            else
                            {
                                sb.AppendFormat("'{0}' ,", row[col]);
                            }
                        }

                        sb.Remove(sb.Length - 1, 1);

                        sb.AppendLine(" ) ");

                        cmd.CommandText = sb.ToString();

                        cmd.ExecuteNonQuery();

                        sb.Remove(0, sb.Length);
                    }


                    #endregion

                    cmd.Dispose();


                }
                else
                {
                    DataTable dt = dsExport.Tables[0];

                    dt.Columns.Remove("c_corno");
                    dt.Columns.Remove("c_iteno");
                    dt.Columns.Remove("c_itenopri");
                    dt.Columns.Remove("c_nosp");
                    dt.Columns.Remove("c_via");

                    int[] maxLengths = new int[dt.Columns.Count];

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        maxLengths[i] = dt.Columns[i].ColumnName.Length;

                        foreach (DataRow rows in dt.Rows)
                        {
                            if (!rows.IsNull(i))
                            {
                                int length = rows[i].ToString().Length;

                                if (length > maxLengths[i])
                                {
                                    maxLengths[i] = length;
                                }
                            }
                        }
                    }

                    using (StreamWriter sw = new StreamWriter(folderPath + sNama, false))
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
                        }

                        sw.WriteLine();

                        foreach (DataRow rows in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (!rows.IsNull(i))
                                {
                                    sw.Write(rows[i].ToString().PadRight(maxLengths[i] + 2));
                                }
                                else
                                {
                                    sw.Write(new string(' ', maxLengths[i] + 2));
                                }
                            }

                            sw.WriteLine();
                        }

                        sw.Close();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return true;
        }

        private bool ExportDBF(System.Data.DataSet dsExport, string folderPath, string sNama, bool isHeader, bool isText, string sNosup)
        {
            if (!string.IsNullOrEmpty(folderPath) && !Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string tableName = sNama;
            string tempFolder = System.Environment.GetEnvironmentVariable("TEMP");
            string tempFile = tempFolder + "\\" + sNama + ".DBF";
            string destination = string.Empty;
            string source = string.Empty;

            try
            {
                string createStatement = "Create Table " + tableName + " ( ";
                string insertStatement = "Insert Into " + tableName + " Values ( ";
                string insertTemp = string.Empty;
                OleDbCommand cmd = new OleDbCommand();
                OleDbConnection conn = null;
                if (dsExport.Tables[0].Columns.Count <= 0) { throw new Exception(); }

                StringBuilder sb = new StringBuilder();
                int nLoop = 0,
                  nLen = 0,
                  nLenC = 0,
                  nLoopC = 0;
                System.Data.DataColumn col = null;
                System.Data.DataRow row = null;
                string reslt = null;

                bool bData = false;
                DateTime d_corda;

                if (!isText)
                {
                    #region Create Table

                    conn = new System.Data.OleDb.OleDbConnection(string.Format("Provider=vfpoledb;Data Source='{0}';Collating Sequence=general;", tempFolder));
                    conn.Open();

                    cmd = conn.CreateCommand();

                    DataTable table = dsExport.Tables[0];

                    sb.AppendFormat("CREATE TABLE {0} (", tableName);

                    for (nLoop = 0, nLen = table.Columns.Count; nLoop < nLen; nLoop++)
                    {
                        if ((nLoop + 1) >= nLen)
                        {
                            sb.AppendFormat(" {0}", DbfColumnParser(table.Columns[nLoop], table.Columns[nLoop].Caption));
                        }
                        else
                        {
                            sb.AppendFormat(" {0},", DbfColumnParser(table.Columns[nLoop], table.Columns[nLoop].Caption));
                        }
                    }
                    sb.Append(" )");

                    cmd.CommandText = sb.ToString();
                    cmd.ExecuteNonQuery();

                    cmd.Dispose();

                    sb.Remove(0, sb.Length);

                    #endregion

                    #region Populate Data

                    cmd = conn.CreateCommand();

                    nLenC = table.Columns.Count;

                    for (nLoopC = 0; nLoopC < nLenC; nLoopC++)
                    {
                        col = table.Columns[nLoopC];

                        reslt = string.Concat(reslt, ",", col.ColumnName);
                    }

                    reslt = (reslt.StartsWith(",", StringComparison.OrdinalIgnoreCase) ?
                      reslt.Remove(0, 1) : reslt);

                    for (nLoop = 0, nLen = table.Rows.Count; nLoop < nLen; nLoop++)
                    {
                        row = table.Rows[nLoop];

                        sb.AppendFormat("Insert Into {0} ({1}) Values (", tableName, reslt);

                        for (nLoopC = 0; nLoopC < nLenC; nLoopC++)
                        {
                            col = table.Columns[nLoopC];

                            if (col.DataType.Equals(typeof(float)) ||
                               col.DataType.Equals(typeof(double)) ||
                               col.DataType.Equals(typeof(decimal)))
                            {
                                sb.AppendFormat("{0} ,", decimal.Parse(row[col].ToString()));
                            }
                            else if (col.DataType.Equals(typeof(ushort)) ||
                               col.DataType.Equals(typeof(short)) ||
                               col.DataType.Equals(typeof(uint)) ||
                               col.DataType.Equals(typeof(int)) ||
                               col.DataType.Equals(typeof(ulong)) ||
                               col.DataType.Equals(typeof(long)))
                            {
                                sb.AppendFormat("{0} ,", int.Parse(row[col].ToString()));
                            }
                            else if (col.DataType.Equals(typeof(DateTime)))
                            {
                                if (!string.IsNullOrEmpty(row[col].ToString()))
                                {
                                    d_corda = DateTime.Parse(row[col].ToString());
                                    sb.AppendFormat("Date({0},{1},{2}) ,", d_corda.Year, d_corda.Month, d_corda.Day);
                                }
                                else
                                {
                                    sb.AppendFormat("NULL ,");
                                }
                            }
                            else if (col.DataType.Equals(typeof(bool)))
                            {
                                bData = bool.Parse(row[col].ToString());
                                sb.AppendFormat("{0} ,", (bData ? ".t." : ".f."));
                            }
                            else
                            {
                                sb.AppendFormat("'{0}' ,", row[col]);
                            }
                        }

                        sb.Remove(sb.Length - 1, 1);

                        sb.AppendLine(" ) ");

                        cmd.CommandText = sb.ToString();

                        cmd.ExecuteNonQuery();

                        sb.Remove(0, sb.Length);
                    }


                    #endregion

                    cmd.Dispose();

                    string query = "select C_PARAMETERNM from tbl_parameter_sales where C_PARAMETERCD = ?";
                    Dictionary<string, object> parameters = new Dictionary<string, object>();
                    parameters.Add("@C_PARAMETERCD", "SCMS_SEND_PO_NEW_PRINCIPAL");
                    DataRowCollection dataRows = ExecuteQuery(query, parameters);
                    string sSupplier = string.Empty;
                    string[] sSupplierArray = new string[] { "" };
                    if (dataRows.Count > 0)
                    {
                        sSupplier = dataRows[0].Field<string>("C_PARAMETERNM");
                    }

                    sSupplierArray = sSupplier.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                    if (isHeader)
                    {
                        source = sSupplierArray.Contains(sNosup) ? tempFile.Replace("Header.DBF", ".SP1") : tempFile;
                        destination = folderPath + Path.GetFileName(source);
                        
                        if (File.Exists(destination))
                        {
                            File.Delete(destination);
                        }

                        File.Copy(tempFile, destination);
                    }
                    if (!isHeader)
                    {
                        source = sSupplierArray.Contains(sNosup) ? tempFile.Replace("Detil.DBF", ".SP2") : tempFile;
                        destination = folderPath + Path.GetFileName(source);
                        
                        if (File.Exists(destination))
                        {
                            File.Delete(destination);
                        }
                        
                        File.Copy(tempFile, destination);
                    }

                }
                else
                {
                    DataTable dt = dsExport.Tables[0];

                    dt.Columns.Remove("c_corno");
                    dt.Columns.Remove("c_iteno");
                    dt.Columns.Remove("c_itenopri");
                    dt.Columns.Remove("c_nosp");
                    dt.Columns.Remove("c_via");

                    int[] maxLengths = new int[dt.Columns.Count];

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        maxLengths[i] = dt.Columns[i].ColumnName.Length;

                        foreach (DataRow rows in dt.Rows)
                        {
                            if (!rows.IsNull(i))
                            {
                                int length = rows[i].ToString().Length;

                                if (length > maxLengths[i])
                                {
                                    maxLengths[i] = length;
                                }
                            }
                        }
                    }

                    using (StreamWriter sw = new StreamWriter(folderPath + sNama, false))
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
                        }

                        sw.WriteLine();

                        foreach (DataRow rows in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (!rows.IsNull(i))
                                {
                                    sw.Write(rows[i].ToString().PadRight(maxLengths[i] + 2));
                                }
                                else
                                {
                                    sw.Write(new string(' ', maxLengths[i] + 2));
                                }
                            }

                            sw.WriteLine();
                        }

                        sw.Close();
                    }
                }
            }
            catch (Exception e)
            {
                string message = e.Message;
                throw e;
            }

            return true;
        }

        private string DbfColumnParser(System.Data.DataColumn column, string caption)
        {
            string rets = null;

            if (column.DataType.Equals(typeof(float)) ||
              column.DataType.Equals(typeof(double)) ||
              column.DataType.Equals(typeof(decimal)))
            {
                int i = 18;
                switch (caption.ToLower())
                {
                    case "n_qty":
                        i = 10;
                        break;
                    case "n_salpri":
                        i = 13;
                        break;
                }
                rets = string.Format("[{0}] numeric({1},2) {2}",
                  column.ColumnName, i,
                  (column.AllowDBNull ? "NULL" : "NOT NULL"));
            }
            else if (column.DataType.Equals(typeof(ushort)) ||
              column.DataType.Equals(typeof(short)) ||
              column.DataType.Equals(typeof(uint)) ||
              column.DataType.Equals(typeof(int)) ||
              column.DataType.Equals(typeof(ulong)) ||
              column.DataType.Equals(typeof(long)))
            {
                rets = string.Format("[{0}] int {1}",
                  column.ColumnName,
                  (column.AllowDBNull ? "NULL" : "NOT NULL"));
            }
            else if (column.DataType.Equals(typeof(bool)))
            {
                rets = string.Format("[{0}] logical {1}",
                  column.ColumnName,
                  (column.AllowDBNull ? "NULL" : "NOT NULL"));
            }
            else if (column.DataType.Equals(typeof(DateTime)))
            {
                rets = string.Format("[{0}] date {1}",
                  column.ColumnName,
                  (column.AllowDBNull ? "NULL" : "NOT NULL"));
            }
            else
            {
                int i = 1;
                switch (caption.ToLower())
                {
                    case "c_corno":
                        i = 8;
                        break;
                    case "c_komen1":
                    case "c_itnam":
                        i = 50;
                        break;
                    case "c_kdcab":
                        i = 3;
                        break;
                    case "c_iteno":
                        i = 4;
                        break;
                    case "c_itenopri":
                        i = 15;
                        break;
                    case "c_undes":
                        i = 10;
                        break;
                }
                rets = string.Format("[{0}] char({1}) {2}",
                  column.ColumnName, i,
                  (column.AllowDBNull ? "NULL" : "NOT NULL"));
            }

            return rets;
        }

        private DataRowCollection ExecuteQuery(string query)
        {
            string cs = ConfigurationManager.ConnectionStrings["MainConnection"].ConnectionString;

            using (OdbcConnection cn = new OdbcConnection(cs))
            {
                using (OdbcCommand cmd = new OdbcCommand(query, cn))
                {
                    if (cn.State == ConnectionState.Closed) cn.Open();

                    DataTable dt = new DataTable();
                    OdbcDataAdapter adapter = new OdbcDataAdapter(cmd);
                    adapter.Fill(dt);
                    return dt.Rows;
                }
            }
        }

        private DataRowCollection ExecuteQuery(string query, Dictionary<string, object> parameters)
        {
            string cs = ConfigurationManager.ConnectionStrings["MainConnection"].ConnectionString;

            using (OdbcConnection cn = new OdbcConnection(cs))
            {
                using (OdbcCommand cmd = new OdbcCommand(query, cn))
                {
                    if (cn.State == ConnectionState.Closed) cn.Open();

                    foreach (KeyValuePair<string, object> p in parameters)
                    {
                        cmd.Parameters.AddWithValue(p.Key, p.Value);
                    }

                    DataTable dt = new DataTable();
                    OdbcDataAdapter adapter = new OdbcDataAdapter(cmd);
                    adapter.Fill(dt);
                    return dt.Rows;
                }
            }
        }

    }
}

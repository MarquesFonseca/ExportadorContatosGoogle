using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
//using DevExpress.XtraEditors;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Threading;
using System.Reflection;
using System.Data.OleDb;
using Microsoft.Office.Interop;
using System.Globalization;

namespace ExportadorContatosGoogle
{
    public class ImportarArquivos
    {
        /// <summary>
        /// Metodo usado para retornar um DataTable de um arquivo *.xls
        /// </summary>
        /// <param name="NomeArquivo">Endereço completo do arquivo xls</param>
        /// <param name="NomePlanilha">Nome da planilha onde deseja buscar o arquivo.</param>
        /// <returns>Retorna um DataTable</returns>
        public DataTable ImportarXLS(string NomeArquivo)
        {
            return ImportarXLS(NomeArquivo, ListSheetInExcel(String.Format(@"{0}", NomeArquivo))[0]);
        }
        public DataTable ImportarXLS(string NomeArquivo, string NomePlanilha)
        {
            return ImportarXLS(NomeArquivo, ListSheetInExcel(String.Format(@"{0}", NomeArquivo))[0], 0);
        }
        public DataTable ImportarXLS(string NomeArquivo, string NomePlanilha, int QtdTopSelect)
        {
            string topSelect = "";
            if (QtdTopSelect == 0)
                topSelect = "";
            else
                topSelect = string.Format("TOP {0} ", QtdTopSelect);

            if (NomePlanilha.Contains("$")) NomePlanilha = NomePlanilha.Replace("$", "");
            string cnnString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0", NomeArquivo);
            string sql = string.Format("select {0}* from [{1}$]", topSelect, NomePlanilha);
            OleDbConnection cnn = new OleDbConnection(cnnString);
            OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            try
            {
                cnn.Open();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            finally
            {
                cnn.Close();
                cnn.Dispose();
                da.Dispose();
                ds.Dispose();
            }
            return dt;
        }

        /// <summary>
        /// Metodo usado para retornar um DataTable de um arquivo *.xlsx
        /// </summary>
        /// <param name="NomeArquivo">Endereço completo do arquivo xls</param>
        /// <param name="NomePlanilha">Nome da planilha onde deseja buscar o arquivo.</param>
        /// <returns>Retorna um DataTable</returns>
        public DataTable ImportarXLSX(string NomeArquivo)
        {
            return ImportarXLSX(NomeArquivo, ListSheetInExcel(String.Format(@"{0}", NomeArquivo))[0]);
        }
        public DataTable ImportarXLSX(string NomeArquivo, string NomePlanilha)
        {
            return ImportarXLSX(NomeArquivo, ListSheetInExcel(String.Format(@"{0}", NomeArquivo))[0], 0);
        }
        public DataTable ImportarXLSX(string NomeArquivo, string NomePlanilha, int QtdTopSelect)
        {
            string topSelect = "";
            if (QtdTopSelect == 0)
                topSelect = "";
            else
                topSelect = string.Format("TOP {0} ", QtdTopSelect);

            if (NomePlanilha.Contains("$")) NomePlanilha = NomePlanilha.Replace("$", "");
            //- Para abrir um arquivo Excel com linha de cabeçalho: Data Source =c:\ExcelArq.xlsx;HDR=yes;Format=xlsx;
            //- Para abrir um arquivo Excel sem linha de cabeçalho: Data Source =c:\ExcelArq.xlsx;HDR=no;Format=xlsx;
            string cnnString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0", NomeArquivo);
            string sql = string.Format("select {0}* from [{1}$]", topSelect, NomePlanilha);
            OleDbConnection cnn = new OleDbConnection(cnnString);
            OleDbDataAdapter da = new OleDbDataAdapter(sql, cnn);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            try
            {
                cnn.Open();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            finally
            {
                cnn.Close();
                cnn.Dispose();
                da.Dispose();
                ds.Dispose();
            }
            return dt;
        }


        /// <summary>
        /// Metodo usado para retornar um DataTable de um arquivo *.xls
        /// </summary>
        /// <param name="NomeArquivo">Endereço completo do arquivo xls</param>
        /// <param name="NomePlanilha">Nome da planilha onde deseja buscar o arquivo.</param>
        /// <returns>Retorna um DataTable</returns>
        public DataTable ImportarSCV(string NomeArquivo)
        {
            return ImportarSCV(NomeArquivo, ListSheetInExcel(String.Format(@"{0}", NomeArquivo))[0]);
        }
        public DataTable ImportarSCV(string NomeArquivo, string NomePlanilha)
        {
            return ImportarSCV(NomeArquivo, ListSheetInExcel(String.Format(@"{0}", NomeArquivo))[0], 0);
        }
        public DataTable ImportarSCV(string NomeArquivo, string NomePlanilha, int QtdTopSelect)
        {
            string topSelect = "";
            if (QtdTopSelect == 0)
                topSelect = "";
            else
                topSelect = string.Format("TOP {0}", QtdTopSelect);

            if (NomePlanilha.Contains("$")) NomePlanilha = NomePlanilha.Replace("$", "");
            string cnnString = String.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=""Excel 8.0;", NomeArquivo);
            OleDbConnection cnn = new OleDbConnection(cnnString);
            OleDbDataAdapter da = new OleDbDataAdapter(String.Format("select {0} * from [{1}$]", topSelect, NomePlanilha), cnn);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            try
            {
                cnn.Open();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            finally
            {
                cnn.Close();
                cnn.Dispose();
                da.Dispose();
                ds.Dispose();
            }
            return dt;
        }

        /// <summary>
        /// Metodo usado para retornar um DataTable de um arquivo *.XML
        /// </summary>
        /// <param name="NomeArquivo">Endereço completo do arquivo xls</param>
        /// <param name="NomePlanilha">Nome da planilha onde deseja buscar o arquivo.</param>
        /// <returns>Retorna um DataTable</returns>
        public DataTable ImportarXML(string NomeArquivo, string NomePlanilha)
        {
            if (NomePlanilha.Contains("$")) NomePlanilha = NomePlanilha.Replace("$", "");
            string cnnString = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;", NomeArquivo);
            string isql = "select * from [{0}$]";
            OleDbConnection cnn = new OleDbConnection(cnnString);
            OleDbDataAdapter da = new OleDbDataAdapter(String.Format(isql, NomePlanilha), cnn);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            try
            {
                cnn.Open();
                da.Fill(ds);
                dt = ds.Tables[0];
            }
            finally
            {
                cnn.Close();
                cnn.Dispose();
                da.Dispose();
                ds.Dispose();
            }
            return dt;
        }

        public DataTable ImportarXLSXNovo(string NomeArquivo, string NomePlanilha, int QtdTopSelect = 0)
        {
            string topSelect = "";
            if (QtdTopSelect == 0)
                topSelect = "";
            else
                topSelect = string.Format("TOP {0} ", QtdTopSelect);

            //if (NomePlanilha.Contains("$")) NomePlanilha = NomePlanilha.Replace("$", "");

            
            OleDbConnectionStringBuilder sbConnection = new OleDbConnectionStringBuilder();
            sbConnection.DataSource = NomeArquivo;
            string strExtendedProperties = string.Empty;

            #region Verifica a extensão do arquivo e aplica a Connection String apropriada
            if (Path.GetExtension(NomeArquivo).Equals(".xls")) // Excel 97-03
            {
                sbConnection.Provider = "Microsoft.Jet.OLEDB.4.0";
                strExtendedProperties = "Excel 8.0;HDR=Yes;IMEX=1";
            }
            else if (Path.GetExtension(NomeArquivo).Equals(".xlsx")) // Excel 2007
            {
                sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            }

            sbConnection.Add("Extended Properties", strExtendedProperties);
            #endregion

            using (OleDbConnection conn = new OleDbConnection(sbConnection.ToString()))
            {
                string sql = string.Format("select {0}* from [{1}]", topSelect, NomePlanilha);
                OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                try
                {
                    conn.Open();
                    da.Fill(ds);
                    dt = ds.Tables[0];
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                    da.Dispose();
                    ds.Dispose();
                }
                return dt;
            }
        }

        public List<string> ListSheetInExcel(string filePath)
        {
            OleDbConnectionStringBuilder sbConnection = new OleDbConnectionStringBuilder();
            sbConnection.DataSource = filePath;
            string strExtendedProperties = string.Empty;

            // Verifica a extensão do arquivo e aplica a Connection String apropriada
            if (Path.GetExtension(filePath).Equals(".xls")) // Excel 97-03
            {
                sbConnection.Provider = "Microsoft.Jet.OLEDB.4.0";
                strExtendedProperties = "Excel 8.0;HDR=Yes;IMEX=1";
            }
            else if (Path.GetExtension(filePath).Equals(".xlsx")) // Excel 2007
            {
                sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            }

            sbConnection.Add("Extended Properties", strExtendedProperties);
            List<string> listSheet = new List<string>();

            using (OleDbConnection conn = new OleDbConnection(sbConnection.ToString()))
            {
                conn.Open();
                System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                foreach (DataRow row in dt.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Contains("$"))
                    {
                        listSheet.Add(row["TABLE_NAME"].ToString());
                    }
                }
            }
            return listSheet;
        }
    }
}


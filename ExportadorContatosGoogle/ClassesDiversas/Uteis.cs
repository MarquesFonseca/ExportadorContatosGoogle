using System;
using System.Data;
using System.Drawing;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExportadorContatosGoogle
{
    public static class Uteis
    {
        public static void CordeFundo(this NumericUpDown Txt)
        {
            Txt.GotFocus += new EventHandler(Uteis.Txt_GotFocus);
            Txt.LostFocus += new EventHandler(Uteis.Txt_LostFocus);
        }

        public static bool EData(this object Data)
        {
            try
            {
                Convert.ToDateTime(Data);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static string ENullo(this string Valor, string Retorno)
        {
            if (string.IsNullOrEmpty(Valor))
            {
                return Retorno;
            }
            return Valor;
        }

        public static bool ENumero(this object Texto)
        {
            try
            {
                Convert.ToDouble(Texto);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static string FormataCNPJ(this string CNPJ)
        {
            if (CNPJ.Length == 14)
            {
                CNPJ = string.Format("{0}.{1}.{2}/{3}-{4}", new object[] { CNPJ.Substring(0, 2), CNPJ.Substring(2, 3), CNPJ.Substring(5, 3), CNPJ.Substring(8, 4), CNPJ.Substring(12, 2) });
            }
            return CNPJ;
        }

        public static string FormataDataBanco(this DateTime Data, TipoBanco Tp)
        {
            return ((Tp == TipoBanco.SqlServer) ? Data.ToString("yyyy-MM-dd") : Data.ToString("dd/MM/yyyy"));
        }

        public static string FormataMAC(this string MAC)
        {
            return MAC.Replace(":", "").Replace("-", "");
        }

        public static string FormatCNPJSoNumeros(this string CNPJ)
        {
            return CNPJ.Replace(".", "").Replace("/", "");
        }

        public static bool IsValidEmail(this string email)
        {
            return Regex.IsMatch(email, "(?<user>[^@]+)@(?<host>.+)");
        }

        public static bool NulloOuVazio(this string Valor)
        {
            return string.IsNullOrEmpty(Valor.Trim());
        }

        public static string PegaNomePC()
        {
            return Dns.GetHostName().ToUpper();
        }

        public static string RemoveLetras(this string Texto)
        {
            string Retorno = string.Empty;
            foreach (char item in Texto)
            {
                if (!char.IsLetter(item))
                {
                    Retorno = Retorno + item.ToString();
                }
            }
            return Retorno;
        }

        public static string RemovePontuacao(this string Texto)
        {
            string Retorno = string.Empty;
            foreach (char item in Texto)
            {
                if (!char.IsPunctuation(item))
                {
                    Retorno = Retorno + item.ToString();
                }
            }
            return Retorno;
        }

        public static string RemoveSpecialChars(this string str)
        {
            // Create  a string array and add the special characters you want to remove
            string[] chars = new string[] { ",", ".", "/", "!", "@", "#", "$", "%", "^", "&", "*", "'", "\"", ";", "_", "(", ")", ":", "|", "[", "]" };
            //Iterate the number of times based on the String array length.
            for (int i = 0; i < chars.Length; i++)
            {
                if (str.Contains(chars[i]))
                {
                    str = str.Replace(chars[i], "");
                }
            }
            return str;
        }

        public static string RemoverEspaco(this string Texto)
        {
            Texto = Texto.Trim();
            while (Texto.Contains(" "))
            {
                Texto = Texto.Remove(Texto.IndexOf(' '), 1);
            }
            return Texto;
        }

        public static string RemoveSimbolos(this string Texto)
        {
            string Retorno = string.Empty;
            foreach (char item in Texto)
            {
                if (!(char.IsSymbol(item) || item.Equals('\x00ba')))
                {
                    Retorno = Retorno + item.ToString();
                }
            }
            return Retorno;
        }

        public static string RemoveAcentos(this string text)
        {
            System.Text.StringBuilder sbReturn = new System.Text.StringBuilder();
            var arrayText = text.Normalize(System.Text.NormalizationForm.FormD).ToCharArray();
            foreach (char letter in arrayText)
            {
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(letter) != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sbReturn.Append(letter);
            }
            return sbReturn.ToString();
        } 

        public static DataSet RetornaDataSet(string Nome, string[] NomesColunas, string[] ValoresColunas)
        {
            int i;
            DataSet DsRetorno = new DataSet(Nome);
            DataTable TblTemp = new DataTable(Nome);
            for (i = 0; i < NomesColunas.Length; i++)
            {
                TblTemp.Columns.Add(new DataColumn(NomesColunas[i], Type.GetType("System.String")));
            }
            DataRow Linha = TblTemp.NewRow();
            for (i = 0; i < Linha.Table.Columns.Count; i++)
            {
                Linha[i] = ValoresColunas[i];
            }
            TblTemp.Rows.Add(Linha);
            DsRetorno.Tables.Add(TblTemp);
            return DsRetorno;
        }

        public static string RetornaNomeProduto(this string Produto)
        {
            string Retorno = string.Empty;
            Produto = Produto.PadLeft(3, '0');
            switch (Produto)
            {
                case "001":
                    return "SuperSCE 4.3.0";

                case "002":
                    return "SCE Mobile";

                case "003":
                    return "SCE Balc\x00e3o Mobile";

                case "004":
                    return "DTEF Tef Dedicado";

                case "005":
                    return "TEF Discado";

                case "006":
                    return "TEF IP";

                case "007":
                    return "SiTEF";

                case "008":
                    return "Sistema Atendimento";

                case "009":
                    return "SEDSCE NFE";

                case "010":
                    return "SuperSCE Ponto";

                case "011":
                    return "SuperSCE Frente de Loja";

                case "012":
                    return "SEDFiscal EFD";

                case "013":
                    return "OS ECF";

                case "014":
                    return "SCECRM";

                case "015":
                    return "Rensoftware Intelecto";
            }
            return Retorno;
        }

        public static string RetornaValoresColunasRow(this DataRow Row)
        {
            string Temp = string.Empty;
            for (int i = 0; i < Row.Table.Columns.Count; i++)
            {
                Temp = Temp + string.Format("{0}: {1} ", Row.Table.Columns[i].ColumnName, Row[i]);
            }
            return Temp;
        }

        public static void SelecionaCaixaTexto(this TextBox Txt)
        {
            Txt.Focus();
            Txt.SelectAll();
        }

        public static void SelecionaControle(this Control Controle)
        {
            Controle.Focus();
            Controle.Select();
        }

        public static DateTime ToDateTime(this object Valor)
        {
            return Convert.ToDateTime(Valor);
        }

        public static double ToDouble(this object Valor)
        {
            return Convert.ToDouble(Valor);
        }

        public static int ToInt(this object Valor)
        {
            return Convert.ToInt32(Valor);
        }

        public static void Txt_GotFocus(object sender, EventArgs e)
        {
            if (!(!((NumericUpDown)sender).Enabled || ((NumericUpDown)sender).ReadOnly))
            {
                ((NumericUpDown)sender).BackColor = SystemColors.Info;
            }
        }

        public static void Txt_LostFocus(object sender, EventArgs e)
        {
            if (!(!((NumericUpDown)sender).Enabled || ((NumericUpDown)sender).ReadOnly))
            {
                ((NumericUpDown)sender).BackColor = Color.White;
            }
        }

        public static bool ValidaCNPJ(string CNPJ)
        {
            string cnpj = CNPJ.Replace(".", "").Replace("/", "").Replace("-", "");
            int[] resultado = new int[2];
            string ftmt = "6543298765432";
            bool[] cnpjok = new bool[2];
            int[] digitos = new int[14];
            int[] soma = new int[] { 0, 0 };
            resultado[0] = 0;
            resultado[1] = 0;
            cnpjok[0] = false;
            cnpjok[1] = false;
            try
            {
                int nrdig;
                for (nrdig = 0; nrdig < 14; nrdig++)
                {
                    digitos[nrdig] = int.Parse(cnpj.Substring(nrdig, 1));
                    if (nrdig <= 11)
                    {
                        soma[0] += digitos[nrdig] * int.Parse(ftmt.Substring(nrdig + 1, 1));
                    }
                    if (nrdig <= 12)
                    {
                        soma[1] += digitos[nrdig] * int.Parse(ftmt.Substring(nrdig, 1));
                    }
                }
                for (nrdig = 0; nrdig < 2; nrdig++)
                {
                    resultado[nrdig] = soma[nrdig] % 11;
                    if ((resultado[nrdig] == 0) || (resultado[nrdig] == 1))
                    {
                        cnpjok[nrdig] = digitos[12 + nrdig] == 0;
                    }
                    else
                    {
                        cnpjok[nrdig] = digitos[12 + nrdig] == (11 - resultado[nrdig]);
                    }
                }
                return (cnpjok[0] && cnpjok[1]);
            }
            catch
            {
                return false;
            }
        }

        public static bool VerificaSePar(this int valor)
        {
            bool retorno = true;
            if (valor % 2 == 0)
                retorno = true;
            else
                retorno = false;

            return retorno;
        }

        public static string TrocaEstadoPorSegla(this string Estado)
        {
            string retorno = Estado;
            string[] siglasEstados;
            siglasEstados = new string[] { "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO" };

            string[] nomesEstados;
            nomesEstados = new string[] { "Acre", "Alagoas", "Amapá", "Amazonas", "Bahia", "Ceará", "Distrito Federal", "Espírito Santo", "Goiás", "Maranhão", "Mato Grosso", "Mato Grosso do Sul", "Minas Gerais", "Pará", "Paraíba", "Paraná", "Pernambuco", "Piauí", "Rio de Janeiro", "Rio Grande do Norte", "Rio Grande do Sul", "Rondônia", "Roraima", "Santa Catarina", "São Paulo", "Sergipe", "Tocantins" };

            for (int i = 0; i < nomesEstados.Length; i++)
            {
                if (nomesEstados[i].ToUpper().RemoveSpecialChars().RemoveSimbolos().RemovePontuacao().RemoveAcentos() == Estado.ToUpper().RemoveSpecialChars().RemoveSimbolos().RemovePontuacao().RemoveAcentos())
                {
                    retorno = siglasEstados[i];
                    break;
                }
            }

            return retorno;
        }
    }
}

using System;
using System.Windows.Forms;

namespace ExportadorContatosGoogle
{
    public static class Mensagens
    {
        private static string CabecalhoMensagens = "iQUES Desenvolvimento de Sistemas";

        public static void Alerta(string mensagem)
        {
            MessageBox.Show(mensagem, CabecalhoMensagens, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
        }

        public static void Erro(string mensagem)
        {
            MessageBox.Show(mensagem, CabecalhoMensagens, MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Informa(string mensagem)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, MessageBoxButtons.OK, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Informa(string mensagem, MessageBoxIcon Icone)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, MessageBoxButtons.OK, Icone, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Informa(string mensagem, MessageBoxIcon Icone, MessageBoxButtons Botoes)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, Botoes, Icone, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Informa(string mensagem, MessageBoxIcon Icone, MessageBoxButtons Botoes, MessageBoxDefaultButton BotaoPadrao)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, Botoes, Icone, BotaoPadrao);
        }

        public static DialogResult Pergunta(string mensagem)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Pergunta(string mensagem, MessageBoxButtons Botoes)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, Botoes, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Pergunta(string mensagem, MessageBoxButtons Botoes, MessageBoxIcon Icone)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, Botoes, Icone, MessageBoxDefaultButton.Button1);
        }

        public static DialogResult Pergunta(string mensagem, MessageBoxButtons Botoes, MessageBoxIcon Icone, MessageBoxDefaultButton BotaoPadrao)
        {
            return MessageBox.Show(mensagem, CabecalhoMensagens, Botoes, Icone, BotaoPadrao);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// TXT de importação no layout do lanctos.txt (sem cabeçalho, Windows-1252, CRLF):
    /// Data;Débito;Crédito;Valor;Cód.Histórico;Complemento;Inicia Lote;Matriz/Filial;CC Débito;CC Crédito
    /// </summary>
    public static class GravadorTxt
    {
        public static void Gravar(string dest, List<Lancamento> lancs, Dictionary<Lancamento, string> complementos, bool acrescentar)
        {
            var sb = new StringBuilder();
            if (acrescentar && File.Exists(dest) && new FileInfo(dest).Length > 0)
            {
                var b = File.ReadAllBytes(dest);
                if (b[b.Length - 1] != '\n') sb.Append("\r\n");
            }
            foreach (var l in lancs)
            {
                string c;
                if (!complementos.TryGetValue(l, out c)) c = l.historico;
                sb.Append(Linha(l, c)).Append("\r\n");
            }
            var bytes = J.Cp1252.GetBytes(sb.ToString());
            using (var fs = new FileStream(dest, acrescentar ? FileMode.Append : FileMode.Create, FileAccess.Write))
                fs.Write(bytes, 0, bytes.Length);
        }

        internal static string Linha(Lancamento l, string complemento)
        {
            return l.data + ";" + l.debito + ";" + l.credito + ";" + Valor(l.valor) + ";;" + Texto(complemento) + ";;;;";
        }

        /// <summary>6409,62 | 3613 | 12344,6 | 0,35</summary>
        public static string Valor(double v)
        {
            decimal d = Math.Round((decimal)Math.Abs(v), 2, MidpointRounding.AwayFromZero);
            return d.ToString("0.##", CultureInfo.InvariantCulture).Replace('.', ',');
        }

        /// <summary>Complemento numa linha só, sem ';', só com caracteres do Windows-1252.</summary>
        public static string Texto(string s)
        {
            if (s == null) return "";
            s = Regex.Replace(s.Normalize(NormalizationForm.FormC), @"[\r\n\t]+", " ").Replace(';', ',').Replace(' ', ' ');
            s = Regex.Replace(s, @"\p{Cc}", "");
            var o = new StringBuilder();
            var enc = J.Cp1252;
            foreach (char c in s)
            {
                string cs = c.ToString();
                if (enc.GetString(enc.GetBytes(cs)) == cs) o.Append(c);
                else
                {
                    string b = J.SemAcento(cs);
                    o.Append(b.Length > 0 && enc.GetString(enc.GetBytes(b)) == b ? b : " ");
                }
            }
            return Regex.Replace(o.ToString(), " {2,}", " ").Trim();
        }
    }
}

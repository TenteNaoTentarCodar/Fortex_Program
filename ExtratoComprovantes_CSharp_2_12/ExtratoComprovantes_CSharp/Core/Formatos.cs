using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Datas e valores em qualquer formato usado por bancos:
    ///  - valores: 1.234,56 | 1,234.56 | 1 234,56 | 1'234.56 | 1234.56 | (1.234,56) | -1.234,56 | 1.234,56- | 1.234,56 D/C | 1,234.56 CR/DR, com ou sem moeda;
    ///  - datas: dd/mm/aaaa, mm/dd/aaaa (detectado pelo arquivo), aaaa-mm-dd, dd.mm.aaaa, aaaammdd, 12 Jan 2026, Jan 12, 2026, 12-JAN-26,
    ///    nomes de mês em português, inglês, espanhol, francês, alemão e italiano.
    /// </summary>
    public static class Formatos
    {
        public const int DMY = 0, MDY = 1, YMD = 2;

        static readonly Regex MOEDA = new Regex(@"R\$|US\$|U\$S|C\$|A\$|NZ\$|HK\$|S\$|MX\$|\$|€|£|¥|₹|₩|₽|₺|₪|₱|฿|zł|kr\.?|\b(BRL|USD|EUR|GBP|JPY|CHF|CAD|AUD|ARS|CLP|COP|MXN|PEN|UYU|PYG|BOB|CNY|INR|ZAR|SEK|NOK|DKK|PLN|CZK|HUF|RUB|TRY|ILS|AED|SAR|KRW|SGD|HKD|NZD|MZN|AOA|CVE)\b", RegexOptions.IgnoreCase);
        static readonly Regex MARCA = new Regex(@"^(.*\d\)?)\s*(CR|DR|DB|DEB|CRED|DEBIT|CREDIT|D|C)\.?$", RegexOptions.IgnoreCase | RegexOptions.Singleline);

        static string SemMoeda(string t) { return MOEDA.Replace(t, " ").Trim(); }

        /// <summary>Valor com sinal; null se não for número.</summary>
        public static double? Valor(string s)
        {
            if (s == null) return null;
            string t = s.Replace(' ', ' ').Replace('−', '-').Replace(' ', ' ').Replace(' ', ' ').Trim();
            if (t.Length == 0) return null;
            bool neg = false;
            var mk = MARCA.Match(t);
            if (mk.Success) { if (mk.Groups[2].Value.ToUpperInvariant().StartsWith("D")) neg = true; t = mk.Groups[1].Value.Trim(); }
            t = SemMoeda(t);
            if (t.StartsWith("(") && t.EndsWith(")")) { neg = true; t = t.Substring(1, t.Length - 2).Trim(); }
            t = SemMoeda(t);
            if (t.StartsWith("-")) { neg = true; t = t.Substring(1).Trim(); }
            else if (t.StartsWith("+")) t = t.Substring(1).Trim();
            if (t.EndsWith("-")) { neg = true; t = t.Substring(0, t.Length - 1).Trim(); }
            else if (t.EndsWith("+")) t = t.Substring(0, t.Length - 1).Trim();
            t = SemMoeda(t);
            if (t.StartsWith("-")) { neg = true; t = t.Substring(1).Trim(); }
            if (t.Length == 0 || !t.Matches(@"[0-9.,' ]+") || !char.IsDigit(t[0]) || !char.IsDigit(t[t.Length - 1])) return null;
            string num;
            int ip = t.LastIndexOf('.'), iv = t.LastIndexOf(',');
            int ult = Math.Max(ip, iv);
            if (ult < 0) num = Regex.Replace(t, "[' ]", "");
            else
            {
                char sep = t[ult];
                int casas = t.Length - ult - 1;
                bool outroSep = (sep == '.' ? iv : ip) >= 0;
                bool repetido = t.IndexOf(sep) != ult;
                bool dec;
                if (outroSep) dec = true;
                else if (repetido) dec = false;
                else dec = casas != 3;
                string inteiro = t.Substring(0, ult), frac = t.Substring(ult + 1);
                if (!dec) { inteiro = t; frac = ""; }
                inteiro = Regex.Replace(inteiro, "[.,' ]", "");
                if (Regex.IsMatch(frac, @"[^0-9]")) return null;
                num = frac.Length == 0 ? inteiro : inteiro + "." + frac;
            }
            if (!num.Matches(@"[0-9]+(\.[0-9]+)?")) return null;
            double v;
            if (!double.TryParse(num, NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return null;
            return neg ? -v : v;
        }

        /// <summary>Parece valor monetário (tem casas decimais ou marcador)?</summary>
        public static bool PareceDinheiro(string s)
        {
            if (s == null) return false;
            string t = s.Trim();
            if (t.Length > 24) return false;
            if (!Regex.IsMatch(t, @"^(?:.*[0-9][.,][0-9]{2}(\s*-|\)|\s*(CR|DR|D|C))?\s*)$", RegexOptions.IgnoreCase)
                && !Regex.IsMatch(t, @"^(?:.*[0-9][.,][0-9]{2}\s*[A-Z]{0,3})$", RegexOptions.IgnoreCase)) return false;
            return Valor(t) != null;
        }

        // ------------------------------------------------------------------ datas

        static readonly Dictionary<string, int> MES = new Dictionary<string, int>();
        static Formatos()
        {
            string[][] nomes = {
                new[]{"JAN", "JANEIRO", "JANUARY", "ENERO", "JANVIER", "JANUAR", "GENNAIO", "GEN", "ENE", "JANV", "JAENNER"},
                new[]{"FEV", "FEVEREIRO", "FEB", "FEBRUARY", "FEBRERO", "FEVRIER", "FEBRUAR", "FEBBRAIO", "FEVR"},
                new[]{"MAR", "MARCO", "MARCH", "MARZO", "MARS", "MAERZ", "MARZ", "MRZ"},
                new[]{"ABR", "ABRIL", "APR", "APRIL", "AVRIL", "APRILE", "AVR"},
                new[]{"MAI", "MAIO", "MAY", "MAYO", "MAGGIO", "MAG"},
                new[]{"JUN", "JUNHO", "JUNE", "JUNIO", "JUIN", "JUNI", "GIUGNO", "GIU"},
                new[]{"JUL", "JULHO", "JULY", "JULIO", "JUILLET", "JULI", "LUGLIO", "LUG", "JUIL"},
                new[]{"AGO", "AGOSTO", "AUG", "AUGUST", "AOUT", "AUGUSTO"},
                new[]{"SET", "SETEMBRO", "SEP", "SEPT", "SEPTEMBER", "SEPTIEMBRE", "SETIEMBRE", "SEPTEMBRE", "SETTEMBRE"},
                new[]{"OUT", "OUTUBRO", "OCT", "OCTOBER", "OCTUBRE", "OCTOBRE", "OKT", "OKTOBER", "OTTOBRE", "OTT"},
                new[]{"NOV", "NOVEMBRO", "NOVEMBER", "NOVIEMBRE", "NOVEMBRE"},
                new[]{"DEZ", "DEZEMBRO", "DEC", "DECEMBER", "DICIEMBRE", "DECEMBRE", "DEZEMBER", "DICEMBRE", "DIC"}};
            for (int i = 0; i < 12; i++) foreach (var n in nomes[i]) if (!MES.ContainsKey(n)) MES[n] = i + 1;
        }

        static readonly Regex NUM = new Regex(@"(?<![0-9])([0-9]{1,4})[/.\-]([0-9]{1,2})(?:[/.\-]([0-9]{2,4}))?(?![0-9])");
        static readonly Regex COMPACTA = new Regex(@"^([0-9]{4})([0-9]{2})([0-9]{2})(?:[0-9]{6})?(?:\.[0-9]+)?(?:\[.*\])?$");
        static readonly Regex EXT1 = new Regex(@"(?<![A-Z0-9])([0-9]{1,2})(?:ST|ND|RD|TH|O)?[\s\-./]*(?:DE\s+)?([A-Z]{3,10})\.?[\s\-./,]*(?:DE\s+)?([0-9]{4}|[0-9]{2})?(?![0-9])");
        static readonly Regex EXT2 = new Regex(@"(?<![A-Z])([A-Z]{3,10})\.?\s+([0-9]{1,2})(?:ST|ND|RD|TH)?,?\s*([0-9]{4})?(?![0-9])");

        static string G(Match m, int i) { return m.Groups[i].Success ? m.Groups[i].Value : null; }

        /// <summary>Data em dd/MM/aaaa ou "" se não reconhecer. anoPadrao usado quando a data vem sem ano (null = ano corrente).</summary>
        public static string Data(string s, int ordem, string anoPadrao)
        {
            if (s == null) return "";
            string t = s.Trim();
            if (t.Length == 0 || t.Length > 40) return "";
            var mc = COMPACTA.Match(t);
            if (mc.Success) return Monta(G(mc, 3), G(mc, 2), G(mc, 1), anoPadrao);
            var m = NUM.Match(t);
            if (m.Success && m.Index <= 3)
            {
                string a = G(m, 1), b = G(m, 2), c = G(m, 3);
                if (a.Length == 4) return Monta(c, b, a, anoPadrao);
                if (c == null && a.Length > 2) return "";
                if (ordem == MDY) return Monta(b, a, c, anoPadrao);
                if (ordem == YMD && c != null && a.Length == 2 && c.Length == 2) return Monta(c, b, "20" + a, anoPadrao);
                return Monta(a, b, c, anoPadrao);
            }
            string n = J.SemAcento(t).ToUpperInvariant();
            var e1 = EXT1.Match(n);
            if (e1.Success && MES.ContainsKey(G(e1, 2))) return Monta(G(e1, 1), MES[G(e1, 2)].ToString(), G(e1, 3), anoPadrao);
            var e2 = EXT2.Match(n);
            if (e2.Success && MES.ContainsKey(G(e2, 1))) return Monta(G(e2, 2), MES[G(e2, 1)].ToString(), G(e2, 3), anoPadrao);
            return "";
        }

        public static string Data(string s) { return Data(s, DMY, null); }

        static string Monta(string d, string m, string a, string anoPadrao)
        {
            int dd, mm, aa;
            if (!int.TryParse(d, out dd) || !int.TryParse(m, out mm)) return "";
            if (string.IsNullOrEmpty(a)) { if (anoPadrao != null) { if (!int.TryParse(anoPadrao, out aa)) return ""; } else aa = DateTime.Now.Year; }
            else { if (!int.TryParse(a, out aa)) return ""; if (aa < 100) aa += aa < 70 ? 2000 : 1900; }
            if (dd < 1 || dd > 31 || mm < 1 || mm > 12 || aa < 1970 || aa > 2100) return "";
            return dd.ToString("00") + "/" + mm.ToString("00") + "/" + aa.ToString("0000");
        }

        /// <summary>Ordem das datas numéricas de uma coluna/arquivo.</summary>
        public static int Ordem(IEnumerable<string> datas)
        {
            int primeiraMaior = 0, segundaMaior = 0, ymd2 = 0;
            foreach (var s in datas)
            {
                if (s == null) continue;
                var m = NUM.Match(s.Trim());
                if (!m.Success || m.Index > 3 || m.Groups[1].Value.Length == 4) continue;
                int a = int.Parse(m.Groups[1].Value), b = int.Parse(m.Groups[2].Value);
                if (a > 12 && a <= 31) primeiraMaior++;
                if (b > 12 && b <= 31) segundaMaior++;
                if (a > 31) ymd2++;
            }
            if (ymd2 > 0 && primeiraMaior == 0) return YMD;
            if (segundaMaior > 0 && primeiraMaior == 0) return MDY;
            return DMY;
        }

        /// <summary>O texto novo já está no acumulado? (sem acento/caixa; texto em outro alfabeto compara como está)</summary>
        public static bool JaContem(string acumulado, string novo)
        {
            string n = Regras.Normalizar(novo);
            if (Regex.Replace(n, "[^A-Z0-9]", "").Length == 0) return novo.Trim().Length > 0 && acumulado.Contains(novo.Trim()) || novo.Trim().Length == 0;
            return Regras.Normalizar(acumulado).Contains(n);
        }

        public static string Normalizar(string s)
        {
            string n = J.SemAcento(s ?? "").Replace("ø", "o").Replace("Ø", "O").Replace("æ", "ae").Replace("Æ", "AE").Replace("ł", "l").Replace("Ł", "L").Replace("đ", "d").Replace("ı", "i");
            n = n.ToUpperInvariant().Replace("ß", "SS");
            return Regex.Replace(Regex.Replace(n, "[^A-Z0-9/ ]", " "), @"\s+", " ").Trim();
        }
    }
}

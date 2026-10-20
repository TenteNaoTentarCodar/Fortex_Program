using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Converte PDF, XLSX/XLSM, XLS (97-2003), ODS, HTML/XML disfarçado de .xls, CSV ou TXT em linhas de células.
    /// O formato é identificado pelo conteúdo, não pela extensão.
    /// </summary>
    public static class LeitorTabular
    {
        [ThreadStatic] public static string ultimoFormato;

        public static List<string[]> Linhas(string arq)
        {
            var b = File.ReadAllBytes(arq);
            if (b.Length >= 4 && b[0] == '%' && b[1] == 'P' && b[2] == 'D' && b[3] == 'F')
            {
                ultimoFormato = "PDF";
                return TextoPDF.Linhas(arq).Select(l => new[] { l }).ToList();
            }
            if (b.Length >= 2 && b[0] == 'P' && b[1] == 'K')
            {
                if (Ods(b)) { ultimoFormato = "OpenDocument (.ods)"; return LerOds(b); }
                ultimoFormato = "Excel (.xlsx)"; return Xlsx(b);
            }
            if (b.Length >= 2 && b[0] == 0xD0 && b[1] == 0xCF) { ultimoFormato = "Excel 97-2003 (.xls)"; return LeitorXls.Linhas(b); }
            string txt = J.Texto(b);
            string ini = txt.Length > 5000 ? txt.Substring(0, 5000) : txt;
            if (ini.Contains("urn:schemas-microsoft-com:office:spreadsheet")) { ultimoFormato = "Excel XML 2003"; return Xml2003(txt); }
            if (Regex.IsMatch(ini, @"(?i)<table|<tr[\s>]") || Regex.IsMatch(ini, "(?i)<html") && Regex.IsMatch(txt, @"(?i)<tr[\s>]"))
            {
                ultimoFormato = "HTML (tabela)"; return Html(txt);
            }
            ultimoFormato = "Texto/CSV";
            return Csv(txt);
        }

        static bool Ods(byte[] b)
        {
            string ini = Encoding.GetEncoding("iso-8859-1").GetString(b, 0, Math.Min(b.Length, 200));
            return ini.Contains("application/vnd.oasis.opendocument.spreadsheet");
        }

        // ---------------------------------------------------------------- CSV / TXT

        public static List<string[]> Csv(string txt)
        {
            txt = txt.Replace("\uFEFF", "");
            var ls = txt.SplitRe(@"\r?\n");
            char sep = '\0';
            int ini = 0;
            if (ls.Length > 0 && ls[0].Trim().ToLowerInvariant().Matches("\"?sep=.\"?")) { sep = ls[0].Trim().Replace("\"", "")[4]; ini = 1; }
            if (sep == '\0')
            {
                char melhor = '\0'; double nota = 0;
                foreach (char c in new[] { ';', '\t', ',', '|' })
                {
                    var freq = new Dictionary<int, int>();
                    int usadas = 0;
                    for (int i = 0; i < ls.Length && usadas < 60; i++)
                    {
                        if (ls[i].Trim().Length == 0) continue;
                        usadas++;
                        int n = Dividir(ls[i], c).Count - 1;
                        if (n > 0) freq.Somar(n, 1);
                    }
                    int moda = 0, vezes = 0;
                    foreach (var e in freq) if (e.Value > vezes || e.Value == vezes && e.Key > moda) { moda = e.Key; vezes = e.Value; }
                    double nt = vezes * Math.Min(moda, 12) * (c == ',' ? 0.9 : 1);
                    if (nt > nota) { nota = nt; melhor = c; }
                }
                sep = melhor;
                if (sep == ',')
                {
                    int virg = 0, dec = 0;
                    for (int i = 0; i < ls.Length && i < 200; i++)
                    {
                        virg += ls[i].Count(ch => ch == ',');
                        dec += Regex.Matches(ls[i], @"\d,\d{2}(?!\d)").Count;
                    }
                    if (virg > 0 && dec * 2 >= virg && !txt.Contains("\"")) sep = '\0';
                }
            }
            var r = new List<string[]>();
            StringBuilder pend = null;
            for (int i = ini; i < ls.Length; i++)
            {
                string l = ls[i];
                if (pend != null) { pend.Append(' ').Append(l); l = pend.ToString(); }
                if (sep != '\0' && AspasAbertas(l)) { pend = new StringBuilder(l); continue; }
                pend = null;
                if (l.Trim().Length == 0) continue;
                var c = sep == '\0' ? new List<string> { l } : Dividir(l, sep);
                r.Add(c.Select(x => x.Trim()).ToArray());
            }
            return r;
        }

        static bool AspasAbertas(string l) { return l.Count(ch => ch == '"') % 2 == 1; }

        static List<string> Dividir(string l, char sep)
        {
            var r = new List<string>();
            var c = new StringBuilder();
            bool asp = false;
            for (int i = 0; i < l.Length; i++)
            {
                char ch = l[i];
                if (ch == '"')
                {
                    if (asp && i + 1 < l.Length && l[i + 1] == '"') { c.Append('"'); i++; }
                    else asp = !asp;
                }
                else if (ch == sep && !asp) { r.Add(c.ToString()); c.Length = 0; }
                else c.Append(ch);
            }
            r.Add(c.ToString());
            for (int i = 0; i < r.Count; i++) { string x = r[i].Trim(); if (x.Matches("=\".*\"")) r[i] = x.Substring(2, x.Length - 3); }
            return r;
        }

        // ---------------------------------------------------------------- HTML disfarçado de .xls

        static readonly string[][] ENT = { new[]{"aacute","á"},new[]{"eacute","é"},new[]{"iacute","í"},new[]{"oacute","ó"},new[]{"uacute","ú"},new[]{"atilde","ã"},new[]{"otilde","õ"},
            new[]{"ccedil","ç"},new[]{"acirc","â"},new[]{"ecirc","ê"},new[]{"ocirc","ô"},new[]{"agrave","à"},new[]{"uuml","ü"},new[]{"ntilde","ñ"},new[]{"auml","ä"},new[]{"ouml","ö"},new[]{"szlig","ß"},
            new[]{"egrave","è"},new[]{"ordm","º"},new[]{"ordf","ª"},new[]{"deg","°"},new[]{"euro","€"},new[]{"pound","£"},new[]{"yen","¥"},new[]{"cent","¢"},new[]{"ndash","-"},new[]{"mdash","-"},new[]{"minus","-"} };

        static string Entidades(string s)
        {
            if (s.IndexOf('&') < 0) return s;
            foreach (var e in ENT)
            {
                s = s.Replace("&" + e[0] + ";", e[1]);
                string up = e[0].Substring(0, 1).ToUpperInvariant() + e[0].Substring(1);
                s = s.Replace("&" + up + ";", e[1].ToUpperInvariant());
            }
            return s;
        }

        internal static string Desesc(string s)
        {
            if (s.IndexOf('&') < 0) return s;
            s = Regex.Replace(s, "&#(x?)([0-9A-Fa-f]+);", m =>
            {
                try { return char.ConvertFromUtf32(Convert.ToInt32(m.Groups[2].Value, m.Groups[1].Value.Length == 0 ? 10 : 16)); }
                catch { return m.Value; }
            });
            return s.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&amp;", "&");
        }

        public static List<string[]> Html(string txt)
        {
            var r = new List<string[]>();
            string t = Regex.Replace(txt, "(?is)<script.*?</script>|<style.*?</style>|<!--.*?-->", " ");
            var td = new Regex(@"(?is)<t([dh])\b([^>]*)>(.*?)(?=</t[dh]>|<t[dh]\b|$)");
            foreach (Match tr in Regex.Matches(t, @"(?is)<tr\b[^>]*>(.*?)(?=</tr>|<tr\b|</table>)"))
            {
                var cel = new List<string>();
                foreach (Match m in td.Matches(tr.Groups[1].Value))
                {
                    string v = Entidades(m.Groups[3].Value);
                    v = Regex.Replace(v, @"(?i)<br\s*/?>", " ");
                    v = Regex.Replace(v, "<[^>]*>", " ").Replace("&nbsp;", " ");
                    v = Regex.Replace(Desesc(v).Replace('\u00A0', ' '), @"\s+", " ").Trim();
                    cel.Add(v);
                    var cs = Regex.Match(m.Groups[2].Value, "(?i)colspan\\s*=\\s*\"?(\\d+)");
                    if (cs.Success) for (int k = 1; k < Math.Min(int.Parse(cs.Groups[1].Value), 30); k++) cel.Add("");
                }
                if (cel.Any(c => c.Length > 0)) r.Add(cel.ToArray());
            }
            return r;
        }

        // ---------------------------------------------------------------- Excel XML 2003

        public static List<string[]> Xml2003(string txt)
        {
            var r = new List<string[]>();
            string t = Regex.Replace(Regex.Replace(txt, "<(/?)[A-Za-z]+:", "<$1"), @"\s[A-Za-z]+:(Index|Type)=", " $1=");
            var cell = new Regex(@"(?s)<Cell\b([^>]*?)(?:/>|>(.*?)</Cell>)");
            foreach (Match row in Regex.Matches(t, @"(?s)<Row\b[^>]*?(?:/>|>(.*?)</Row>)"))
            {
                if (!row.Groups[1].Success) continue;
                var cel = new List<string>();
                foreach (Match c in cell.Matches(row.Groups[1].Value))
                {
                    var ix = Regex.Match(c.Groups[1].Value, "Index=\"(\\d+)\"");
                    if (ix.Success) while (cel.Count < int.Parse(ix.Groups[1].Value) - 1) cel.Add("");
                    string corpo = c.Groups[2].Success ? c.Groups[2].Value : "";
                    var d = Regex.Match(corpo, @"(?s)<Data\b([^>]*)>(.*?)</Data>");
                    string v = "";
                    if (d.Success)
                    {
                        v = Desesc(Regex.Replace(d.Groups[2].Value, "<[^>]*>", "")).Trim();
                        if (d.Groups[1].Value.Contains("DateTime") && v.Matches(@"\d{4}-\d{2}-\d{2}.*")) v = v.Substring(8, 2) + "/" + v.Substring(5, 2) + "/" + v.Substring(0, 4);
                    }
                    cel.Add(v);
                }
                if (cel.Any(x => x.Length > 0)) r.Add(cel.ToArray());
            }
            return r;
        }

        // ---------------------------------------------------------------- XLSX sem bibliotecas

        static Dictionary<string, byte[]> Zip(byte[] arquivo, Func<string, bool> quero)
        {
            var p = new Dictionary<string, byte[]>();
            using (var ms = new MemoryStream(arquivo))
            using (var z = new ZipArchive(ms, ZipArchiveMode.Read))
                foreach (var e in z.Entries)
                {
                    if (!quero(e.FullName)) continue;
                    using (var s = e.Open()) using (var o = new MemoryStream()) { s.CopyTo(o); p[e.FullName] = o.ToArray(); }
                }
            return p;
        }

        static string SemPrefixo(string x)
        {
            if (!Regex.IsMatch(x, @"<[A-Za-z0-9]+:(row|c|si|t|sheet|xf|v)\b")) return x;
            return Regex.Replace(Regex.Replace(x, "<(/?)[A-Za-z0-9]+:", "<$1"), @"\s[A-Za-z0-9]+:id=", " r:id=");
        }

        static readonly Regex P_SI = new Regex("<si>(.*?)</si>", RegexOptions.Singleline);
        static readonly Regex P_T = new Regex("<t[^>]*>(.*?)</t>", RegexOptions.Singleline);
        static readonly Regex P_ROW = new Regex("<row[^>]*>(.*?)</row>", RegexOptions.Singleline);
        static readonly Regex P_CEL = new Regex("<c\\b([^>]*?)(?:/>|>(.*?)</c>)", RegexOptions.Singleline);
        static readonly Regex P_V = new Regex("<v>(.*?)</v>", RegexOptions.Singleline);

        static List<string[]> Xlsx(byte[] arquivo)
        {
            var p = Zip(arquivo, n => n == "xl/sharedStrings.xml" || n.StartsWith("xl/worksheets/sheet") || n == "xl/workbook.xml"
                                       || n == "xl/_rels/workbook.xml.rels" || n == "xl/styles.xml");
            Func<string, string> txt = k => p.ContainsKey(k) ? SemPrefixo(Encoding.UTF8.GetString(p[k])) : "";
            var ss = new List<string>();
            if (p.ContainsKey("xl/sharedStrings.xml"))
                foreach (Match si in P_SI.Matches(txt("xl/sharedStrings.xml")))
                {
                    var sb = new StringBuilder();
                    foreach (Match t in P_T.Matches(si.Groups[1].Value)) sb.Append(t.Groups[1].Value);
                    ss.Add(Desesc(sb.ToString()));
                }
            var estilosData = EstilosData(txt("xl/styles.xml"));
            var abas = new List<string>();
            string wb = txt("xl/workbook.xml"), rels = txt("xl/_rels/workbook.xml.rels");
            foreach (Match sh in Regex.Matches(wb, "<sheet [^>]*r:id=\"([^\"]+)\""))
            {
                var rl = Regex.Match(rels, "<Relationship [^>]*Id=\"" + Regex.Escape(sh.Groups[1].Value) + "\"[^>]*>");
                if (rl.Success)
                {
                    var t = Regex.Match(rl.Value, "Target=\"([^\"]+)\"");
                    if (t.Success) abas.Add(t.Groups[1].Value.StartsWith("/") ? t.Groups[1].Value.Substring(1) : "xl/" + t.Groups[1].Value);
                }
            }
            if (abas.Count == 0) abas.AddRange(p.Keys.Where(k => k.StartsWith("xl/worksheets/sheet")).OrderBy(k => k, StringComparer.Ordinal));
            var r = new List<string[]>();
            foreach (var aba in abas)
            {
                if (!p.ContainsKey(aba)) continue;
                foreach (Match mr in P_ROW.Matches(txt(aba)))
                {
                    var cel = new SortedDictionary<int, string>();
                    int colAnt = -1;
                    foreach (Match mc in P_CEL.Matches(mr.Groups[1].Value))
                    {
                        string attrs = mc.Groups[1].Value, corpo = mc.Groups[2].Success ? mc.Groups[2].Value : "";
                        string rf = Attr(attrs, "r");
                        int col = rf != null && rf.Matches(@"[A-Z]+\d*") ? Coluna(Regex.Replace(rf, @"\d", "")) : colAnt + 1;
                        colAnt = col;
                        string tipo = Attr(attrs, "t"), v = "";
                        var mv = P_V.Match(corpo);
                        if (tipo == "s") { if (mv.Success) { int i = int.Parse(mv.Groups[1].Value.Trim()); v = i < ss.Count ? ss[i] : ""; } }
                        else if (tipo == "inlineStr") { var sb = new StringBuilder(); foreach (Match t in P_T.Matches(corpo)) sb.Append(t.Groups[1].Value); v = Desesc(sb.ToString()); }
                        else if (mv.Success)
                        {
                            v = Desesc(mv.Groups[1].Value.Trim());
                            string s = Attr(attrs, "s");
                            int si;
                            if (s != null && int.TryParse(s, out si) && estilosData.Contains(si) && v.Matches(@"\d+(\.\d+)?")) v = SerialParaData(v);
                            else if (v.Matches(@"-?\d+\.0+")) v = v.Substring(0, v.IndexOf('.'));
                        }
                        if (v.Length > 0) cel[col] = v.Trim();
                    }
                    if (cel.Count == 0) continue;
                    var linha = Enumerable.Repeat("", cel.Keys.Last() + 1).ToArray();
                    foreach (var e in cel) linha[e.Key] = e.Value;
                    r.Add(linha);
                }
            }
            return r;
        }

        static HashSet<int> EstilosData(string x)
        {
            var r = new HashSet<int>();
            if (x.Length == 0) return r;
            var fmtData = new HashSet<int> { 14, 15, 16, 17, 22 };
            foreach (Match nf in Regex.Matches(x, "<numFmt numFmtId=\"(\\d+)\" formatCode=\"([^\"]*)\""))
            {
                string f = Regex.Replace(nf.Groups[2].Value.ToLowerInvariant(), @"\[[^]]*]", "");
                if (f.Contains("d") && f.Contains("m") && f.Contains("y") || f.Contains("dd/mm")) fmtData.Add(int.Parse(nf.Groups[1].Value));
            }
            var cx = Regex.Match(x, "<cellXfs[^>]*>(.*?)</cellXfs>", RegexOptions.Singleline);
            if (!cx.Success) return r;
            int i = 0;
            foreach (Match xf in Regex.Matches(cx.Groups[1].Value, "<xf [^>]*?numFmtId=\"(\\d+)\""))
            {
                if (fmtData.Contains(int.Parse(xf.Groups[1].Value))) r.Add(i);
                i++;
            }
            return r;
        }

        internal static string SerialParaData(string v)
        {
            try
            {
                var d = new DateTime(1899, 12, 30).AddDays(Math.Floor(J.ParseD(v)));
                return d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            catch { return v; }
        }

        static string Attr(string attrs, string nome)
        {
            var m = Regex.Match(attrs, @"\b" + nome + "=\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        static int Coluna(string letras)
        {
            int n = 0;
            foreach (char c in letras) n = n * 26 + (c - 'A' + 1);
            return n - 1;
        }

        // ---------------------------------------------------------------- ODS

        static string AttrOds(string at, string nome)
        {
            var m = Regex.Match(at, Regex.Escape(nome) + "=\"([^\"]*)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        static List<string[]> LerOds(byte[] arquivo)
        {
            var p = Zip(arquivo, n => n == "content.xml");
            if (!p.ContainsKey("content.xml")) throw new LeituraException("arquivo .ods sem content.xml");
            string x = Encoding.UTF8.GetString(p["content.xml"]);
            var r = new List<string[]>();
            var cel = new Regex(@"(?s)<table:(?:covered-)?table-cell\b([^>]*?)(?:/>|>(.*?)</table:(?:covered-)?table-cell>)");
            foreach (Match row in Regex.Matches(x, @"(?s)<table:table-row\b([^>]*?)(?:/>|>(.*?)</table:table-row>)"))
            {
                if (!row.Groups[2].Success) continue;
                var c = new List<string>();
                foreach (Match m in cel.Matches(row.Groups[2].Value))
                {
                    string at = m.Groups[1].Value, corpo = m.Groups[2].Success ? m.Groups[2].Value : "";
                    string v;
                    string tipo = AttrOds(at, "office:value-type");
                    if (tipo == "date" && AttrOds(at, "office:date-value") != null)
                    {
                        string d = AttrOds(at, "office:date-value");
                        v = d.Length >= 10 ? d.Substring(8, 2) + "/" + d.Substring(5, 2) + "/" + d.Substring(0, 4) : d;
                    }
                    else if ((tipo == "float" || tipo == "currency" || tipo == "percentage") && AttrOds(at, "office:value") != null) v = AttrOds(at, "office:value");
                    else
                    {
                        var sb = new StringBuilder();
                        foreach (Match pp in Regex.Matches(corpo, @"(?s)<text:p\b[^>]*>(.*?)</text:p>"))
                            sb.Append(sb.Length > 0 ? " " : "").Append(Regex.Replace(Regex.Replace(pp.Groups[1].Value, @"<text:s\b[^>]*/>", " "), "<[^>]*>", ""));
                        v = Desesc(sb.ToString()).Trim();
                    }
                    string rep = AttrOds(at, "table:number-columns-repeated");
                    int n = rep == null ? 1 : Math.Min(int.Parse(rep), v.Length == 0 ? 1 : 50);
                    for (int i = 0; i < n; i++) c.Add(v);
                }
                while (c.Count > 0 && c[c.Count - 1].Length == 0) c.RemoveAt(c.Count - 1);
                if (c.Count > 0) r.Add(c.ToArray());
            }
            return r;
        }
    }

    /// <summary>Excel 97-2003 (.xls / BIFF8) sem bibliotecas: primeira planilha como linhas de texto.</summary>
    public static class LeitorXls
    {
        static int I16(byte[] b, int p) { return b[p] | (b[p + 1] << 8); }
        static int I32(byte[] b, int p) { return b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24); }

        public static List<string[]> Linhas(byte[] f) { return Biff(StreamWorkbook(f)); }

        static byte[] StreamWorkbook(byte[] f)
        {
            if (f.Length < 512 || f[0] != 0xD0 || f[1] != 0xCF) throw new LeituraException("não é um .xls do Excel 97-2003 (abra no Excel e salve como .xlsx).");
            int ssz = 1 << (short)I16(f, 0x1E), mssz = 1 << (short)I16(f, 0x20);
            int nFat = I32(f, 0x2C), dir0 = I32(f, 0x30), corte = I32(f, 0x38), mini0 = I32(f, 0x3C), difat0 = I32(f, 0x44);
            int nSet = (f.Length - 512 + ssz - 1) / ssz;
            var dados = new byte[512 + nSet * ssz + ssz];
            Array.Copy(f, dados, f.Length);
            int off0 = ssz == 512 ? 512 : ssz;
            var fatSet = new List<int>();
            for (int i = 0; i < 109 && fatSet.Count < nFat; i++) fatSet.Add(I32(dados, 0x4C + i * 4));
            int d = difat0;
            while (fatSet.Count < nFat && d >= 0 && d < nSet)
            {
                int bs = off0 + d * ssz;
                for (int i = 0; i < ssz / 4 - 1 && fatSet.Count < nFat; i++) fatSet.Add(I32(dados, bs + i * 4));
                d = I32(dados, bs + ssz - 4);
            }
            var fat = new int[fatSet.Count * (ssz / 4)];
            for (int i = 0; i < fatSet.Count; i++)
            {
                int bs = off0 + fatSet[i] * ssz;
                for (int j = 0; j < ssz / 4; j++) fat[i * (ssz / 4) + j] = bs >= 0 && bs + j * 4 + 4 <= dados.Length ? I32(dados, bs + j * 4) : -2;
            }
            var dir = Cadeia(dados, fat, dir0, ssz, off0);
            int rootIni = -1;
            for (int e = 0; e + 128 <= dir.Length; e += 128)
            {
                int nl = I16(dir, e + 0x40);
                int tipo = dir[e + 0x42];
                if (nl < 2 || tipo == 0) continue;
                string nome = Encoding.Unicode.GetString(dir, e, Math.Min(64, nl - 2));
                int ini = I32(dir, e + 0x74);
                long tam = (uint)I32(dir, e + 0x78);
                if (tipo == 5) { rootIni = ini; continue; }
                if (nome == "Workbook" || nome == "Book")
                {
                    if (tam >= corte) { var c = Cadeia(dados, fat, ini, ssz, off0); Array.Resize(ref c, (int)tam); return c; }
                    var mini = Cadeia(dados, fat, rootIni, ssz, off0);
                    var mfatB = Cadeia(dados, fat, mini0, ssz, off0);
                    var mfat = new int[mfatB.Length / 4];
                    for (int i = 0; i < mfat.Length; i++) mfat[i] = I32(mfatB, i * 4);
                    var r = new byte[(int)tam];
                    int s = ini, p = 0;
                    while (p < tam && s >= 0 && s < mfat.Length)
                    {
                        int n = (int)Math.Min(mssz, tam - p);
                        int cp = Math.Min(n, mini.Length - s * mssz);
                        if (cp > 0) Array.Copy(mini, s * mssz, r, p, cp);
                        p += n; s = mfat[s];
                    }
                    return r;
                }
            }
            throw new LeituraException("planilha sem a parte \"Workbook\" (arquivo .xls inválido).");
        }

        static byte[] Cadeia(byte[] dados, int[] fat, int ini, int ssz, int off0)
        {
            var o = new MemoryStream();
            int s = ini, guarda = 0;
            while (s >= 0 && s < fat.Length && guarda++ < fat.Length)
            {
                int bs = off0 + s * ssz;
                if (bs + ssz > dados.Length) break;
                o.Write(dados, bs, ssz);
                s = fat[s];
            }
            return o.ToArray();
        }

        static List<string[]> Biff(byte[] w)
        {
            var sst = new List<string>();
            var formatos = new Dictionary<int, string>();
            var xfFmt = new List<int>();
            var cel = new SortedDictionary<int, SortedDictionary<int, string>>();
            int bofs = 0, p = 0;
            while (p + 4 <= w.Length)
            {
                int tipo = I16(w, p), tam = I16(w, p + 2), d = p + 4;
                int prox = d + tam;
                if (prox > w.Length) break;
                switch (tipo)
                {
                    case 0x809: bofs++; break;
                    case 0x0A: if (bofs >= 2) return Montar(cel); break;
                    case 0x41E: formatos[I16(w, d)] = XlString(w, d + 2, true).Item1; break;
                    case 0xE0: xfFmt.Add(I16(w, d + 2)); break;
                    case 0xFC:
                        {
                            var pedacos = new List<int[]> { new[] { d, prox } };
                            int q = prox;
                            while (q + 4 <= w.Length && I16(w, q) == 0x3C)
                            {
                                int t2 = I16(w, q + 2);
                                pedacos.Add(new[] { q + 4, q + 4 + t2 });
                                q += 4 + t2;
                            }
                            LerSst(w, pedacos, sst);
                            prox = q;
                            break;
                        }
                    case 0xFD: if (bofs >= 2) Por(cel, w, d, Idx(sst, I32(w, d + 6))); break;
                    case 0x204: if (bofs >= 2) Por(cel, w, d, XlString(w, d + 6, true).Item1); break;
                    case 0x203: if (bofs >= 2) Por(cel, w, d, Numero(BitConverter.ToDouble(w, d + 6), Fmt(I16(w, d + 4), xfFmt, formatos))); break;
                    case 0x27E: if (bofs >= 2) Por(cel, w, d, Numero(Rk(I32(w, d + 6)), Fmt(I16(w, d + 4), xfFmt, formatos))); break;
                    case 0xBD:
                        if (bofs >= 2)
                        {
                            int lin = I16(w, d), c0 = I16(w, d + 2), n = (tam - 6) / 6;
                            for (int i = 0; i < n; i++)
                            {
                                int xf = I16(w, d + 4 + i * 6);
                                double v = Rk(I32(w, d + 6 + i * 6));
                                cel.ObterOuCriar(lin)[c0 + i] = Numero(v, Fmt(xf, xfFmt, formatos));
                            }
                        }
                        break;
                    case 0x06:
                        if (bofs >= 2)
                        {
                            bool especial = I16(w, d + 12) == 0xFFFF;
                            if (!especial) Por(cel, w, d, Numero(BitConverter.ToDouble(w, d + 6), Fmt(I16(w, d + 4), xfFmt, formatos)));
                            else if (w[d + 6] == 0 && prox + 4 <= w.Length && I16(w, prox) == 0x207) Por(cel, w, d, XlString(w, prox + 4, true).Item1);
                        }
                        break;
                }
                p = prox;
            }
            return Montar(cel);
        }

        static void Por(SortedDictionary<int, SortedDictionary<int, string>> cel, byte[] w, int d, string v)
        {
            if (v == null || v.Trim().Length == 0) return;
            cel.ObterOuCriar(I16(w, d))[I16(w, d + 2)] = v;
        }

        static string Idx(List<string> sst, int i) { return i >= 0 && i < sst.Count ? sst[i] : ""; }

        static List<string[]> Montar(SortedDictionary<int, SortedDictionary<int, string>> cel)
        {
            var r = new List<string[]>();
            foreach (var l in cel.Values)
            {
                var a = Enumerable.Repeat("", l.Keys.Last() + 1).ToArray();
                foreach (var e in l) a[e.Key] = e.Value.Trim();
                r.Add(a);
            }
            return r;
        }

        static string Fmt(int xf, List<int> xfFmt, Dictionary<int, string> formatos)
        {
            if (xf < 0 || xf >= xfFmt.Count) return "";
            int id = xfFmt[xf];
            string f;
            if (!formatos.TryGetValue(id, out f) && (id >= 14 && id <= 17 || id == 22)) return "dd/mm/yyyy";
            return f ?? "";
        }

        static bool FormatoData(string f)
        {
            string s = Regex.Replace(f.ToLowerInvariant(), "\\[[^]]*]|\"[^\"]*\"", "");
            return s.Contains("d") && s.Contains("m") && s.Contains("y") || s.Contains("dd/mm") || s.Contains("mm/dd");
        }

        static bool FormatoDC(string f)
        {
            string s = f.ToLowerInvariant();
            return s.Contains(";") && (s.Contains("\\d") || s.Contains("\"d\"")) && (s.Contains("\\c") || s.Contains("\"c\""));
        }

        static string Numero(double v, string formato)
        {
            if (FormatoData(formato) && v >= 1 && v < 80000)
                return new DateTime(1899, 12, 30).AddDays(Math.Floor(v)).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            if (FormatoDC(formato))
                return Math.Abs(v).ToString("N2", J.BR) + (v > 0.004 ? "D" : v < -0.004 ? "C" : "");
            if (v == Math.Round(v) && Math.Abs(v) < 1e15) return ((long)v).ToString(CultureInfo.InvariantCulture);
            return ((decimal)v).ToString(CultureInfo.InvariantCulture);
        }

        static double Rk(int rk)
        {
            double v;
            if ((rk & 2) != 0) v = rk >> 2;
            else v = BitConverter.Int64BitsToDouble(((long)(rk & unchecked((int)0xFFFFFFFC))) << 32);
            return (rk & 1) != 0 ? v / 100.0 : v;
        }

        static Tuple<string, int> XlString(byte[] b, int p, bool cch16)
        {
            int cch = cch16 ? I16(b, p) : b[p];
            p += cch16 ? 2 : 1;
            int fl = b[p++];
            int runs = 0, ext = 0;
            if ((fl & 8) != 0) { runs = I16(b, p); p += 2; }
            if ((fl & 4) != 0) { ext = I32(b, p); p += 4; }
            string s;
            if ((fl & 1) != 0) { s = Encoding.Unicode.GetString(b, p, Math.Min(cch * 2, b.Length - p)); p += cch * 2; }
            else { s = Encoding.GetEncoding("iso-8859-1").GetString(b, p, Math.Min(cch, b.Length - p)); p += cch; }
            p += runs * 4 + ext;
            return Tuple.Create(s, p);
        }

        static void LerSst(byte[] b, List<int[]> ped, List<string> sst)
        {
            var latin = Encoding.GetEncoding("iso-8859-1");
            int pi = 0, p = ped[0][0] + 8;
            int fimP = ped[0][1];
            while (true)
            {
                if (p >= fimP) { if (++pi >= ped.Count) return; p = ped[pi][0]; fimP = ped[pi][1]; }
                int cch = I16(b, p); p += 2;
                int fl = b[p++];
                int runs = 0, ext = 0;
                if ((fl & 8) != 0) { runs = I16(b, p); p += 2; }
                if ((fl & 4) != 0) { ext = I32(b, p); p += 4; }
                bool uni = (fl & 1) != 0;
                var sb = new StringBuilder();
                int falta = cch;
                while (falta > 0)
                {
                    if (p >= fimP)
                    {
                        if (++pi >= ped.Count) { sst.Add(sb.ToString()); return; }
                        p = ped[pi][0]; fimP = ped[pi][1];
                        uni = (b[p++] & 1) != 0;
                    }
                    int cabe = uni ? (fimP - p) / 2 : fimP - p;
                    int n = Math.Min(falta, cabe);
                    if (n <= 0) { p = fimP; continue; }
                    sb.Append(uni ? Encoding.Unicode.GetString(b, p, n * 2) : latin.GetString(b, p, n));
                    p += uni ? n * 2 : n;
                    falta -= n;
                }
                sst.Add(sb.ToString());
                int pular = runs * 4 + ext;
                while (pular > 0)
                {
                    if (p >= fimP) { if (++pi >= ped.Count) return; p = ped[pi][0]; fimP = ped[pi][1]; }
                    int n = Math.Min(pular, fimP - p);
                    p += n; pular -= n;
                }
            }
        }
    }
}

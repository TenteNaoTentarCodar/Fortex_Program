using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Gravador .xlsx simples (sem formatação): várias abas, células de texto, número e data.
    /// Só a data recebe o formato de data padrão do Excel (dd/mm/aaaa), para ficar filtrável e ordenável.
    /// </summary>
    public class GravadorXlsx
    {
        public class Aba
        {
            public string nome;
            public readonly List<object[]> linhas = new List<object[]>();
            public Aba(string nome) { this.nome = nome; }
            public void Add(params object[] valores) { linhas.Add(valores); }
        }

        public readonly List<Aba> abas = new List<Aba>();

        public Aba NovaAba(string nome)
        {
            var a = new Aba(NomeAba(nome));
            abas.Add(a);
            return a;
        }

        string NomeAba(string n)
        {
            var sb = new StringBuilder();
            foreach (char c in n) if ("[]:*?/\\".IndexOf(c) < 0) sb.Append(c);
            string r = sb.ToString().Trim();
            if (r.Length == 0) r = "Aba";
            if (r.Length > 31) r = r.Substring(0, 31);
            string b = r;
            int i = 2;
            while (abas.Exists(a => string.Equals(a.nome, r, StringComparison.OrdinalIgnoreCase)))
            {
                string suf = " " + i++;
                r = (b.Length + suf.Length > 31 ? b.Substring(0, 31 - suf.Length) : b) + suf;
            }
            return r;
        }

        public void Gravar(string arquivo)
        {
            if (abas.Count == 0) NovaAba("Planilha1");
            string tmp = arquivo + ".tmp";
            if (File.Exists(tmp)) File.Delete(tmp);
            using (var fs = new FileStream(tmp, FileMode.CreateNew))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                Escrever(zip, "[Content_Types].xml", ContentTypes());
                Escrever(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                var wb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
                var rels = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                for (int i = 0; i < abas.Count; i++)
                {
                    wb.Append("<sheet name=\"").Append(Esc(abas[i].nome)).Append("\" sheetId=\"").Append(i + 1).Append("\" r:id=\"rId").Append(i + 1).Append("\"/>");
                    rels.Append("<Relationship Id=\"rId").Append(i + 1).Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet").Append(i + 1).Append(".xml\"/>");
                }
                rels.Append("<Relationship Id=\"rId").Append(abas.Count + 1).Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
                wb.Append("</sheets></workbook>");
                rels.Append("</Relationships>");
                Escrever(zip, "xl/workbook.xml", wb.ToString());
                Escrever(zip, "xl/_rels/workbook.xml.rels", rels.ToString());
                Escrever(zip, "xl/styles.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                    + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                    + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                    + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                    + "<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                    + "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>"
                    + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");
                for (int i = 0; i < abas.Count; i++) Escrever(zip, "xl/worksheets/sheet" + (i + 1) + ".xml", Planilha(abas[i]));
            }
            if (File.Exists(arquivo)) File.Delete(arquivo);
            File.Move(tmp, arquivo);
        }

        string ContentTypes()
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 0; i < abas.Count; i++)
                sb.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i + 1).Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            return sb.Append("</Types>").ToString();
        }

        static string Planilha(Aba a)
        {
            var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (int r = 0; r < a.linhas.Count; r++)
            {
                sb.Append("<row r=\"").Append(r + 1).Append("\">");
                var lin = a.linhas[r];
                for (int c = 0; c < lin.Length; c++)
                {
                    object v = lin[c];
                    if (v == null) continue;
                    string refc = Coluna(c) + (r + 1);
                    if (v is DateTime)
                    {
                        double serial = ((DateTime)v).Date.ToOADate();
                        sb.Append("<c r=\"").Append(refc).Append("\" s=\"1\"><v>").Append(serial.ToString(CultureInfo.InvariantCulture)).Append("</v></c>");
                    }
                    else if (v is double || v is int || v is long || v is decimal || v is float)
                    {
                        double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                        if (double.IsNaN(d) || double.IsInfinity(d)) continue;
                        sb.Append("<c r=\"").Append(refc).Append("\"><v>").Append(Math.Round(d, 6).ToString("R", CultureInfo.InvariantCulture)).Append("</v></c>");
                    }
                    else
                    {
                        string s = Convert.ToString(v, CultureInfo.InvariantCulture);
                        if (s.Length == 0) continue;
                        if (s.Length > 32767) s = s.Substring(0, 32767);
                        sb.Append("<c r=\"").Append(refc).Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">").Append(Esc(s)).Append("</t></is></c>");
                    }
                }
                sb.Append("</row>");
            }
            return sb.Append("</sheetData></worksheet>").ToString();
        }

        static string Coluna(int i)
        {
            string s = "";
            i++;
            while (i > 0) { int m = (i - 1) % 26; s = (char)('A' + m) + s; i = (i - 1) / 26; }
            return s;
        }

        static string Esc(string s)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
                if (ch == '&') sb.Append("&amp;");
                else if (ch == '<') sb.Append("&lt;");
                else if (ch == '>') sb.Append("&gt;");
                else if (ch == '"') sb.Append("&quot;");
                else if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') sb.Append(' ');
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        static void Escrever(ZipArchive zip, string nome, string conteudo)
        {
            var e = zip.CreateEntry(nome, CompressionLevel.Optimal);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false))) w.Write(conteudo);
        }
    }
}

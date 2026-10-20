using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ExtratoComprovantes
{
    /// <summary>Lê OFX 1.x (SGML, sem tags de fechamento) e 2.x (XML).</summary>
    public static class LeitorOFX
    {
        static readonly Regex BLOCO = new Regex(@"<STMTTRN>(.*?)(</STMTTRN>|(?=<STMTTRN>)|</BANKTRANLIST>)", RegexOptions.Singleline | RegexOptions.IgnoreCase);

        public static List<Lancamento> Ler(string arquivo)
        {
            byte[] bytes = File.ReadAllBytes(arquivo);
            string cab = Encoding.GetEncoding("iso-8859-1").GetString(bytes, 0, Math.Min(bytes.Length, 600)).ToUpperInvariant();
            string txt = (cab.Contains("UTF-8") || cab.Contains("CHARSET:UTF") || Utf8Valido(bytes)) ? new UTF8Encoding(false).GetString(bytes) : J.Cp1252.GetString(bytes);
            txt = txt.Replace("﻿", "");
            var lista = new List<Lancamento>();
            foreach (Match m in BLOCO.Matches(txt))
            {
                string b = m.Groups[1].Value;
                string dt = Tag(b, "DTPOSTED"), amt = Tag(b, "TRNAMT");
                if (dt == null || amt == null || dt.Length < 8) continue;
                string memo = Tag(b, "MEMO"), name = Tag(b, "NAME");
                string hist;
                if (string.IsNullOrEmpty(memo)) hist = name ?? "";
                else if (!string.IsNullOrEmpty(name) && !memo.ToUpperInvariant().Contains(name.ToUpperInvariant())) hist = name + " - " + memo;
                else hist = memo;
                string nh = Regras.Normalizar(hist).Trim();
                if (Regex.IsMatch(nh, @"^(SALDO|S A L D O)\b") || nh.Contains("SALDO ANTERIOR") || nh.Contains("SALDO DO DIA") || nh.Contains("SALDO TOTAL")) continue;
                string doc = Tag(b, "CHECKNUM");
                if (!string.IsNullOrEmpty(doc) && !doc.Matches("0+") && !hist.Contains(doc)) hist = hist + " DOC " + doc;
                double v;
                if (!double.TryParse(amt.Replace(" ", "").Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out v)) continue;
                if (v == 0) continue;
                string data = dt.Substring(6, 2) + "/" + dt.Substring(4, 2) + "/" + dt.Substring(0, 4);
                var l = new Lancamento(data, Limpar(hist), v, v < 0);
                l.origem = Path.GetFileName(arquivo);
                lista.Add(l);
            }
            return lista;
        }

        static string Tag(string bloco, string nome)
        {
            var m = Regex.Match(bloco, "<" + nome + ">([^<\\r\\n]*)", RegexOptions.IgnoreCase);
            return m.Success ? Decodificar(m.Groups[1].Value.Trim()) : null;
        }

        static string Decodificar(string s)
        {
            return s.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'");
        }

        static bool Utf8Valido(byte[] b)
        {
            if (!b.Any(x => x >= 0x80)) return false;
            try { new UTF8Encoding(false, true).GetString(b); return true; } catch (DecoderFallbackException) { return false; }
        }

        public static string Limpar(string s)
        {
            return Regex.Replace(Regex.Replace(s ?? "", @"[\r\n\t]+", " "), " {2,}", " ").Trim();
        }
    }

    /// <summary>Quicken Interchange Format (QIF).</summary>
    public static class LeitorQIF
    {
        public static LeitorExtrato.Resultado Ler(string t, string origem)
        {
            var r = new LeitorExtrato.Resultado { formato = "QIF" };
            var regs = new List<Dictionary<char, string>>();
            var atual = new Dictionary<char, string>();
            bool investimento = false;
            foreach (var l0 in t.SplitRe(@"\r?\n"))
            {
                string l = l0.Trim();
                if (l.Length == 0) continue;
                if (l.StartsWith("!"))
                {
                    string u = l.ToUpperInvariant();
                    if (u.StartsWith("!TYPE:")) investimento = u.Contains("INVST");
                    continue;
                }
                char c = l[0];
                string v = l.Substring(1).Trim();
                if (c == '^') { if (atual.Count > 0) regs.Add(atual); atual = new Dictionary<char, string>(); continue; }
                if (c == 'M' && atual.ContainsKey('M')) v = atual['M'] + " " + v;
                if (c == 'S' || c == 'E' || c == '$') continue;
                if (!atual.ContainsKey(c)) atual[c] = v;
                if (c == 'M') atual['M'] = v;
            }
            if (atual.Count > 0) regs.Add(atual);
            var datas = new List<string>();
            foreach (var g in regs) if (g.ContainsKey('D')) datas.Add(DataQif(g['D']));
            int ordem = Formatos.Ordem(datas);
            bool ambigua = ordem == Formatos.DMY;
            foreach (var d in datas) { var p = d.Split('/'); int x; if (p.Length >= 2 && p[0].Matches("[0-9]+") && int.TryParse(p[0], out x) && x > 12) ambigua = false; }
            foreach (var g in regs)
            {
                string d = Formatos.Data(DataQif(g.GetOr('D', "")), ordem, null);
                string sv = g.ContainsKey('T') ? g['T'] : g.GetOr('U', null);
                double? v = Formatos.Valor(sv);
                if (d.Length == 0 || v == null || Math.Abs(v.Value) < 0.005) continue;
                var h = new StringBuilder();
                foreach (char c in new[] { 'P', 'M', 'L' })
                {
                    string x = g.GetOr(c, null);
                    if (!string.IsNullOrEmpty(x) && !Formatos.JaContem(h.ToString(), x)) h.Append(h.Length > 0 ? " - " : "").Append(x);
                }
                string n = g.GetOr('N', null);
                if (!string.IsNullOrEmpty(n) && !Regex.IsMatch(n, "^(?:DEP|ATM|EFT|XFER|TXFR)$", RegexOptions.IgnoreCase)) h.Append(" DOC ").Append(n);
                r.lancamentos.Add(LeitorExtrato.Novo(d, h.ToString(), v.Value, origem));
            }
            if (investimento) r.aviso = "QIF de investimentos: lidos só os valores totais de cada operação.";
            else if (ambigua && regs.Count > 0) r.aviso = "QIF: datas lidas como dia/mês (nenhum dia acima de 12 permite confirmar) — confira.";
            return r;
        }

        static string DataQif(string s) { return s.Replace('\'', '/').Replace(" ", "").Replace('-', '/'); }
    }

    /// <summary>SWIFT MT940 / MT942.</summary>
    public static class LeitorMT940
    {
        static readonly Regex L61 = new Regex(@"^([0-9]{6})([0-9]{4})?(R?[DC])([A-Z])?([0-9]+,[0-9]*)([A-Z][A-Za-z0-9_]{3})?([^/]*?)(?://(.*))?$", RegexOptions.Singleline);
        static readonly Regex SALDO = new Regex(@"^([DC])([0-9]{6})([A-Z]{3})([0-9]+,[0-9]*)");
        static readonly Regex CAMPO = new Regex(@"^:([0-9]{2}[A-Z]?):(.*)$");

        static double Num(string s) { return J.ParseD(s.Replace(',', '.') + (s.EndsWith(",") ? "0" : "")); }

        public static LeitorExtrato.Resultado Ler(string t, string origem)
        {
            var r = new LeitorExtrato.Resultado { formato = "SWIFT MT940" };
            var campos = new List<string[]>();
            foreach (var l0 in t.Replace("\r", "").SplitRe("\n"))
            {
                string l = l0;
                if (l.StartsWith("{") || l == "-}" || l == "-")
                {
                    int i = l.IndexOf(":20:", StringComparison.Ordinal);
                    if (i < 0) continue;
                    l = l.Substring(i);
                }
                var m = CAMPO.Match(l);
                if (m.Success) campos.Add(new[] { m.Groups[1].Value, m.Groups[2].Value });
                else if (campos.Count > 0) campos[campos.Count - 1][1] += "\n" + l;
            }
            int extratos = 0;
            Lancamento ultimo = null;
            foreach (var c in campos)
            {
                string tag = c[0], v = c[1].Trim();
                switch (tag)
                {
                    case "20": extratos++; break;
                    case "60F":
                    case "60M":
                        {
                            double? s = Saldo(v);
                            if (r.saldoInicial == null && s != null && (tag == "60F" || extratos <= 1)) r.saldoInicial = s;
                            break;
                        }
                    case "62F":
                    case "62M": { double? s = Saldo(v); if (s != null) r.saldoFinal = s; break; }
                    case "61":
                        {
                            string linha1 = v.Split('\n')[0];
                            string sub = v.Contains("\n") ? v.Substring(v.IndexOf('\n') + 1).Trim() : "";
                            var m = L61.Match(linha1);
                            if (!m.Success) { ultimo = null; break; }
                            string vd = m.Groups[1].Value;
                            int ano = 2000 + int.Parse(vd.Substring(0, 2));
                            string mm = vd.Substring(2, 2), dd = vd.Substring(4, 2);
                            if (m.Groups[2].Success)
                            {
                                string em = m.Groups[2].Value.Substring(0, 2), ed = m.Groups[2].Value.Substring(2);
                                if (em == "12" && mm == "01") ano--; else if (em == "01" && mm == "12") ano++;
                                mm = em; dd = ed;
                            }
                            string data = dd + "/" + mm + "/" + ano;
                            double valor = Num(m.Groups[5].Value);
                            string dc = m.Groups[3].Value;
                            bool saida = dc == "D" || dc == "RC";
                            var h = new StringBuilder();
                            string rf = m.Groups[7].Success ? m.Groups[7].Value.Trim() : "";
                            if (rf.Length > 0 && !rf.Equals("NONREF", StringComparison.OrdinalIgnoreCase)) h.Append(rf);
                            if (sub.Length > 0) h.Append(h.Length > 0 ? " " : "").Append(sub.Replace("\n", " "));
                            if (valor < 0.005) { ultimo = null; break; }
                            ultimo = LeitorExtrato.Novo(data, h.ToString(), saida ? -valor : valor, origem);
                            if (dc.StartsWith("R")) ultimo.historico = LeitorOFX.Limpar("ESTORNO " + ultimo.historico);
                            r.lancamentos.Add(ultimo);
                            break;
                        }
                    case "86":
                        {
                            if (ultimo == null) break;
                            string info = Info86(v);
                            if (info.Length > 0)
                            {
                                string h = ultimo.historico == "(sem histórico)" ? "" : ultimo.historico;
                                ultimo.historico = LeitorOFX.Limpar(info + (h.Length == 0 ? "" : " " + h));
                            }
                            ultimo = null;
                            break;
                        }
                }
            }
            if (extratos > 1) r.aviso = "MT940 com " + extratos + " extratos no mesmo arquivo.";
            return r;
        }

        static double? Saldo(string v)
        {
            var m = SALDO.Match(v.Replace("\n", ""));
            if (!m.Success) return null;
            double x = Num(m.Groups[4].Value);
            return m.Groups[1].Value == "D" ? -x : x;
        }

        static readonly Regex CAMPOS86 = new Regex(@"/(NAME|REMI|EREF|ORDP|BENM|PURP|TRCD|CNTP|IREF)/([^/]*(?:/(?!NAME/|REMI/|EREF/|ORDP/|BENM/|PURP/|TRCD/|CNTP/|IREF/)[^/]*)*)");

        internal static string Info86(string v)
        {
            string s = v.Replace("\n", "");
            if (Regex.IsMatch(s, @"^[0-9]{3}\?", RegexOptions.Singleline))
            {
                StringBuilder nome = new StringBuilder(), texto = new StringBuilder(), tipo = new StringBuilder();
                foreach (var p in s.Substring(3).SplitRe(@"\?"))
                {
                    if (p.Length < 2) continue;
                    string k = p.Substring(0, 2), x = p.Substring(2).Trim();
                    if (x.Length == 0) continue;
                    if (k == "00") tipo.Append(x).Append(' ');
                    else if (k.Matches("2[0-9]|6[0-3]")) texto.Append(x);
                    else if (k.Matches("3[23]")) nome.Append(x).Append(' ');
                }
                return LeitorOFX.Limpar(Regex.Replace(tipo + " " + nome + " " + texto, @"\b(EREF|KREF|MREF|CRED|DEBT|SVWZ|ABWA|ABWE)\+", " "));
            }
            if (s.Contains("/"))
            {
                var f = new Dictionary<string, string>();
                foreach (Match m in CAMPOS86.Matches(s)) f[m.Groups[1].Value] = m.Groups[2].Value.Trim();
                if (f.Count > 0)
                {
                    var b = new StringBuilder();
                    foreach (var k in new[] { "NAME", "ORDP", "BENM", "CNTP", "REMI", "PURP", "EREF" })
                        if (f.ContainsKey(k) && f[k].Length > 0) b.Append(b.Length > 0 ? " - " : "").Append(f[k].Replace("/", " "));
                    if (b.Length > 0) return LeitorOFX.Limpar(b.ToString());
                }
            }
            return LeitorOFX.Limpar(s);
        }
    }

    /// <summary>ISO 20022 CAMT.052 / 053 / 054 — qualquer versão do namespace.</summary>
    public static class LeitorCAMT
    {
        public static LeitorExtrato.Resultado Ler(byte[] xml, string origem)
        {
            var r = new LeitorExtrato.Resultado();
            var doc = new XmlDocument();
            try
            {
                var cfg = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (var ms = new MemoryStream(xml))
                using (var rd = XmlReader.Create(ms, cfg)) doc.Load(rd);
            }
            catch (Exception e) { throw new LeituraException(origem + ": XML CAMT inválido (" + e.Message + ")", e); }
            var raizEl = doc.DocumentElement;
            string raiz = "";
            foreach (var k in new[] { "BkToCstmrStmt", "BkToCstmrAcctRpt", "BkToCstmrDbtCdtNtfctn" })
                if (Todos(raizEl, k).Count > 0 || raizEl.LocalName == k) { raiz = k; break; }
            r.formato = raiz == "BkToCstmrAcctRpt" ? "CAMT.052" : raiz == "BkToCstmrDbtCdtNtfctn" ? "CAMT.054" : "CAMT.053";
            var stmts = new List<XmlElement>();
            foreach (var k in new[] { "Stmt", "Rpt", "Ntfctn" }) stmts.AddRange(Todos(raizEl, k));
            int pendentes = 0;
            foreach (var st in stmts)
            {
                if (stmts.Count == 1)
                {
                    foreach (var bal in Filhos(st, "Bal"))
                    {
                        string tp = Texto(Caminho(bal, "Tp", "CdOrPrtry", "Cd"));
                        double? v = Valor(bal);
                        if (v == null) continue;
                        if ((tp == "OPBD" || tp == "PRCD") && r.saldoInicial == null) r.saldoInicial = v;
                        if (tp == "CLBD") r.saldoFinal = v;
                    }
                }
                foreach (var e in Filhos(st, "Ntry"))
                {
                    string sts = Texto(Caminho(e, "Sts", "Cd"));
                    if (sts.Length == 0) sts = Texto(Caminho(e, "Sts"));
                    if (sts == "PDNG" || sts == "INFO") { pendentes++; continue; }
                    double? v = Valor(e);
                    if (v == null || Math.Abs(v.Value) < 0.005) continue;
                    string data = DataIso(Texto(Caminho(e, "BookgDt", "Dt")));
                    if (data.Length == 0) data = DataIso(Texto(Caminho(e, "BookgDt", "DtTm")));
                    if (data.Length == 0) data = DataIso(Texto(Caminho(e, "ValDt", "Dt")));
                    if (data.Length == 0) data = DataIso(Texto(Caminho(e, "ValDt", "DtTm")));
                    if (data.Length == 0) continue;
                    bool debito = v < 0;
                    string lado = debito ? "Cdtr" : "Dbtr";
                    var partes = new List<string>();
                    Action<string> add = x => { if (!partes.Contains(x)) partes.Add(x); };
                    foreach (var tx in Todos(e, "TxDtls"))
                    {
                        string cp = Texto(Caminho(tx, "RltdPties", lado, "Nm"));
                        if (cp.Length == 0) cp = Texto(Caminho(tx, "RltdPties", lado, "Pty", "Nm"));
                        if (cp.Length > 0) add(cp);
                        string d2 = Texto(Caminho(tx, "RltdPties", lado, "Id", "OrgId", "Othr", "Id"));
                        if (d2.Length == 0) d2 = Texto(Caminho(tx, "RltdPties", lado, "Pty", "Id", "OrgId", "Othr", "Id"));
                        if (d2.Length > 0) add(d2);
                        foreach (var u in Todos(tx, "Ustrd")) if (Texto(u).Length > 0) add(Texto(u));
                        string rf = Texto(Caminho(tx, "RmtInf", "Strd", "CdtrRefInf", "Ref"));
                        if (rf.Length > 0) add(rf);
                        string inf = Texto(Caminho(tx, "AddtlTxInf"));
                        if (inf.Length > 0) add(inf);
                    }
                    string adic = Texto(Caminho(e, "AddtlNtryInf"));
                    if (adic.Length > 0) add(adic);
                    if (partes.Count == 0)
                    {
                        string cod = Texto(Caminho(e, "BkTxCd", "Prtry", "Cd"));
                        if (cod.Length == 0) cod = Texto(Caminho(e, "BkTxCd", "Domn", "Fmly", "SubFmlyCd"));
                        if (cod.Length > 0) add(cod);
                    }
                    string acct = Texto(Caminho(e, "AcctSvcrRef"));
                    var h = new StringBuilder();
                    foreach (var p in partes) if (!Formatos.JaContem(h.ToString(), p)) h.Append(h.Length > 0 ? " - " : "").Append(p);
                    if (acct.Length > 0 && h.Length < 200) h.Append(" DOC ").Append(acct);
                    var l = LeitorExtrato.Novo(data, h.ToString(), v.Value, origem);
                    if ("true".Equals(Texto(Caminho(e, "RvslInd")), StringComparison.OrdinalIgnoreCase)) l.historico = LeitorOFX.Limpar("ESTORNO " + l.historico);
                    r.lancamentos.Add(l);
                }
            }
            if (stmts.Count > 1) r.aviso = r.formato + " com " + stmts.Count + " contas/extratos no mesmo arquivo — lidos todos.";
            if (pendentes > 0) r.aviso = (r.aviso.Length == 0 ? "" : r.aviso + " ") + pendentes + " lançamento(s) pendente(s) (PDNG) ignorado(s).";
            return r;
        }

        static double? Valor(XmlElement e)
        {
            var amt = Filhos(e, "Amt").FirstOrDefault();
            if (amt == null) return null;
            double v;
            if (!double.TryParse(Texto(amt), NumberStyles.Float, CultureInfo.InvariantCulture, out v)) return null;
            return Texto(Caminho(e, "CdtDbtInd")) == "DBIT" ? -v : v;
        }

        static string DataIso(string s)
        {
            if (s.Length < 10 || !s.Substring(0, 10).Matches("[0-9]{4}-[0-9]{2}-[0-9]{2}")) return "";
            return s.Substring(8, 2) + "/" + s.Substring(5, 2) + "/" + s.Substring(0, 4);
        }

        static List<XmlElement> Filhos(XmlElement e, string n)
        {
            var r = new List<XmlElement>();
            foreach (XmlNode c in e.ChildNodes) { var el = c as XmlElement; if (el != null && el.LocalName == n) r.Add(el); }
            return r;
        }

        static List<XmlElement> Todos(XmlElement e, string n)
        {
            var r = new List<XmlElement>();
            foreach (XmlNode c in e.GetElementsByTagName(n, "*")) r.Add((XmlElement)c);
            if (r.Count == 0) foreach (XmlNode c in e.GetElementsByTagName(n)) r.Add((XmlElement)c);
            return r;
        }

        static XmlElement Caminho(XmlElement e, params string[] ns)
        {
            var c = e;
            foreach (var n in ns)
            {
                var f = Filhos(c, n);
                if (f.Count == 0) return null;
                c = f[0];
            }
            return c;
        }

        static string Texto(XmlElement e) { return e == null ? "" : Regex.Replace(e.InnerText, @"\s+", " ").Trim(); }
    }

    /// <summary>FEBRABAN CNAB 240 — extrato de conta corrente (serviço 04, segmento E).</summary>
    public static class LeitorCNAB240
    {
        public static LeitorExtrato.Resultado Ler(string t, string origem)
        {
            var r = new LeitorExtrato.Resultado { formato = "CNAB 240 (extrato)" };
            int lotes = 0, outros = 0;
            string banco = "";
            foreach (var l in t.SplitRe(@"\r?\n"))
            {
                if (l.Length < 240) continue;
                char tipo = l[7];
                if (tipo == '0') banco = l.Sub(102, 132).Trim();
                else if (tipo == '1')
                {
                    lotes++;
                    double? s = Saldo(l);
                    if (r.saldoInicial == null && s != null) r.saldoInicial = s;
                }
                else if (tipo == '5') { double? s = Saldo(l); if (s != null) r.saldoFinal = s; }
                else if (tipo == '3')
                {
                    if (l[13] != 'E') { outros++; continue; }
                    string dl = l.Sub(142, 150), dc = l.Sub(134, 142);
                    string d = dl.Matches("[0-9]{8}") && dl != "00000000" ? dl : dc;
                    if (!d.Matches("[0-9]{8}") || d == "00000000") continue;
                    string data = d.Substring(0, 2) + "/" + d.Substring(2, 2) + "/" + d.Substring(4, 4);
                    string sv = l.Sub(150, 168).Trim();
                    long n;
                    if (!long.TryParse(sv.Length == 0 ? "0" : sv, out n)) continue;
                    double v = n / 100.0;
                    if (v < 0.005) continue;
                    bool saida = l[168] == 'D';
                    string desc = l.Sub(176, 201).Trim(), docu = new Regex("^0+(?=[0-9])").Replace(l.Sub(201, 240).Trim(), "", 1), compl = l.Sub(113, 133).Trim();
                    var h = new StringBuilder(desc);
                    if (compl.Length > 0 && !compl.Matches("0+")) h.Append(" ").Append(compl);
                    if (docu.Length > 0 && !docu.Matches("0+")) h.Append(" DOC ").Append(docu);
                    r.lancamentos.Add(LeitorExtrato.Novo(data, h.ToString(), saida ? -v : v, origem));
                }
            }
            if (r.lancamentos.Count == 0 && outros > 0)
                throw new LeituraException(origem + ": CNAB 240 sem segmento E — é arquivo de cobrança/pagamento (retorno), não extrato. Peça ao banco o extrato CNAB 240 (serviço 04) ou OFX.");
            if (banco.Length > 0) r.formato += " - " + banco;
            if (lotes > 1) { r.aviso = "CNAB 240 com " + lotes + " lotes (contas) — prova de saldo pelo primeiro e último."; r.saldoInicial = null; r.saldoFinal = null; }
            return r;
        }

        static double? Saldo(string l)
        {
            string v = l.Sub(150, 168).Trim();
            long n;
            if (!v.Matches("[0-9]+") || !long.TryParse(v, out n)) return null;
            double x = n / 100.0;
            return l[168] == 'D' ? -x : x;
        }
    }
}

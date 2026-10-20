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
    /// Porta de entrada de qualquer extrato: identifica o formato pelo CONTEÚDO e chama o leitor certo.
    /// OFX/OFC · QIF · MT940/942 · CAMT.052/053/054 · CNAB 240 · xlsx/xls/ods/html/xml/csv · PDF (texto ou colunas) · layout mapeado.
    /// </summary>
    public static class LeitorExtrato
    {
        public class Resultado
        {
            public List<Lancamento> lancamentos = new List<Lancamento>();
            public string formato = "";
            public double? saldoInicial, saldoFinal;
            public string aviso = "";

            /// <summary>"" se não há saldos no arquivo; senão o texto da prova.</summary>
            public string Prova()
            {
                if (saldoInicial == null || saldoFinal == null) return "";
                double mov = lancamentos.Sum(l => l.saida ? -l.valor : l.valor);
                double calc = saldoInicial.Value + mov, dif = saldoFinal.Value - calc;
                if (Math.Abs(dif) < 0.01) return "prova do saldo confere (" + J.Moeda(saldoInicial.Value) + " + movimento " + J.Moeda(mov) + " = " + J.Moeda(saldoFinal.Value) + ")";
                return "PROVA DO SALDO NÃO CONFERE: inicial " + J.Moeda(saldoInicial.Value) + " + movimento " + J.Moeda(mov) + " = " + J.Moeda(calc)
                    + ", extrato diz " + J.Moeda(saldoFinal.Value) + " (diferença " + J.Moeda(dif) + ") — confira se faltou alguma linha";
            }
        }

        /// <summary>Formatos padrão que dispensam a pergunta "é extrato ou planilha?".</summary>
        public static string FormatoPadrao(string p)
        {
            try
            {
                var b = File.ReadAllBytes(p);
                string t = J.Texto(b.Length <= 4096 ? b : b.Take(4096).ToArray());
                if (Ofx(t)) return "OFX";
                if (Qif(t)) return "QIF";
                if (Mt940(t)) return "MT940";
                if (Camt(t)) return "CAMT";
                if (Cnab240(b)) return "CNAB240";
            }
            catch (IOException) { }
            return null;
        }

        public static Resultado Ler(string p, bool juntarContinuacao)
        {
            byte[] b = File.ReadAllBytes(p);
            string nome = Path.GetFileName(p);
            Resultado r;
            if (b.Length >= 4 && b[0] == '%' && b[1] == 'P' && b[2] == 'D' && b[3] == 'F')
            {
                r = LeitorPDFColunas.Melhor(p, juntarContinuacao);
            }
            else
            {
                string t = (b.Length >= 2 && b[0] == 0xD0 && b[1] == 0xCF) || (b.Length >= 2 && b[0] == 'P' && b[1] == 'K') ? "" : J.Texto(b);
                if (Ofx(t)) { r = new Resultado { formato = "OFX", lancamentos = LeitorOFX.Ler(p), saldoFinal = SaldoOFX(t) }; }
                else if (Qif(t)) r = LeitorQIF.Ler(t, nome);
                else if (Mt940(t)) r = LeitorMT940.Ler(t, nome);
                else if (Camt(t)) r = LeitorCAMT.Ler(b, nome);
                else if (Cnab240(b)) r = LeitorCNAB240.Ler(t, nome);
                else
                {
                    r = new Resultado();
                    var rows = LeitorTabular.Linhas(p);
                    r.formato = LeitorTabular.ultimoFormato ?? "";
                    var lay = LayoutsExtrato.Procurar(rows);
                    if (lay != null) { r.lancamentos = LayoutsExtrato.Aplicar(lay, rows, nome, r); r.formato += " (layout salvo: " + lay.nome + ")"; }
                    else r.lancamentos = LeitorExtratoExcel.Ler(p, rows, juntarContinuacao, r);
                }
            }
            foreach (var l in r.lancamentos) if (string.IsNullOrEmpty(l.origem)) l.origem = nome;
            if (r.lancamentos.Count == 0) throw new LeituraException(nome + ": nenhum lançamento reconhecido (" + r.formato + ").");
            return r;
        }

        static string Ini(string t) { return t.Length > 4000 ? t.Substring(0, 4000) : t; }

        public static bool Ofx(string t)
        {
            string u = Ini(t).ToUpperInvariant();
            return u.Contains("OFXHEADER") || u.Contains("<OFX>") || u.Contains("<OFC>") || u.Contains("<STMTTRN>");
        }

        public static bool Qif(string t)
        {
            string s = t.Replace("﻿", "").Trim();
            return s.StartsWith("!Type:", StringComparison.OrdinalIgnoreCase) || s.StartsWith("!Account", StringComparison.OrdinalIgnoreCase) || s.StartsWith("!Option", StringComparison.OrdinalIgnoreCase);
        }

        public static bool Mt940(string t)
        {
            return Regex.IsMatch(t, @"^:(20|25):", RegexOptions.Multiline) && Regex.IsMatch(t, @"^:(61|60[FM]|62[FM]|34F):", RegexOptions.Multiline);
        }

        public static bool Camt(string t)
        {
            string s = Ini(t);
            return s.Contains("BkToCstmrStmt") || s.Contains("BkToCstmrAcctRpt") || s.Contains("BkToCstmrDbtCdtNtfctn") || Regex.IsMatch(s, @"camt\.05[234]");
        }

        public static bool Cnab240(byte[] b)
        {
            string t = Encoding.GetEncoding("iso-8859-1").GetString(b);
            int ok = 0, n = 0;
            foreach (var l in t.SplitRe(@"\r?\n"))
            {
                if (l.Trim().Length == 0) continue;
                n++;
                if (l.Length == 240 && char.IsDigit(l[0]) && char.IsDigit(l[7])) ok++;
            }
            return n >= 3 && ok * 10 >= n * 9;
        }

        static double? SaldoOFX(string t)
        {
            double? r = null;
            foreach (Match m in Regex.Matches(t, @"<LEDGERBAL>.*?<BALAMT>\s*([-+]?[0-9.,]+)", RegexOptions.IgnoreCase | RegexOptions.Singleline)) r = Formatos.Valor(m.Groups[1].Value);
            return r;
        }

        public static Lancamento Novo(string data, string hist, double valorComSinal, string origem)
        {
            return new Lancamento(data, LeitorOFX.Limpar(hist.Length == 0 ? "(sem histórico)" : hist), Math.Abs(valorComSinal), valorComSinal < 0) { origem = origem };
        }
    }

    /// <summary>
    /// Layouts de extrato ensinados pelo usuário ("Mapear colunas"), em layouts_extrato.txt (pasta do programa):
    /// nome ; assinatura ; data ; históricos ; valor ; crédito ; débito ; D/C ; saldo ; documento ; ordem ; inverter
    /// </summary>
    public static class LayoutsExtrato
    {
        public class Layout
        {
            public string nome = "", assinatura = "";
            public int data, valor, cred, deb, dc, saldo, doc, ordem = -1;
            public List<int> hist = new List<int>();
            public bool inverter;
            internal int inicio;

            internal string Linha()
            {
                return string.Join(";", nome.Replace(";", ","), assinatura.Replace(";", ","), "" + data, string.Join(",", hist), "" + valor, "" + cred, "" + deb,
                    "" + dc, "" + saldo, "" + doc, ordem == Formatos.MDY ? "MDY" : ordem == Formatos.YMD ? "YMD" : ordem == Formatos.DMY ? "DMY" : "AUTO", inverter ? "S" : "N");
            }

            internal static Layout De(string l)
            {
                var c = l.Split(';');
                if (c.Length < 12) return null;
                var x = new Layout { nome = c[0], assinatura = c[1], data = I(c[2]) };
                foreach (var h in c[3].Split(',')) if (I(h) > 0) x.hist.Add(I(h));
                x.valor = I(c[4]); x.cred = I(c[5]); x.deb = I(c[6]); x.dc = I(c[7]); x.saldo = I(c[8]); x.doc = I(c[9]);
                x.ordem = c[10] == "MDY" ? Formatos.MDY : c[10] == "YMD" ? Formatos.YMD : c[10] == "DMY" ? Formatos.DMY : -1;
                x.inverter = c[11].Trim() == "S";
                return x.data > 0 && (x.valor > 0 || x.cred > 0 || x.deb > 0) ? x : null;
            }

            static int I(string s) { int v; return int.TryParse(s.Trim(), out v) ? v : 0; }
        }

        public static string Arquivo() { return Path.Combine(Log.PastaApp(), "layouts_extrato.txt"); }

        public static List<Layout> Todos()
        {
            var r = new List<Layout>();
            try
            {
                if (!File.Exists(Arquivo())) return r;
                foreach (var l in File.ReadAllLines(Arquivo(), Encoding.UTF8))
                {
                    if (l.Trim().Length == 0 || l.StartsWith("#")) continue;
                    var x = Layout.De(l);
                    if (x != null) r.Add(x);
                }
            }
            catch (IOException e) { Log.Erro("layouts_extrato", e); }
            return r;
        }

        public static void Salvar(Layout novo)
        {
            var ls = new List<string> { "# Layouts de extrato ensinados em \"Mapear colunas\" — nome;assinatura;data;históricos;valor;crédito;débito;D/C;saldo;documento;ordem;inverter" };
            foreach (var x in Todos()) if (x.assinatura != novo.assinatura) ls.Add(x.Linha());
            ls.Add(novo.Linha());
            Directory.CreateDirectory(Path.GetDirectoryName(Arquivo()));
            File.WriteAllLines(Arquivo(), ls, new UTF8Encoding(false));
        }

        public static string Assinatura(string[] cel)
        {
            var b = new StringBuilder();
            foreach (var c in cel)
            {
                string n = Formatos.Normalizar(c);
                if (n.Length == 0 && c != null) n = c.Trim().Replace("|", "/").Replace(";", ",");
                b.Append(b.Length > 0 ? "|" : "").Append(n);
            }
            return Regex.Replace(b.ToString(), @"\|+$", "");
        }

        public static Layout Procurar(List<string[]> rows)
        {
            var ls = Todos();
            if (ls.Count == 0) return null;
            foreach (var x in ls)
            {
                if (x.assinatura.StartsWith("#") || x.assinatura.Replace("|", "").Trim().Length == 0) continue;
                for (int i = 0; i < rows.Count && i < 200; i++)
                    if (Assinatura(rows[i]) == x.assinatura) { x.inicio = i + 1; return x; }
            }
            foreach (var x in ls)
            {
                if (!x.assinatura.StartsWith("#")) continue;
                int n;
                if (!int.TryParse(x.assinatura.Substring(1), out n)) continue;
                int ok = 0, total = 0, primeira = -1;
                for (int i = 0; i < rows.Count; i++)
                {
                    if (rows[i].Length != n) continue;
                    total++;
                    string d = x.data - 1 < rows[i].Length ? rows[i][x.data - 1] ?? "" : "";
                    if (Formatos.Data(d).Length > 0 || d.Matches(@"[0-9]{5}(\.[0-9]+)?")) { ok++; if (primeira < 0) primeira = i; }
                }
                if (total > 0 && ok * 10 >= total * 8) { x.inicio = primeira; return x; }
            }
            return null;
        }

        public static LeitorExtratoExcel.Mapa Mapa(Layout x)
        {
            var m = new LeitorExtratoExcel.Mapa { data = x.data - 1 };
            if (x.hist.Count > 0) { m.hist = x.hist[0] - 1; m.histForte = true; }
            if (x.hist.Count > 1) m.hist2 = x.hist[1] - 1;
            for (int i = 2; i < x.hist.Count; i++) m.extras.Add(x.hist[i] - 1);
            m.valor = x.valor - 1; m.cred = x.cred - 1; m.deb = x.deb - 1; m.dc = x.dc - 1; m.saldo = x.saldo - 1; m.doc = x.doc - 1;
            m.credBruto = m.cred; m.debBruto = m.deb;
            return m;
        }

        public static List<Lancamento> Aplicar(Layout x, List<string[]> rows, string origem, LeitorExtrato.Resultado res)
        {
            bool xlsx = LeitorExtratoExcel.EhPlanilha(LeitorTabular.ultimoFormato ?? "");
            var r = LeitorExtratoExcel.Processar(rows, Math.Max(0, x.inicio), Mapa(x), x.ordem, origem, xlsx, res, false);
            if (x.inverter) foreach (var l in r) l.saida = !l.saida;
            return r;
        }
    }
}

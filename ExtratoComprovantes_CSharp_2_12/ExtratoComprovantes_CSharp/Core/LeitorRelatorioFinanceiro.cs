using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Relatório "Contas e cartões" de sistema financeiro em PDF (Data | Terceiro | Documento | Histórico | Entrada | Saída | Saldo).
    /// Lê pela posição de cada texto na página, coluna por coluna.
    /// </summary>
    public static class LeitorRelatorioFinanceiro
    {
        public class Registro
        {
            public string data = "", terceiro = "", cnpj = "", documento = "", categoria = "", historico = "";
            public double valor;
            public bool saida;
            public double? saldo;
            internal List<string>[] partes;
        }

        static readonly Regex DATA = new Regex(@"^[0-9]{2}/[0-9]{2}/[0-9]{4}$");
        static readonly Regex DIN = new Regex(@"^-?\(?([0-9]{1,3}(?:\.[0-9]{3})*|[0-9]+),([0-9]{2})\)?\s*[DC]?$");
        static readonly Regex DOC = new Regex(@"([0-9]{2}\.?[0-9]{3}\.?[0-9]{3}/?[0-9]{4}-?\s?[0-9]{2}|[0-9]{3}\.?[0-9]{3}\.?[0-9]{3}-?\s?[0-9]{2})");
        static readonly Regex CATEGORIA = new Regex(@"^[0-9]{1,2}(?:\.[0-9]{1,3})*\s+-\s+\S");
        static readonly string[] COLUNAS = { "DATA", "TERCEIRO", "DOCUMENTO", "HISTORICO", "ENTRADA", "SAIDA", "SALDO" };

        public static List<Registro> Ler(string arq)
        {
            var r = new List<Registro>();
            Registro atual = null;
            float[] lim = null;
            float yCab = -1;
            foreach (var pagina in TextoPDF.Trechos(arq))
            {
                var ts = new List<TextoPDF.Trecho>(pagina);
                ts.OrdenarEstavel((a, b) => Math.Abs(a.y - b.y) > 0.5f ? a.y.CompareTo(b.y) : a.x.CompareTo(b.x));
                float[] l2 = Colunas(ts);
                if (l2 != null) { lim = new float[l2.Length - 1]; Array.Copy(l2, lim, lim.Length); yCab = l2[l2.Length - 1]; }
                else yCab = -1;
                if (lim == null) continue;
                var vistos = new HashSet<string>();
                var ySaldo = new List<float>();
                foreach (var t in ts) { string n0 = Regras.Normalizar(t.Texto()); if (n0.StartsWith("SALDO EM") || n0.StartsWith("TOTAL")) ySaldo.Add(t.y); }
                foreach (var t in ts)
                {
                    if (ySaldo.Any(ys => Math.Abs(t.y - ys) < 3)) { atual = null; continue; }
                    if (yCab >= 0 && t.y <= yCab + 6) continue;
                    string s = t.Texto();
                    if (s.Length == 0) continue;
                    if (!vistos.Add(TextoPDF.Rnd(t.x) + "|" + TextoPDF.Rnd(t.y) + "|" + s)) continue;
                    string n = Regras.Normalizar(s);
                    if (n.StartsWith("SALDO EM") || n.StartsWith("TOTAL")) { atual = null; continue; }
                    int col = DIN.IsMatch(s) ? Coluna(lim, t.fim - 2) : Coluna(lim, t.x);
                    if (col == 0)
                    {
                        if (DATA.IsMatch(s))
                        {
                            atual = new Registro { data = s, partes = new List<string>[7] };
                            for (int i = 0; i < 7; i++) atual.partes[i] = new List<string>();
                            r.Add(atual);
                        }
                        continue;
                    }
                    if (atual == null || col < 0) continue;
                    if (col == 1 && t.x < lim[0] + 1) continue;
                    atual.partes[col].Add(s);
                }
            }
            if (lim == null) throw new LeituraException(Path.GetFileName(arq) + ": cabeçalho Data / Terceiro / Documento / Histórico / Entrada / Saída / Saldo não encontrado.");
            var ok = new List<Registro>();
            foreach (var g in r)
            {
                var p = g.partes;
                var ter = new StringBuilder();
                foreach (var s in p[1])
                {
                    if (s.Matches("[A-Z]{1,3}") && p[1].Count > 1) continue;
                    char u = ter.Length > 0 ? ter[ter.Length - 1] : ' ';
                    bool cola = ter.Length > 0 && (u == '-' || u == '/' || u == '.') && Regex.IsMatch(s, "^[0-9]");
                    ter.Append(cola || ter.Length == 0 ? "" : " ").Append(s);
                }
                string t = ter.ToString();
                var md = DOC.Match(t);
                if (md.Success) { g.cnpj = md.Groups[1].Value.Replace(" ", ""); t = (t.Substring(0, md.Index) + " " + t.Substring(md.Index + md.Length)).Trim(); }
                g.terceiro = Regex.Replace(t, @"\s{2,}", " ").Trim();
                StringBuilder doc = new StringBuilder(), cat = new StringBuilder();
                foreach (var s in p[2])
                {
                    if (cat.Length == 0 && !CATEGORIA.IsMatch(s)) doc.Append(doc.Length > 0 ? " " : "").Append(s);
                    else cat.Append(cat.Length > 0 ? " " : "").Append(s);
                }
                g.documento = doc.ToString().Trim();
                g.categoria = cat.ToString().Trim();
                g.historico = Regex.Replace(string.Join(" ", p[3]), @"\s{2,}", " ").Trim();
                double? e = Primeiro(p[4]), sa = Primeiro(p[5]);
                g.saldo = UltimoAssinado(p[6]);
                if (e != null && Math.Abs(e.Value) > 0.004) { g.valor = Math.Abs(e.Value); g.saida = false; }
                else if (sa != null && Math.Abs(sa.Value) > 0.004) { g.valor = Math.Abs(sa.Value); g.saida = true; }
                else continue;
                ok.Add(g);
            }
            if (ok.Count == 0) throw new LeituraException(Path.GetFileName(arq) + ": nenhum lançamento reconhecido no relatório financeiro.");
            return ok;
        }

        static float[] Colunas(List<TextoPDF.Trecho> ts)
        {
            foreach (var a in ts)
            {
                if (Regras.Normalizar(a.Texto()) != "DATA") continue;
                float[] ini = new float[7], fim = new float[7];
                for (int i = 0; i < 7; i++) ini[i] = -1;
                foreach (var b in ts)
                {
                    if (Math.Abs(b.y - a.y) > 3) continue;
                    string n = Regras.Normalizar(b.Texto());
                    for (int i = 0; i < 7; i++) if (n == COLUNAS[i]) { ini[i] = b.x; fim[i] = b.fim; }
                }
                if (ini.Any(v => v < 0)) continue;
                var lim = new float[8];
                for (int i = 0; i < 7; i++)
                {
                    float prox = i < 6 ? ini[i + 1] : float.MaxValue;
                    float sep = -1;
                    foreach (var b in ts)
                    {
                        if (Math.Abs(b.y - a.y) > 6 || b.Texto().Length > 2) continue;
                        if (b.x > fim[i] && b.x < prox && (sep < 0 || b.x < sep)) sep = b.fim;
                    }
                    lim[i] = sep > 0 ? sep : (i < 6 ? (fim[i] + ini[i + 1]) / 2 : float.MaxValue);
                }
                lim[7] = a.y;
                return lim;
            }
            return null;
        }

        static int Coluna(float[] lim, float x)
        {
            for (int i = 0; i < lim.Length; i++) if (x < lim[i] - 1) return i;
            return -1;
        }

        static double? Primeiro(List<string> l) { foreach (var s in l) { var v = Num(s); if (v != null) return v; } return null; }

        static double? UltimoAssinado(List<string> l) { double? r = null; foreach (var s in l) { var v = Num(s); if (v != null) r = v; } return r; }

        static double? Num(string s)
        {
            s = s.Trim();
            if (!DIN.IsMatch(s)) return null;
            bool neg = s.StartsWith("-") || s.StartsWith("(");
            string t = Regex.Replace(s, "[^0-9,]", "").Replace(',', '.');
            double v;
            if (!double.TryParse(t, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out v)) return null;
            return neg ? -v : v;
        }

        public static List<ItemPlanilha> Itens(string arq)
        {
            var r = new List<ItemPlanilha>();
            int i = 0;
            foreach (var g in Ler(arq))
            {
                var it = new ItemPlanilha { arquivo = Path.GetFileName(arq), linha = ++i, valor = g.valor, saida = g.saida, favorecido = g.terceiro, documento = g.cnpj, descricao = (g.documento + " " + g.historico).Trim(), categoria = g.categoria };
                it.nfs.AddRange(Conciliador.NumerosDaCelula(g.documento));
                it.datas.Add(g.data);
                r.Add(it);
            }
            return r;
        }

        public static List<Lancamento> Lancamentos(string arq)
        {
            var r = new List<Lancamento>();
            foreach (var g in Ler(arq))
            {
                var h = new StringBuilder(g.terceiro);
                if (g.historico.Length > 0) h.Append(h.Length > 0 ? " - " : "").Append(g.historico);
                if (g.cnpj.Length > 0) h.Append(" ").Append(g.cnpj);
                r.Add(new Lancamento(g.data, LeitorOFX.Limpar(h.ToString()), g.valor, g.saida) { origem = Path.GetFileName(arq) });
            }
            return r;
        }
    }
}

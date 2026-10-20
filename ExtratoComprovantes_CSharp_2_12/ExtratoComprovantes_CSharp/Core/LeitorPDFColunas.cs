using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Extrato em PDF lido como TABELA, pela posição de cada texto (cabeçalho Data / Histórico / Valor | Débito + Crédito / Saldo).
    /// Melhor(): lê das duas formas (linhas de texto e colunas) e fica com a que fecha melhor a cadeia de saldo.
    /// </summary>
    public static class LeitorPDFColunas
    {
        class Linha
        {
            public float y, h;
            public readonly List<TextoPDF.Trecho> ts = new List<TextoPDF.Trecho>();
        }

        class Col
        {
            public string papel;
            public float x, fim;
            public bool Num() { return papel == "valor" || papel == "cred" || papel == "deb" || papel == "saldo"; }
        }

        class Pend { public float y, h; public string texto; }

        public static LeitorExtrato.Resultado Melhor(string p, bool juntar)
        {
            string nome = Path.GetFileName(p);
            if (LeitorPDF.PareceRelatorioFinanceiro(p))
            {
                var r = new LeitorExtrato.Resultado { formato = "PDF relatório financeiro (colunas)" };
                r.lancamentos = LeitorPDF.Ler(p, juntar);
                try
                {
                    var c = Ler(p);
                    if (c.lancamentos.Count == r.lancamentos.Count && MesmoTotal(c.lancamentos, r.lancamentos)) { r.saldoInicial = c.saldoInicial; r.saldoFinal = c.saldoFinal; }
                }
                catch (Exception) { }
                return r;
            }
            List<Lancamento> texto = null;
            IOException erroTexto = null;
            try { texto = LeitorPDF.Ler(p, juntar); } catch (IOException e) { erroTexto = e; }
            LeitorExtrato.Resultado col = null;
            try { col = Ler(p); } catch (Exception e) { Log.Erro("PDF colunas " + nome, e); }
            if (texto == null && (col == null || col.lancamentos.Count == 0))
            {
                if (erroTexto != null) throw erroTexto;
                throw new LeituraException(nome + ": nenhum lançamento reconhecido.");
            }
            var rt = new LeitorExtrato.Resultado { formato = "PDF (linhas de texto)", lancamentos = texto ?? new List<Lancamento>() };
            if (col == null || col.lancamentos.Count == 0) return rt;
            if (rt.lancamentos.Count == 0) return col;
            int nt = rt.lancamentos.Count, nc = col.lancamentos.Count;
            int dt = Duvidosos(rt.lancamentos), dcol = Duvidosos(col.lancamentos);
            string prova = col.Prova();
            bool colFecha = prova.Length > 0 && !prova.StartsWith("PROVA DO SALDO NÃO");
            if (colFecha && nc >= nt) return col;
            if (nc > nt * 1.3 && dcol <= dt) return col;
            if (nc >= nt && dcol < dt) return col;
            if (colFecha && nc >= nt * 0.95 && MesmoTotal(rt.lancamentos, col.lancamentos)) return col;
            return rt;
        }

        static int Duvidosos(List<Lancamento> l) { return l.Count(x => x.conferirSinal); }

        static bool MesmoTotal(List<Lancamento> a, List<Lancamento> b)
        {
            double sa = a.Sum(x => x.saida ? -x.valor : x.valor), sb = b.Sum(x => x.saida ? -x.valor : x.valor);
            return Math.Abs(sa - sb) < 0.02;
        }

        static readonly Regex ANO = new Regex(@"\b[0-9]{1,2}[/.\-][0-9]{1,2}[/.\-](20[0-9]{2})\b");

        public static LeitorExtrato.Resultado Ler(string p)
        {
            string nome = Path.GetFileName(p);
            var res = new LeitorExtrato.Resultado { formato = "PDF (tabela por colunas)" };
            List<Col> cols = null;
            var r = new List<Lancamento>();
            var saldos = new List<double?>();
            var semSinal = new List<bool>();
            string ano = null;
            var paginas = TextoPDF.Trechos(p);
            foreach (var pg in paginas)
            {
                foreach (var t in pg) { var m = ANO.Match(t.Texto()); if (m.Success) { ano = m.Groups[1].Value; break; } }
                if (ano != null) break;
            }
            var todasDatas = new List<string>();
            foreach (var pg in paginas) foreach (var t in pg) if (t.Texto().Matches(@"[0-9]{1,2}[/.\-][0-9]{1,2}([/.\-][0-9]{2,4})?")) todasDatas.Add(t.Texto());
            int ordem = Formatos.Ordem(todasDatas);
            bool negValor = false;
            var viaValor = new List<bool>();
            string ultimaData = null;
            Lancamento ultimo = null;
            float yUlt = -1;
            bool ignorar = false;
            var pend = new List<Pend>();
            foreach (var pg in paginas)
            {
                var linhas = Linhas(pg);
                Descarregar(pend, ultimo, yUlt);
                ultimo = null;
                foreach (var ln in linhas)
                {
                    var c2 = Cabecalho(ln);
                    if (c2 != null) { Descarregar(pend, ultimo, yUlt); cols = c2; ultimo = null; continue; }
                    if (cols == null) continue;
                    var txt = new Dictionary<string, StringBuilder>();
                    var num = new Dictionary<string, double?>();
                    var numTxt = new Dictionary<string, string>();
                    string data = "";
                    foreach (var t in ln.ts)
                    {
                        string s = t.Texto();
                        if (s.Length == 0) continue;
                        if (Formatos.PareceDinheiro(s))
                        {
                            var cn = MaisPerto(cols, t, true);
                            if (cn != null && !num.ContainsKey(cn.papel)) { num[cn.papel] = Formatos.Valor(s); numTxt[cn.papel] = s; continue; }
                        }
                        var c = MaisPerto(cols, t, false);
                        string papel = c == null ? "hist" : c.papel;
                        if (papel == "data" && data.Length == 0)
                        {
                            var partes = Regex.Split(s, @"\s+");
                            string p0 = partes[0], p1 = partes.Length > 1 ? s.Substring(s.IndexOf(p0, StringComparison.Ordinal) + p0.Length).Trim() : null;
                            string d1 = Formatos.Data(p0, ordem, ano);
                            if (d1.Length > 0 && Regex.IsMatch(p0, "[0-9]")) { data = d1; if (!string.IsNullOrEmpty(p1)) App(txt, "hist", p1); continue; }
                            string d = Formatos.Data(s, ordem, ano);
                            if (d.Length > 0 && s.Length <= 24) { data = d; continue; }
                            papel = "hist";
                        }
                        if (papel.Matches("data|valor|cred|deb|saldo")) papel = "hist";
                        App(txt, papel, s);
                    }
                    string hist = LeitorOFX.Limpar(Str(txt, "hist") + (Str(txt, "doc").Length == 0 ? "" : " " + Str(txt, "doc")));
                    string nh = Regras.Normalizar(hist);
                    double? sl = num.GetOr("saldo", null);
                    bool temMov = num.ContainsKey("valor") || num.ContainsKey("cred") || num.ContainsKey("deb");
                    float xHist = float.MaxValue;
                    foreach (var c in cols) if (c.papel == "hist") xHist = Math.Min(xHist, c.x);
                    if (!temMov && sl == null && data.Length == 0 && ln.ts.Count > 0 && ln.ts[0].x < xHist - 15 && hist.Length <= 60)
                    {
                        if (nh.Matches(".*(FUTUR|AGENDAD|PROGRAMAD|PREVIST|A VENCER|SCHEDULED|UPCOMING|PENDING|PENDENTE|PROXIMOS LANC|PROXIMO DIA|BLOQUEAD).*")) { Descarregar(pend, ultimo, yUlt); ignorar = true; ultimo = null; continue; }
                        if (nh.Matches(".*(LANCAMENTOS|MOVIMENTA|EXTRATO|TRANSAC|TRANSACTIONS|STATEMENT|ACTIVITY|MOVIMIENTOS|UMSATZE|OPERATIONS).*")) ignorar = false;
                    }
                    if (ignorar) continue;
                    bool linhaSaldo = nh.Matches(@".*\b(SALDO|BALANCE|SOLDE|KONTOSTAND|TOTAL|SUBTOTAL|BROUGHT FORWARD|CARRIED FORWARD)\b.*") && !nh.Matches(".*(APLIC|RESGATE).*")
                        && (hist.Length < 60 || !temMov);
                    if (linhaSaldo)
                    {
                        double? sv = sl ?? num.GetOr("valor", null);
                        if (sv != null && r.Count == 0 && nh.Matches(".*(ANTERIOR|INICIAL|OPENING|PREVIOUS|BROUGHT|ANFANG|PRECEDENT).*")) res.saldoInicial = sv;
                        if (sv != null && nh.Matches(".*(SALDO FINAL|CLOSING|CARRIED|ENDSALDO|SALDO ATUAL|SALDO DISPONIVEL).*")) res.saldoFinal = sv;
                        if (data.Length > 0) ultimaData = data;
                        Descarregar(pend, ultimo, yUlt);
                        ultimo = null;
                        continue;
                    }
                    if (!temMov)
                    {
                        if (data.Length > 0) ultimaData = data;
                        if (hist.Length > 0 && !EhRodape(nh) && pend.Count < 4) pend.Add(new Pend { y = ln.y, texto = hist, h = ln.h });
                        else if (sl != null && ultimo != null && saldos[saldos.Count - 1] == null && ln.y - yUlt < ln.h * 2.6f) saldos[saldos.Count - 1] = sl;
                        else if (EhRodape(nh)) { Descarregar(pend, ultimo, yUlt); ultimo = null; }
                        continue;
                    }
                    double? v = null;
                    bool saida = false, explicito = false;
                    if (num.ContainsKey("cred") && num["cred"] != null && Math.Abs(num["cred"].Value) > 0.004) { v = num["cred"]; saida = v < 0; explicito = true; }
                    else if (num.ContainsKey("deb") && num["deb"] != null && Math.Abs(num["deb"].Value) > 0.004) { v = num["deb"]; saida = true; explicito = true; }
                    else if (num.ContainsKey("valor"))
                    {
                        v = num["valor"];
                        string sv = numTxt["valor"];
                        saida = v != null && v < 0;
                        if (saida) negValor = true;
                        explicito = saida || Regex.IsMatch(sv, @"^(?:.*(\+|\bC|CR)\s*$|^\+.*)$", RegexOptions.IgnoreCase);
                    }
                    if (v == null || Math.Abs(v.Value) < 0.005) { Descarregar(pend, ultimo, yUlt); ultimo = null; continue; }
                    string dcTxt = Str(txt, "dc");
                    bool? sdc = LeitorExtratoExcel.Sentido(dcTxt);
                    if (sdc != null) { saida = sdc.Value; explicito = true; }
                    else if (dcTxt.Length > 0) hist = LeitorOFX.Limpar(dcTxt + " " + hist);
                    var antes = new StringBuilder();
                    foreach (var pd in pend)
                    {
                        bool doAnterior = ultimo != null && (pd.y - yUlt) < (ln.y - pd.y) && pd.y - yUlt < pd.h * 2.6f;
                        if (doAnterior) ultimo.historico = LeitorOFX.Limpar(ultimo.historico + " " + pd.texto);
                        else if (ln.y - pd.y < pd.h * 2.6f) antes.Append(pd.texto).Append(' ');
                    }
                    pend.Clear();
                    if (antes.Length > 0) hist = LeitorOFX.Limpar(antes + hist);
                    if (data.Length == 0) data = ultimaData ?? "";
                    if (data.Length == 0) { ultimo = null; continue; }
                    ultimaData = data;
                    var l = new Lancamento(data, hist.Length == 0 ? "(sem histórico)" : hist, Math.Abs(v.Value), saida) { origem = nome };
                    r.Add(l);
                    saldos.Add(sl);
                    semSinal.Add(!explicito);
                    viaValor.Add(!num.ContainsKey("cred") && !num.ContainsKey("deb"));
                    ultimo = l;
                    yUlt = ln.y;
                }
            }
            Descarregar(pend, ultimo, yUlt);
            if (negValor) for (int i = 0; i < semSinal.Count; i++) if (viaValor[i]) semSinal[i] = false;
            bool cadeia = LeitorExtratoExcel.CadeiaDeSaldo(r, saldos, semSinal, res);
            if (!cadeia && r.Count > 0 && !semSinal.Contains(false))
                foreach (var l in r) { l.conferirSinal = true; l.obs = "Sinal presumido (entrada)"; }
            res.lancamentos = r;
            return res;
        }

        static void Descarregar(List<Pend> pend, Lancamento ultimo, float yUlt)
        {
            if (ultimo != null)
            {
                float yRef = yUlt;
                foreach (var pd in pend)
                    if (pd.y - yRef < pd.h * 2.6f && pd.y > yRef) { ultimo.historico = LeitorOFX.Limpar(ultimo.historico + " " + pd.texto); yRef = pd.y; }
            }
            pend.Clear();
        }

        static bool EhRodape(string n)
        {
            return n.Matches(@".*(PAGINA|PAGE|SEITE|PAG\.? [0-9]|OUVIDORIA|SAC |FALE CONOSCO|EMITIDO|GERADO EM|CENTRAL DE ATENDIMENTO|WWW\.|HTTP).*");
        }

        static void App(Dictionary<string, StringBuilder> m, string k, string s)
        {
            var b = m.ObterOuCriar(k);
            if (b.Length > 0) b.Append(' ');
            b.Append(s);
        }

        static string Str(Dictionary<string, StringBuilder> m, string k) { StringBuilder b; return m.TryGetValue(k, out b) ? b.ToString().Trim() : ""; }

        static Col MaisPerto(List<Col> cols, TextoPDF.Trecho t, bool numero)
        {
            Col melhor = null;
            float dist = float.MaxValue;
            float ct = (t.x + t.fim) / 2;
            foreach (var c in cols)
            {
                if (numero != c.Num()) continue;
                float cc = (c.x + c.fim) / 2;
                float d = numero ? Math.Min(Math.Abs(t.fim - c.fim), Math.Min(Math.Abs(ct - cc), Math.Abs(t.x - c.x)))
                    : (t.x >= c.x - 4 && t.x <= c.fim + 4 ? 0 : Math.Min(Math.Abs(t.x - c.x), Math.Abs(ct - cc)));
                if (d < dist) { dist = d; melhor = c; }
            }
            if (!numero)
            {
                Col esq = null;
                foreach (var c in cols) if (c.x <= t.x + 4 && (esq == null || c.x > esq.x)) esq = c;
                if (esq != null && !esq.Num()) return esq;
                if (esq != null) return null;
                return melhor;
            }
            return dist < 90 ? melhor : null;
        }

        static List<Col> Cabecalho(Linha ln)
        {
            var ts = new List<TextoPDF.Trecho>(ln.ts);
            ts.OrdenarEstavel((a, b) => a.x.CompareTo(b.x));
            var cel = ts.Select(t => t.Texto()).ToArray();
            var m = LeitorExtratoExcel.Cabecalho(cel);
            if (!m.Ok() || !LeitorExtratoExcel.SemDado(cel)) return null;
            if (m.hist < 0 && m.doc < 0 && m.extras.Count == 0) return null;
            var r = new List<Col>();
            Add(r, ts, m.data, "data");
            Add(r, ts, m.hist, "hist");
            Add(r, ts, m.hist2, "hist");
            Add(r, ts, m.doc, "doc");
            foreach (int e in m.extras) Add(r, ts, e, "hist");
            Add(r, ts, m.valor, "valor");
            Add(r, ts, m.cred, "cred");
            Add(r, ts, m.deb, "deb");
            Add(r, ts, m.saldo, "saldo");
            Add(r, ts, m.dc, "dc");
            return r;
        }

        static void Add(List<Col> r, List<TextoPDF.Trecho> ts, int i, string papel)
        {
            if (i < 0 || i >= ts.Count) return;
            r.Add(new Col { papel = papel, x = ts[i].x, fim = ts[i].fim });
        }

        static List<Linha> Linhas(List<TextoPDF.Trecho> pg)
        {
            var ts = new List<TextoPDF.Trecho>(pg);
            ts.OrdenarEstavel((a, b) => a.y.CompareTo(b.y));
            var r = new List<Linha>();
            var vistos = new HashSet<string>();
            Linha L = null;
            foreach (var t in ts)
            {
                if (t.Texto().Length == 0) continue;
                if (!vistos.Add(TextoPDF.Rnd(t.x) + "|" + TextoPDF.Rnd(t.y) + "|" + t.Texto())) continue;
                float tol = Math.Max(1.5f, Math.Min(t.h, 12) * 0.45f);
                if (L != null && Math.Abs(t.y - L.y) <= tol) { L.ts.Add(t); L.h = Math.Max(L.h, t.h); continue; }
                L = new Linha { y = t.y, h = Math.Max(t.h, 6) };
                L.ts.Add(t);
                r.Add(L);
            }
            foreach (var l in r) l.ts.OrdenarEstavel((a, b) => a.x.CompareTo(b.x));
            return r;
        }
    }
}

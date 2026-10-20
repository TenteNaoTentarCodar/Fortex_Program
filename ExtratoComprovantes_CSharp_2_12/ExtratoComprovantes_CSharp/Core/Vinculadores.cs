using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>Liga cada comprovante a uma linha do extrato: mesmo valor e data mais próxima (até N dias).</summary>
    public static class Vinculador
    {
        public static int Vincular(List<Lancamento> extrato, List<Comprovante> comps, int toleranciaDias)
        {
            foreach (var l in extrato) l.comprovante = null;
            foreach (var c in comps) c.vinculo = null;
            var pares = new List<Tuple<Comprovante, Lancamento, long>>();
            foreach (var c in comps)
            {
                var dc = J.Data(c.data);
                if (c.valor <= 0 || c.exportado.Length > 0) continue;
                foreach (var l in extrato)
                {
                    if (Math.Abs(l.valor - c.valor) > 0.005) continue;
                    var dl = J.Data(l.data);
                    long dif = (dc == null || dl == null) ? toleranciaDias : Math.Abs(J.Dias(dc.Value, dl.Value));
                    if (dif > toleranciaDias) continue;
                    pares.Add(Tuple.Create(c, l, dif * 2 + (l.saida ? 0 : 1)));
                }
            }
            pares.OrdenarEstavel((a, b) => a.Item3.CompareTo(b.Item3));
            int n = 0;
            foreach (var p in pares)
            {
                if (p.Item1.vinculo != null || p.Item2.comprovante != null) continue;
                p.Item1.vinculo = p.Item2;
                p.Item2.comprovante = p.Item1;
                n++;
            }
            return n;
        }
    }

    /// <summary>
    /// Confere o extrato com as notas fiscais (XML): título de mesmo valor, sentido compatível (compra = saída, venda = entrada) e data próxima.
    /// Depois tenta vários títulos do mesmo participante somando um único lançamento.
    /// </summary>
    public static class VinculadorNotas
    {
        public static void Vincular(List<Lancamento> extrato, List<NotaFiscal> notas, string cnpjEmpresa, int tolVenc, int prazo)
        {
            foreach (var l in extrato) l.notas.Clear();
            var tits = new List<NotaFiscal.Titulo>();
            foreach (var nf in notas) foreach (var t in nf.titulos) { t.vinculo = null; if (!nf.cancelada && t.exportado.Length == 0) tits.Add(t); }
            var tokens = new Dictionary<NotaFiscal, HashSet<string>>();
            foreach (var nf in notas) tokens[nf] = Classificador.TokSet(nf.ContraparteNome(cnpjEmpresa));

            var pares = new List<Tuple<NotaFiscal.Titulo, Lancamento, long>>();
            foreach (var t in tits)
            {
                bool? compra = t.nf.Compra(cnpjEmpresa);
                foreach (var l in extrato)
                {
                    if (l.descartada) continue;
                    if (compra != null && compra != l.saida) continue;
                    int alt = Math.Abs(l.valor - t.valor) <= 0.005 ? 0 : -1;
                    if (alt < 0) foreach (var a in t.alternativos) if (Math.Abs(l.valor - a) <= 0.005) { alt = 1; break; }
                    if (alt < 0) continue;
                    bool nfPlan = NfNaPlanilha(l, t.nf);
                    bool nome = CitaParticipante(l, t.nf, tokens[t.nf], cnpjEmpresa) || nfPlan;
                    long? dif = Diferenca(t, l, nome, tolVenc, prazo);
                    if (dif == null) continue;
                    long score = (nome ? 0 : 100000) + (nfPlan ? 0 : 50000) + dif.Value * 4 + alt * 2 + (compra == null ? 1 : 0);
                    pares.Add(Tuple.Create(t, l, score));
                }
            }
            pares.OrdenarEstavel((a, b) => a.Item3.CompareTo(b.Item3));
            foreach (var p in pares)
            {
                if (p.Item1.vinculo != null || p.Item2.notas.Count > 0) continue;
                p.Item1.vinculo = p.Item2;
                p.Item2.notas.Add(p.Item1);
            }
            foreach (var l in extrato)
            {
                if (l.notas.Count > 0 || l.descartada) continue;
                var grupos = new Dictionary<string, List<NotaFiscal.Titulo>>();
                var ordem = new List<string>();
                foreach (var t in tits)
                {
                    if (t.vinculo != null || t.valor >= l.valor - 0.005) continue;
                    bool? compra = t.nf.Compra(cnpjEmpresa);
                    if (compra != null && compra != l.saida) continue;
                    bool nome = CitaParticipante(l, t.nf, tokens[t.nf], cnpjEmpresa) || NfNaPlanilha(l, t.nf);
                    if (Diferenca(t, l, nome, tolVenc, prazo) == null) continue;
                    string k = t.nf.ContraparteCnpj(cnpjEmpresa);
                    if (k.Length == 0) continue;
                    if (!grupos.ContainsKey(k)) { grupos[k] = new List<NotaFiscal.Titulo>(); ordem.Add(k); }
                    grupos[k].Add(t);
                }
                foreach (var k in ordem)
                {
                    var g = grupos[k];
                    if (g.Count < 2 || g.Count > 25) continue;
                    var sol = Soma(g, J.RoundJ(l.valor * 100));
                    if (sol != null) { foreach (var t in sol) { t.vinculo = l; l.notas.Add(t); } break; }
                }
            }
        }

        static long? Diferenca(NotaFiscal.Titulo t, Lancamento l, bool nome, int tolVenc, int prazo)
        {
            var dl = J.Data(l.data);
            if (dl == null) return null;
            if (t.vencimento != null)
            {
                long d0 = Math.Abs(J.Dias(t.vencimento.Value, dl.Value));
                if (d0 <= tolVenc) return d0;
                if (!nome) return null;
            }
            if (t.nf.emissao == null) return nome ? (long?)0 : null;
            long d = J.Dias(t.nf.emissao.Value, dl.Value);
            if (d < -5) return null;
            if (d > (nome ? prazo : 10)) return null;
            return Math.Abs(d) + (t.vencimento != null ? 30 : 0);
        }

        /// <summary>A planilha financeira vinculada à linha cita o número desta nota.</summary>
        static bool NfNaPlanilha(Lancamento l, NotaFiscal nf)
        {
            if (l.planilha.Count == 0) return false;
            string n = Conciliador.NumNF(nf.numero);
            return n.Length > 0 && l.planilha.Any(it => it.nfs.Contains(n));
        }

        public static bool CitaParticipante(Lancamento l, NotaFiscal nf, HashSet<string> tkNome, string cnpjEmpresa)
        {
            string h = l.historico;
            string digitos = Regex.Replace(h, "[^0-9]", "");
            string cnpj = nf.ContraparteCnpj(cnpjEmpresa);
            if (cnpj.Length == 14 && digitos.Contains(cnpj.Substring(0, 8))) return true;
            if (cnpj.Length == 11 && digitos.Contains(cnpj)) return true;
            if (tkNome == null || tkNome.Count == 0) return false;
            var th = Classificador.TokSet(h);
            int comum = tkNome.Count(t => t.Length >= 3 && th.Contains(t));
            return comum >= Math.Min(2, tkNome.Count);
        }

        static List<NotaFiscal.Titulo> Soma(List<NotaFiscal.Titulo> g, long alvo)
        {
            long tot = g.Sum(t => J.RoundJ(t.valor * 100));
            if (tot == alvo) return new List<NotaFiscal.Titulo>(g);
            var ord = new List<NotaFiscal.Titulo>(g);
            ord.OrdenarEstavel((a, b) => b.valor.CompareTo(a.valor));
            return Busca(ord, 0, alvo, new List<NotaFiscal.Titulo>());
        }

        static List<NotaFiscal.Titulo> Busca(List<NotaFiscal.Titulo> l, int i, long resto, List<NotaFiscal.Titulo> atual)
        {
            if (resto == 0 && atual.Count >= 2) return new List<NotaFiscal.Titulo>(atual);
            if (resto <= 0 || i >= l.Count || atual.Count >= 8) return null;
            for (int j = i; j < l.Count; j++)
            {
                long v = J.RoundJ(l[j].valor * 100);
                if (v > resto) continue;
                atual.Add(l[j]);
                var r = Busca(l, j + 1, resto - v, atual);
                atual.RemoveAt(atual.Count - 1);
                if (r != null) return r;
            }
            return null;
        }

        /// <summary>CNPJ da empresa: o que mais aparece como emitente ou destinatário nos XMLs.</summary>
        public static string DetectarEmpresa(List<NotaFiscal> notas)
        {
            var c = new Dictionary<string, int>();
            var ordem = new List<string>();
            foreach (var nf in notas)
            {
                if (nf.cancelada && nf.tipo == "Evento") continue;
                var s = new HashSet<string>();
                if (nf.emitCnpj.Length > 0) s.Add(nf.emitCnpj);
                if (nf.destCnpj.Length > 0) s.Add(nf.destCnpj);
                foreach (var x in s) { if (!c.ContainsKey(x)) ordem.Add(x); c.Somar(x, 1); }
            }
            string melhor = "";
            int n = 0;
            foreach (var k in J.OrdemHashJava(ordem, J.HashJava)) if (c[k] > n) { n = c[k]; melhor = k; }
            return melhor;
        }
    }

    /// <summary>
    /// Concilia a planilha financeira com o extrato: um para um (mesmo valor e data próxima), recebimento líquido de gateway,
    /// e agrupado (vários itens do mesmo favorecido somando uma linha do extrato). A quantidade de lançamentos segue o extrato.
    /// </summary>
    public static class VinculadorPlanilha
    {
        static bool Eh(bool? b, bool v) { return b != null && b.Value == v; }

        public static void Vincular(List<Lancamento> extrato, List<ItemPlanilha> itens, int tol)
        {
            foreach (var l in extrato) l.planilha.Clear();
            foreach (var it in itens) it.vinculo = null;

            var pares = new List<Tuple<ItemPlanilha, Lancamento, long>>();
            foreach (var it in itens)
            {
                if (it.exportado.Length > 0) continue;
                foreach (var l in extrato)
                {
                    if (l.descartada) continue;
                    if (it.saida != null && it.saida != l.saida) continue;
                    int alt = Math.Abs(l.valor - it.valor) <= 0.005 ? 0 : -1;
                    if (alt < 0) foreach (var a in it.alternativos) if (Math.Abs(l.valor - a) <= 0.005) { alt = 1; break; }
                    if (alt < 0 && l.valorBruto > 0 && Math.Abs(l.valorBruto - it.valor) <= 0.005) alt = 1;
                    if (alt < 0) continue;
                    long dif = Dif(it, l);
                    if (dif > tol) continue;
                    // empate (mesma distância de data): prefere o lançamento cuja descrição/NF bate com o item
                    pares.Add(Tuple.Create(it, l, (dif * 4 + alt * 2 + (it.saida == null ? 1 : 0)) * 4 + (3 - Conciliador.Afinidade(it, l))));
                }
            }
            pares.OrdenarEstavel((a, b) => a.Item3.CompareTo(b.Item3));
            foreach (var p in pares)
            {
                if (p.Item1.vinculo != null || p.Item2.planilha.Count > 0) continue;
                p.Item1.vinculo = p.Item2;
                p.Item2.planilha.Add(p.Item1);
            }

            // recebimento líquido de gateway: planilha traz o bruto e a taxa em linhas separadas
            foreach (var l in extrato)
            {
                if (l.planilha.Count > 0 || l.descartada || l.saida) continue;
                ItemPlanilha achou = null, taxa = null;
                bool citada = false;
                foreach (var it in itens)
                {
                    if (it.vinculo != null || it.exportado.Length > 0 || Eh(it.saida, true) || Dif(it, l) > tol) continue;
                    long falta = J.RoundJ((it.valor - l.valor) * 100);
                    if (falta <= 0 || falta > Math.Max(500, J.RoundJ(it.valor * 8))) continue;
                    string bruto = J.Moeda(it.valor);
                    foreach (var f in itens)
                    {
                        if (f == it || f.vinculo != null || f.exportado.Length > 0 || Eh(f.saida, false) || J.RoundJ(f.valor * 100) != falta) continue;
                        bool c = f.descricao.Contains(bruto);
                        if (achou == null || c && !citada) { achou = it; taxa = f; citada = c; }
                        if (c) break;
                    }
                    if (citada) break;
                }
                if (achou != null) { achou.vinculo = l; taxa.vinculo = l; l.planilha.Add(achou); }
            }

            // vários recebimentos líquidos (bruto − taxa citada) liquidados num único crédito
            foreach (var l in extrato)
            {
                if (l.planilha.Count > 0 || l.descartada || l.saida) continue;
                var rec = new List<ItemPlanilha>();
                var tax = new List<ItemPlanilha>();
                var liq = new List<long>();
                var usadasTaxa = new HashSet<ItemPlanilha>();
                foreach (var it in itens)
                {
                    if (it.vinculo != null || it.exportado.Length > 0 || Eh(it.saida, true) || Dif(it, l) > tol || it.valor > l.valor + 1000) continue;
                    string bruto = J.Moeda(it.valor);
                    ItemPlanilha f = null;
                    foreach (var x in itens)
                        if (x != it && x.vinculo == null && x.exportado.Length == 0 && !Eh(x.saida, false) && !usadasTaxa.Contains(x)
                            && x.valor < it.valor * 0.08 + 5 && x.descricao.Contains(bruto)) { f = x; break; }
                    if (f == null) continue;
                    usadasTaxa.Add(f);
                    rec.Add(it); tax.Add(f); liq.Add(J.RoundJ(it.valor * 100) - J.RoundJ(f.valor * 100));
                    if (rec.Count >= 20) break;
                }
                if (rec.Count < 2) continue;
                var sol = SomaIdx(liq, 0, J.RoundJ(l.valor * 100), new List<int>());
                if (sol == null) continue;
                foreach (int i in sol) { rec[i].vinculo = l; tax[i].vinculo = l; l.planilha.Add(rec[i]); }
            }

            // agrupados (vários boletos do mesmo fornecedor pagos num único débito)
            foreach (var l in extrato)
            {
                if (l.planilha.Count > 0 || l.descartada) continue;
                var cand = itens.Where(it => it.vinculo == null && it.exportado.Length == 0 && it.valor < l.valor - 0.005
                                             && (it.saida == null || it.saida == l.saida) && Dif(it, l) <= tol).ToList();
                if (cand.Count < 2) continue;
                var grupos = new Dictionary<string, List<ItemPlanilha>>();
                var ordem = new List<string>();
                foreach (var it in cand)
                {
                    string dg = Regex.Replace(it.documento, "[^0-9]", "");
                    string k = it.conta != null ? "#" + it.conta : dg.Length > 0 ? dg : Regras.Normalizar(it.favorecido).Trim();
                    if (k.Length == 0) continue;
                    if (!grupos.ContainsKey(k)) { grupos[k] = new List<ItemPlanilha>(); ordem.Add(k); }
                    grupos[k].Add(it);
                }
                foreach (var k in ordem)
                {
                    var g = grupos[k];
                    if (g.Count < 2 || g.Count > 20) continue;
                    var sol = Soma(g, J.RoundJ(l.valor * 100));
                    if (sol != null) { foreach (var it in sol) { it.vinculo = l; l.planilha.Add(it); } break; }
                }
            }
        }

        static List<int> SomaIdx(List<long> v, int i, long resto, List<int> atual)
        {
            if (resto == 0 && atual.Count >= 2) return new List<int>(atual);
            if (resto <= 0 || i >= v.Count || atual.Count >= 8) return null;
            for (int j = i; j < v.Count; j++)
            {
                if (v[j] > resto) continue;
                atual.Add(j);
                var r = SomaIdx(v, j + 1, resto - v[j], atual);
                atual.RemoveAt(atual.Count - 1);
                if (r != null) return r;
            }
            return null;
        }

        static List<ItemPlanilha> Soma(List<ItemPlanilha> g, long alvo)
        {
            long tot = g.Sum(it => J.RoundJ(it.valor * 100));
            if (tot == alvo) return new List<ItemPlanilha>(g);
            var ord = new List<ItemPlanilha>(g);
            ord.OrdenarEstavel((a, b) => b.valor.CompareTo(a.valor));
            return Busca(ord, 0, alvo, new List<ItemPlanilha>());
        }

        static List<ItemPlanilha> Busca(List<ItemPlanilha> l, int i, long resto, List<ItemPlanilha> atual)
        {
            if (resto == 0 && atual.Count >= 2) return new List<ItemPlanilha>(atual);
            if (resto <= 0 || i >= l.Count || atual.Count >= 8) return null;
            for (int j = i; j < l.Count; j++)
            {
                long v = J.RoundJ(l[j].valor * 100);
                if (v > resto) continue;
                atual.Add(l[j]);
                var r = Busca(l, j + 1, resto - v, atual);
                atual.RemoveAt(atual.Count - 1);
                if (r != null) return r;
            }
            return null;
        }

        public static long Dif(ItemPlanilha it, Lancamento l)
        {
            var dl = J.Data(l.data);
            long best = long.MaxValue;
            foreach (var s in it.datas)
            {
                var d = J.Data(s);
                if (d != null && dl != null) best = Math.Min(best, Math.Abs(J.Dias(d.Value, dl.Value)));
            }
            return best;
        }
    }
}

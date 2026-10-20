using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Planilha financeira do cliente (Excel, CSV ou PDF). Com cabeçalho: datas, valor, entrada/saída, favorecido, CNPJ, descrição,
    /// categoria e conta contábil. Sem cabeçalho (ou PDF): cada linha com data e valor vira um item.
    /// </summary>
    public static partial class LeitorPlanilhaFinanceira
    {
        static readonly Regex MOEDA = new Regex(@"(?<![0-9.,/])(\(?-?\s*(?:R\$\s*)?-?\s*(?:[0-9]{1,3}(?:\.[0-9]{3})+|[0-9]+),[0-9]{2}\)?)(?:\s*([DC])\b)?(?![0-9/])");
        static readonly Regex DOC = new Regex(@"\b[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}/?[0-9]{4}-?[0-9]{2}\b|\b[0-9]{3}\.[0-9]{3}\.[0-9]{3}-[0-9]{2}\b");

        public static List<ItemPlanilha> Ler(string arq)
        {
            string nome = Path.GetFileName(arq);
            bool pdf = nome.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
            if (pdf && LeitorPDF.PareceRelatorioFinanceiro(arq))
            {
                try { return ComNF(LeitorRelatorioFinanceiro.Itens(arq)); }
                catch (IOException e) { Log.Erro("relatorio financeiro " + nome, e); }
            }
            var rows = LeitorTabular.Linhas(arq);
            var r = new List<ItemPlanilha>();
            if (!pdf) r = PorColunas(rows, nome, Regex.IsMatch(nome.ToLowerInvariant(), @"\.xls[xm]$"));
            if (r.Count == 0) r = PorLinha(rows, nome);
            if (r.Count == 0) throw new LeituraException(nome + ": nenhuma linha com data e valor reconhecida na planilha financeira.");
            return ComNF(r);
        }

        /// <summary>Completa os números de NF a partir da descrição (itens lidos sem coluna de NF).</summary>
        static List<ItemPlanilha> ComNF(List<ItemPlanilha> r)
        {
            foreach (var it in r)
                foreach (var n in Conciliador.NumerosNoTexto(it.descricao)) if (!it.nfs.Contains(n)) it.nfs.Add(n);
            return r;
        }

        class Mapa
        {
            public int dPag = -1, dVenc = -1, dGeral = -1, valor = -1, entrada = -1, saida = -1, tipo = -1;
            public int fav = -1, doc = -1, descr = -1, cat = -1, conta = -1, contaAmbigua = -1, nf = -1;
            public List<int> outrosValores = new List<int>();
            public bool Ok() { return (dPag >= 0 || dVenc >= 0 || dGeral >= 0) && (valor >= 0 || entrada >= 0 || saida >= 0); }
        }

        static Mapa Cabecalho(string[] cel)
        {
            var m = new Mapa();
            int valorPrio = -1;
            for (int i = 0; i < cel.Length; i++)
            {
                string h = Regex.Replace(Regex.Replace(Regras.Normalizar(cel[i]), "[^A-Z0-9/ ]", " "), @"\s+", " ").Trim();
                if (h.Length == 0 || h.Length > 40) continue;
                bool ehData = h.Matches(@".*\b(DATA|DT|DIA)\b.*") || h.StartsWith("VENC", StringComparison.Ordinal) || h == "PAGO EM" || h == "BAIXA";
                if (ehData)
                {
                    if (h.Contains("VENC")) { if (m.dVenc < 0) m.dVenc = i; }
                    else if (h.Matches(".*(PAG|BAIXA|LIQUID|QUIT|RECEB|COMPENS|CREDITO|DEBITO|EFETIV|MOVIMENT).*")) { if (m.dPag < 0) m.dPag = i; }
                    else if (!h.Contains("EMISS") && !h.Contains("CADASTR") && !h.Contains("COMPET")) { if (m.dGeral < 0) m.dGeral = i; }
                    else if (m.dGeral < 0 && h.Contains("EMISS")) m.dGeral = i;
                    continue;
                }
                if (h.Matches("(VALOR |VLR |VL )?(ENTRADAS?|RECEITAS?|CREDITOS?|RECEBIMENTOS?)")) { m.entrada = i; continue; }
                if (h.Matches("(VALOR |VLR |VL )?(SAIDAS?|DESPESAS?|DEBITOS?|PAGAMENTOS?)")) { m.saida = i; continue; }
                if (h.Matches("(TIPO|E/S|D/C|C/D|NATUREZA|OPERACAO|MOVIMENTO|ES|DC)( .*)?") && !h.Contains("DOC")) { if (m.tipo < 0) m.tipo = i; continue; }
                if (h.Matches(@".*\b(VALOR|VLR|VL|TOTAL|LIQUIDO|PAGO|RECEBIDO|MONTANTE|IMPORTANCIA)\b.*") && !h.Contains("SALDO"))
                {
                    int prio = h.Matches(".*(PAGO|RECEBIDO|BAIXADO|LIQUIDO|QUITADO|EFETIVO|REALIZADO).*") ? 3 : h.Matches("(VALOR|VLR|VL)( R)?") ? 2 : 1;
                    if (h.Matches(".*(JUROS|MULTA|DESCONTO|ACRESC|TAXA|ABATIMENTO).*")) { m.outrosValores.Add(i); continue; }
                    if (prio > valorPrio) { if (m.valor >= 0) m.outrosValores.Add(m.valor); m.valor = i; valorPrio = prio; }
                    else m.outrosValores.Add(i);
                    continue;
                }
                if (h.Matches(@".*\b(CNPJ|CPF|CNPJ/CPF|CPF/CNPJ)\b.*")) { if (m.doc < 0) m.doc = i; continue; }
                if (h.Matches(".*(FORNECEDOR|CLIENTE|FAVORECIDO|BENEFICIARIO|RAZAO SOCIAL|NOME|PARCEIRO|CONTATO|SACADO|CEDENTE|PAGADOR|RECEBEDOR|CREDOR|DEVEDOR|EMPRESA|PESSOA).*")) { if (m.fav < 0) m.fav = i; continue; }
                if (h.Matches(".*(CONTABIL|REDUZID|COD CONTA|CODIGO CONTA|CONTA COD|CONTA RED).*")) { m.conta = i; continue; }
                if (h.Matches(".*(CATEGORIA|PLANO DE CONTA|CENTRO DE CUSTO|CLASSIFICA|GRUPO|SUBGRUPO|NATUREZA FIN).*")) { if (m.cat < 0) m.cat = i; continue; }
                if (h == "CONTA" || h == "CONTAS") { m.contaAmbigua = i; continue; }
                if (m.nf < 0 && h.Matches(@"((N|NO|NR|NUM|NUMERO) )?(NF|NFE|NF E|NFS|NFSE|NFS E|NOTA|NOTA FISCAL|NOTAS|DOCUMENTO|DOC|NUM DOC|DUPLICATA|DUPL|TITULO|FATURA)( (N|NO|NR|NUM|NUMERO|FISCAL))?"))
                    m.nf = i;
                if (h.Matches(".*(DESCRI|HISTOR|OBSERV|REFERENTE|DOCUMENTO|NOTA|NF|TITULO|PARCELA|DETALHE|MOTIVO).*"))
                {
                    if (m.descr < 0) m.descr = i; else m.outrosValores.Add(-1 - i);
                }
            }
            return m;
        }

        static List<ItemPlanilha> PorColunas(List<string[]> rows, string arquivo, bool xlsx)
        {
            var r = new List<ItemPlanilha>();
            Mapa m = null;
            for (int ln = 0; ln < rows.Count; ln++)
            {
                var cel = rows[ln];
                var h = Cabecalho(cel);
                if (h.Ok() && NaoTemDado(cel)) { m = h; ResolverContaAmbigua(m, rows, ln + 1); continue; }
                if (m == null) continue;
                var it = new ItemPlanilha { arquivo = arquivo, linha = ln + 1 };
                foreach (int c in new[] { m.dPag, m.dVenc, m.dGeral }) { string d = Data(Get(cel, c), xlsx); if (d.Length > 0 && !it.datas.Contains(d)) it.datas.Add(d); }
                if (it.datas.Count == 0) continue;
                double? v = null;
                if (m.entrada >= 0 || m.saida >= 0)
                {
                    double? e = Valor(Get(cel, m.entrada), xlsx), s = Valor(Get(cel, m.saida), xlsx);
                    if (e != null && Math.Abs(e.Value) > 0.004) { v = e; it.saida = e < 0; }
                    else if (s != null && Math.Abs(s.Value) > 0.004) { v = s; it.saida = s > 0; }
                }
                if (v == null && m.valor >= 0)
                {
                    v = Valor(Get(cel, m.valor), xlsx);
                    if (v != null && v < 0) it.saida = true;
                }
                foreach (int c in m.outrosValores)
                {
                    if (c < 0) continue;
                    double? o = Valor(Get(cel, c), xlsx);
                    if (o == null || Math.Abs(o.Value) < 0.005) continue;
                    if (v == null || Math.Abs(v.Value) < 0.005) v = o; else it.alternativos.Add(Math.Abs(o.Value));
                }
                if (v == null || Math.Abs(v.Value) < 0.005) continue;
                it.valor = Math.Abs(v.Value);
                if (m.tipo >= 0) { bool? t = Tipo(Get(cel, m.tipo)); if (t != null) it.saida = t; }
                it.favorecido = Limpar(Get(cel, m.fav));
                it.documento = Get(cel, m.doc).Trim();
                var ds = new StringBuilder(Limpar(Get(cel, m.descr)));
                foreach (int c in m.outrosValores) if (c < 0) { string t = Limpar(Get(cel, -1 - c)); if (t.Length > 0) ds.Append(ds.Length > 0 ? " " : "").Append(t); }
                it.descricao = ds.ToString();
                it.categoria = Limpar(Get(cel, m.cat));
                string ct = Regex.Replace(Get(cel, m.conta), @"\s", "");
                if (ct.Matches("[0-9]{1,7}")) it.conta = int.Parse(ct);
                if (it.documento.Length == 0)
                {
                    var d = DOC.Match(it.favorecido + " " + it.descricao);
                    if (d.Success) it.documento = d.Value;
                }
                if (Regras.Normalizar(it.favorecido + it.descricao).Matches("(SUB)?TOTA.*|SALDO.*")) continue;
                foreach (var n in Conciliador.NumerosDaCelula(Get(cel, m.nf))) if (!it.nfs.Contains(n)) it.nfs.Add(n);
                r.Add(it);
            }
            return r;
        }

        static void ResolverContaAmbigua(Mapa m, List<string[]> rows, int ini)
        {
            if (m.contaAmbigua < 0) return;
            int num = 0, tot = 0;
            for (int i = ini; i < rows.Count && tot < 30; i++)
            {
                string v = Get(rows[i], m.contaAmbigua).Trim();
                if (v.Length == 0) continue;
                tot++;
                if (v.Matches("[0-9]{1,7}")) num++;
            }
            if (tot > 0 && num * 10 >= tot * 8) { if (m.conta < 0) m.conta = m.contaAmbigua; }
            else if (m.cat < 0) m.cat = m.contaAmbigua;
        }

        static bool NaoTemDado(string[] cel)
        {
            foreach (var c in cel) if (DATA.IsMatch(c ?? "") || (c ?? "").Trim().Matches("-?[0-9.,]+")) return false;
            return true;
        }

        static List<ItemPlanilha> PorLinha(List<string[]> rows, string arquivo)
        {
            var r = new List<ItemPlanilha>();
            for (int ln = 0; ln < rows.Count; ln++)
            {
                string t = Regex.Replace(string.Join(" ", rows[ln]).Replace(' ', ' '), @"\s+", " ").Trim();
                string n = Regras.Normalizar(t);
                if (n.Matches(@"(SUB)?TOTA.*|SALDO.*|.*\bSALDO (ANTERIOR|INICIAL|FINAL|DO DIA)\b.*")) continue;
                var it = new ItemPlanilha();
                foreach (Match d in DATA.Matches(t)) { string x = Data(d.Value, false); if (x.Length > 0 && !it.datas.Contains(x)) it.datas.Add(x); }
                if (it.datas.Count == 0) continue;
                var vals = new List<double>();
                bool? sinal = null;
                foreach (Match mv in MOEDA.Matches(t))
                {
                    double? v = Valor(mv.Groups[1].Value, false);
                    if (v == null || Math.Abs(v.Value) < 0.005) continue;
                    if (vals.Count == 0)
                    {
                        if (v < 0) sinal = true;
                        if (mv.Groups[2].Success) sinal = mv.Groups[2].Value == "D";
                    }
                    vals.Add(Math.Abs(v.Value));
                }
                if (vals.Count == 0) continue;
                it.arquivo = arquivo;
                it.linha = ln + 1;
                it.valor = vals[0];
                for (int i = 1; i < vals.Count; i++) it.alternativos.Add(vals[i]);
                it.saida = sinal;
                if (sinal == null)
                {
                    if (n.Matches(@".*\b(PAGAR|PAGO|PAGAMENTO|DESPESA|SAIDA)\b.*")) it.saida = true;
                    else if (n.Matches(@".*\b(RECEBER|RECEBIDO|RECEBIMENTO|RECEITA|ENTRADA)\b.*")) it.saida = false;
                }
                var doc = DOC.Match(t);
                if (doc.Success) it.documento = doc.Value;
                string desc = MOEDA.Replace(DATA.Replace(t, " "), " ");
                if (it.documento.Length > 0) desc = desc.Replace(it.documento, " ");
                it.descricao = Limpar(Regex.Replace(Regex.Replace(desc, @"\bR\$", " ", RegexOptions.IgnoreCase), @"^[\s\-|;:]+|[\s\-|;:]+$", ""));
                r.Add(it);
            }
            return r;
        }

        static string Get(string[] c, int i) { return i >= 0 && i < c.Length && c[i] != null ? c[i] : ""; }

        static string Limpar(string s) { return Regex.Replace(Regex.Replace(s, @"[\r\n\t]+", " "), @"\s{2,}", " ").Trim(); }

        static bool? Tipo(string s)
        {
            string n = Regras.Normalizar(s).Trim();
            if (n.Length == 0) return null;
            if (n.Matches(@"(S|D|SAIDA.*|DESPESA.*|DEBITO.*|PAGAR|PAGO|PAGAMENTO.*|CP|-)")) return true;
            if (n.Matches(@"(E|C|ENTRADA.*|RECEITA.*|CREDITO.*|RECEBER|RECEBIDO|RECEBIMENTO.*|CR|\+)")) return false;
            return null;
        }
    }
}

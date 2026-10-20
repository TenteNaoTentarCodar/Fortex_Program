using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Resultado da conciliação de uma linha do extrato com o(s) item(ns) da planilha financeira vinculados:
    /// confere valor, data, descrição (favorecido/CNPJ) e número da nota fiscal.
    /// </summary>
    public class Conciliacao
    {
        public bool valor, data, descricao;
        public long dias;                                  // maior diferença de dias entre a planilha e o extrato
        public int nf;                                     // 1 = NF confere, 0 = planilha sem NF, -1 = NF não confirmada
        public readonly List<string> nfsPlanilha = new List<string>();
        public readonly List<string> nfsNaoConfirmadas = new List<string>();

        public bool Completa(bool exigirNF)
        {
            return valor && data && descricao && (nf == 1 || nf == 0 && !exigirNF);
        }

        /// <summary>Critérios que não bateram (vazio se todos bateram).</summary>
        public List<string> Divergencias(bool exigirNF)
        {
            var r = new List<string>();
            if (!valor) r.Add("valor");
            if (!data) r.Add("data (" + dias + " dia" + (dias == 1 ? "" : "s") + ")");
            if (!descricao) r.Add("descrição/favorecido");
            if (nf == -1) r.Add("NF " + string.Join(", ", nfsNaoConfirmadas) + " não confirmada");
            else if (nf == 0 && exigirNF) r.Add("planilha sem NF");
            return r;
        }

        public string Texto(bool exigirNF)
        {
            if (Completa(exigirNF))
                return "✔ Conciliado (valor, data, descrição" + (nf == 1 ? ", NF " + string.Join(", ", nfsPlanilha) : "; planilha sem NF") + ")";
            return "Divergente: " + string.Join("; ", Divergencias(exigirNF));
        }
    }

    public static class Conciliador
    {
        // NF citada num texto livre: "NF 1234", "NF-e 001234", "NFS-e nº 55", "Nota Fiscal 1234", "Dupl 1234/2", "Fatura 77"
        static readonly Regex NF_TEXTO = new Regex(
            @"(?<![A-Z0-9])(?:NFS?\s?-?\s?E|NFC\s?-?\s?E|CT\s?-?\s?E|NFS|NF|N\.\s?F\.|NOTA\s+FISCAL(?:\s+ELETRONICA)?|NOTA|DANFE|DUPL(?:ICATA)?|DUP|FATURA|FAT|TITULO|TIT)"
            + @"\s*(?:N\s+|NO\s+|NR\.?\s*|NUM(?:ERO)?\.?\s*|[:#.\-]\s*)?0*([1-9][0-9]{0,8})(?![0-9])");
        static readonly HashSet<string> GUIA = new HashSet<string>(new[] { "DARF", "GPS", "GARE", "DAE", "DARE", "FGTS", "GRF", "GRRF", "IPTU", "IPVA", "DAMSP", "DAM" });
        static readonly Regex NUM = new Regex(@"(?<![0-9.,/])0*([1-9][0-9]{2,8})(?![0-9]|[.,][0-9])");

        /// <summary>Número da NF só com dígitos e sem zeros à esquerda ("000123/2" → "123").</summary>
        public static string NumNF(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var m = Regex.Match(s, "[0-9]+");
            if (!m.Success) return "";
            string d = m.Value.TrimStart('0');
            return d;
        }

        /// <summary>Números de NF citados num texto livre (descrição, histórico).</summary>
        public static List<string> NumerosNoTexto(string texto)
        {
            var r = new List<string>();
            if (string.IsNullOrEmpty(texto)) return r;
            foreach (Match m in NF_TEXTO.Matches(Regras.Normalizar(texto)))
                if (!r.Contains(m.Groups[1].Value)) r.Add(m.Groups[1].Value);
            return r;
        }

        /// <summary>Números de NF de uma célula da coluna NF/Documento: "1234", "1234/1", "1234, 1235", "NF 1234".</summary>
        public static List<string> NumerosDaCelula(string celula)
        {
            var r = NumerosNoTexto(celula);
            if (r.Count > 0 || string.IsNullOrEmpty(celula)) return r;
            string c = celula.Trim();
            if (Regex.IsMatch(c, "[A-Za-z]{3,}")) return r;      // texto sem padrão de NF
            if (Classificador.Documentos(c).Count > 0) return r;  // é CNPJ/CPF
            foreach (var p in Regex.Split(c, @"[,;|]|\se\s"))
            {
                string n = NumNF(p);
                if (n.Length > 0 && n.Length <= 9 && !r.Contains(n)) r.Add(n);
            }
            return r;
        }

        /// <summary>
        /// Avalia a conciliação da linha do extrato com seus itens da planilha. taxas = itens de taxa de gateway
        /// vinculados à linha (bruto − taxa = líquido). Retorna null quando a linha não tem planilha.
        /// </summary>
        public static Conciliacao Avaliar(Lancamento l, IList<ItemPlanilha> taxas, string cnpjEmpresa)
        {
            if (l.planilha.Count == 0 || l.descartada) return null;
            var c = new Conciliacao();

            // valor
            long alvo = J.RoundJ(l.valor * 100);
            if (l.planilha.Count == 1 && (taxas == null || taxas.Count == 0))
            {
                var it = l.planilha[0];
                c.valor = Math.Abs(it.valor - l.valor) <= 0.005 || it.alternativos.Any(a => Math.Abs(a - l.valor) <= 0.005)
                          || l.valorBruto > 0 && Math.Abs(it.valor - l.valorBruto) <= 0.005;
            }
            else
            {
                long soma = l.planilha.Sum(it => J.RoundJ(it.valor * 100)) - (taxas == null ? 0 : taxas.Sum(t => J.RoundJ(t.valor * 100)));
                c.valor = soma == alvo || l.valorBruto > 0 && soma == J.RoundJ(l.valorBruto * 100);
            }

            // data: todos os itens com uma das datas igual à do extrato
            c.dias = 0;
            foreach (var it in l.planilha) c.dias = Math.Max(c.dias, VinculadorPlanilha.Dif(it, l));
            c.data = c.dias == 0;

            // descrição / favorecido
            string texto = TextoExtrato(l, cnpjEmpresa);
            var docs = DocumentosExtrato(l, cnpjEmpresa);
            c.descricao = l.planilha.All(it => DescricaoBate(it, texto, docs));

            // NF
            var nfsLanc = new HashSet<string>(l.notas.Select(t => NumNF(t.nf.numero)).Where(n => n.Length > 0));
            var nfsTexto = new HashSet<string>(NumerosNoTexto(l.historico + " " + (l.comprovante == null ? "" : l.comprovante.descricao + " " + l.comprovante.texto)));
            var numsHist = new HashSet<string>(NUM.Matches(l.historico).Cast<Match>().Select(m => m.Groups[1].Value));
            bool algumaNF = false, todasOk = true;
            foreach (var it in l.planilha)
            {
                foreach (var n in it.nfs)
                {
                    algumaNF = true;
                    if (!c.nfsPlanilha.Contains(n)) c.nfsPlanilha.Add(n);
                    if (!(nfsLanc.Contains(n) || nfsTexto.Contains(n) || numsHist.Contains(n)))
                    {
                        todasOk = false;
                        if (!c.nfsNaoConfirmadas.Contains(n)) c.nfsNaoConfirmadas.Add(n);
                    }
                }
            }
            c.nf = !algumaNF ? 0 : todasOk ? 1 : -1;
            return c;
        }

        /// <summary>
        /// Afinidade de um item da planilha com a linha do extrato antes de vincular as notas (desempate):
        /// 2 = NF do item citada no histórico/comprovante, +1 = descrição/favorecido bate.
        /// </summary>
        public static int Afinidade(ItemPlanilha it, Lancamento l)
        {
            int r = 0;
            if (it.nfs.Count > 0)
            {
                var nt = new HashSet<string>(NumerosNoTexto(l.historico + " " + (l.comprovante == null ? "" : l.comprovante.descricao + " " + l.comprovante.texto)));
                foreach (Match m in NUM.Matches(l.historico)) nt.Add(m.Groups[1].Value);
                if (it.nfs.Any(nt.Contains)) r += 2;
            }
            var docs = new List<string>(Classificador.Documentos(l.historico));
            if (l.comprovante != null && l.comprovante.docFavorecido.Length > 0) docs.Add(Regex.Replace(l.comprovante.docFavorecido, "[^0-9]", ""));
            string tx = l.historico + (l.comprovante == null ? "" : " " + l.comprovante.favorecido + " " + l.comprovante.descricao);
            if (DescricaoBate(it, tx, docs)) r += 1;
            return r;
        }

        /// <summary>Texto do lado do extrato: histórico, comprovante e participante das notas vinculadas.</summary>
        static string TextoExtrato(Lancamento l, string cnpjEmpresa)
        {
            var sb = new StringBuilder(l.historico);
            if (l.buscaSalva != null) sb.Append(' ').Append(l.buscaSalva);
            if (l.comprovante != null) sb.Append(' ').Append(l.comprovante.favorecido).Append(' ').Append(l.comprovante.descricao);
            foreach (var t in l.notas) sb.Append(' ').Append(t.nf.ContraparteNome(cnpjEmpresa));
            return sb.ToString();
        }

        static List<string> DocumentosExtrato(Lancamento l, string cnpjEmpresa)
        {
            var r = new List<string>(Classificador.Documentos(l.historico));
            if (l.comprovante != null)
            {
                string d = Regex.Replace(l.comprovante.docFavorecido, "[^0-9]", "");
                if (d.Length > 0) r.Add(d);
            }
            foreach (var t in l.notas)
            {
                string d = t.nf.ContraparteCnpj(cnpjEmpresa);
                if (d.Length > 0) r.Add(d);
            }
            return r;
        }

        static bool DescricaoBate(ItemPlanilha it, string texto, List<string> docsExtrato)
        {
            string doc = Regex.Replace(it.documento, "[^0-9]", "");
            if (doc.Length == 14 || doc.Length == 11)
            {
                string raiz = doc.Length == 14 ? doc.Substring(0, 8) : doc;
                foreach (var d in docsExtrato)
                    if (d.Length == doc.Length && d.StartsWith(raiz, StringComparison.Ordinal)) return true;
            }
            var te = Classificador.Tokens(texto);
            if (NomesBatem(Classificador.Tokens(it.favorecido), te)) return true;
            var td = Classificador.Tokens(Regex.Replace(it.descricao, @"\b(?:NFS?-?E?|NOTA|PARC(?:ELA)?|DUPL?)\b[\s0-9/.\-]*", " ", RegexOptions.IgnoreCase));
            return NomesBatem(td, te);
        }

        /// <summary>Palavras do nome da planilha encontradas no extrato (aceita nome cortado pelo banco: prefixo de 4+ letras).</summary>
        static bool NomesBatem(List<string> nome, List<string> extrato)
        {
            if (nome.Count == 0 || extrato.Count == 0) return false;
            // guia de tributo: o tipo da guia basta (DARF IRPJ × "DARF NUMERADO")
            if (nome.Any(t => GUIA.Contains(t) && extrato.Contains(t))) return true;
            int comum = 0;
            foreach (var t in nome)
                if (extrato.Any(e => e == t || e.Length >= 4 && t.StartsWith(e, StringComparison.Ordinal) || t.Length >= 4 && e.StartsWith(t, StringComparison.Ordinal))) comum++;
            return comum >= Math.Min(2, nome.Count);
        }
    }
}

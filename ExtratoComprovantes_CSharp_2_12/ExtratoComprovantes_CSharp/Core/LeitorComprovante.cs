using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Comprovantes bancários em PDF (PIX, TED/DOC, transferência, boleto, DARF/DAS/GPS/FGTS, convênios).
    /// Cada página com valor vira um comprovante. Extração por rótulos comuns dos bancos.
    /// </summary>
    public static class LeitorComprovante
    {
        static readonly Regex DINHEIRO = new Regex(@"(?<![0-9.,])(?:R\$\s*)?([0-9]{1,3}(?:\.[0-9]{3})+|[0-9]+),([0-9]{2})(?![0-9])");
        static readonly Regex DATA_NUM = new Regex(@"\b([0-9]{2})/([0-9]{2})/([0-9]{4}|[0-9]{2})\b");
        static readonly Regex DATA_EXT = new Regex(@"\b([0-9]{1,2})\s*(?:DE\s+)?(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[A-Z]*\.?\s*(?:DE\s+)?([0-9]{4})\b");
        static readonly Regex CNPJ = new Regex(@"\b[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}/?[0-9]{4}-?[0-9]{2}\b");
        static readonly Regex CPF = new Regex(@"(?<![0-9*])[0-9*]{3}\.[0-9*]{3}\.[0-9*]{3}-[0-9*]{2}(?![0-9*])|\b[0-9]{11}\b");
        static readonly Regex E2E = new Regex(@"\bE[0-9]{8}[0-9]{12}[A-Za-z0-9]{11}\b");
        static readonly Regex BARRAS = new Regex(@"[0-9][0-9 .\-]{40,70}[0-9]");
        static readonly string[] MESES = { "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ" };

        static readonly string[] ROT_VALOR = {"VALOR PAGO", "VALOR DO PAGAMENTO", "VALOR TOTAL PAGO", "VALOR TOTAL", "VALOR COBRADO",
            "VALOR DA TRANSFERENCIA", "VALOR TRANSFERIDO", "VALOR DO PIX", "VALOR DEBITADO", "VALOR DA OPERACAO", "TOTAL PAGO",
            "VALOR DO DOCUMENTO", "VALOR"};
        static readonly string[] ROT_DATA = {"DATA DO PAGAMENTO", "DATA DE PAGAMENTO", "DATA DO DEBITO", "DATA DE DEBITO",
            "DATA DA TRANSFERENCIA", "DATA DA TRANSACAO", "DATA DA OPERACAO", "DATA DA EFETIVACAO", "PAGO EM", "REALIZADO EM",
            "EFETUADO EM", "DEBITADO EM", "DATA/HORA", "DATA E HORA", "DATA"};
        static readonly string[] ROT_FAVORECIDO = {"DADOS DO RECEBEDOR", "DADOS DO FAVORECIDO", "DADOS DO BENEFICIARIO", "DADOS DO DESTINO",
            "QUEM RECEBEU", "NOME DO FAVORECIDO", "NOME DO RECEBEDOR", "NOME DO BENEFICIARIO", "RAZAO SOCIAL DO BENEFICIARIO",
            "BENEFICIARIO FINAL", "BENEFICIARIO", "FAVORECIDO", "RECEBEDOR", "DESTINATARIO", "DESTINO", "CEDENTE", "PARA",
            "EMPRESA", "CONVENIO", "CONCESSIONARIA", "ORGAO"};
        static readonly string[] ROT_PAGADOR = {"DADOS DO PAGADOR", "DADOS DA ORIGEM", "QUEM PAGOU", "NOME DO PAGADOR", "PAGADOR",
            "ORIGEM", "REMETENTE", "SACADO", "DEBITADO DE", "CONTA DEBITADA", "DE"};
        static readonly HashSet<string> GENERICOS = new HashSet<string> {
            "DE", "PARA", "DESTINO", "ORIGEM", "EMPRESA", "CONVENIO", "CONCESSIONARIA", "ORGAO", "CEDENTE", "RECEBEDOR", "REMETENTE" };
        static readonly string[] NAO_NOME = {"CPF", "CNPJ", "INSTITUICAO", "BANCO", "AGENCIA", "CONTA", "CHAVE", "TIPO", "VALOR", "DATA",
            "NOME", "ISPB", "ID", "AUTENTICACAO", "CODIGO", "DOCUMENTO", "SALDO", "HORA"};

        static int Ix(string s, string t) { return s.IndexOf(t, StringComparison.Ordinal); }

        public static List<Comprovante> Ler(string arquivo)
        {
            var lista = new List<Comprovante>();
            string nome = Path.GetFileName(arquivo);
            List<List<string>> paginas;
            try { paginas = TextoPDF.LinhasPorPagina(arquivo); }
            catch (Exception e) { throw new LeituraException("Não foi possível ler " + nome + " (" + e.Message + ")", e); }
            for (int p = 0; p < paginas.Count; p++)
            {
                string txt = string.Join("\n", paginas[p]);
                if (txt.Trim().Length == 0) continue;
                var c = Interpretar(txt);
                if (c == null) continue;
                c.arquivo = nome;
                c.pagina = p + 1;
                lista.Add(c);
            }
            if (lista.Count == 0) throw new LeituraException(nome + ": nenhum comprovante com texto encontrado (PDF escaneado/imagem?).");
            return lista;
        }

        /// <summary>Interpreta o texto de uma página. null se não houver valor monetário.</summary>
        public static Comprovante Interpretar(string txt)
        {
            var b = new List<string>();
            var n = new List<string>();
            foreach (var l in txt.SplitRe(@"\r?\n"))
            {
                string t = Regex.Replace(l.Replace('\u00A0', ' '), @"\s{2,}", " ").Trim();
                if (t.Length == 0) continue;
                b.Add(t);
                n.Add(Regras.Normalizar(t));
            }
            string tudo = string.Join(" \n", n);
            if (!DINHEIRO.IsMatch(txt)) return null;

            var c = new Comprovante();
            c.texto = string.Join("\n", b);
            c.tipo = Tipo(tudo);
            c.valor = Valor(b, n);
            c.data = Data(b, n);
            var fav = Parte(b, n, ROT_FAVORECIDO, ROT_PAGADOR);
            c.favorecido = fav[0]; c.docFavorecido = fav[1];
            var pag = Parte(b, n, ROT_PAGADOR, ROT_FAVORECIDO);
            c.pagador = pag[0]; c.docPagador = pag[1];

            if (c.tipo == "DARF" || c.tipo == "DAS" || c.tipo == "GPS" || c.tipo == "FGTS")
            {
                string cod = DepoisDe(b, n, new[] { "CODIGO DA RECEITA", "COD. RECEITA", "CODIGO RECEITA", "CODIGO DE PAGAMENTO" });
                string per = DepoisDe(b, n, new[] { "PERIODO DE APURACAO", "PERIODO APURACAO", "COMPETENCIA", "PA " });
                if (c.favorecido.Length == 0) c.favorecido = c.tipo == "FGTS" ? "FGTS - CAIXA" : (c.tipo == "GPS" ? "INSS - GPS" : "RECEITA FEDERAL");
                var d = new StringBuilder(c.tipo);
                if (cod.Length > 0) d.Append(" cód. ").Append(Regex.Replace(cod, @"[^0-9-]", ""));
                if (per.Length > 0) d.Append(" PA ").Append(per);
                c.descricao = d.ToString();
                if (c.docPagador.Length == 0)
                {
                    string contrib = DepoisDe(b, n, new[] { "CNPJ", "CPF", "CONTRIBUINTE" });
                    var m = CNPJ.Match(contrib);
                    if (m.Success) c.docPagador = m.Value;
                }
            }
            c.codigoBarras = CodigoBarras(b);
            var e2e = E2E.Match(c.texto);
            if (e2e.Success) c.idTransacao = e2e.Value;
            else
            {
                string aut = DepoisDe(b, n, new[] { "AUTENTICACAO", "ID DA TRANSACAO", "ID DA TRANSFERENCIA", "NUMERO DE CONTROLE", "NSU" });
                c.idTransacao = aut.Length > 60 ? aut.Substring(0, 60) : aut;
            }
            if (c.descricao.Length == 0)
            {
                string desc = DepoisDe(b, n, new[] { "DESCRICAO", "MENSAGEM", "INFORMACAO AO RECEBEDOR", "IDENTIFICACAO", "HISTORICO", "FINALIDADE" });
                c.descricao = desc.Length > 120 ? desc.Substring(0, 120) : desc;
            }
            c.conferir = c.valor == 0 || c.data.Length == 0 || c.favorecido.Length == 0;
            return c;
        }

        static string Tipo(string t)
        {
            if (t.Contains("DARF")) return "DARF";
            if (t.Contains("SIMPLES NACIONAL") || Regex.IsMatch(t, @"\bDAS\b")) return "DAS";
            if (t.Contains("GUIA DA PREVIDENCIA") || Regex.IsMatch(t, @"\bGPS\b")) return "GPS";
            if (t.Contains("FGTS")) return "FGTS";
            if (t.Contains("PIX")) return "PIX";
            if (Regex.IsMatch(t, @"\bTED\b")) return "TED";
            if (Regex.IsMatch(t, @"\bDOC\b") && t.Contains("TRANSF")) return "DOC";
            if (t.Contains("BOLETO") || t.Contains("TITULO") || t.Contains("LINHA DIGITAVEL") || t.Contains("CODIGO DE BARRAS")) return "Boleto";
            if (t.Contains("CONVENIO") || t.Contains("CONCESSIONARIA") || t.Contains("CONTA DE CONSUMO") || t.Contains("ENERGIA") || t.Contains("AGUA")) return "Convênio";
            if (t.Contains("TRANSFERENCIA")) return "Transferência";
            if (t.Contains("DEPOSITO")) return "Depósito";
            return "Outro";
        }

        static double Num(Match m) { return J.ParseD(m.Groups[1].Value.Replace(".", "") + "." + m.Groups[2].Value); }

        static double Valor(List<string> b, List<string> n)
        {
            foreach (var rot in ROT_VALOR)
                for (int i = 0; i < n.Count; i++)
                {
                    int p = Ix(n[i], rot);
                    if (p < 0) continue;
                    if (rot == "VALOR" && n[i].Matches(@".*VALOR (DO DOCUMENTO|ORIGINAL|NOMINAL|DO TITULO|DE FACE)?\s*(DESCONTO|JUROS|MULTA|ABATIMENTO|ENCARGO).*")) continue;
                    for (int k = i; k <= Math.Min(i + 2, n.Count - 1); k++)
                    {
                        string alvo = k == i ? b[k].Substring(Math.Min(b[k].Length, p)) : b[k];
                        var m = DINHEIRO.Match(alvo);
                        if (m.Success) return Num(m);
                        if (k > i && Regras.Normalizar(b[k]).StartsWith("VALOR", StringComparison.Ordinal)) break;
                    }
                }
            double max = 0;
            foreach (var l in b) foreach (Match m in DINHEIRO.Matches(l)) max = Math.Max(max, Num(m));
            return max;
        }

        static string Data(List<string> b, List<string> n)
        {
            foreach (var rot in ROT_DATA)
                for (int i = 0; i < n.Count; i++)
                {
                    string l = n[i];
                    int p = Ix(l, rot);
                    if (p < 0 || l.Contains("VENC") || l.Contains("EMISSAO") || l.Contains("AGENDAMENTO EM")) continue;
                    for (int k = i; k <= Math.Min(i + 1, n.Count - 1); k++)
                    {
                        string d = AcharData(k == i ? n[k].Substring(p) : n[k]);
                        if (d.Length > 0) return d;
                    }
                }
            foreach (var l in n)
            {
                if (l.Contains("VENC")) continue;
                string d = AcharData(l);
                if (d.Length > 0) return d;
            }
            return "";
        }

        static string AcharData(string s)
        {
            var m = DATA_NUM.Match(s);
            if (m.Success)
            {
                int dd = int.Parse(m.Groups[1].Value), mm = int.Parse(m.Groups[2].Value);
                if (dd >= 1 && dd <= 31 && mm >= 1 && mm <= 12)
                    return m.Groups[1].Value + "/" + m.Groups[2].Value + "/" + (m.Groups[3].Value.Length == 2 ? "20" + m.Groups[3].Value : m.Groups[3].Value);
            }
            m = DATA_EXT.Match(s);
            if (m.Success)
                for (int i = 0; i < 12; i++)
                    if (MESES[i] == m.Groups[2].Value) return int.Parse(m.Groups[1].Value).ToString("00") + "/" + (i + 1).ToString("00") + "/" + m.Groups[3].Value;
            return "";
        }

        static readonly Regex TIRA_INICIO = new Regex(@"^[\s:\-]+");
        static readonly Regex ROT_NOME = new Regex(@"^\s*(nome( do (favorecido|recebedor|benefici[aá]rio|pagador))?|raz[aã]o social)\s*:?\s*", RegexOptions.IgnoreCase);

        static string[] Parte(List<string> b, List<string> n, string[] rotulos, string[] outraParte)
        {
            foreach (var rot in rotulos)
                for (int i = 0; i < n.Count; i++)
                {
                    string l = n[i];
                    if (!(l == rot || l.StartsWith(rot + " ", StringComparison.Ordinal) || l.StartsWith(rot + ":", StringComparison.Ordinal) || l.StartsWith(rot + " :", StringComparison.Ordinal))) continue;
                    if (GENERICOS.Contains(rot) && l != rot)
                    {
                        int dois = b[i].IndexOf(':');
                        if (dois < 0 || dois > rot.Length + 3) continue;
                    }
                    string nome = "", docu = "";
                    string resto = b[i].Length > rot.Length ? TIRA_INICIO.Replace(b[i].Substring(Math.Min(b[i].Length, IndiceDepois(b[i], rot))), "", 1).Trim() : "";
                    if (resto.Length > 0)
                    {
                        string d = Documento(resto);
                        if (d.Length > 0) { docu = d; resto = Regex.Replace(resto.Replace(d, ""), @"\b(CPF|CNPJ)\b[:\s]*", "", RegexOptions.IgnoreCase).Trim(); }
                        if (EhNome(resto)) nome = resto;
                    }
                    for (int k = i + 1; k < Math.Min(i + 8, n.Count); k++)
                    {
                        string nk = n[k];
                        if (Inicia(nk, outraParte)) break;
                        if (docu.Length == 0) docu = Documento(b[k]);
                        if (nome.Length == 0)
                        {
                            if (nk.StartsWith("NOME", StringComparison.Ordinal) || nk.StartsWith("RAZAO SOCIAL", StringComparison.Ordinal))
                            {
                                string v = ROT_NOME.Replace(b[k], "", 1).Trim();
                                if (EhNomeLivre(v)) nome = v;
                                else if (k + 1 < n.Count && EhNomeLivre(b[k + 1])) nome = b[k + 1].Trim();
                            }
                            else if (!Inicia(nk, NAO_NOME) && EhNome(b[k])) nome = b[k].Trim();
                        }
                        if (nome.Length > 0 && docu.Length > 0) break;
                    }
                    if (nome.Length > 0) return new[] { LimparNome(nome), docu };
                }
            return new[] { "", "" };
        }

        static int IndiceDepois(string bruta, string rotNorm)
        {
            string n = Regras.Normalizar(bruta);
            int p = Ix(n, rotNorm);
            return p < 0 ? bruta.Length : Math.Min(bruta.Length, p + rotNorm.Length);
        }

        static bool Inicia(string norm, string[] lista)
        {
            foreach (var r in lista) if (norm == r || norm.StartsWith(r + " ", StringComparison.Ordinal) || norm.StartsWith(r + ":", StringComparison.Ordinal)) return true;
            return false;
        }

        static int Letras(string s) { return Regex.Replace(s, "[^A-Za-zÀ-ú]", "").Length; }

        static bool EhNome(string s)
        {
            if (s == null) return false;
            s = s.Trim();
            if (s.Length < 3) return false;
            int letras = Letras(s);
            if (letras < 3) return false;
            if (DINHEIRO.IsMatch(s) && letras < 6) return false;
            string n = Regras.Normalizar(s);
            return !Inicia(n, NAO_NOME) && !n.StartsWith("R ", StringComparison.Ordinal) && !n.Contains("COMPROVANTE");
        }

        static bool EhNomeLivre(string s)
        {
            if (s == null) return false;
            string n = Regras.Normalizar(s);
            if (n.Matches("(CPF|CNPJ|CHAVE|INSTITUICAO|AGENCIA|CONTA|VALOR|DATA|TIPO)( .*)?:?")) return n.Split(' ').Length > 2 && !n.Contains(":");
            return Letras(s) >= 3 && !DINHEIRO.IsMatch(s);
        }

        static string LimparNome(string s)
        {
            return Regex.Replace(Regex.Replace(s, @"\b(CPF|CNPJ)\b.*$", "", RegexOptions.IgnoreCase), @"[\s:\-]+$", "").Trim();
        }

        static string Documento(string s)
        {
            var m = CNPJ.Match(s);
            if (m.Success) return m.Value;
            m = CPF.Match(s);
            return m.Success ? m.Value : "";
        }

        static string DepoisDe(List<string> b, List<string> n, string[] rotulos)
        {
            foreach (var rot in rotulos)
                for (int i = 0; i < n.Count; i++)
                {
                    int p = Ix(n[i], rot);
                    if (p < 0) continue;
                    string resto = TIRA_INICIO.Replace(b[i].Substring(Math.Min(b[i].Length, IndiceDepois(b[i], rot))), "", 1).Trim();
                    if (resto.Length > 0) return resto;
                    if (i + 1 < b.Count) return b[i + 1].Trim();
                }
            return "";
        }

        static string CodigoBarras(List<string> b)
        {
            foreach (var l in b)
                foreach (Match m in BARRAS.Matches(l))
                {
                    string dig = Regex.Replace(m.Value, "[^0-9]", "");
                    if (dig.Length == 44 || dig.Length == 47 || dig.Length == 48) return dig;
                }
            for (int i = 0; i + 1 < b.Count; i++)
            {
                string dig = Regex.Replace(b[i] + b[i + 1], "[^0-9]", "");
                if ((dig.Length == 47 || dig.Length == 48) && Regex.Replace(b[i], "[^0-9]", "").Length >= 20) return dig;
            }
            return "";
        }
    }
}

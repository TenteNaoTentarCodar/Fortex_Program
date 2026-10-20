using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Contrapartida de cada linha do extrato:
    ///  0) regra manual do cliente; conta da planilha financeira;
    ///  1) funcionário do quadro (salário 1–10 → Salários a Pagar; demais → Adiantamento);
    ///  2) aprendido (CNPJ/CPF, palavras do histórico, categoria);
    ///  3) Razão: lançamento com banco, mesmo sentido, histórico semelhante (mesmo banco tem preferência);
    ///  4) Plano: conta do terceiro no lado correto, categoria = nome de conta, natureza da operação;
    ///  5) Valores a Identificar (685). Toda conta precisa existir no Plano e ser analítica.
    /// </summary>
    public class Classificador
    {
        public Regras regras;
        public Plano plano;
        public Razao razao;
        public List<string> funcionarios = new List<string>();
        public ICollection<int> contasBanco = new HashSet<int>();
        public Dictionary<string, int> categoriaSessao = new Dictionary<string, int>();

        List<Razao.Ref> refs = new List<Razao.Ref>();
        readonly Dictionary<string, List<int>> indice = new Dictionary<string, List<int>>();
        readonly Dictionary<string, double> idf = new Dictionary<string, double>();
        readonly HashSet<int> contasNoRazao = new HashSet<int>();
        readonly List<KeyValuePair<Plano.Conta, HashSet<string>>> terceiros = new List<KeyValuePair<Plano.Conta, HashSet<string>>>();
        /// <summary>Contas de adiantamento a sócio com o nome do sócio (ex.: "ADIANTAMENTO A SOCIO - ENOCH ALVES") e as palavras do nome.</summary>
        /// <summary>Relação de líquidos da folha e do adiantamento do cliente (null = não carregada).</summary>
        public FolhaPagamento folha;
        readonly List<KeyValuePair<Plano.Conta, List<string>>> socios = new List<KeyValuePair<Plano.Conta, List<string>>>();
        static readonly HashSet<string> PALAVRA_SOCIO = new HashSet<string>(new[] { "ADIANTAMENTO", "ADIANTAMENTOS", "ADTO", "ADIANT", "SOCIO", "SOCIOS", "SOCIA", "SOCIAS", "CONTA", "CORRENTE", "EMPRESTIMO", "EMPRESTIMOS", "MUTUO" });
        readonly List<KeyValuePair<string, List<string>>> funcs = new List<KeyValuePair<string, List<string>>>();
        Plano.Conta contaSalarios, contaAdiantamento;

        static readonly HashSet<string> STOP = new HashSet<string>(new[] {
            "PIX", "ENVIADO", "ENVIADA", "RECEBIDO", "RECEBIDA", "RECEB", "PAGAMENTO", "PAGTO", "PGTO", "PAG", "EFETUADO", "EFETUADA",
            "TED", "DOC", "TEF", "TRANSF", "TRANSFERENCIA", "ENVIO", "CP", "LTDA", "EPP", "EIRELI", "S/A", "SA", "ME", "MEI", "NF", "NFE",
            "NOTA", "FISCAL", "REF", "REFERENTE", "PARCELA", "PARC", "VALOR", "DEBITO", "CREDITO", "CONTA", "BOLETO", "TITULO", "TIT",
            "DE", "DA", "DO", "DAS", "DOS", "EM", "COM", "PARA", "POR", "VIA", "DEB", "CRED", "LIQUIDACAO", "COMPRA", "DIVERSOS",
            "COMERCIO", "INDUSTRIA", "SERVICOS", "SERVICO", "CIA", "COMPANHIA", "INSTITUICAO", "PAGAMENTOS", "AGENDADO", "AGENDAMENTO",
            "INTERNET", "BANKING", "APP", "CONF", "CONFORME", "MES", "COMPETENCIA", "OUTROS", "AUTOMATICO", "AUT",
            "EMPRESA", "BRASIL", "THE", "AND", "INTER", "NUMERADO", "VR", "VLR", "ADM", "TRANSFERENCIAS", "REALIZADO", "REALIZADA",
            "PAGO", "PAGA", "ELETRON", "ELETRONICO", "COBRANCA", "REM", "DEST", "RECEBIMENTO", "RECEBIMENTOS", "SISPAG", "FAVORECIDO",
            "TRANSFERIDO", "DEBITADO", "CREDITADO", "LIQUIDO", "BRUTO", "OPERACAO", "TRANSACAO", "DIG" });
        static readonly HashSet<string> TRIBUTO = new HashSet<string>(new[] {
            "DARF", "GPS", "DAS", "GUIA", "TRIBUTO", "TRIBUTOS", "IMPOSTO", "IMPOSTOS", "SIMPLES", "NACIONAL", "RECEITA", "FEDERAL", "GARE",
            "DAE", "ESTADUAL", "MUNICIPAL", "ARRECADACAO" });
        static readonly HashSet<string> GENERICO_CONTA = new HashSet<string>(new[] {
            "FORNECEDOR", "FORNECEDORES", "CLIENTE", "CLIENTES", "PAGAR", "RECEBER", "DUPLICATAS", "DUPLICATA", "CONTAS", "TITULOS",
            "ADIANTAMENTO", "ADIANTAMENTOS", "NACIONAIS", "DIVERSOS", "OUTRAS", "OUTROS", "CURTO", "LONGO", "PRAZO", "RECOLHER", "RECUPERAR" });

        static readonly string[][][] NATUREZA = {
            new[] { new[] { "TARIFA", "TAR", "CESTA", "PACOTE", "MANUTENCAO CONTA", "TAXA BANCARIA", "ANUIDADE" }, new[] { "S" }, new[] { "TARIFA", "DESPESAS BANCARIAS", "DESPESA BANCARIA", "TAXAS BANCARIAS" } },
            new[] { new[] { "IOF" }, new[] { "S" }, new[] { "IOF" } },
            new[] { new[] { "JUROS", "MORA", "ENCARGO" }, new[] { "S" }, new[] { "JUROS PASSIVOS", "JUROS PAGOS", "JUROS E MULTAS", "JUROS S/", "ENCARGOS FINANCEIROS" } },
            new[] { new[] { "RENDIMENTO", "REND", "RENDIMENTOS", "RENTAB", "RENTABILIDADE", "REMUNERACAO APLIC" }, new[] { "E" }, new[] { "RENDIMENTO DE APLIC", "RENDIMENTOS DE APLIC", "RENDIMENTO", "JUROS DE APLIC", "RECEITAS FINANCEIRAS", "RECEITA FINANCEIRA", "JUROS ATIVOS" } },
        };

        static readonly Regex NAO_ALFANUM = new Regex("[^A-Z0-9]+");
        static readonly Regex NAO_LETRA = new Regex("[^A-Z]+");

        /// <summary>Palavras significativas (3+ letras, sem números puros nem palavras de banco), na ordem.</summary>
        public static List<string> Tokens(string s)
        {
            var r = new List<string>();
            if (s == null) return r;
            var vistos = new HashSet<string>();
            foreach (var t in NAO_ALFANUM.Split(Regras.Normalizar(s)))
            {
                if (t.Length < 3 || t.All(char.IsDigit) || STOP.Contains(t)) continue;
                if (vistos.Add(t)) r.Add(t);
            }
            return r;
        }

        public static HashSet<string> TokSet(string s) { return new HashSet<string>(Tokens(s)); }

        public void Preparar()
        {
            indice.Clear(); idf.Clear(); contasNoRazao.Clear(); terceiros.Clear(); funcs.Clear(); socios.Clear();
            var bancos = new HashSet<int>(contasBanco);
            if (regras != null && regras.contaBanco != null) bancos.Add(regras.contaBanco.Value);
            refs = razao == null ? new List<Razao.Ref>() : razao.Referencias(plano, bancos);
            for (int i = 0; i < refs.Count; i++)
            {
                var r = refs[i];
                r.tokens = TokSet(r.historico);
                contasNoRazao.Add(r.contra);
                foreach (var t in Tokens(r.historico)) indice.ObterOuCriar(t).Add(i);
            }
            int n = Math.Max(1, refs.Count);
            foreach (var e in indice) idf[e.Key] = Math.Log(1.0 + n / (double)e.Value.Count);
            contaSalarios = null; contaAdiantamento = null;
            if (plano != null)
            {
                foreach (var c in plano.lista)
                {
                    if (!c.Analitica() || c.Disponivel()) continue;
                    string nm = Regras.Normalizar(c.nome);
                    if (contaSalarios == null && c.Grupo() == '2' && nm.Contains("SALARI") && (nm.Contains("PAGAR") || nm.Contains("ORDENADO"))) contaSalarios = c;
                    if (contaAdiantamento == null && c.Grupo() == '1' && nm.Contains("ADIANTAMENTO") && nm.Contains("SALARI")) contaAdiantamento = c;
                    if (c.Grupo() == '1')
                    {
                        // conta de adiantamento a sócio com o nome dele: as palavras do nome, sem "ADIANTAMENTO", "SOCIO" etc.
                        var ns = NAO_LETRA.Split(nm).Where(t => t.Length >= 2).ToList();
                        if (ns.Any(t => t == "SOCIO" || t == "SOCIOS" || t == "SOCIA") && ns.Any(t => t.StartsWith("ADIANT", StringComparison.Ordinal) || t == "ADTO"))
                        {
                            var nome = ns.Where(t => t.Length >= 3 && !PALAVRA_SOCIO.Contains(t) && !STOP.Contains(t)).ToList();
                            if (nome.Count >= 2) socios.Add(new KeyValuePair<Plano.Conta, List<string>>(c, nome));
                        }
                    }
                    if (c.Grupo() != '1' && c.Grupo() != '2') continue;
                    var tk = Tokens(c.nome).Where(t => !GENERICO_CONTA.Contains(t) && !TRIBUTO.Contains(t)).ToList();
                    if (tk.Any(t => t.Length >= 4)) terceiros.Add(new KeyValuePair<Plano.Conta, HashSet<string>>(c, new HashSet<string>(tk)));
                }
            }
            foreach (var f in funcionarios)
            {
                var tk = NAO_LETRA.Split(Regras.Normalizar(f)).Where(t => t.Length >= 3 && !STOP.Contains(t)).ToList();
                if (tk.Count >= 2) funcs.Add(new KeyValuePair<string, List<string>>(f, tk));
            }
        }

        public int QtdReferencias() { return refs.Count; }

        double Peso(string t) { double v; return idf.TryGetValue(t, out v) ? v : Math.Log(1.0 + Math.Max(1, refs.Count)); }

        public void Classificar(Lancamento l, string textoBusca)
        {
            if (l.manual || l.descartada) return;
            int? banco = l.contaBanco ?? regras.contaBanco;
            string prefixo = l.conferirSinal ? (l.obs.StartsWith("Sinal") ? Regex.Replace(l.obs, @" \|.*", "") + " | " : "Sinal deduzido | ") : "";
            if (banco == null || plano != null && !plano.Valida(banco))
            {
                l.debito = null; l.credito = null;
                l.obs = prefixo + (banco == null ? "Sem banco/caixa identificado — informe a conta banco" : "Conta banco " + banco + " inexistente/sintética no Plano");
                return;
            }
            int? contra = null;
            string motivo;
            var tk = TokSet(textoBusca);
            var tkOrdem = Tokens(textoBusca);
            bool soTributo = tk.Count > 0 && tk.All(TRIBUTO.Contains) || tk.Count == 0 && ContemTributo(textoBusca);

            var rm = regras.Procurar(textoBusca, l.saida);
            var func = l.saida ? Funcionario(textoBusca) : null;
            int? cpl = ContaPlanilha(l);
            object[] ap;
            Plano.Conta sc;
            var fr = folha == null ? null : folha.Encontrar(l, w => Contem(tk, w));
            if (fr != null && ContaOk(fr.conta)) { contra = fr.conta; motivo = fr.Motivo(); }
            else if (rm != null && ContaOk(rm.conta)) { contra = rm.conta; motivo = "Regra: " + rm.texto; }
            else if (rm != null) motivo = "Regra \"" + rm.texto + "\" usa conta " + rm.conta + " inexistente/sintética → Valores a Identificar";
            else if (cpl != null && ContaOk(cpl.Value) && !contasBanco.Contains(cpl.Value) && cpl != banco)
            {
                contra = cpl; motivo = "Planilha financeira: conta " + contra + " informada na planilha";
            }
            else if (l.saida && (sc = Socio(tk)) != null)
            {
                contra = sc.codigo; motivo = "Plano: adiantamento a sócio (" + sc.nome + ")";
            }
            else if (func != null)
            {
                bool salario = TermoSalario(textoBusca) && Dia(l.data) >= 1 && Dia(l.data) <= 10;
                var c = salario ? contaSalarios : contaAdiantamento;
                if (c != null || plano == null)
                {
                    contra = c != null ? c.codigo : (int?)null;
                    motivo = "Funcionário " + func + (salario ? ": salário (dias 1–10)" : ": adiantamento");
                }
                else motivo = "Funcionário " + func + ": conta de " + (salario ? "Salários a Pagar" : "Adiantamento de Salário") + " não encontrada no Plano → Valores a Identificar";
            }
            else if (!soTributo && (ap = Aprendido(l, textoBusca, tk, banco)) != null) { contra = (int)ap[0]; motivo = (string)ap[1]; }
            else if (soTributo && (ap = CategoriaAprendida(l, banco)) != null) { contra = (int)ap[0]; motivo = (string)ap[1]; }
            else if (soTributo) motivo = "Tributo sem identificação no histórico → Valores a Identificar";
            else
            {
                var rz = RazaoRef(tkOrdem, l.saida, banco);
                var pl = Terceiro(tk, l.saida);
                Plano.Conta terc = pl != null && pl.Value.Value ? pl.Value.Key : null;
                Razao.Ref rf = rz;
                Plano.Conta contaRz = rf == null || plano == null ? null : plano.Get(rf.contra);
                bool rzDespesaDireta = contaRz != null && contaRz.Grupo() != '1' && contaRz.Grupo() != '2';
                Plano.Conta cplano;
                if (rf != null && terc != null && rf.contra != terc.codigo && rzDespesaDireta)
                {
                    contra = terc.codigo;
                    motivo = "Plano: conta do fornecedor " + terc.nome + " (Razão usava despesa direta " + rf.contra + ")";
                }
                else if (rf != null) { contra = rf.contra; motivo = string.Format("Razão: \"{0}\" ({1})", Curto(rf.historico), rf.data); }
                else if (terc != null) { contra = terc.codigo; motivo = "Plano: conta do " + (l.saida ? "fornecedor " : "cliente ") + terc.nome; }
                else if ((cplano = CategoriaNoPlano(l)) != null)
                {
                    contra = cplano.codigo;
                    motivo = "Planilha financeira: categoria \"" + CategoriaPlanilha(l) + "\" = conta " + cplano.nome;
                }
                else if ((ap = CategoriaAprendida(l, banco)) != null) { contra = (int)ap[0]; motivo = (string)ap[1]; }
                else
                {
                    var nat = Natureza(textoBusca, l.saida);
                    if (nat != null && nat.Item1 != null) { contra = nat.Item1.codigo; motivo = "Plano (natureza): " + nat.Item2; }
                    else if (nat != null) motivo = "Dúvida entre contas de " + nat.Item2 + " → Valores a Identificar";
                    else if (pl != null) motivo = "Terceiro sem conta no lado correto (" + pl.Value.Key.nome + ") → Valores a Identificar";
                    else motivo = (refs.Count == 0 && plano == null) ? "Sem Razão/Plano carregados → Valores a Identificar" : "Não localizado no Razão/Plano → Valores a Identificar";
                }
            }
            if (contra == null) contra = regras.contaVAI;
            if (l.saida) { l.debito = contra; l.credito = banco; }
            else { l.debito = banco; l.credito = contra; }
            var pc = plano == null ? null : plano.Get(contra);
            l.obs = prefixo + motivo + (pc != null && contra != regras.contaVAI ? "  [" + pc.codigo + " " + pc.nome + "]" : "");
        }

        // ---------------------------------------------------------------- aprendizado

        static readonly Regex DOCS = new Regex(@"(?<!\d)(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}|\d{3}\.?\d{3}\.?\d{3}-?\d{2})(?!\d)");

        /// <summary>CNPJ/CPF válidos citados no texto (só dígitos).</summary>
        public static List<string> Documentos(string texto)
        {
            var r = new List<string>();
            if (texto == null) return r;
            foreach (Match m in DOCS.Matches(texto))
            {
                string d = Regex.Replace(m.Groups[1].Value, @"\D", "");
                if ((d.Length == 11 ? CpfOk(d) : d.Length == 14 && CnpjOk(d)) && !r.Contains(d)) r.Add(d);
            }
            return r;
        }

        internal static bool CpfOk(string d)
        {
            if (d.Distinct().Count() == 1) return false;
            for (int n = 9; n <= 10; n++)
            {
                int s = 0;
                for (int i = 0; i < n; i++) s += (d[i] - '0') * (n + 1 - i);
                int dv = s * 10 % 11 % 10;
                if (dv != d[n] - '0') return false;
            }
            return true;
        }

        internal static bool CnpjOk(string d)
        {
            if (d.Distinct().Count() == 1) return false;
            int[] p = { 6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2 };
            for (int n = 12; n <= 13; n++)
            {
                int s = 0;
                for (int i = 0; i < n; i++) s += (d[i] - '0') * p[i + 13 - n];
                int dv = s % 11 < 2 ? 0 : 11 - s % 11;
                if (dv != d[n] - '0') return false;
            }
            return true;
        }

        public static string ChaveHistorico(string historico)
        {
            var t = Tokens(historico);
            return t.Any(x => x.Length >= 4) ? string.Join(" ", t) : "";
        }

        public static string ChaveCategoria(string cat)
        {
            return cat == null ? "" : Regras.Normalizar(Regex.Replace(cat, @"\(\s*[\d.,]+\s*%\s*\)", "")).Trim();
        }

        public static string CategoriaDaLinha(Lancamento l)
        {
            string c = CategoriaPlanilha(l);
            return c == null ? null : ChaveCategoria(c);
        }

        bool ContaAprendidaOk(int? c, int? banco) { return c != null && c != regras.contaVAI && c != banco && ContaOk(c.Value); }

        object[] Aprendido(Lancamento l, string textoBusca, HashSet<string> tk, int? banco)
        {
            if (regras.aprendido.Count == 0) return null;
            var docs = Documentos(textoBusca);
            foreach (var it in l.planilha) foreach (var d in Documentos(it.documento)) if (!docs.Contains(d)) docs.Add(d);
            foreach (var d in docs)
            {
                var c = regras.Aprendida(Regras.DOC, l.saida, d);
                if (ContaAprendidaOk(c, banco)) return new object[] { c.Value, "Aprendido (CNPJ/CPF " + d + ")" };
            }
            int? melhor = null; string melhorTxt = null; int melhorN = 0; bool empate = false;
            foreach (var a in regras.HistoricosAprendidos(l.saida))
            {
                var ws = a.Key.Split(' ');
                bool todas = ws.All(w => Contem(tk, w));
                if (!todas) continue;
                int c = a.Value;
                if (!ContaAprendidaOk(c, banco)) continue;
                if (ws.Length > melhorN) { melhor = c; melhorTxt = a.Key; melhorN = ws.Length; empate = false; }
                else if (ws.Length == melhorN && c != melhor) empate = true;
            }
            if (melhor != null && !empate) return new object[] { melhor.Value, "Aprendido: \"" + melhorTxt + "\"" };
            return null;
        }

        object[] CategoriaAprendida(Lancamento l, int? banco)
        {
            string cat = CategoriaDaLinha(l);
            if (string.IsNullOrEmpty(cat)) return null;
            var c = regras.Aprendida(Regras.CAT, l.saida, cat);
            if (ContaAprendidaOk(c, banco)) return new object[] { c.Value, "Aprendido (categoria da planilha \"" + cat + "\")" };
            int cs;
            if (categoriaSessao.TryGetValue((l.saida ? "S|" : "E|") + cat, out cs) && ContaAprendidaOk(cs, banco))
                return new object[] { cs, "Categoria da planilha \"" + cat + "\" (mesma conta das outras linhas deste extrato)" };
            return null;
        }

        static int Dia(string data) { int d; return data != null && data.Length >= 2 && int.TryParse(data.Substring(0, 2), out d) ? d : 0; }

        static bool TermoSalario(string s)
        {
            string n = Regras.Normalizar(s);
            return n.Contains("SALAR") || Regex.IsMatch(n, @"\bSAL\b") || n.Contains("FOLHA") || n.Contains("ORDENADO") || n.Contains("PROVENTO");
        }

        string Funcionario(string texto)
        {
            var tk = new HashSet<string>(NAO_LETRA.Split(Regras.Normalizar(texto)));
            string melhor = null;
            int melhorN = 0;
            foreach (var f in funcs)
            {
                var nome = f.Value;
                if (!tk.Contains(nome[0])) continue;
                int n = nome.Count(t => tk.Contains(t));
                if (n >= 2 && n > melhorN) { melhor = f.Key; melhorN = n; }
            }
            return melhor;
        }

        bool ContemTributo(string s) { return NAO_ALFANUM.Split(Regras.Normalizar(s)).Any(TRIBUTO.Contains); }

        static int? ContaPlanilha(Lancamento l)
        {
            int? c = null;
            foreach (var it in l.planilha)
            {
                if (it.conta == null) return null;
                if (c != null && c != it.conta) return null;
                c = it.conta;
            }
            return c;
        }

        static string CategoriaPlanilha(Lancamento l)
        {
            string c = null;
            foreach (var it in l.planilha)
            {
                if (it.categoria.Length == 0) return null;
                if (c != null && Regras.Normalizar(c) != Regras.Normalizar(it.categoria)) return null;
                c = it.categoria;
            }
            return c;
        }

        Plano.Conta CategoriaNoPlano(Lancamento l)
        {
            string cat = CategoriaPlanilha(l);
            if (cat == null || plano == null) return null;
            var tc = TokSet(cat);
            if (tc.Count == 0) return null;
            Plano.Conta melhor = null;
            double mf = 0;
            bool empate = false;
            foreach (var c in plano.lista)
            {
                if (!c.Analitica() || c.Disponivel() || c.codigo == regras.contaVAI) continue;
                if (l.saida && c.Grupo() == '3' || !l.saida && c.Grupo() == '4') continue;
                var tn = TokSet(c.nome);
                if (tn.Count == 0) continue;
                int inter = tc.Count(t => tn.Contains(t));
                if (inter == 0) continue;
                double p = (double)inter / tc.Count, r = (double)inter / tn.Count, f = 2 * p * r / (p + r);
                if (f > mf + 1e-9) { mf = f; melhor = c; empate = false; }
                else if (Math.Abs(f - mf) < 1e-9) empate = true;
            }
            return mf >= 0.75 && !empate ? melhor : null;
        }

        bool ContaOk(int c) { return plano == null || plano.Valida(c); }

        static string Curto(string s) { return s.Length > 60 ? s.Substring(0, 60) + "…" : s; }

        Tuple<Plano.Conta, string> Natureza(string texto, bool saida)
        {
            if (plano == null) return null;
            string n = " " + Regras.Normalizar(texto) + " ";
            foreach (var regra in NATUREZA)
            {
                string sentido = regra[1][0];
                if (sentido == "S" && !saida || sentido == "E" && saida) continue;
                bool casa = regra[0].Any(w => n.Contains(" " + w + " ") || n.Contains(" " + w));
                if (!casa) continue;
                var cand = new List<Plano.Conta>();
                int melhorRank = int.MaxValue;
                foreach (var c in plano.lista)
                {
                    if (!c.Analitica() || c.Disponivel()) continue;
                    if (saida ? c.Grupo() != '4' : c.Grupo() != '3') continue;
                    string nm = Regras.Normalizar(c.nome);
                    for (int ai = 0; ai < regra[2].Length; ai++) if (nm.Contains(regra[2][ai]))
                    {
                        if (ai < melhorRank) { melhorRank = ai; cand.Clear(); }
                        if (ai == melhorRank) cand.Add(c);
                        break;
                    }
                }
                string desc = regra[0][0].ToLowerInvariant();
                if (cand.Count == 0) return null;
                if (cand.Count == 1) return Tuple.Create(cand[0], desc);
                var noRz = cand.Where(c => contasNoRazao.Contains(c.codigo)).ToList();
                return noRz.Count == 1 ? Tuple.Create(noRz[0], desc) : Tuple.Create((Plano.Conta)null, desc);
            }
            return null;
        }

        KeyValuePair<Plano.Conta, bool>? Terceiro(HashSet<string> tk, bool saida)
        {
            if (plano == null || tk.Count == 0) return null;
            Plano.Conta melhor = null, melhorErrado = null;
            double melhorF = 0, melhorFErrado = 0;
            bool melhorNoRazao = false;
            foreach (var o in terceiros)
            {
                var c = o.Key;
                bool ladoCorreto = saida ? c.Grupo() == '2' : c.Grupo() == '1';
                var nome = o.Value;
                int achou = 0;
                bool forte = false;
                foreach (var w in nome) if (Contem(tk, w)) { achou++; if (w.Length >= 4) forte = true; }
                if (achou == 0 || !forte) continue;
                double p = achou / (double)nome.Count, r = achou / (double)tk.Count, f = 2 * p * r / (p + r);
                if (p < 0.5 || f < 0.6) continue;
                if (!ladoCorreto) { if (f > melhorFErrado) { melhorErrado = c; melhorFErrado = f; } continue; }
                bool noRazao = contasNoRazao.Contains(c.codigo);
                if (f > melhorF + 1e-9 || Math.Abs(f - melhorF) < 1e-9 && noRazao && !melhorNoRazao) { melhor = c; melhorF = f; melhorNoRazao = noRazao; }
            }
            if (melhor != null) return new KeyValuePair<Plano.Conta, bool>(melhor, true);
            if (melhorErrado != null) return new KeyValuePair<Plano.Conta, bool>(melhorErrado, false);
            return null;
        }

        /// <summary>
        /// Conta "Adiantamento a sócio NOME" cujo nome aparece no lançamento: o primeiro nome e pelo menos
        /// todas as palavras (ou 3, se o nome for longo). Entre duas, vale a que casou mais palavras; empate = nenhuma.
        /// </summary>
        Plano.Conta Socio(HashSet<string> tk)
        {
            Plano.Conta melhor = null; int melhorN = 0; bool empate = false;
            foreach (var e in socios)
            {
                var nome = e.Value;
                if (!Contem(tk, nome[0])) continue;
                int n = nome.Count(w => Contem(tk, w));
                if (n < Math.Min(3, nome.Count)) continue;
                if (n > melhorN) { melhor = e.Key; melhorN = n; empate = false; }
                else if (n == melhorN && e.Key != melhor) empate = true;
            }
            return empate ? null : melhor;
        }

        static bool Contem(HashSet<string> tk, string w)
        {
            if (tk.Contains(w)) return true;
            if (w.Length >= 5) foreach (var t in tk) if (t.Length >= 5 && (t.StartsWith(w, StringComparison.Ordinal) || w.StartsWith(t, StringComparison.Ordinal))) return true;
            return false;
        }

        /// <summary>Melhor contrapartida no Razão (mesmo sentido; mesmo banco tem preferência).</summary>
        Razao.Ref RazaoRef(List<string> tk, bool saida, int? banco)
        {
            if (refs.Count == 0 || tk.Count == 0) return null;
            double totalX = tk.Sum(t => Peso(t));
            var score = new Dictionary<int, double>();
            var forteAprox = new HashSet<int>();
            foreach (var t in tk)
            {
                List<int> lista;
                if (indice.TryGetValue(t, out lista))
                {
                    foreach (int i in lista) if (refs[i].saidaBanco == saida) score.Somar(i, Peso(t));
                    continue;
                }
                if (t.Length < 5) continue;
                foreach (var e in J.OrdemHashJava(indice.Keys, J.HashJava).Select(k0 => new KeyValuePair<string, List<int>>(k0, indice[k0])))
                {
                    string k = e.Key;
                    if (k.Length < 5 || !(k.StartsWith(t, StringComparison.Ordinal) || t.StartsWith(k, StringComparison.Ordinal))) continue;
                    foreach (int i in e.Value) if (refs[i].saidaBanco == saida) { score.Somar(i, Peso(k)); forteAprox.Add(i); }
                }
            }
            var porConta = new Dictionary<int, object[]>();   // contra -> {Ref, melhorScore, qtd}
            foreach (var e in J.OrdemHashJava(score.Keys, x => x).Select(k0 => new KeyValuePair<int, double>(k0, score[k0])))
            {
                var r = refs[e.Key];
                double totalR = r.tokens.Sum(t => Peso(t));
                bool forte = tk.Any(t => r.tokens.Contains(t) && t.Length >= 4) || forteAprox.Contains(e.Key);
                if (!forte) continue;
                double s = 2 * e.Value / (totalX + totalR);
                double cobertura = e.Value / totalX;
                if (s < 0.5 && cobertura < 0.75) continue;
                if (plano != null && !plano.Valida(r.contra)) continue;
                if (banco != null && r.banco == banco) s += 0.05;
                object[] atual;
                if (!porConta.TryGetValue(r.contra, out atual)) porConta[r.contra] = new object[] { r, s, 1 };
                else { atual[2] = (int)atual[2] + 1; if (s > (double)atual[1]) { atual[0] = r; atual[1] = s; } }
            }
            object[] melhor = null, segundo = null;
            foreach (var o in J.OrdemHashJava(porConta.Keys, x => x).Select(k0 => porConta[k0]))
            {
                if (melhor == null || (double)o[1] > (double)melhor[1] + 1e-9
                    || Math.Abs((double)o[1] - (double)melhor[1]) < 1e-9 && (int)o[2] > (int)melhor[2]) { segundo = melhor; melhor = o; }
                else if (segundo == null || (double)o[1] > (double)segundo[1]) segundo = o;
            }
            if (melhor != null && segundo != null && Math.Abs((double)melhor[1] - (double)segundo[1]) < 1e-9 && (int)melhor[2] == (int)segundo[2]) return null;
            return melhor == null ? null : (Razao.Ref)melhor[0];
        }
    }
}

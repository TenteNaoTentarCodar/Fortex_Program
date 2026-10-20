using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Regras de classificação de um cliente, gravadas em CSV (;):
    ///   #CLIENTE;Nome  #BANCO;596  #VAI;685  #PLANO;caminho  #RAZAO;caminho  #FUNC;caminho
    ///   TEXTO;CONTA;TIPO (E entrada, S saída, A ambos)
    ///   #APRENDIDO;tipo;sentido;chave;conta;vezes
    /// Mesmo arquivo do programa em Java (compatível nos dois sentidos).
    /// </summary>
    public class Regras
    {
        public class Regra
        {
            public string texto; public int conta; public string tipo;
            public Regra(string texto, int conta, string tipo) { this.texto = texto; this.conta = conta; this.tipo = tipo; }
        }

        public string cliente = "";
        public int? contaBanco;
        public int contaVAI = 685;
        public readonly List<string> razoes = new List<string>();
        public string plano = "";
        public string funcionarios = "";
        /// <summary>Folha e adiantamento: relações de líquidos (PDF) e configuração (contas e intervalos de pagamento).</summary>
        public readonly List<string> folhaArquivos = new List<string>();
        public FolhaPagamento.Config folhaCfg;
        public readonly List<Regra> lista = new List<Regra>();

        public const string HIST = "H", DOC = "D", CAT = "C";
        /// <summary>tipo|sentido|chave → (conta → vezes), na ordem em que apareceram.</summary>
        public readonly Dictionary<string, Dictionary<int, int>> aprendido = new Dictionary<string, Dictionary<int, int>>();
        private readonly List<string> ordemAprendido = new List<string>();

        public static string ChaveAprendido(string tipo, bool saida, string chave) { return tipo + "|" + (saida ? "S" : "E") + "|" + chave; }

        private Dictionary<int, int> Mapa(string k)
        {
            Dictionary<int, int> m;
            if (!aprendido.TryGetValue(k, out m)) { m = new Dictionary<int, int>(); aprendido[k] = m; ordemAprendido.Add(k); }
            return m;
        }

        public void Aprender(string tipo, bool saida, string chave, int conta, int peso)
        {
            if (chave == null || chave.Trim().Length == 0) return;
            Mapa(ChaveAprendido(tipo, saida, chave.Trim())).Somar(conta, peso);
        }

        public int? Aprendida(string tipo, bool saida, string chave)
        {
            Dictionary<int, int> m;
            aprendido.TryGetValue(ChaveAprendido(tipo, saida, chave), out m);
            return Dominante(m, 0.6);
        }

        public static int? Dominante(Dictionary<int, int> m, double minimo)
        {
            if (m == null || m.Count == 0) return null;
            int tot = 0, mv = 0; int? melhor = null;
            foreach (var e in m) { tot += e.Value; if (e.Value > mv) { mv = e.Value; melhor = e.Key; } }
            return mv >= tot * minimo ? melhor : null;
        }

        /// <summary>{palavras, conta} dos históricos aprendidos no sentido dado.</summary>
        public List<KeyValuePair<string, int>> HistoricosAprendidos(bool saida)
        {
            var r = new List<KeyValuePair<string, int>>();
            string pre = HIST + "|" + (saida ? "S" : "E") + "|";
            foreach (var k in ordemAprendido)
            {
                if (!k.StartsWith(pre, StringComparison.Ordinal)) continue;
                var c = Dominante(aprendido[k], 0.6);
                if (c != null) r.Add(new KeyValuePair<string, int>(k.Substring(pre.Length), c.Value));
            }
            return r;
        }

        public IEnumerable<string> ChavesAprendidas() { return ordemAprendido; }

        public void Esquecer(string chave) { aprendido.Remove(chave); ordemAprendido.Remove(chave); }

        private static readonly Regex NAO_PERMITIDO = new Regex(@"[^A-Z0-9/.,*\- ]");
        private static readonly Regex ESPACOS = new Regex(@"\s+");

        /// <summary>Maiúsculas, sem acento, só letras/números/ / . , * - e espaço.</summary>
        public static string Normalizar(string s)
        {
            if (s == null) return "";
            string n = J.SemAcento(s).ToUpperInvariant();
            n = NAO_PERMITIDO.Replace(n, " ");
            return ESPACOS.Replace(n, " ").Trim();
        }

        /// <summary>Regra cujo texto aparece no histórico — a de texto mais longo vence.</summary>
        public Regra Procurar(string historico, bool saida)
        {
            string h = Normalizar(historico);
            Regra melhor = null;
            foreach (var r in lista)
            {
                if (r.tipo == "E" && saida) continue;
                if (r.tipo == "S" && !saida) continue;
                string t = Normalizar(r.texto);
                if (t.Length > 0 && h.Contains(t) && (melhor == null || t.Length > Normalizar(melhor.texto).Length)) melhor = r;
            }
            return melhor;
        }

        public static Regras Carregar(string arquivo)
        {
            var rg = new Regras();
            if (!File.Exists(arquivo)) return rg;
            string txt = J.Texto(File.ReadAllBytes(arquivo)).Replace("﻿", "");
            foreach (var linha in Regex.Split(txt, "\r?\n"))
            {
                var p = linha.Split(';');
                if (p.Length < 2 || p[0].Trim().Length == 0) continue;
                string k = p[0].Trim();
                try
                {
                    string ku = k.ToUpperInvariant();
                    if (ku == "#CLIENTE") rg.cliente = p[1].Trim();
                    else if (ku == "#BANCO") rg.contaBanco = p[1].Trim().Length == 0 ? (int?)null : int.Parse(p[1].Trim());
                    else if (ku == "#VAI") rg.contaVAI = int.Parse(p[1].Trim());
                    else if (ku == "#RAZAO") { if (p[1].Trim().Length > 0) rg.razoes.Add(p[1].Trim()); }
                    else if (ku == "#PLANO") rg.plano = p[1].Trim();
                    else if (ku == "#FUNC") rg.funcionarios = p[1].Trim();
                    else if (ku == "#FOLHAARQ") { if (p[1].Trim().Length > 0 && !rg.folhaArquivos.Contains(p[1].Trim())) rg.folhaArquivos.Add(p[1].Trim()); }
                    else if (ku == "#FOLHACFG") rg.folhaCfg = FolhaPagamento.Config.Ler(p, 1);
                    else if (ku == "#APRENDIDO" && p.Length >= 6)
                        rg.Mapa(p[1].Trim() + "|" + p[2].Trim() + "|" + p[3].Trim()).Somar(int.Parse(p[4].Trim()), int.Parse(p[5].Trim()));
                    else if (k.StartsWith("#") || ku == "TEXTO") continue;
                    else
                    {
                        string tipo = p.Length > 2 ? p[2].Trim().ToUpperInvariant() : "A";
                        if (!tipo.Matches("[ESA]")) tipo = "A";
                        rg.lista.Add(new Regra(k, int.Parse(p[1].Trim()), tipo));
                    }
                }
                catch (FormatException) { }
                catch (OverflowException) { }
            }
            return rg;
        }

        public void Salvar(string arquivo)
        {
            var sb = new StringBuilder();
            sb.Append("#CLIENTE;").Append(cliente).Append("\r\n");
            sb.Append("#BANCO;").Append(contaBanco == null ? "" : contaBanco.ToString()).Append("\r\n");
            sb.Append("#VAI;").Append(contaVAI).Append("\r\n");
            sb.Append("#PLANO;").Append(plano).Append("\r\n");
            sb.Append("#FUNC;").Append(funcionarios).Append("\r\n");
            foreach (var r in razoes) sb.Append("#RAZAO;").Append(r).Append("\r\n");
            if (folhaCfg != null) sb.Append("#FOLHACFG;").Append(folhaCfg.Gravar()).Append("\r\n");
            foreach (var f in folhaArquivos) sb.Append("#FOLHAARQ;").Append(f).Append("\r\n");
            sb.Append("TEXTO;CONTA;TIPO\r\n");
            foreach (var r in lista) sb.Append(r.texto.Replace(";", ",")).Append(';').Append(r.conta).Append(';').Append(r.tipo).Append("\r\n");
            foreach (var k0 in ordemAprendido)
            {
                var k = k0.Split(new[] { '|' }, 3);
                foreach (var c in aprendido[k0])
                    sb.Append("#APRENDIDO;").Append(k[0]).Append(';').Append(k[1]).Append(';').Append(k[2].Replace(";", ",")).Append(';')
                      .Append(c.Key).Append(';').Append(c.Value).Append("\r\n");
            }
            var dir = Path.GetDirectoryName(Path.GetFullPath(arquivo));
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(arquivo, sb.ToString(), new UTF8Encoding(true));
        }
    }
}

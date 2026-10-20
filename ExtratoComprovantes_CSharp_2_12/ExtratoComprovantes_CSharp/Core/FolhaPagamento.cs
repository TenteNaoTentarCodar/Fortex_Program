using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Folha de pagamento e adiantamento salarial a partir da "Relação Geral dos Líquidos" (Domínio):
    /// separa salário e adiantamento de salário pelo nome do empregado, pelo líquido e pelos intervalos de pagamento do cliente.
    /// </summary>
    public class FolhaPagamento
    {
        public enum Tipo { Folha, Adiantamento }

        public class Item
        {
            public Tipo tipo;
            public int ano, mes;              // competência
            public string nome = "";
            public List<string> palavras = new List<string>();
            public double valor;
            public bool contribuinte;         // pró-labore (seção "Contribuintes")
            public string arquivo = "";
            public string Competencia() { return mes.ToString("00") + "/" + ano; }
            public string Rotulo() { return (tipo == Tipo.Folha ? (contribuinte ? "Pró-labore " : "Folha ") : "Adiantamento ") + Competencia(); }
        }

        /// <summary>Configuração do cliente (gravada no arquivo de regras).</summary>
        public class Config
        {
            public int? contaFolha, contaAdiantamento, contaProlabore;
            public int adIni = 15, adFim = 25;      // adiantamento: dias do mês da competência
            public int folIni = 1, folFim = 10;     // folha: dias do mês seguinte à competência

            public string Gravar()
            {
                return string.Join(";", new[] { N(contaFolha), N(contaAdiantamento), N(contaProlabore), adIni.ToString(), adFim.ToString(), folIni.ToString(), folFim.ToString() });
            }

            public static Config Ler(string[] p, int ini)
            {
                var c = new Config();
                Func<int, int?> num = i => { int v; return ini + i < p.Length && int.TryParse(p[ini + i].Trim(), out v) ? (int?)v : null; };
                c.contaFolha = num(0); c.contaAdiantamento = num(1); c.contaProlabore = num(2);
                c.adIni = num(3) ?? 15; c.adFim = num(4) ?? 25; c.folIni = num(5) ?? 1; c.folFim = num(6) ?? 10;
                return c;
            }

            static string N(int? v) { return v == null ? "" : v.ToString(); }
        }

        public readonly List<Item> itens = new List<Item>();
        public readonly List<string> arquivos = new List<string>();
        public readonly List<string> ignorados = new List<string>();   // páginas de outros cálculos (13º, férias, rescisão...)
        public Config cfg = new Config();

        static readonly Regex CALCULO = new Regex(@"C[aá]lculo\s*:\s*(.+?)(?:\s+Horas\s*:.*)?$", RegexOptions.IgnoreCase);
        static readonly Regex COMPET = new Regex(@"Compet[eê]ncia\s*:\s*(\d{2})/(\d{4})", RegexOptions.IgnoreCase);
        static readonly Regex LINHA = new Regex(@"^\s*(\d{1,6})\s+(.+?)\s+((?:\d{1,3}(?:\.\d{3})+|\d+),\d{2})\s*$");

        /// <summary>Lê uma "Relação Geral dos Líquidos" (PDF, uma competência por página). Devolve quantos itens leu.</summary>
        public int Carregar(string arq)
        {
            var paginas = TextoPDF.LinhasPorPagina(arq);
            return CarregarTexto(paginas, Path.GetFileName(arq));
        }

        public int CarregarTexto(List<List<string>> paginas, string nomeArquivo)
        {
            int n = 0;
            bool reconhecido = false;
            foreach (var pg in paginas)
            {
                Tipo? tipo = null; int ano = 0, mes = 0; bool contrib = false; string calc = "";
                foreach (var l0 in pg)
                {
                    string l = l0.Trim();
                    var mc = CALCULO.Match(l);
                    if (mc.Success)
                    {
                        calc = mc.Groups[1].Value.Trim();
                        string c = Regras.Normalizar(calc);
                        if (c.StartsWith("FOLHA", StringComparison.Ordinal) && !c.Contains("13")) tipo = Tipo.Folha;
                        else if (c.StartsWith("ADIANTAMENTO", StringComparison.Ordinal) && !c.Contains("13")) tipo = Tipo.Adiantamento;
                        else tipo = null;
                        continue;
                    }
                    var mp = COMPET.Match(l);
                    if (mp.Success) { mes = int.Parse(mp.Groups[1].Value); ano = int.Parse(mp.Groups[2].Value); continue; }
                    string nl = Regras.Normalizar(l);
                    if (nl.StartsWith("RELACAO GERAL DOS LIQUIDOS", StringComparison.Ordinal)) reconhecido = true;
                    if (nl == "EMPREGADOS" || nl == "ESTAGIARIOS") { contrib = false; continue; }
                    if (nl == "CONTRIBUINTES") { contrib = true; continue; }
                    if (nl.StartsWith("EMPREGADOS:", StringComparison.Ordinal) || nl.StartsWith("EMPREGADOS ", StringComparison.Ordinal) && nl.Contains("TOTAL")) continue;
                    if (tipo == null || mes == 0) continue;
                    var m = LINHA.Match(l);
                    if (!m.Success) continue;
                    // nome = palavras sem dígito; a identidade (RG) vem depois e tem dígitos
                    var toks = m.Groups[2].Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    var nome = new List<string>();
                    foreach (var t in toks) { if (t.Any(char.IsDigit)) break; nome.Add(t); }
                    if (nome.Count < 2) continue;
                    var it = new Item
                    {
                        tipo = tipo.Value, ano = ano, mes = mes, nome = string.Join(" ", nome), contribuinte = contrib, arquivo = nomeArquivo,
                        valor = double.Parse(m.Groups[3].Value.Replace(".", "").Replace(',', '.'), CultureInfo.InvariantCulture)
                    };
                    it.palavras = Regex.Split(Regras.Normalizar(it.nome), "[^A-Z]+").Where(w => w.Length >= 2 && w != "DE" && w != "DA" && w != "DO" && w != "DOS" && w != "DAS" && w != "E").ToList();
                    if (it.palavras.Count < 2) continue;
                    // a mesma linha lida duas vezes (arquivo repetido) não entra de novo
                    if (itens.Any(x => x.tipo == it.tipo && x.ano == it.ano && x.mes == it.mes && x.nome == it.nome && Math.Abs(x.valor - it.valor) < 0.005)) continue;
                    itens.Add(it);
                    n++;
                }
                if (tipo == null && calc.Length > 0 && mes > 0) ignorados.Add(calc + " " + mes.ToString("00") + "/" + ano);
            }
            if (!reconhecido && n == 0) throw new LeituraException(nomeArquivo + ": não parece uma \"Relação Geral dos Líquidos\" (Folha Mensal ou Adiantamento).");
            if (!arquivos.Contains(nomeArquivo)) arquivos.Add(nomeArquivo);
            return n;
        }

        public class Resultado
        {
            public Item item;
            public int conta;
            public bool valorConfere;
            public string Motivo()
            {
                return item.Rotulo() + " – " + item.nome + (valorConfere ? ": líquido R$ " + J.Moeda(item.valor) + " confere"
                       : ": valor diferente do líquido (R$ " + J.Moeda(item.valor) + ") — conferir");
            }
        }

        /// <summary>Data de pagamento esperada: adiantamento no mês da competência, folha no mês seguinte.</summary>
        public bool NoIntervalo(Item it, DateTime d)
        {
            int ano = it.ano, mes = it.mes, ini, fim;
            if (it.tipo == Tipo.Adiantamento) { ini = cfg.adIni; fim = cfg.adFim; }
            else { ini = cfg.folIni; fim = cfg.folFim; mes++; if (mes > 12) { mes = 1; ano++; } }
            if (d.Year != ano || d.Month != mes) return false;
            int ult = DateTime.DaysInMonth(ano, mes);
            return d.Day >= Math.Min(ini, ult) && d.Day <= Math.Min(fim, ult);
        }

        int? ContaDo(Item it)
        {
            if (it.tipo == Tipo.Adiantamento) return cfg.contaAdiantamento;
            return it.contribuinte ? cfg.contaProlabore : cfg.contaFolha;
        }

        /// <summary>
        /// Saída do extrato paga a um empregado da relação, dentro do intervalo de pagamento. Prefere o item cujo líquido
        /// confere com o valor; sem valor igual, usa o item do intervalo (marcado para conferir). null = não se aplica.
        /// </summary>
        public Resultado Encontrar(Lancamento l, Func<string, bool> citado)
        {
            if (!l.saida || itens.Count == 0) return null;
            var d = J.Data(l.data);
            if (d == null) return null;
            Resultado melhor = null;
            int melhorPal = 0;
            foreach (var it in itens)
            {
                if (!NoIntervalo(it, d.Value)) continue;
                var conta = ContaDo(it);
                if (conta == null) continue;
                // primeiro nome e todas as palavras do nome (ou 3, em nomes longos)
                if (!citado(it.palavras[0])) continue;
                int n = it.palavras.Count(citado);
                if (n < Math.Min(3, it.palavras.Count)) continue;
                bool igual = Math.Abs(it.valor - l.valor) < 0.005;
                // valor diferente só vale se estiver na mesma faixa (metade a uma vez e meia do líquido): PIX pequenos da semana não entram
                if (!igual && (l.valor < it.valor * 0.5 || l.valor > it.valor * 1.5)) continue;
                if (melhor == null || igual && !melhor.valorConfere || igual == melhor.valorConfere && n > melhorPal)
                {
                    melhor = new Resultado { item = it, conta = conta.Value, valorConfere = igual };
                    melhorPal = n;
                }
            }
            return melhor;
        }
    }
}

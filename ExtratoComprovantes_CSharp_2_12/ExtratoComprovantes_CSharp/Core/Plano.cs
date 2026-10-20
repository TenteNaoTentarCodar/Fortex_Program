using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Plano de Contas (Domínio e similares): Código | Classificação | Descrição [| Tipo].
    /// Analítica: marcada "A" no Plano ou, sem a coluna Tipo, a conta sem nenhuma conta abaixo dela.
    /// </summary>
    public class Plano
    {
        public class Conta
        {
            public readonly int codigo;
            public readonly string classificacao, nome;
            internal char? tipo;
            internal bool? folha;
            internal Conta(int codigo, string classificacao, string nome) { this.codigo = codigo; this.classificacao = classificacao; this.nome = nome; }
            public int Niveis() { return classificacao.Split('.').Length; }
            public bool Analitica()
            {
                if (tipo != null) return tipo == 'A';
                if (folha != null) return folha.Value;
                return Niveis() == 5;
            }
            public char Grupo() { return classificacao.Length == 0 ? '?' : classificacao[0]; }
            int Nivel(int i)
            {
                var p = classificacao.Split('.');
                int v;
                return i < p.Length && int.TryParse(p[i], out v) ? v : -1;
            }
            /// <summary>Disponível: 1.1.1 em qualquer máscara, ou conta do Ativo com nome de banco/caixa/aplicação.</summary>
            public bool Disponivel()
            {
                if (Nivel(0) == 1 && Nivel(1) == 1 && Nivel(2) == 1) return true;
                if (Nivel(0) != 1) return false;
                string n = " " + Regex.Replace(Regras.Normalizar(nome), "[^A-Z0-9/]+", " ") + " ";
                foreach (var x in new[] { "EMPRESTIMO", "FINANCIAMENTO", "CONSORCIO", "CAPITAL SOCIAL", "CLIENTE", "DUPLICATA", "RECEBER",
                    "ADIANTAMENTO", "FORNECEDOR", "TITULO", "CARTAO", "CARTOES", "IMOBILIZ", "INVESTIMENTO EM", "PARTICIPAC" }) if (n.Contains(x)) return false;
                if (n.StartsWith(" BANCO") || n.StartsWith(" BCO ") || n.StartsWith(" CAIXA") || n.Contains("APLICAC")) return true;
                foreach (var b in BANCOS) if (n.Contains(" " + b + " ")) return true;
                return false;
            }
            public override string ToString() { return codigo + " - " + nome; }
        }

        static readonly string[] BANCOS = { "C/C", "CONTA CORRENTE", "CONTA MOVIMENTO", "ITAU", "BRADESCO", "SANTANDER", "BANCO DO BRASIL", "BB", "CEF",
            "CAIXA ECONOMICA", "SICOOB", "SICREDI", "UNICRED", "CRESOL", "NUBANK", "NU PAGAMENTOS", "INTER", "C6", "BTG", "SAFRA", "BANRISUL",
            "ORIGINAL", "SOFISA", "DAYCOVAL", "MERCADO PAGO", "PAGSEGURO", "PAGBANK", "STONE", "CORA", "ASAAS", "XP", "BS2", "SISPRIME",
            "AILOS", "BANESE", "BANESTES", "BRB", "BMG", "PAN", "MODAL", "NEON", "PICPAY", "INFINITEPAY", "EFI", "GERENCIANET", "TOPAZIO", "BANPARA",
            "PAGARME", "PAGAR ME", "CIELO", "GETNET", "IUGU", "PJBANK", "VINDI", "GALAX PAY", "BANCOOB", "UNIBANCO" };

        /// <summary>Contas na ordem do arquivo.</summary>
        public readonly List<Conta> lista = new List<Conta>();
        public readonly Dictionary<int, Conta> contas = new Dictionary<int, Conta>();
        public string arquivo = "";

        void MarcarFolhas()
        {
            var pais = new HashSet<string>();
            foreach (var c in lista)
            {
                string cl = c.classificacao;
                for (int i = cl.LastIndexOf('.'); i > 0; i = i - 1 >= 0 ? cl.LastIndexOf('.', i - 1) : -1) pais.Add(cl.Substring(0, i));
            }
            int comTipo = 0;
            foreach (var c in lista) { c.folha = !pais.Contains(c.classificacao); if (c.tipo != null) comTipo++; }
            if (comTipo < lista.Count / 2) foreach (var c in lista) c.tipo = null;
        }

        public Conta Get(int? codigo) { Conta c; return codigo != null && contas.TryGetValue(codigo.Value, out c) ? c : null; }

        public bool Valida(int? codigo) { var c = Get(codigo); return c != null && c.Analitica(); }

        public static Plano Carregar(string arq)
        {
            var p = new Plano();
            p.arquivo = Path.GetFileName(arq);
            foreach (var cel in LeitorTabular.Linhas(arq))
            {
                Conta c = cel.Length >= 3 ? PorColunas(cel) : null;
                if (c == null) c = PorTexto(string.Join(" ", cel).Trim());
                if (c != null && !p.contas.ContainsKey(c.codigo)) { p.contas[c.codigo] = c; p.lista.Add(c); }
            }
            p.MarcarFolhas();
            if (p.contas.Count == 0) throw new LeituraException(Path.GetFileName(arq) + ": nenhuma conta reconhecida (esperado: código, classificação e descrição).");
            return p;
        }

        static int Letras(string s) { return Regex.Replace(s, "[^A-Za-zÀ-ú]", "").Length; }

        static Conta PorColunas(string[] cel)
        {
            int iClass = -1, iCod = -1;
            for (int i = 0; i < cel.Length; i++)
            {
                string v = cel[i].Trim();
                if (iClass < 0 && v.Matches(@"\d{1,2}(\.\d{1,8}){1,5}")) iClass = i;
                else if (iCod < 0 && v.Matches(@"\d{1,7}")) iCod = i;
            }
            if (iClass < 0 || iCod < 0) return null;
            string nome = "";
            for (int i = 0; i < cel.Length; i++)
            {
                if (i == iClass || i == iCod) continue;
                string v = cel[i].Trim();
                if (Letras(v) >= 3) { nome = v; break; }
            }
            if (nome.Length == 0) return null;
            var c = new Conta(int.Parse(cel[iCod].Trim()), cel[iClass].Trim(), nome);
            for (int i = 0; i < cel.Length; i++)
            {
                if (i == iClass || i == iCod) continue;
                string v = Regras.Normalizar(cel[i]).Trim();
                if (v == "A" || v == "ANALITICA") c.tipo = 'A';
                else if (v == "S" || v == "SINTETICA") c.tipo = 'S';
            }
            return c;
        }

        static readonly Regex DOMINIO_GRAU = new Regex(@"^(?:1\s+)?(\d{1,7})\s+(?:([SA])\s+)?(\d{1,2}(?:\.\d{1,8}){0,5})\s+(.+?)\s+(\d{1,2})$");
        static readonly Regex GERAL = new Regex(@"^(?:(\d{1,7})\s+)?(\d{1,2}(?:\.\d{1,8}){1,5})\s+(?:(\d{1,7})\s+)?(.+)$");
        static readonly Regex TIPO_FIM = new Regex(@"\s+(A|S|ANALITICA|SINTETICA|Analítica|Sintética)$");

        static Conta PorTexto(string l)
        {
            var mg = DOMINIO_GRAU.Match(l.Trim());
            if (mg.Success)
            {
                string cl = mg.Groups[3].Value;
                int grau = int.Parse(mg.Groups[5].Value);
                if (grau == cl.Split('.').Length && Letras(mg.Groups[4].Value) >= 2)
                {
                    var c = new Conta(int.Parse(mg.Groups[1].Value), cl, mg.Groups[4].Value.Trim());
                    c.tipo = mg.Groups[2].Value == "S" ? 'S' : 'A';
                    return c;
                }
            }
            var m = GERAL.Match(l);
            if (!m.Success) return null;
            string cod = m.Groups[1].Success ? m.Groups[1].Value : (m.Groups[3].Success ? m.Groups[3].Value : null);
            if (cod == null) return null;
            string bruto = m.Groups[4].Value.Trim();
            char? tipo = null;
            var mt = TIPO_FIM.Match(bruto);
            if (mt.Success) tipo = char.ToUpperInvariant(mt.Groups[1].Value[0]);
            string nome = Regex.Replace(Regex.Replace(bruto, @"\s+[AS]$", ""), @"\s+(ANALITICA|SINTETICA|Analítica|Sintética)$", "").Trim();
            if (Letras(nome) < 3) return null;
            var c2 = new Conta(int.Parse(cod), m.Groups[2].Value, nome);
            c2.tipo = tipo;
            return c2;
        }

        public List<Conta> ContasDisponivel()
        {
            return lista.Where(c => c.Analitica() && c.Disponivel()).ToList();
        }
    }
}

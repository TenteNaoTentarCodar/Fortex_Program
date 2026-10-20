using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Razão contábil (Domínio e similares) em PDF, XLSX ou CSV: movimentos (conta, data, histórico, contrapartida, D/C, valor)
    /// e, a partir deles, os lançamentos com banco/caixa — a referência para classificar o extrato.
    /// </summary>
    public class Razao
    {
        public class Mov
        {
            public int conta;
            public int? contrapartida;
            public string data = "", historico = "";
            public bool debito;
            public double valor;
        }

        public class Ref
        {
            public string historico, data;
            public int contra;
            public int banco;
            public bool saidaBanco;
            public double valor;
            public HashSet<string> tokens;
        }

        public readonly List<Mov> movs = new List<Mov>();
        public readonly List<string> arquivos = new List<string>();
        public string empresa = "";

        static readonly Regex DATA = new Regex(@"^(\d{2}/\d{2}/(?:\d{4}|\d{2}))\b");
        static readonly Regex DIN = new Regex(@"(?<![\d.,])(\d{1,3}(?:\.\d{3})*|\d+),(\d{2})(?!\d)\s*([DC])?(?![A-Za-z])");
        static readonly Regex CONTA_CAB = new Regex(@"\bCONTA\s*(?:CONTABIL)?\s*:?\s*(\d{1,7})\b");

        public void Carregar(string arq, Plano plano)
        {
            var linhas = LeitorTabular.Linhas(arq);
            int antes = movs.Count;
            if (!PorCabecalho(linhas)) PorTexto(linhas, plano);
            if (movs.Count == antes) throw new LeituraException(Path.GetFileName(arq) + ": nenhum lançamento reconhecido no Razão.");
            arquivos.Add(Path.GetFileName(arq));
            string e = Empresa(linhas);
            if (e.Length > 0) empresa = e;
        }

        static readonly Regex SUFIXO_EMPRESA = new Regex(@"\b(LTDA|S\.?/A\.?|EIRELI|EPP|ME|MEI|SLU|SOCIEDADE|ASSOCIACAO|ASSOCIAÇÃO|COOPERATIVA|CONDOMINIO|CONDOMÍNIO|INSTITUTO|FUNDACAO|FUNDAÇÃO)\b", RegexOptions.IgnoreCase);
        static readonly Regex ROT = new Regex(@"^.*?\b(EMPRESA|RAZ[AÃ]O SOCIAL|NOME EMPRESARIAL|CONTRIBUINTE|CLIENTE)\s*:\s*(.+)$", RegexOptions.IgnoreCase);

        internal static string Empresa(List<string[]> linhas)
        {
            var topo = new List<string>();
            for (int i = 0; i < linhas.Count && topo.Count < 15; i++)
            {
                var sb = new System.Text.StringBuilder();
                foreach (var c in linhas[i]) if (c != null && c.Trim().Length > 0) sb.Append(sb.Length > 0 ? "  " : "").Append(c.Trim());
                if (sb.Length > 0) topo.Add(sb.ToString());
            }
            foreach (var l in topo)
            {
                var m = ROT.Match(l);
                if (m.Success) { string n = LimparNome(m.Groups[2].Value); if (n.Length >= 3) return n; }
            }
            for (int i = 1; i < topo.Count; i++)
            {
                string n = Regras.Normalizar(topo[i]);
                if (n.StartsWith("CNPJ") || n.StartsWith("C.N.P.J") || n.StartsWith("CPF"))
                {
                    string nome = LimparNome(topo[i - 1]);
                    if (nome.Length >= 3 && !EhCabecalho(Regras.Normalizar(nome))) return nome;
                }
            }
            foreach (var l in topo)
            {
                string nome = LimparNome(l);
                string nn = Regras.Normalizar(l);
                if (nn.StartsWith("CONTA") || nn.Contains("CONTA:") || EhCabecalho(nn) || nome.Contains(":")) continue;
                if (SUFIXO_EMPRESA.IsMatch(nome) && !DATA.IsMatch(nome)) return nome;
            }
            return "";
        }

        static string LimparNome(string s)
        {
            s = Regex.Replace(s, @"\s+(CNPJ|C\.N\.P\.J|CPF|INSCRI|P[AÁ]GINA|FOLHA|FL\.|PER[IÍ]ODO|DATA|EMISS[AÃ]O|HORA|CONTA\s*:|RAZ[AÃ]O\b|LIVRO).*$", "", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^\s*\d{1,6}\s*[-–]\s*", "");
            s = Regex.Replace(s, @"\s*\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}.*$", "");
            return Regex.Replace(Regex.Replace(s, @"[\s:;|-]+$", ""), @"\s{2,}", " ").Trim();
        }

        bool PorCabecalho(List<string[]> linhas)
        {
            int iData = -1, iHist = -1, iCp = -1, iDeb = -1, iCred = -1, iConta = -1, ini = -1;
            for (int r = 0; r < linhas.Count && ini < 0; r++)
            {
                var c = linhas[r];
                if (c.Length < 3) continue;
                int h = -1, d = -1, cr = -1;
                for (int i = 0; i < c.Length; i++)
                {
                    string n = Regras.Normalizar(c[i]);
                    if (n.StartsWith("HISTORICO")) h = i;
                    else if (n.StartsWith("DEBITO")) d = i;
                    else if (n.StartsWith("CREDITO")) cr = i;
                }
                if (h >= 0 && (d >= 0 || cr >= 0))
                {
                    ini = r + 1; iHist = h; iDeb = d; iCred = cr;
                    for (int i = 0; i < c.Length; i++)
                    {
                        string n = Regras.Normalizar(c[i]);
                        if (n.StartsWith("DATA")) iData = i;
                        else if (n.Contains("CONTRA") || n.StartsWith("C.PARTIDA") || n.StartsWith("C PARTIDA") || n.StartsWith("CPARTIDA")) iCp = i;
                        else if (n.StartsWith("CONTA") || n.StartsWith("COD") && n.Contains("CONTA")) iConta = i;
                    }
                }
            }
            if (ini < 0 || iData < 0) return false;
            int? contaAtual = null;
            for (int r = 0; r < ini; r++)
            {
                var mc0 = CONTA_CAB.Match(Regras.Normalizar(string.Join(" ", linhas[r])));
                if (mc0.Success) contaAtual = int.Parse(mc0.Groups[1].Value);
            }
            for (int r = ini; r < linhas.Count; r++)
            {
                var c = linhas[r];
                string todo = Regras.Normalizar(string.Join(" ", c));
                var mc = CONTA_CAB.Match(todo);
                string dt = Cel(c, iData);
                if (!dt.Matches(@"\d{2}/\d{2}/\d{2,4}.*"))
                {
                    if (mc.Success) contaAtual = int.Parse(mc.Groups[1].Value);
                    else if (iConta < 0) { var cab = ContaDeCabecalho(string.Join(" ", c)); if (cab != null) contaAtual = cab; }
                    continue;
                }
                int? conta = iConta >= 0 && Cel(c, iConta).Matches(@"\d{1,7}") ? int.Parse(Cel(c, iConta)) : contaAtual;
                if (conta == null) continue;
                double deb = Num(Cel(c, iDeb)), cred = Num(Cel(c, iCred));
                if (deb == 0 && cred == 0) continue;
                var m = new Mov();
                m.conta = conta.Value;
                m.data = dt.Substring(0, Math.Min(10, dt.Length));
                m.historico = Cel(c, iHist);
                string cp = Regex.Replace(Cel(c, iCp), @"\D", "");
                m.contrapartida = cp.Length == 0 || cp.Length > 7 ? (int?)null : int.Parse(cp);
                m.debito = deb != 0;
                m.valor = deb != 0 ? deb : cred;
                movs.Add(m);
            }
            return true;
        }

        static string Cel(string[] c, int i) { return i >= 0 && i < c.Length && c[i] != null ? c[i].Trim() : ""; }

        static readonly Regex PERIODO = new Regex(@"^\d{2}/\d{2}/\d{2,4}\s+(?:a|à|ate|até|ao|to)\s+\d{2}/\d{2}/\d{2,4}", RegexOptions.IgnoreCase);
        static readonly Regex CP_FIM = new Regex(@"(?:^|\s)(\d{1,7})$");

        void PorTexto(List<string[]> linhas, Plano plano)
        {
            int? contaAtual = null;
            double? saldoAnt = null;
            Mov ultimo = null;
            int cont = 0;
            foreach (var cel in linhas)
            {
                string l = Regex.Replace(string.Join(" ", cel), @"\s{2,}", " ").Trim();
                string n = Regras.Normalizar(l);
                var md = DATA.Match(l);
                // "01/05/2026 até 31/05/2026 ..." no histórico quebrado é texto do lançamento anterior, não lançamento
                if (md.Success && PERIODO.IsMatch(l)) md = Match.Empty;
                if (!md.Success)
                {
                    var mc = CONTA_CAB.Match(n);
                    int? cab = mc.Success ? int.Parse(mc.Groups[1].Value) : (DIN.IsMatch(l) ? (int?)null : ContaDeCabecalho(l));
                    if (cab != null && (plano == null || plano.Get(cab) != null || mc.Success)) { if (cab != contaAtual) saldoAnt = null; contaAtual = cab; ultimo = null; continue; }   // mesma conta na página seguinte: a cadeia de saldo continua
                    if (n.Contains("SALDO") && n.Contains("ANTERIOR"))
                    {
                        double? s = null;
                        foreach (Match mv0 in DIN.Matches(l)) s = Assinado(mv0);
                        if (s != null) saldoAnt = s;
                        continue;
                    }
                    if (ultimo != null && cont < 2 && !DIN.IsMatch(l) && !EhCabecalho(n)) { ultimo.historico += " " + l; cont++; }
                    else ultimo = null;
                    continue;
                }
                if (contaAtual == null) continue;
                string data = md.Groups[1].Value;
                if (data.Length == 8) data = data.Substring(0, 6) + "20" + data.Substring(6);
                string resto = l.Substring(md.Index + md.Length).Trim();
                var vals = new List<double[]>();
                foreach (Match mv in DIN.Matches(resto))
                {
                    double v = J.ParseD(mv.Groups[1].Value.Replace(".", "") + "." + mv.Groups[2].Value);
                    vals.Add(new double[] { v, mv.Index, mv.Groups[3].Value == "D" ? 1 : (mv.Groups[3].Value == "C" ? -1 : 0) });
                }
                if (vals.Count == 0) continue;
                string antes = resto.Substring(0, (int)vals[0][1]).Trim();
                antes = new Regex(@"^\d{1,9}\s+").Replace(antes, "", 1);
                int? cp = null;
                var mcp = CP_FIM.Match(antes);
                if (mcp.Success)
                {
                    int cod = int.Parse(mcp.Groups[1].Value);
                    if (plano == null || plano.Get(cod) != null) { cp = cod; antes = antes.Substring(0, mcp.Index).Trim(); }
                }
                var m = new Mov();
                m.conta = contaAtual.Value;
                m.data = data;
                m.historico = antes;
                m.contrapartida = cp;
                m.valor = vals[0][0];
                double? saldo = null;
                if (vals.Count >= 2) { var s = vals[vals.Count - 1]; saldo = s[2] < 0 ? -s[0] : s[0]; }
                // saldo zerado desenhado antes do valor (histórico sobreposto): "...PR0,00 159,39"
                if (vals.Count == 2 && vals[0][0] == 0 && vals[1][0] > 0 && vals[1][2] == 0) { m.valor = vals[1][0]; saldo = 0; }
                bool? deb = null;
                if (vals.Count >= 3 && vals[0][0] == 0) { m.valor = vals[1][0]; deb = false; }
                else if (vals.Count >= 3 && vals[1][0] == 0) deb = true;
                if (deb == null && saldo != null && saldoAnt != null)
                {
                    double dif = saldo.Value - saldoAnt.Value;
                    if (Math.Abs(dif - m.valor) < 0.011) deb = true;
                    else if (Math.Abs(dif + m.valor) < 0.011) deb = false;
                }
                if (deb == null && vals.Count == 1 && vals[0][2] != 0) deb = vals[0][2] > 0;
                if (deb == null)
                {
                    string h = Regras.Normalizar(antes);
                    bool contaBanco = plano != null && plano.Get(contaAtual) != null && plano.Get(contaAtual).Disponivel();
                    bool recebe = h.Contains("RECEB") || h.Contains("DEPOSITO") || h.Contains("CREDITO") || h.Contains("RESGATE");
                    deb = contaBanco == recebe;
                }
                m.debito = deb.Value;
                if (saldo != null) saldoAnt = saldo;
                else if (saldoAnt != null) saldoAnt += deb.Value ? m.valor : -m.valor;
                if (m.valor > 0) { movs.Add(m); ultimo = m; cont = 0; }
            }
        }

        static double? Assinado(Match mv)
        {
            double v = J.ParseD(mv.Groups[1].Value.Replace(".", "") + "." + mv.Groups[2].Value);
            return mv.Groups[3].Value == "C" ? -v : v;
        }

        static readonly Regex CAB1 = new Regex(@"^(\d{1,7})\s*[-–]?\s*(?:\d{1,2}(?:\.\d{1,8}){1,5})?\s*[-–]?\s*[A-Za-zÀ-ú].{2,}$");
        static readonly Regex CAB2 = new Regex(@"\d(?:\.\d+){2,}");

        static int? ContaDeCabecalho(string l)
        {
            var m = CAB1.Match(l.Trim());
            if (m.Success && CAB2.IsMatch(l)) return int.Parse(m.Groups[1].Value);
            return null;
        }

        static bool EhCabecalho(string n)
        {
            return n.StartsWith("DATA") || n.StartsWith("PAGINA") || n.Contains("FOLHA") || n.StartsWith("EMPRESA") || n.StartsWith("CNPJ")
                   || n.StartsWith("PERIODO") || n.StartsWith("RAZAO") || n.StartsWith("TOTAL") || n.StartsWith("SALDO") || n.Contains("SISTEMA")
                   || n.StartsWith("EMISSAO") || n.StartsWith("HISTORICO");
        }

        internal static double Num(string s)
        {
            if (s == null) return 0;
            s = Regex.Replace(s.Replace("R$", ""), "[DC]$", "").Trim();
            if (s.Length == 0) return 0;
            if (s.Contains(",")) s = s.Replace(".", "").Replace(',', '.');
            double v;
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? Math.Abs(v) : 0;
        }

        /// <summary>Lançamentos com banco/caixa. Linhas sem banco (NF) são ignoradas.</summary>
        public List<Ref> Referencias(Plano plano, ICollection<int> contasBanco)
        {
            var r = new List<Ref>();
            var vistos = new HashSet<string>();
            foreach (var m in movs)
            {
                if (m.contrapartida == null || m.contrapartida == m.conta) continue;
                bool contaEhBanco = EhBanco(m.conta, plano, contasBanco), cpEhBanco = EhBanco(m.contrapartida.Value, plano, contasBanco);
                var x = new Ref { historico = m.historico, data = m.data, valor = m.valor };
                if (contaEhBanco && !cpEhBanco) { x.contra = m.contrapartida.Value; x.banco = m.conta; x.saidaBanco = !m.debito; }
                else if (!contaEhBanco && cpEhBanco) { x.contra = m.conta; x.banco = m.contrapartida.Value; x.saidaBanco = m.debito; }
                else continue;
                string chave = x.data + "|" + x.valor.ToString("R", CultureInfo.InvariantCulture) + "|" + x.contra + "|" + x.saidaBanco + "|" + Regras.Normalizar(x.historico);
                if (!vistos.Add(chave)) continue;
                r.Add(x);
            }
            return r;
        }

        static bool EhBanco(int conta, Plano plano, ICollection<int> contasBanco)
        {
            if (contasBanco != null && contasBanco.Contains(conta)) return true;
            var c = plano == null ? null : plano.Get(conta);
            return c != null && c.Disponivel();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Arquivo de progresso (.ecprog, texto UTF-8) — mesmo formato do programa em Java: lançamentos da tela e já exportados,
    /// com contas, observação, complemento e marca de exportação. Plano/Razão anexados (#ANEXO) também são lidos.
    /// </summary>
    public class Progresso
    {
        public string cliente = "", regras = "", salvoEm = "";
        public int? contaBanco;
        public int contaVAI = 685;
        public readonly Dictionary<string, int?> bancoPorArquivo = new Dictionary<string, int?>();
        public readonly List<Lancamento> lancamentos = new List<Lancamento>();
        /// <summary>Linhas #ANEXO (tipo;original;copia;paginas) guardadas como vieram.</summary>
        public readonly List<string[]> anexos = new List<string[]>();

        public static string PastaAnexos(string p) { return Path.Combine(Path.GetDirectoryName(Path.GetFullPath(p)), "anexos_progresso"); }

        public static void Gravar(string p, Progresso pr, Dictionary<Lancamento, string[]> extra)
        {
            var sb = new StringBuilder();
            sb.Append("#PROGRESSO-EXTRATO;1\n");
            sb.Append("#SALVO;").Append(Esc(pr.salvoEm)).Append('\n');
            sb.Append("#CLIENTE;").Append(Esc(pr.cliente)).Append('\n');
            sb.Append("#REGRAS;").Append(Esc(pr.regras)).Append('\n');
            sb.Append("#BANCO;").Append(pr.contaBanco == null ? "" : pr.contaBanco.ToString()).Append('\n');
            sb.Append("#VAI;").Append(pr.contaVAI).Append('\n');
            foreach (var e in pr.bancoPorArquivo)
                sb.Append("#BANCOARQ;").Append(Esc(e.Key)).Append(';').Append(e.Value == null ? "" : e.Value.ToString()).Append('\n');
            foreach (var a in pr.anexos) sb.Append("#ANEXO;").Append(string.Join(";", a.Select(Esc))).Append('\n');
            sb.Append("data;historico;valor;saida;debito;credito;obs;manual;conferirSinal;origem;contaBanco;descartada;exportadoEm;complemento;busca;seq\n");
            foreach (var l in pr.lancamentos)
            {
                string[] x;
                if (!extra.TryGetValue(l, out x)) x = new[] { l.complementoSalvo, l.buscaSalva };
                sb.Append(Esc(l.data)).Append(';').Append(Esc(l.historico)).Append(';').Append(DoubleJava(l.valor)).Append(';').Append(l.saida ? 1 : 0).Append(';')
                  .Append(l.debito).Append(';').Append(l.credito).Append(';').Append(Esc(l.obs)).Append(';')
                  .Append(l.manual ? 1 : 0).Append(';').Append(l.conferirSinal ? 1 : 0).Append(';').Append(Esc(l.origem)).Append(';')
                  .Append(l.contaBanco).Append(';').Append(l.descartada ? 1 : 0).Append(';').Append(Esc(l.exportadoEm)).Append(';')
                  .Append(Esc(x[0])).Append(';').Append(Esc(x[1])).Append(';').Append(l.seq).Append('\n');
            }
            string tmp = p + ".tmp";
            File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(false));
            if (File.Exists(p)) File.Delete(p);
            File.Move(tmp, p);
        }

        public static Progresso Ler(string p)
        {
            var ls = File.ReadAllLines(p, Encoding.UTF8);
            if (ls.Length == 0 || !ls[0].StartsWith("#PROGRESSO-EXTRATO")) throw new LeituraException(Path.GetFileName(p) + " não é um arquivo de progresso do programa.");
            var pr = new Progresso();
            bool dados = false;
            foreach (var s in ls.Skip(1))
            {
                if (s.Length == 0) continue;
                var c = s.Split(';');
                if (!dados)
                {
                    switch (c[0])
                    {
                        case "#SALVO": pr.salvoEm = Des(c[1]); continue;
                        case "#CLIENTE": pr.cliente = Des(c[1]); continue;
                        case "#REGRAS": pr.regras = Des(c[1]); continue;
                        case "#BANCO": pr.contaBanco = J.Int(c[1]); continue;
                        case "#VAI": pr.contaVAI = int.Parse(c[1]); continue;
                        case "#BANCOARQ": pr.bancoPorArquivo[Des(c[1])] = c.Length > 2 ? J.Int(c[2]) : null; continue;
                        case "#ANEXO": pr.anexos.Add(c.Skip(1).Select(Des).ToArray()); continue;
                    }
                    if (c[0] == "data") dados = true;
                    continue;
                }
                if (c.Length < 15) continue;
                var l = new Lancamento(Des(c[0]), Des(c[1]), J.ParseD(c[2]), c[3] == "1");
                l.debito = J.Int(c[4]);
                l.credito = J.Int(c[5]);
                l.obs = Des(c[6]);
                l.manual = c[7] == "1";
                l.conferirSinal = c[8] == "1";
                l.origem = Des(c[9]);
                l.contaBanco = J.Int(c[10]);
                l.descartada = c[11] == "1";
                l.exportadoEm = Des(c[12]);
                string comp = Des(c[13]), busca = Des(c[14]);
                l.complementoSalvo = comp.Length == 0 ? null : comp;
                l.buscaSalva = busca.Length == 0 ? null : busca;
                long sq;
                l.seq = c.Length > 15 && long.TryParse(c[15], out sq) ? sq : pr.lancamentos.Count;
                pr.lancamentos.Add(l);
            }
            return pr;
        }

        /// <summary>Como o Double.toString do Java ("10000.0", "1932.5").</summary>
        internal static string DoubleJava(double v)
        {
            string r = v.ToString("R", CultureInfo.InvariantCulture);
            return r.IndexOf('.') < 0 && r.IndexOf('E') < 0 && r.IndexOf("Infinity", StringComparison.Ordinal) < 0 && r != "NaN" ? r + ".0" : r;
        }

        public static string Esc(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\").Replace(";", "\\s").Replace("\r", "").Replace("\n", "\\n");
        }

        public static string Des(string s)
        {
            var o = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch == '\\' && i + 1 < s.Length)
                {
                    char n = s[++i];
                    o.Append(n == 's' ? ';' : n == 'n' ? '\n' : n);
                }
                else o.Append(ch);
            }
            return o.ToString();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>Datas e valores como a planilha financeira grava (usados também pelos leitores de extrato).</summary>
    public static partial class LeitorPlanilhaFinanceira
    {
        static readonly Regex DATA = new Regex(@"\b([0-9]{2})[/.-]([0-9]{2})[/.-]([0-9]{4}|[0-9]{2})\b|\b([0-9]{4})-([0-9]{2})-([0-9]{2})\b");

        public static string Data(string s, bool xlsx)
        {
            s = (s ?? "").Trim();
            if (s.Length == 0) return "";
            if (xlsx && s.Matches(@"[0-9]{5}(\.[0-9]+)?"))
            {
                var d = new DateTime(1899, 12, 30).AddDays((long)J.ParseD(s));
                return d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            }
            var m = DATA.Match(s);
            if (!m.Success) return "";
            int dd, mm, aa;
            if (m.Groups[4].Success) { aa = int.Parse(m.Groups[4].Value); mm = int.Parse(m.Groups[5].Value); dd = int.Parse(m.Groups[6].Value); }
            else { dd = int.Parse(m.Groups[1].Value); mm = int.Parse(m.Groups[2].Value); aa = int.Parse(m.Groups[3].Value); if (aa < 100) aa += 2000; }
            if (dd < 1 || dd > 31 || mm < 1 || mm > 12) return "";
            return dd.ToString("00") + "/" + mm.ToString("00") + "/" + aa.ToString("0000");
        }

        static readonly Regex LIMPA_VALOR = new Regex(@"R\$|[()\s+\-]|\b[DC]$", RegexOptions.IgnoreCase);

        public static double? Valor(string s, bool xlsx)
        {
            s = (s ?? "").Trim().Replace(' ', ' ');
            if (s.Length == 0) return null;
            bool neg = s.StartsWith("(") && s.EndsWith(")") || s.Matches(@"^(R\$\s*)?-.*|.*-$|.*\bD$");
            string t = LIMPA_VALOR.Replace(s, "");
            double v;
            var inv = CultureInfo.InvariantCulture;
            if (t.Matches(@"[0-9]{1,3}(\.[0-9]{3})*,[0-9]{1,2}|[0-9]+,[0-9]{1,2}")) v = double.Parse(t.Replace(".", "").Replace(',', '.'), inv);
            else if (t.Matches(@"[0-9]{1,3}(,[0-9]{3})*\.[0-9]{1,2}")) v = double.Parse(t.Replace(",", ""), inv);
            else if (xlsx && t.Matches(@"[0-9]+(\.[0-9]+)?(E-?[0-9]+)?")) { if (!double.TryParse(t, NumberStyles.Float, inv, out v)) return null; }
            else if (t.Matches(@"[0-9]{1,3}(\.[0-9]{3})+")) v = double.Parse(t.Replace(".", ""), inv);
            else if (t.Matches(@"[0-9]+(\.[0-9]{1,2})?")) v = double.Parse(t, inv);
            else return null;
            return neg ? -v : v;
        }
    }

    /// <summary>
    /// Extrato bancário em planilha (xlsx/xls/ods/csv/html). Com cabeçalho: Data + Histórico + (Valor com sinal | Crédito e Débito | Valor + D/C).
    /// Sem cabeçalho reconhecido: cada linha é lida como linha de texto de extrato (mesma leitura do PDF).
    /// </summary>
    public static class LeitorExtratoExcel
    {
        public class Mapa
        {
            public int data = -1, hist = -1, hist2 = -1, doc = -1, valor = -1, cred = -1, deb = -1, dc = -1, saldo = -1, credBruto = -1, debBruto = -1;
            public bool histForte;
            public List<int> extras = new List<int>();
            public bool Ok() { return data >= 0 && (valor >= 0 || cred >= 0 || deb >= 0); }
        }

        public static bool EhPlanilha(string f) { return f.StartsWith("Excel (") || f.StartsWith("Excel 97") || f.StartsWith("OpenDocument"); }

        public static List<Lancamento> Ler(string arq, List<string[]> rows, bool juntarContinuacao, LeitorExtrato.Resultado res)
        {
            string nome = System.IO.Path.GetFileName(arq);
            string f = LeitorTabular.ultimoFormato ?? "";
            bool xlsx = EhPlanilha(f);
            var r = PorColunas(rows, nome, xlsx, res);
            if (r.Count == 0) { r = PorTexto(rows, nome, xlsx, juntarContinuacao); res.formato = f + " sem cabeçalho (lido como texto)"; }
            else res.formato = f + " por colunas";
            if (r.Count == 0) throw new LeituraException(nome + ": nenhum lançamento reconhecido (" + f + "). Use \"Mapear colunas\" para ensinar o layout deste banco.");
            return r;
        }

        const string DATA_H = "(DATA|DT|DIA|DATE|FECHA|FEC|DATUM|BUCHUNGSTAG|BUCHUNGSDATUM|BUCHUNG|WERTSTELLUNG|VALUTA|DATE OPERATION|DATE DE L OPERATION|DATE COMPTABLE|POSTED|POSTING|TRANSACTION DATE|BOOKING DATE|POSTING DATE|DATE POSTED|TRANS DATE|TRAN DATE|VALUE DATE|DATE DE VALEUR|DATA VALUTA|DATA CONTABILE|DATA OPERAZIONE|FECHA OPERACION|FECHA VALOR|F VALOR|F OPERACION|BOEKDATUM|RENTEDATUM|MOVIMENTO|DATA MOVIMENTO|DATA OPERACJI|DATA KSIEGOWANIA|BOKFORINGSDAG|BOKFORINGSDATUM|TRANSAKTIONSDAG|DATO|BOKFORT|TARIH|ISLEM TARIHI|TANGGAL|TGL|NGAY|DATUM ZAUCTOVANI)";
        const string VALOR_DATA = ".*(VALUTA|VALOR|VALUE|WERTSTELLUNG|VALEUR|RENTE).*";
        const string CRED_H = "(VALOR |VLR |IMPORTE |MONTANT |BETRAG )?(CREDITOS?|ENTRADAS?|RECEITAS?|CREDIT|CREDITS|CREDIT AMOUNT|DEPOSITS?|DEPOSIT AMOUNT|PAID IN|MONEY IN|IN|INFLOW|RECEIPTS?|ABONOS?|HABER|INGRESOS?|HABEN|GUTSCHRIFT|EINGANG|ENTRATE|AVERE|ACCREDITI|BIJ|ENCAISSEMENT|RECETTES|CREDITO|UZNANIA|INSATTNING|INDSAT|INN|ALACAK|KREDIT|MASUK|PRIJEM)";
        const string DEB_H = "(VALOR |VLR |IMPORTE |MONTANT |BETRAG )?(DEBITOS?|SAIDAS?|DESPESAS?|DEBIT|DEBITS|DEBIT AMOUNT|WITHDRAWALS?|WITHDRAWAL AMOUNT|PAID OUT|MONEY OUT|OUT|OUTFLOW|PAYMENTS?|CARGOS?|DEBE|EGRESOS?|RETIROS?|SOLL|LASTSCHRIFT|BELASTUNG|AUSGANG|USCITE|DARE|ADDEBITI|AF|DECAISSEMENT|DEPENSES|DEBITO|OBCIAZENIA|UTTAG|HAEVET|UT|BORC|KELUAR|VYDAJ)";
        const string VALOR_H = "(VALOR|VLR|VL|VALOR LANCAMENTO|VALOR DO LANCAMENTO|VALOR MOVIMENTO|VALOR TRANSACAO|MONTANTE|QUANTIA|AMOUNT|TRANSACTION AMOUNT|AMT|VALUE|NET AMOUNT|IMPORTE|IMPORTE EUR|MONTO|CANTIDAD|BETRAG|UMSATZ|BETRAG EUR|MONTANT|IMPORTO|BEDRAG|SUM|SUMA|TOTAL MOVIMENTO|KWOTA|KWOTA OPERACJI|BELOPP|BELOB|BELOP|CASTKA|TUTAR|ISLEM TUTARI|JUMLAH|NOMINAL|SO TIEN|OSSZEG|SUMA OPERACJI)";
        const string HIST_FORTE = ".*(HISTOR|DESCRI|DESCRIPTION|DETAILS|NARRATIVE|PARTICULARS|CONCEPTO|DETALLE|VERWENDUNGSZWECK|BUCHUNGSTEXT|LIBELLE|DESCRIZIONE|CAUSALE|OMSCHRIJVING|MEDEDELINGEN|OPIS|TYTUL|TYTULEM|TEXT|TEKST|BESKRIVELSE|BESKRIVNING|POPIS|ACIKLAMA|KETERANGAN|URAIAN|NOI DUNG|DIEN GIAI|KOZLEMENY|ZPRAVA).*";
        const string HIST_FRACO = ".*(LANCAMENTO|MOVIMENTACAO|TRANSACAO|OPERACAO|DETALHE|COMPLEMENTO|FAVORECIDO|ESTABELECIMENTO|MEMO|INFORMAC|PAYEE|MERCHANT|COUNTERPARTY|BENEFICIARY|NAME|NOMBRE|BEGUNSTIGTER|AUFTRAGGEBER|EMPFANGER|ZAHLUNGSPFLICHTIGER|BENEFICIAIRE|TIERS|BENEFICIARIO|ORDINANTE|TEGENREKENING NAAM|NAAM|TRANSACTION|TYPE DESCRIPTION|REFERENCE|REFERENCIA|RIFERIMENTO).*";

        public static Mapa Cabecalho(string[] cel)
        {
            var m = new Mapa();
            bool dataEhValor = false;
            for (int i = 0; i < cel.Length; i++)
            {
                string h = Formatos.Normalizar(cel[i]).ReplaceAll(" (R|EUR|USD|BRL|GBP|CHF|CAD|AUD|MXN|ARS|CLP|COP|PEN)$", "").ReplaceAll("^(R|EUR|USD) ", "").Trim();
                if (h.Length == 0 || h.Length > 40) continue;
                if (h.Matches(".*(SALDO|BALANCE|SOLDE|KONTOSTAND|SALDO EUR).*")) { if (m.saldo < 0 || h.Matches(".*(RUNNING|CORRIDO|ATUAL|APOS).*")) m.saldo = i; continue; }
                if (h.Matches(DATA_H + "( .*)?") || h.Matches(@".*\b(DATE|FECHA|DATUM)\b.*"))
                {
                    bool ev = h.Matches(VALOR_DATA);
                    if (m.data < 0 || dataEhValor && !ev || h.Matches(".*(LANC|MOVIM|TRANSAC|OPERAC|POSTED|POSTING|BOOKING|BUCHUNG|CONTAB).*") && !ev) { m.data = i; dataEhValor = ev; }
                    continue;
                }
                if (h.Matches("(D/C|C/D|DC|CD|TIPO|NATUREZA|SINAL|E/S|CR/DR|DR/CR|DEBIT/CREDIT|CREDIT/DEBIT|D/H|S/H|SOLL/HABEN|TYPE|SENS|DEBIT CREDIT|CREDIT DEBIT|TRANSACTION TYPE|TIPO MOVIMENTO|TIPO LANCAMENTO|CARGO/ABONO|AF BIJ|AF/BIJ)")) { m.dc = i; continue; }
                if (h.Contains("TAXA") || h.Contains("ESTORN") || h.Matches(@".*\b(FEE|FEES|COMISION)\b.*")) continue;
                if (h.Matches("(VALOR )?(CREDITOS?|ENTRADAS?) LIQUID[AO]S?")) { m.cred = i; continue; }
                if (h.Matches("(VALOR )?(DEBITOS?|SAIDAS?) LIQUID[AO]S?")) { m.deb = i; continue; }
                if (h.Matches("(VALOR )?(CREDITOS?|ENTRADAS?) BRUT[AO]S?")) { m.credBruto = i; continue; }
                if (h.Matches("(VALOR )?(DEBITOS?|SAIDAS?) BRUT[AO]S?")) { m.debBruto = i; continue; }
                if (h.Matches(CRED_H)) { m.cred = i; continue; }
                if (h.Matches(DEB_H)) { m.deb = i; continue; }
                if (h.Matches("TIPO D[AOE] (OPERACAO|TRANSACAO|MOVIMENTACAO)|METODO( DE PAGAMENTO)?|FORMA DE PAGAMENTO|CATEGORIA|CATEGORY|CATEGORIA OPERACION")) { m.extras.Add(i); continue; }
                if (h.Matches("ID( D[AOE] .*)?|TRANSACTION ID|FITID")) { if (m.doc < 0) m.doc = i; continue; }
                if (h.Matches(VALOR_H)) { m.valor = i; continue; }
                if (h.Matches(".*(DOCUMENTO|DOC|N DOC|NUMERO|NR|CHECK|CHEQUE|CHQ|BELEG|REF NO|REFERENCE NO|NUM).*") && !h.Matches(HIST_FORTE)) { if (m.doc < 0) m.doc = i; continue; }
                if (h.Matches(HIST_FORTE))
                {
                    if (m.hist >= 0 && !m.histForte) m.hist2 = m.hist;
                    if (m.hist < 0 || !m.histForte) { m.hist = i; m.histForte = true; } else if (m.hist2 < 0) m.hist2 = i;
                    continue;
                }
                if (h.Matches(HIST_FRACO))
                {
                    if (m.hist < 0) m.hist = i; else if (m.hist2 < 0) m.hist2 = i; else m.extras.Add(i);
                }
            }
            if (m.cred < 0) m.cred = m.credBruto;
            if (m.deb < 0) m.deb = m.debBruto;
            return m;
        }

        static List<Lancamento> PorColunas(List<string[]> rows, string origem, bool xlsx, LeitorExtrato.Resultado res)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                var h = Cabecalho(rows[i]);
                if (h.Ok() && SemDado(rows[i])) return Processar(rows, i + 1, h, -1, origem, xlsx, res, true);
            }
            return new List<Lancamento>();
        }

        static readonly string SALDO_LINHA = "(SALDO|S A L D O|TOTAL|SUBTOTAL|OPENING BALANCE|CLOSING BALANCE|BALANCE|SALDO INICIAL|SALDO FINAL|SOLDE|KONTOSTAND|ANFANGSSALDO|ENDSALDO|BALANCE BROUGHT FORWARD|BALANCE CARRIED FORWARD|BROUGHT FORWARD|CARRIED FORWARD).*";

        /// <summary>Lê as linhas a partir de ini com o mapa de colunas m. ordem &lt; 0 = detectar.</summary>
        public static List<Lancamento> Processar(List<string[]> rows, int ini, Mapa m, int ordem, string origem, bool xlsx, LeitorExtrato.Resultado res, bool recabecalho)
        {
            if (ordem < 0)
            {
                var datas = new List<string>();
                for (int i = ini; i < rows.Count; i++) datas.Add(Get(rows[i], m.data));
                ordem = Formatos.Ordem(datas);
            }
            bool valorComSinal = false;
            if (m.valor >= 0) for (int i = ini; i < rows.Count; i++) { double? v0 = Valor(Get(rows[i], m.valor), xlsx); if (v0 != null && v0 < 0) { valorComSinal = true; break; } }
            var r = new List<Lancamento>();
            var saldos = new List<double?>();
            var semSinal = new List<bool>();
            string ultimaData = null;
            Lancamento ultimo = null;
            int cont = 0;
            for (int i = ini; i < rows.Count; i++)
            {
                var cel = rows[i];
                if (recabecalho) { var h2 = Cabecalho(cel); if (h2.Ok() && SemDado(cel)) { m = h2; continue; } }
                else if (SemDado(cel) && i > ini && cel.SequenceEqual(rows[ini - 1 >= 0 ? ini - 1 : 0])) continue;
                string data = Data(Get(cel, m.data), xlsx, ordem);
                var hb = new StringBuilder();
                var partes = new List<int>(m.extras);
                partes.Insert(0, m.hist);
                partes.Insert(1, m.hist2);
                foreach (int c in partes)
                {
                    string t = Inteiro(Get(cel, c).Trim());
                    if (t.Length == 0 || t.Matches("-+") || Formatos.JaContem(hb.ToString(), t)) continue;
                    hb.Append(hb.Length > 0 ? " " : "").Append(t);
                }
                string tipoTxt = Get(cel, m.dc);
                bool? sentidoDC = Sentido(tipoTxt);
                if (sentidoDC == null && tipoTxt.Length > 3 && !Formatos.JaContem(hb.ToString(), tipoTxt)) hb.Insert(0, tipoTxt + " ");
                string dc0 = Inteiro(Get(cel, m.doc));
                if (dc0.Length > 0 && !dc0.Matches("-+") && !hb.ToString().Contains(dc0)) hb.Append(" ").Append(dc0);
                string hist = LeitorOFX.Limpar(hb.ToString());
                string nh = Regras.Normalizar(hist);
                double? sl = m.saldo >= 0 ? Valor(Get(cel, m.saldo), xlsx) : null;
                if (nh.Matches(SALDO_LINHA) || nh.Contains("SALDO ANTERIOR") || nh.Contains("SALDO DO DIA") || nh.Contains("OPENING BALANCE") || nh.Contains("PREVIOUS BALANCE"))
                {
                    double? sv = sl ?? (m.valor >= 0 ? Valor(Get(cel, m.valor), xlsx) : null);
                    if (sv != null && r.Count == 0 && nh.Matches(".*(ANTERIOR|INICIAL|OPENING|PREVIOUS|BROUGHT|ANFANG).*")) res.saldoInicial = sv;
                    if (sv != null && nh.Matches(".*(FINAL|CLOSING|CARRIED|ENDSALDO).*")) res.saldoFinal = sv;
                    ultimo = null;
                    continue;
                }
                double? v = null;
                bool saida = false;
                bool explicito = false;
                if (m.cred >= 0 || m.deb >= 0)
                {
                    double? c = Valor(Get(cel, m.cred), xlsx), d = Valor(Get(cel, m.deb), xlsx);
                    if (c != null && Math.Abs(c.Value) > 0.004) { v = c; saida = c < 0; explicito = true; }
                    else if (d != null && Math.Abs(d.Value) > 0.004) { v = d; saida = true; explicito = true; }
                }
                if (v == null && m.valor >= 0)
                {
                    string sv = Get(cel, m.valor).Trim();
                    v = Valor(sv, xlsx);
                    if (v != null) { saida = v < 0; explicito = v < 0 || valorComSinal || Regex.IsMatch(sv, @"^(?:.*(\+|\bC|CR)\s*$|^\+.*)$", RegexOptions.IgnoreCase); }
                }
                if (v == null || Math.Abs(v.Value) < 0.005)
                {
                    if (ultimo != null && data.Length == 0 && hist.Length > 0 && cont < 3 && SoHistorico(cel, m))
                    {
                        ultimo.historico = LeitorOFX.Limpar(ultimo.historico + " " + hist);
                        cont++;
                    }
                    continue;
                }
                if (sentidoDC != null) { saida = sentidoDC.Value; explicito = true; }
                if (data.Length == 0) data = ultimaData;
                if (data == null || hist.Length == 0 && m.hist >= 0) continue;
                ultimaData = data;
                var l = new Lancamento(data, hist.Length == 0 ? "(sem histórico)" : hist, Math.Abs(v.Value), saida);
                l.origem = origem;
                if (m.cred != m.credBruto || m.deb != m.debBruto)
                {
                    double? b = Valor(Get(cel, saida ? (m.debBruto >= 0 ? m.debBruto : m.credBruto) : m.credBruto), xlsx);
                    if (b != null && Math.Abs(Math.Abs(b.Value) - l.valor) > 0.004) l.valorBruto = Math.Abs(b.Value);
                }
                r.Add(l);
                saldos.Add(sl);
                semSinal.Add(!explicito && m.cred < 0 && m.deb < 0);
                ultimo = l;
                cont = 0;
            }
            bool cadeia = CadeiaDeSaldo(r, saldos, semSinal, res);
            if (!cadeia && r.Count > 0 && !semSinal.Contains(false))
                foreach (var l in r) { l.conferirSinal = true; l.obs = "Sinal presumido (entrada)"; }
            return r;
        }

        static void Aviso(LeitorExtrato.Resultado res, string s) { res.aviso = (res.aviso.Length == 0 ? "" : res.aviso + " ") + s; }

        /// <summary>Coluna de saldo: ordem do extrato, sinal dos valores sem sinal e saldos para a prova. true se a cadeia foi usada.</summary>
        public static bool CadeiaDeSaldo(List<Lancamento> r, List<double?> saldos, List<bool> semSinal, LeitorExtrato.Resultado res)
        {
            int n = r.Count;
            int comSaldo = saldos.Count(d => d != null);
            if (n < 2 || comSaldo < n * 0.8) return false;
            int asc = 0, desc = 0, ascAbs = 0, descAbs = 0;
            for (int i = 1; i < n; i++)
            {
                double? a = saldos[i - 1], b = saldos[i];
                if (a == null || b == null) continue;
                double vi = Mov(r[i]), va = Mov(r[i - 1]);
                if (Math.Abs(a.Value + vi - b.Value) < 0.011) asc++;
                if (Math.Abs(b.Value + va - a.Value) < 0.011) desc++;
                if (Math.Abs(Math.Abs(b.Value - a.Value) - r[i].valor) < 0.011) ascAbs++;
                if (Math.Abs(Math.Abs(a.Value - b.Value) - r[i - 1].valor) < 0.011) descAbs++;
            }
            bool decrescente = Math.Max(desc, descAbs) > Math.Max(asc, ascAbs);
            if (decrescente) { r.Reverse(); saldos.Reverse(); semSinal.Reverse(); }
            int corrigidos = 0;
            for (int i = 0; i < n; i++)
            {
                if (!semSinal[i] || saldos[i] == null) continue;
                double? ant = i > 0 ? saldos[i - 1] : null;
                if (ant == null) continue;
                double dif = saldos[i].Value - ant.Value;
                var l = r[i];
                if (Math.Abs(dif + l.valor) < 0.011) { if (!l.saida) corrigidos++; l.saida = true; }
                else if (Math.Abs(dif - l.valor) < 0.011) { if (l.saida) corrigidos++; l.saida = false; }
                else { l.conferirSinal = true; l.obs = "Sinal presumido (saldo não confere)"; }
            }
            for (int i = 1; i < n; i++) if (semSinal[i] && saldos[i] != null && saldos[i - 1] == null) { r[i].conferirSinal = true; r[i].obs = "Sinal presumido (sem saldo anterior)"; }
            if (semSinal[0])
            {
                bool ok = false;
                if (res.saldoInicial != null && saldos[0] != null)
                {
                    double dif = saldos[0].Value - res.saldoInicial.Value;
                    if (Math.Abs(dif + r[0].valor) < 0.011) { r[0].saida = true; ok = true; }
                    else if (Math.Abs(dif - r[0].valor) < 0.011) { r[0].saida = false; ok = true; }
                }
                if (!ok) { r[0].conferirSinal = true; r[0].obs = "Sinal presumido (primeira linha, sem saldo anterior no arquivo)"; }
            }
            if (res.saldoInicial == null && saldos[0] != null) res.saldoInicial = saldos[0] - Mov(r[0]);
            if (saldos[n - 1] != null) res.saldoFinal = saldos[n - 1];
            if (decrescente) Aviso(res, "Extrato em ordem decrescente (mais novo primeiro) — reordenado do mais antigo para o mais novo.");
            if (corrigidos > 0) Aviso(res, corrigidos + " valor(es) sem sinal tiveram entrada/saída definida pelo saldo.");
            return true;
        }

        static double Mov(Lancamento l) { return l.saida ? -l.valor : l.valor; }

        static bool SoHistorico(string[] cel, Mapa m)
        {
            for (int i = 0; i < cel.Length; i++)
            {
                if (cel[i] == null || cel[i].Trim().Length == 0) continue;
                if (i != m.hist && i != m.hist2 && !m.extras.Contains(i)) return false;
            }
            return true;
        }

        public static bool? Sentido(string tipo)
        {
            string t = Formatos.Normalizar(tipo);
            if (t.Length == 0) return null;
            if (t.Matches("D|DB|DR|DEB.*|DEBIT.*|S|SAIDA.*|-|CARGO.*|DEBE|SOLL|AF|RETIRO.*|WITHDRAWAL.*|PAYMENT|OUT|EGRESO.*|ADDEBITO|DEBITO")) return true;
            if (tipo.Trim() == "-") return true;
            if (t.Matches("C|CR|CRED.*|CREDIT.*|E|ENTRADA.*|ABONO.*|HABER|HABEN|H|BIJ|DEPOSIT.*|IN|INGRESO.*|ACCREDITO|CREDITO")) return false;
            if (tipo.Trim() == "+") return false;
            return null;
        }

        public static double? Valor(string s, bool xlsx)
        {
            return LeitorPlanilhaFinanceira.Valor(s, xlsx) ?? Formatos.Valor(s);
        }

        public static string Data(string s, bool xlsx, int ordem)
        {
            s = s == null ? "" : s.Trim();
            if (s.Length == 0) return "";
            if (xlsx && s.Matches(@"[0-9]{5}(\.[0-9]+)?")) return LeitorPlanilhaFinanceira.Data(s, true);
            return Formatos.Data(s, ordem, null);
        }

        static List<Lancamento> PorTexto(List<string[]> rows, string origem, bool xlsx, bool juntar)
        {
            var est = new Dictionary<int, int[]>();
            foreach (var cel in rows)
                for (int i = 0; i < cel.Length; i++)
                {
                    string c = (cel[i] ?? "").Trim();
                    if (c.Length == 0) continue;
                    var e = est.ObterOuCriar3(i);
                    e[2]++;
                    if (c.Matches(@"-?[0-9]+(\.[0-9]+)?")) { e[0]++; if (c.Contains(".")) e[1]++; }
                }
            var numericas = new HashSet<int>(est.Where(e => e.Value[1] > 0 && e.Value[0] * 10 >= e.Value[2] * 6).Select(e => e.Key));
            var sb = new StringBuilder();
            foreach (var cel in rows)
            {
                var partes = new List<string>();
                for (int i = 0; i < cel.Length; i++)
                {
                    string c = (cel[i] ?? "").Trim();
                    if (c.Length == 0) continue;
                    if (partes.Count == 0) { string d = LeitorPlanilhaFinanceira.Data(c, xlsx); if (d.Length > 0 && c.Length <= 19) { partes.Add(d); continue; } }
                    if (xlsx && numericas.Contains(i) && c.Matches(@"-?[0-9]+(\.[0-9]+)?")) c = J.Moeda(J.ParseD(c));
                    partes.Add(c);
                }
                if (partes.Count > 0) sb.Append(string.Join("  ", partes)).Append('\n');
            }
            return LeitorPDF.LerTexto(sb.ToString(), origem, juntar);
        }

        static int[] ObterOuCriar3(this Dictionary<int, int[]> d, int k)
        {
            int[] v;
            if (!d.TryGetValue(k, out v)) { v = new int[3]; d[k] = v; }
            return v;
        }

        /// <summary>Número de identificação lido do Excel como 1.929229312E9 → 1929229312.</summary>
        static string Inteiro(string s)
        {
            if (s.Matches(@"-?[0-9]+(\.[0-9]+)?([Ee][0-9]+)?"))
            {
                double d;
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d == Math.Round(d) && Math.Abs(d) >= 1000) return ((long)d).ToString(CultureInfo.InvariantCulture);
            }
            return s;
        }

        public static bool SemDado(string[] cel)
        {
            foreach (var c in cel) if (LeitorPlanilhaFinanceira.Data(c, false).Length > 0 || (c ?? "").Trim().Matches("-?[0-9.,]+")) return false;
            return true;
        }

        static string Get(string[] c, int i) { return i >= 0 && i < c.Length && c[i] != null ? c[i].Trim() : ""; }
    }
}

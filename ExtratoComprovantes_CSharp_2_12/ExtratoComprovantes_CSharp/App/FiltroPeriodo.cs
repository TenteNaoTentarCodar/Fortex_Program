using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ExtratoComprovantes
{
    /// <summary>
    /// Campos de período como no e-CAC: Data inicial, Data final ou "apenas a data". Usado na alteração em lote e na exportação para Excel.
    /// </summary>
    public class FiltroPeriodo
    {
        public readonly FlowLayoutPanel Painel = new FlowLayoutPanel { AutoSize = true, WrapContents = true };
        public readonly MaskedTextBox Ini, Fim, Unica;
        readonly Label aviso = new Label { AutoSize = true, Padding = new Padding(6, 5, 0, 0), ForeColor = Color.FromArgb(180, 0, 0) };
        public event Action Mudou;

        /// <summary>Período lido (null = sem limite) e se há erro (data incompleta/inválida ou inicial maior que a final).</summary>
        public DateTime? De { get; private set; }
        public DateTime? Ate { get; private set; }
        public bool Erro { get; private set; }

        public FiltroPeriodo(ToolTip dicas)
        {
            Ini = Campo("fDataIni"); Fim = Campo("fDataFim"); Unica = Campo("fDataUnica");
            var limpar = new Button { Text = "Limpar datas", AutoSize = true };
            Painel.Controls.Add(Rot("Data inicial:")); Painel.Controls.Add(Ini);
            Painel.Controls.Add(Rot("  Data final:")); Painel.Controls.Add(Fim);
            Painel.Controls.Add(Rot("     ou apenas a data:")); Painel.Controls.Add(Unica);
            Painel.Controls.Add(limpar); Painel.Controls.Add(aviso);
            if (dicas != null)
            {
                dicas.SetToolTip(Ini, "Lançamentos a partir desta data (pode deixar a final em branco).");
                dicas.SetToolTip(Fim, "Lançamentos até esta data (pode deixar a inicial em branco).");
                dicas.SetToolTip(Unica, "Só os lançamentos deste dia. Preenchida, substitui a data inicial e a final.");
            }
            foreach (var mt in new[] { Ini, Fim, Unica })
            {
                mt.TextChanged += (s, e) => { Ler(); if (Mudou != null) Mudou(); };
                mt.Enter += (s, e) => { var m = (MaskedTextBox)s; if (m.IsHandleCreated) m.BeginInvoke((Action)(() => m.Select(0, 0))); };
            }
            limpar.Click += (s, e) => { Ini.Text = ""; Fim.Text = ""; Unica.Text = ""; };
            Ler();
        }

        static Label Rot(string t) { return new Label { Text = t, AutoSize = true, Padding = new Padding(0, 5, 0, 0) }; }

        static MaskedTextBox Campo(string nome)
        {
            return new MaskedTextBox { Name = nome, Mask = "00/00/0000", Width = 90, TextMaskFormat = MaskFormat.IncludeLiterals, PromptChar = '_' };
        }

        /// <summary>null = em branco; Item2 = preenchida de forma incompleta ou inválida.</summary>
        static Tuple<DateTime?, bool> LerCampo(MaskedTextBox mt)
        {
            string dg = new string(mt.Text.Where(char.IsDigit).ToArray());
            if (dg.Length == 0) return Tuple.Create((DateTime?)null, false);
            DateTime d;
            if (dg.Length == 8 && DateTime.TryParseExact(dg, "ddMMyyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out d))
                return Tuple.Create((DateTime?)d, false);
            return Tuple.Create((DateTime?)null, true);
        }

        void Ler()
        {
            var i = LerCampo(Ini); var f = LerCampo(Fim); var u = LerCampo(Unica);
            Ini.Enabled = Fim.Enabled = !(u.Item1 != null || u.Item2);
            De = null; Ate = null; Erro = false;
            string msg = "";
            if (u.Item1 != null) { De = Ate = u.Item1; }
            else if (u.Item2) { msg = "⚠ data incompleta ou inválida"; Erro = true; }
            else
            {
                if (i.Item2 || f.Item2) { msg = "⚠ data incompleta ou inválida"; Erro = true; }
                De = i.Item1; Ate = f.Item1;
                if (De != null && Ate != null && De > Ate) { msg = "⚠ data inicial maior que a final"; Erro = true; }
            }
            aviso.Text = msg;
        }

        public bool Ativo { get { return De != null || Ate != null; } }

        /// <summary>A data (dd/MM/aaaa) está dentro do período.</summary>
        public bool Contem(string data)
        {
            if (!Ativo) return true;
            var dl = J.Data(data);
            if (dl == null) return false;
            return !(De != null && dl.Value.Date < De.Value || Ate != null && dl.Value.Date > Ate.Value);
        }

        public string Descricao()
        {
            if (!Ativo) return "todo o período";
            if (De != null && Ate != null && De == Ate) return De.Value.ToString("dd/MM/yyyy");
            if (De != null && Ate != null) return De.Value.ToString("dd/MM/yyyy") + " a " + Ate.Value.ToString("dd/MM/yyyy");
            return De != null ? "a partir de " + De.Value.ToString("dd/MM/yyyy") : "até " + Ate.Value.ToString("dd/MM/yyyy");
        }
    }
}

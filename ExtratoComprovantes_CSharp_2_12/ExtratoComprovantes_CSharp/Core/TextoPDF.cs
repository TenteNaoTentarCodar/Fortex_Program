using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
#if !MONO
using UglyToad.PdfPig;
#endif

namespace ExtratoComprovantes
{
    /// <summary>
    /// Texto do PDF em linhas, montado pela posição de cada trecho de texto:
    ///  - junta na mesma linha o que está na mesma altura;
    ///  - descarta texto repetido na mesma linha (relatórios que desenham a mesma célula várias vezes).
    /// </summary>
    public static class TextoPDF
    {
        /// <summary>Trecho de texto contínuo com posição (pontos; y cresce para baixo).</summary>
        public class Trecho
        {
            public readonly StringBuilder s = new StringBuilder();
            public float x, y, fim, h;
            public string Texto() { return s.ToString().Trim(); }
            public override string ToString() { return x.ToString("0.0") + "," + y.ToString("0.0") + "-" + fim.ToString("0.0") + " '" + Texto() + "'"; }
        }

        /// <summary>Uma letra do PDF (posição da linha de base, y de cima para baixo).</summary>
        public struct Letra
        {
            public string ch;
            public float x, y, w, h;
        }

        static readonly List<KeyValuePair<string, List<List<Trecho>>>> CACHE = new List<KeyValuePair<string, List<List<Trecho>>>>();

        public static string Chave(string p)
        {
            try { var f = new FileInfo(p); return f.FullName + "|" + f.Length + "|" + f.LastWriteTimeUtc.Ticks; }
            catch { return Path.GetFullPath(p); }
        }

        internal static int Rnd(float v) { return (int)Math.Floor(v + 0.5f); }

        /// <summary>Trechos de cada página (com cache). Não altere as listas devolvidas.</summary>
        public static List<List<Trecho>> Trechos(string arq)
        {
            string k = Chave(arq);
            lock (CACHE)
            {
                int i = CACHE.FindIndex(e => e.Key == k);
                if (i >= 0) { var e = CACHE[i]; CACHE.RemoveAt(i); CACHE.Add(e); return e.Value; }
            }
            var r = Letras(arq).Select(TrechosDaPagina).ToList();
            lock (CACHE)
            {
                CACHE.Add(new KeyValuePair<string, List<List<Trecho>>>(k, r));
                while (CACHE.Count > 3) CACHE.RemoveAt(0);
            }
            return r;
        }

        /// <summary>Letras de cada página, na ordem em que foram desenhadas.</summary>
        static List<List<Letra>> Letras(string arq)
        {
#if MONO
            throw new LeituraException("Leitura de PDF indisponível nesta compilação de teste.");
#else
            try
            {
                var o = new List<List<Letra>>();
                using (var doc = PdfDocument.Open(arq, new ParsingOptions { UseLenientParsing = true }))
                {
                    foreach (var pg in doc.GetPages())
                    {
                        var ls = new List<Letra>();
                        double alt = pg.Height;
                        foreach (var l in pg.Letters)
                        {
                            float h = (float)l.PointSize;
                            if (h <= 0) h = Math.Max((float)l.GlyphRectangle.Height, 6f);
                            ls.Add(new Letra
                            {
                                ch = l.Value,
                                x = (float)l.StartBaseLine.X,
                                y = (float)(alt - l.StartBaseLine.Y),
                                w = (float)Math.Max(l.EndBaseLine.X - l.StartBaseLine.X, l.Width),
                                h = h
                            });
                        }
                        o.Add(ls);
                    }
                }
                return o;
            }
            catch (LeituraException) { throw; }
            catch (Exception e)
            {
                throw new LeituraException("Não foi possível ler o PDF (" + e.Message + "). Se estiver protegido, abra e use \"Imprimir > Microsoft Print to PDF\".", e);
            }
#endif
        }

        public static List<string> Linhas(string arq)
        {
            var o = new List<string>();
            foreach (var pg in Trechos(arq)) o.AddRange(MontarLinhas(pg));
            return o;
        }

        /// <summary>Linhas de texto de cada página.</summary>
        public static List<List<string>> LinhasPorPagina(string arq)
        {
            return Trechos(arq).Select(MontarLinhas).ToList();
        }

        static List<Trecho> TrechosDaPagina(List<Letra> ps)
        {
            var ts = new List<Trecho>();
            Trecho t = null;
            foreach (var p in ps)
            {
                string ch = p.ch;
                if (string.IsNullOrEmpty(ch)) continue;
                float x = p.x, y = p.y, w = p.w, h = p.h;
                bool novo = t == null || Math.Abs(y - t.y) > Math.Max(0.6f, h * 0.2f) || x < t.fim - h * 0.35f || x > t.fim + h * 1.2f;
                if (novo)
                {
                    if (t != null && t.s.ToString().Trim().Length > 0) ts.Add(t);
                    t = new Trecho { x = x, y = y, h = h, fim = x };
                }
                else if (x - t.fim > h * 0.12f && t.s.Length > 0 && t.s[t.s.Length - 1] != ' ' && !ch.StartsWith(" "))
                {
                    t.s.Append(' ');
                }
                t.s.Append(ch);
                t.fim = Math.Max(t.fim, x + w);
                t.h = Math.Max(t.h, h);
            }
            if (t != null && t.s.ToString().Trim().Length > 0) ts.Add(t);
            return ts;
        }

        static readonly Regex LETRA = new Regex("[A-Za-zÀ-ú]");

        public static List<string> MontarLinhas(List<Trecho> origem)
        {
            var ts = new List<Trecho>(origem);
            ts.OrdenarEstavel((a, b) => a.y != b.y ? a.y.CompareTo(b.y) : a.x.CompareTo(b.x));
            var linhas = new List<List<Trecho>>();
            List<Trecho> L = null;
            float maxY = 0, maxH = 0;
            foreach (var r in ts)
            {
                float Y = r.y, h = r.h * 0.8f;
                if (L != null)
                {
                    bool ov = Math.Abs(Y - maxY) < 0.1f || (maxY <= Y && maxY >= Y - h) || (Y <= maxY && Y >= maxY - maxH);
                    if (ov) { L.Add(r); if (Y > maxY) maxY = Y; if (h > maxH) maxH = h; continue; }
                }
                L = new List<Trecho> { r };
                linhas.Add(L);
                maxY = Y; maxH = h;
            }
            var o = new List<string>();
            foreach (var ln in linhas)
            {
                ln.OrdenarEstavel((a, b) => a.x.CompareTo(b.x));
                var vistos = new HashSet<string>();
                var s = new StringBuilder();
                float? fim = null;
                foreach (var r in ln)
                {
                    string k = r.s.ToString().Trim();
                    if (k.Length >= 6 && LETRA.IsMatch(k) && !vistos.Add(Rnd(r.y) + "|" + k)) continue;
                    if (fim != null)
                    {
                        float gap = r.x - fim.Value;
                        if ((gap > r.h * 0.08f || gap < -r.h * 0.3f) && s.Length > 0 && s[s.Length - 1] != ' ' && !r.s.ToString().StartsWith(" ")) s.Append(' ');
                    }
                    s.Append(r.s);
                    fim = r.fim;
                }
                string txt = Regex.Replace(s.ToString().Replace(' ', ' '), @"\s{2,}", " ").Trim();
                if (txt.Length > 0) o.Add(txt);
            }
            return o;
        }
    }
}

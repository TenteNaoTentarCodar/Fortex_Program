using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtratoComprovantes
{
    /// <summary>Atalhos de texto e expressões regulares (mesma semântica do programa original em Java).</summary>
    public static class J
    {
        static J() { Regex.CacheSize = 400; }

        public static readonly CultureInfo BR = new CultureInfo("pt-BR");

        /// <summary>O texto inteiro casa com a expressão.</summary>
        public static bool Matches(this string s, string re) { return s != null && Regex.IsMatch(s, "^(?:" + re + ")$"); }

        public static string ReplaceAll(this string s, string re, string rep) { return Regex.Replace(s, re, rep); }

        public static string ReplaceFirst(this string s, string re, string rep) { return new Regex(re).Replace(s, rep, 1); }

        /// <summary>Divide por expressão; sem keepTrailing descarta os vazios do fim (como o split do Java).</summary>
        public static string[] SplitRe(this string s, string re, bool keepTrailing = false)
        {
            if (s.Length == 0) return new[] { "" };
            var p = Regex.Split(s, re).ToList();
            if (!keepTrailing) while (p.Count > 1 && p[p.Count - 1].Length == 0) p.RemoveAt(p.Count - 1);
            if (!keepTrailing && p.Count == 1 && p[0].Length == 0) return new string[0];
            return p.ToArray();
        }

        /// <summary>Divide por um caractere mantendo vazios (split com -1 do Java).</summary>
        public static string[] SplitKeep(this string s, char c) { return s.Split(c); }

        public static string SemAcento(string s)
        {
            if (s == null) return "";
            var n = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(n.Length);
            foreach (var ch in n) if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark
                                          && CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.SpacingCombiningMark
                                          && CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.EnclosingMark) sb.Append(ch);
            return sb.ToString();
        }

        public static string Moeda(double v) { return v.ToString("N2", BR); }

        public static int? Int(string s)
        {
            int v;
            return int.TryParse((s ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? (int?)v : null;
        }

        public static double ParseD(string s) { return double.Parse(s, NumberStyles.Float, CultureInfo.InvariantCulture); }

        public static string Sub(this string s, int ini, int fim) { return s.Substring(ini, fim - ini); }

        public static string Nz(string s) { return s ?? ""; }

        public static TV GetOr<TK, TV>(this IDictionary<TK, TV> d, TK k, TV padrao) { TV v; return d.TryGetValue(k, out v) ? v : padrao; }

        public static TV ObterOuCriar<TK, TV>(this IDictionary<TK, TV> d, TK k) where TV : new()
        {
            TV v;
            if (!d.TryGetValue(k, out v)) { v = new TV(); d[k] = v; }
            return v;
        }

        public static void Somar<TK>(this IDictionary<TK, int> d, TK k, int n) { int v; d.TryGetValue(k, out v); d[k] = v + n; }
        public static void Somar<TK>(this IDictionary<TK, double> d, TK k, double n) { double v; d.TryGetValue(k, out v); d[k] = v + n; }

        /// <summary>Bytes → texto: BOM UTF-16/UTF-8, UTF-8 válido ou Windows-1252.</summary>
        public static string Texto(byte[] b)
        {
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
            if (b.Length > 8 && b[1] == 0 && b[3] == 0 && b[0] != 0) return Encoding.Unicode.GetString(b);
            try
            {
                var utf8 = new UTF8Encoding(false, true);
                return utf8.GetString(b).Replace("﻿", "");
            }
            catch (DecoderFallbackException)
            {
                return Cp1252.GetString(b);
            }
        }

        private static Encoding _cp1252;
        public static Encoding Cp1252
        {
            get
            {
                if (_cp1252 == null)
                {
#if !MONO
                    try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); } catch { }
#endif
                    try { _cp1252 = Encoding.GetEncoding(1252); } catch { _cp1252 = Encoding.GetEncoding("iso-8859-1"); }
                }
                return _cp1252;
            }
        }

        /// <summary>Math.round do Java (arredonda 0,5 para cima).</summary>
        public static long RoundJ(double v) { return (long)Math.Floor(v + 0.5); }

        /// <summary>dd/MM/aaaa → data (null se inválida).</summary>
        public static DateTime? Data(string s)
        {
            try { return new DateTime(int.Parse(s.Substring(6, 4)), int.Parse(s.Substring(3, 2)), int.Parse(s.Substring(0, 2))); }
            catch { return null; }
        }

        public static long Dias(DateTime a, DateTime b) { return (long)Math.Round((b.Date - a.Date).TotalDays); }

        /// <summary>hashCode() de String do Java.</summary>
        public static int HashJava(string s) { unchecked { int h = 0; foreach (char c in s) h = 31 * h + c; return h; } }

        /// <summary>
        /// Ordem em que um HashMap/HashSet do Java (capacidade padrão, sem remoções) devolveria estas chaves, dadas na ordem de inserção.
        /// Mantém os desempates do programa original idênticos.
        /// </summary>
        public static List<TK> OrdemHashJava<TK>(IEnumerable<TK> insercao, Func<TK, int> hash)
        {
            var ks = insercao.ToList();
            int cap = 16;
            while (ks.Count > cap * 3 / 4) cap *= 2;
            return ks.Select((k, i) => new { k, i, b = (hash(k) ^ (int)((uint)hash(k) >> 16)) & (cap - 1) })
                     .OrderBy(x => x.b).ThenBy(x => x.i).Select(x => x.k).ToList();
        }

        /// <summary>Ordenação estável (List.Sort do .NET não é estável; o Java era).</summary>
        public static void OrdenarEstavel<T>(this List<T> l, Comparison<T> cmp)
        {
            var o = l.Select((x, i) => new KeyValuePair<int, T>(i, x)).ToList();
            o.Sort((a, b) => { int r = cmp(a.Value, b.Value); return r != 0 ? r : a.Key.CompareTo(b.Key); });
            for (int i = 0; i < o.Count; i++) l[i] = o[i].Value;
        }
    }

    /// <summary>Registro de erros em %LOCALAPPDATA%\ExtratoComprovantes\log.txt.</summary>
    public static class Log
    {
        private static readonly object trava = new object();

        public static string PastaApp()
        {
            string b = Environment.GetEnvironmentVariable("LOCALAPPDATA");
            string p = !string.IsNullOrEmpty(b) ? Path.Combine(b, "ExtratoComprovantes")
                : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".extratocomprovantes");
            try { Directory.CreateDirectory(p); } catch { }
            return p;
        }

        public static string Usuario() { var u = Environment.UserName; return string.IsNullOrEmpty(u) ? "usuario" : u; }
        public static string Maquina() { var m = Environment.MachineName; return string.IsNullOrEmpty(m) ? "maquina" : m; }

        public static void Info(string msg) { Escrever("INFO", msg, null); }
        public static void Erro(string msg, Exception ex) { Escrever("ERRO", msg, ex); }

        private static void Escrever(string nivel, string msg, Exception ex)
        {
            lock (trava)
            {
                try
                {
                    string f = Path.Combine(PastaApp(), "log.txt");
                    if (File.Exists(f) && new FileInfo(f).Length > 2000000) File.Copy(f, Path.Combine(PastaApp(), "log_anterior.txt"), true);
                    if (File.Exists(f) && new FileInfo(f).Length > 2000000) File.Delete(f);
                    var sb = new StringBuilder(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Append(' ').Append(nivel).Append(' ').Append(msg).Append(Environment.NewLine);
                    if (ex != null) sb.Append(ex).Append(Environment.NewLine);
                    File.AppendAllText(f, sb.ToString(), Encoding.UTF8);
                }
                catch { }
            }
        }
    }

    /// <summary>Erro de leitura com mensagem para o usuário (equivale ao IOException do Java).</summary>
    public class LeituraException : IOException
    {
        public LeituraException(string msg) : base(msg) { }
        public LeituraException(string msg, Exception inner) : base(msg, inner) { }
    }
}

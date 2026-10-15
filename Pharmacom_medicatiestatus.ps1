# Leest de medicatiestatus (F3) uit het open dossier in Pharmacom, via de Java Access Bridge.
# Schrijft per regel een tab-gescheiden regel naar standaarduitvoer (UTF-8):
#   OK<tab>venstertitel
#   Srt, Etiketnaam, Afl. Hoev., Dosering, Th. einddatum, Actief, Startdatum, ATC
# of een regel FOUT<tab>reden.
# Leest alleen; stuurt geen toetsen naar Pharmacom en bewaart niets.
# Moet in 32-bit PowerShell draaien (Pharmacom gebruikt 32-bit Java):
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File Pharmacom_medicatiestatus.ps1
# Met -Uit <bestand> gaat de uitvoer naar dat bestand (UTF-8) in plaats van naar het scherm;
# het bestand verschijnt in een keer (eerst .tmp, dan hernoemd), zodat Excel kan wachten tot het er is.
# Met -Proef wordt de koppeling wel gestart maar geen dossier gelezen (vaste, verzonnen uitvoer; voor testen).
# Met -Lab wordt in plaats van de medicatiestatus de nierfunctie gelezen uit het Diagnostisch dossier
# (Laboratorium uitslagen; Alt+I, Alt+D): de nieuwste eGFR (bij voorkeur CKD-EPI). Uitvoer:
#   OK<tab>venstertitel<tab>eGFR<tab>datum<tab>naam bepaling   (of FOUT<tab>reden)
# Gebaseerd op de koppeling van de Etiketten autoprinter (src\Jab.ahk).

param([string]$Uit = "", [switch]$Proef, [switch]$Lab)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Schrijf($Tekst) {
    if ($Uit) {
        [IO.File]::WriteAllText("$Uit.tmp", $Tekst, (New-Object Text.UTF8Encoding $false))
        Move-Item -LiteralPath "$Uit.tmp" -Destination $Uit -Force
    }
    else { [Console]::Out.Write($Tekst) }
}
function Fout($Reden) { Schrijf "FOUT`t$Reden`n"; exit 1 }
trap { Fout ("Onverwachte fout: " + $_.Exception.Message) }

if ([IntPtr]::Size -ne 4) { Fout "Script draait niet in 32-bit PowerShell" }

$Proces = Get-Process javaw -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like 'Pharmacom*' } | Select-Object -First 1
if (-not $Proces) { Fout "Pharmacom is niet geopend" }
$Dll = Join-Path (Split-Path $Proces.Path) 'WindowsAccessBridge-32.dll'
if (-not (Test-Path $Dll)) { Fout "Java Access Bridge niet gevonden ($Dll)" }

Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition (@"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
using System.Text.RegularExpressions;
public static class Jab {
  const string D = @"__DLL__";
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] public static extern void Windows_run();
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int isJavaWindow(IntPtr h);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleContextFromHWND(IntPtr h, out int vm, out long ac);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleContextInfo(int vm, long ac, byte[] buf);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern long getAccessibleChildFromContext(int vm, long ac, int i);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleTableInfo(int vm, long ac, byte[] buf);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleTableColumnHeader(int vm, long ac, byte[] buf);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleTableCellInfo(int vm, long at, int r, int c, byte[] buf);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleTextInfo(int vm, long ac, byte[] buf, int x, int y);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern int getAccessibleTextRange(int vm, long ac, int s, int e, byte[] buf, short n);
  [DllImport(D, CallingConvention=CallingConvention.Cdecl)] static extern void releaseJavaObject(int vm, long o);
  [DllImport("user32.dll")] static extern bool EnumWindows(EP cb, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  delegate bool EP(IntPtr h, IntPtr l);

  static string S(byte[] b, int off) { int n = 0; while (off + n*2 + 1 < b.Length && (b[off+n*2] != 0 || b[off+n*2+1] != 0)) n++; return Encoding.Unicode.GetString(b, off, n*2); }

  // Zichtbare tekst van een element (naam, anders tekstinhoud)
  static string Naam(int vm, long ac) {
    var b = new byte[6188]; if (getAccessibleContextInfo(vm, ac, b) == 0) return "";
    string n = S(b, 0);
    if (n.Trim() != "" || BitConverter.ToInt32(b, 6180) == 0) return n;
    var ti = new byte[12]; if (getAccessibleTextInfo(vm, ac, ti, 0, 0) == 0) return "";
    int len = Math.Min(BitConverter.ToInt32(ti, 0), 1000); if (len < 1) return "";
    var tb = new byte[2*1026]; if (getAccessibleTextRange(vm, ac, 0, len-1, tb, 1025) == 0) return "";
    return S(tb, 0);
  }

  static string Cel(int vm, long at, int r, int c) {
    var ci = new byte[40]; if (getAccessibleTableCellInfo(vm, at, r, c, ci) == 0) return "";
    long cac = BitConverter.ToInt64(ci, 0); if (cac == 0) return "";
    string t = Naam(vm, cac); releaseJavaObject(vm, cac);
    t = Regex.Replace(t, "<[^>]*>", " ");
    t = t.Replace("\t", " ").Replace("\r", " ").Replace("\n", " ").Replace((char)160, ' ');
    return Regex.Replace(t, " {2,}", " ").Trim();
  }

  static readonly string[] Gezocht = { "Srt", "Etiketnaam", "Afl. Hoev.", "Dosering", "Th. einddatum", "Actief", "Startdatum", "ATC" };
  public static string Resultaat = null;
  static int teller;

  static bool Walk(int vm, long ac, int d) {
    if (++teller > 6000 || d > 80) return false;
    var b = new byte[6188]; if (getAccessibleContextInfo(vm, ac, b) == 0) return false;
    string role = S(b, 4608), st = S(b, 5632); int ch = BitConverter.ToInt32(b, 6148);
    if (!("," + st + ",").Contains(",showing,")) return false;
    if (role == "table") return LeesTabel(vm, ac);
    if (role == "menu bar" || role == "popup menu" || role == "menu") return false;
    for (int i = 0; i < Math.Min(ch, 3000); i++) {
      long k = getAccessibleChildFromContext(vm, ac, i); if (k == 0) continue;
      bool klaar = Walk(vm, k, d + 1); releaseJavaObject(vm, k);
      if (klaar) return true;
    }
    return false;
  }

  static bool LeesTabel(int vm, long ac) {
    var ti = new byte[48]; if (getAccessibleTableInfo(vm, ac, ti) == 0) return false;
    int rows = BitConverter.ToInt32(ti, 16), cols = BitConverter.ToInt32(ti, 20);
    long at = BitConverter.ToInt64(ti, 32);
    var koppen = new string[cols];
    var hi = new byte[48];
    if (getAccessibleTableColumnHeader(vm, ac, hi) != 0) {
      long hat = BitConverter.ToInt64(hi, 32);
      if (hat != 0) { for (int c = 0; c < cols; c++) koppen[c] = Cel(vm, hat, 0, c); releaseJavaObject(vm, hat); }
    }
    var kol = new int[Gezocht.Length];
    for (int g = 0; g < Gezocht.Length; g++) { kol[g] = -1; for (int c = 0; c < cols; c++) if (koppen[c] != null && koppen[c].Trim() == Gezocht[g]) { kol[g] = c; break; } }
    // Herkennen aan Etiketnaam + Dosering + Actief (medicatiestatus)
    if (kol[1] < 0 || kol[3] < 0 || kol[5] < 0) { releaseJavaObject(vm, at); return false; }
    var sb = new StringBuilder();
    sb.Append(string.Join("\t", Gezocht)).Append('\n');
    for (int r = 0; r < Math.Min(rows, 500); r++) {
      var v = new string[Gezocht.Length];
      for (int g = 0; g < Gezocht.Length; g++) v[g] = kol[g] >= 0 ? Cel(vm, at, r, kol[g]) : "";
      sb.Append(string.Join("\t", v)).Append('\n');
    }
    releaseJavaObject(vm, at);
    Resultaat = sb.ToString();
    return true;
  }

  // ---- Nierfunctie (Diagnostisch dossier, Laboratorium uitslagen) ----
  // Drie tabellen naast elkaar: namen (kop "Naam"), uitslagen per datum (koppen = datums, nieuwste links), normaalwaarden.
  static readonly Regex Gfr = new Regex("e?GFR", RegexOptions.IgnoreCase);
  static List<int> GfrRegels = new List<int>(); static List<string> GfrNamen = new List<string>();
  static string GfrWaarde, GfrDatum, GfrNaam; static int GfrScore;

  static bool WalkLab(int vm, long ac, int d) {
    if (++teller > 6000 || d > 80) return false;
    var b = new byte[6188]; if (getAccessibleContextInfo(vm, ac, b) == 0) return false;
    string role = S(b, 4608), st = S(b, 5632); int ch = BitConverter.ToInt32(b, 6148);
    if (!("," + st + ",").Contains(",showing,")) return false;
    if (role == "table") { LabTabel(vm, ac); return false; }
    if (role == "menu bar" || role == "popup menu" || role == "menu") return false;
    for (int i = 0; i < Math.Min(ch, 3000); i++) {
      long k = getAccessibleChildFromContext(vm, ac, i); if (k == 0) continue;
      WalkLab(vm, k, d + 1); releaseJavaObject(vm, k);
    }
    return false;
  }

  static void LabTabel(int vm, long ac) {
    var ti = new byte[48]; if (getAccessibleTableInfo(vm, ac, ti) == 0) return;
    int rows = BitConverter.ToInt32(ti, 16), cols = BitConverter.ToInt32(ti, 20);
    long at = BitConverter.ToInt64(ti, 32);
    var koppen = new string[cols];
    var hi = new byte[48];
    if (getAccessibleTableColumnHeader(vm, ac, hi) != 0) {
      long hat = BitConverter.ToInt64(hi, 32);
      if (hat != 0) { for (int c = 0; c < cols; c++) koppen[c] = Cel(vm, hat, 0, c); releaseJavaObject(vm, hat); }
    }
    if (cols == 1 && koppen[0] == "Naam") {
      GfrRegels.Clear(); GfrNamen.Clear();
      for (int r = 0; r < Math.Min(rows, 500); r++) { string n = Cel(vm, at, r, 0); if (Gfr.IsMatch(n)) { GfrRegels.Add(r); GfrNamen.Add(n); } }
    } else if (cols > 1 && GfrRegels.Count > 0 && Regex.IsMatch(koppen[0] ?? "", @"^\d{2}-\d{2}-\d{4}")) {
      for (int i = 0; i < GfrRegels.Count; i++) {
        for (int c = 0; c < cols; c++) {
          string w = Cel(vm, at, GfrRegels[i], c);
          if (w == "") continue;
          // Nieuwste uitslag; bij gelijke datum CKD-EPI boven andere formules
          DateTime dt; DateTime.TryParseExact(koppen[c] ?? "", "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out dt);
          int score = (int)(dt.Ticks / TimeSpan.TicksPerDay) * 2 + (GfrNamen[i].ToUpper().Contains("CKD") ? 1 : 0);
          if (GfrWaarde == null || score > GfrScore) { GfrWaarde = w; GfrDatum = koppen[c]; GfrNaam = GfrNamen[i]; GfrScore = score; }
          break;
        }
      }
      GfrRegels.Clear(); GfrNamen.Clear();
    }
    releaseJavaObject(vm, at);
  }

  public static string LeesLab(uint pid) {
    var hs = new List<IntPtr>();
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p == pid && IsWindowVisible(h)) hs.Add(h); return true; }, IntPtr.Zero);
    GfrWaarde = null;
    foreach (var h in hs) {
      if (isJavaWindow(h) == 0) continue;
      int vm; long ac; if (getAccessibleContextFromHWND(h, out vm, out ac) == 0) continue;
      teller = 0; WalkLab(vm, ac, 0); releaseJavaObject(vm, ac);
      if (GfrWaarde != null) { var sb = new StringBuilder(256); GetWindowText(h, sb, 256); Titel = sb.ToString(); return GfrWaarde + "\t" + GfrDatum + "\t" + GfrNaam; }
    }
    return null;
  }

  public static string Titel = "";
  public static string Lees(uint pid) {
    var hs = new List<IntPtr>();
    EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p == pid && IsWindowVisible(h)) hs.Add(h); return true; }, IntPtr.Zero);
    foreach (var h in hs) {
      if (isJavaWindow(h) == 0) continue;
      int vm; long ac; if (getAccessibleContextFromHWND(h, out vm, out ac) == 0) continue;
      teller = 0; bool ok = Walk(vm, ac, 0); releaseJavaObject(vm, ac);
      if (ok) { var sb = new StringBuilder(256); GetWindowText(h, sb, 256); Titel = sb.ToString(); return Resultaat; }
    }
    return null;
  }
}
"@).Replace('__DLL__', $Dll)

[Jab]::Windows_run()
# De bridge meldt zich via vensterberichten bij Java
$t = [DateTime]::Now
while (([DateTime]::Now - $t).TotalMilliseconds -lt 1200) { [System.Windows.Forms.Application]::DoEvents(); Start-Sleep -Milliseconds 20 }

if ($Lab) {
    if ($Proef) { Schrijf "OK`tPharmacom - Dossier - PROEF (verzonnen)`t72`t01-10-2026`tCKD-EPI, eGFR`n"; exit 0 }
    $Res = [Jab]::LeesLab([uint32]$Proces.Id)
    if (-not $Res) { Start-Sleep -Milliseconds 500; $Res = [Jab]::LeesLab([uint32]$Proces.Id) }
    if (-not $Res) { Fout "Geen eGFR gevonden in het Diagnostisch dossier (Laboratorium uitslagen)" }
    Schrijf ("OK`t" + [Jab]::Titel + "`t" + $Res + "`n")
    exit 0
}

if ($Proef) {
    Schrijf ("OK`tPharmacom - Dossier - PROEF (verzonnen)`nSrt`tEtiketnaam`tAfl. Hoev.`tDosering`tTh. einddatum`tActief`tStartdatum`tATC`n" +
             "G`tOXYCODON CAPS 10MG`t1 ST`tP1CM6`t`t`t`tN02AA05`n")
    exit 0
}

$Res = [Jab]::Lees([uint32]$Proces.Id)
if (-not $Res) {
    $Res = $null
    Start-Sleep -Milliseconds 500
    $Res = [Jab]::Lees([uint32]$Proces.Id)
}
if (-not $Res) { Fout "Medicatiestatus (F3) niet gevonden: open het dossier van de patiënt op het tabblad medicatiestatus" }
Schrijf ("OK`t" + [Jab]::Titel + "`n" + $Res)

#requires -version 1

# ==============================================================================
# FuntTop - התקנת גופנים מהירה
# מבוסס PowerShell ו-Windows Forms
# ==============================================================================

#region Initialization & Assemblies
# טעינת ספריות ליבה הנדרשות לממשק המשתמש (WinForms) ולעיבוד גרפי
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# טעינת ספריות WPF (לצורך קריאת מטא-דאטה של גופנים דרך GlyphTypeface)
try { Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue } catch { }
try { Add-Type -AssemblyName WindowsBase -ErrorAction SilentlyContinue } catch { }
#endregion

#region Win32 API Calls
# הגדרת פונקציות מערכת (Win32 API) להתקנת הגופן במערכת ההפעלה ורענון המטמון
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class Gdi32 {
    // מוסיף את הגופן לטבלת הגופנים של המערכת בזמן ריצה
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int AddFontResourceEx(string lpszFilename, uint fl, IntPtr pdv);
}

public static class User32 {
    // משמש לשליחת הודעת רענון לכל החלונות הפתוחים במערכת
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
        uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    public static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);
    public const uint WM_FONTCHANGE = 0x001D;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
}
"@ -Language CSharp -ReferencedAssemblies 'System.Runtime.InteropServices' -ErrorAction SilentlyContinue

<#
.SYNOPSIS
שולח הודעת מערכת (Broadcast) לכל החלונות כדי לעדכן שמטמון הגופנים השתנה.
#>
function Send-FontChangeBroadcast {
    try {
        $result = [IntPtr]::Zero
        [void][User32]::SendMessageTimeout([User32]::HWND_BROADCAST, [User32]::WM_FONTCHANGE, [IntPtr]::Zero, [IntPtr]::Zero, [User32]::SMTO_ABORTIFHUNG, 2000, [ref]$result)
    } catch { }
}
#endregion

#region Font Processing & Installation Logic

<#
.SYNOPSIS
שולף את הערך המקומי (Localized) הראשון שזמין מתוך מילון, עם עדיפות לעברית ואנגלית.
#>
function Get-FirstLocalizedValue {
    param(
        [Parameter(Mandatory)] $Dict,
        [string[]] $Prefer = @('he-IL','he','en-US','en')
    )
    try {
        if ($null -eq $Dict) { return $null }
        foreach ($k in $Prefer) {
            if ($Dict.ContainsKey($k)) { return [string]$Dict[$k] }
        }
        return [string]($Dict.Values | Select-Object -First 1)
    } catch { return $null }
}

<#
.SYNOPSIS
קורא את המטא-דאטה מתוך קובץ הגופן (כגון גרסה, שם משפחה, מעצב) באמצעות WPF GlyphTypeface.
#>
function Get-FontMetaFromFile {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $uri = [Uri]::new($Path)
        $gt  = [System.Windows.Media.GlyphTypeface]::new($uri)

        $family  = Get-FirstLocalizedValue -Dict $gt.FamilyNames
        $face    = Get-FirstLocalizedValue -Dict $gt.FaceNames
        $ver     = if ($gt.VersionStrings) { Get-FirstLocalizedValue -Dict $gt.VersionStrings } else { $null }
        $manu    = if ($gt.ManufacturerNames) { Get-FirstLocalizedValue -Dict $gt.ManufacturerNames } else { $null }
        $des     = if ($gt.DesignerNames)     { Get-FirstLocalizedValue -Dict $gt.DesignerNames }     else { $null }

        $style   = $gt.Style.ToString()
        $weight  = $gt.Weight.ToString()
        $stretch = $gt.Stretch.ToString()

        $display = if ($face -and $family) { "$family $face" } elseif ($family) { $family } else { [IO.Path]::GetFileNameWithoutExtension($Path) }

        [pscustomobject]@{
            DisplayName  = $display
            Family       = $family
            Face         = $face
            Version      = $ver
            Manufacturer = $manu
            Designer     = $des
            Style        = $style
            Weight       = $weight
            Stretch      = $stretch
            Extension    = [IO.Path]::GetExtension($Path)
        }
    } catch { return $null }
}

<#
.SYNOPSIS
בודק האם הסקריפט רץ בהרשאות מנהל (Administrator).
#>
function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p  = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

<#
.SYNOPSIS
אוסף את כל קבצי הגופנים הנתמכים מתוך תיקייה נתונה.
#>
function Get-FontFilesInFolder([string]$Folder) {
    if (-not (Test-Path -LiteralPath $Folder)) { return @() }
    Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '^\.(ttf|otf|ttc)$' } |
        Sort-Object Name
}

<#
.SYNOPSIS
מבצע את תהליך ההתקנה בפועל: העתקה לתיקיית המערכת/משתמש ורישום ברג'יסטרי.
#>
function Install-FontFile {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [switch]$ForceOverwrite
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "קובץ לא נמצא: $Path" }

    $isAdmin = Test-IsAdmin

    # קביעת תיקיית יעד בהתאם להרשאות
    $destFolder = if ($isAdmin) {
        Join-Path $env:WINDIR 'Fonts'
    } else {
        Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
    }

    if (-not (Test-Path -LiteralPath $destFolder)) {
        New-Item -ItemType Directory -Path $destFolder -Force | Out-Null
    }

    $leaf = Split-Path -Leaf $Path
    $destPath = Join-Path $destFolder $leaf

    # העתקת הקובץ
    try {
        if (-not $ForceOverwrite -and (Test-Path -LiteralPath $destPath)) {
            return # קיים ואין דריסה, יציאה שקטה
        }
        Copy-Item -LiteralPath $Path -Destination $destPath -Force -ErrorAction Stop
    } catch {
        throw "שגיאה בהעתקת הקובץ לתיקיית הגופנים: $($_.Exception.Message)"
    }

    $ext = ([IO.Path]::GetExtension($leaf)).ToLowerInvariant()
    $kind = if ($ext -eq '.otf') { 'OpenType' } else { 'TrueType' }

    # יצירת מפתח רג'יסטרי
    $meta = Get-FontMetaFromFile -Path $destPath
    $display = if ($meta -and $meta.DisplayName) { $meta.DisplayName } else { [IO.Path]::GetFileNameWithoutExtension($leaf) }
    $regValueName = "$display ($kind)"

    $regRoot = if ($isAdmin) { 
        'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' 
    } else { 
        'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts' 
    }

    try {
        New-ItemProperty -Path $regRoot -Name $regValueName -Value $leaf -PropertyType String -Force -ErrorAction Stop | Out-Null
    } catch {
        # שגיאת רג'יסטרי נבלעת כדי לא להפריע אם הקובץ הועתק בהצלחה
    }

    # טעינת הגופן למערכת בזמן ריצה ורענון
    try {
        [void][Gdi32]::AddFontResourceEx($destPath, 0, [IntPtr]::Zero)
        Send-FontChangeBroadcast
    } catch { }
}

<#
.SYNOPSIS
בודק האם הגופן (לפי משפחה או לפי קובץ) כבר מותקן במערכת.
#>
function Test-FontInstalledByFile([string]$fontPath) {
    $pfc = $null
    try {
        Add-Type -AssemblyName System.Drawing | Out-Null
        $pfc = New-Object System.Drawing.Text.PrivateFontCollection
        $pfc.AddFontFile($fontPath)

        if ($pfc.Families.Count -gt 0) {
            $name = $pfc.Families[0].Name
            $installed = New-Object System.Drawing.Text.InstalledFontCollection
            foreach ($fam in $installed.Families) {
                if ($fam.Name -ieq $name) { return $true }
            }
        }
    } catch { 
    } finally {
        # חובה לשחרר כדי למנוע נעילת הקובץ במערכת הפעלה
        if ($null -ne $pfc) { $pfc.Dispose() }
    }

    # בדיקת גיבוי לפי קיום הקובץ בתיקיית Windows
    try {
        $leaf = Split-Path -Leaf $fontPath
        $dst = Join-Path $env:WINDIR ("Fonts\" + $leaf)
        return (Test-Path -LiteralPath $dst)
    } catch { return $false }
}
#endregion

#region C# UI Components (Custom Controls)
# הגדרת פקדים מותאמים אישית: כפתור מעוגל (RoundButton) ורשימת תצוגת גופנים (FontPreviewListView)
$csUI = @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Windows.Forms;

/// <summary>
/// כפתור מעוצב עם פינות מעוגלות ותמיכה באפקט רחרוח (Hover) ולחיצה.
/// </summary>
public class RoundButton : Button
{
    public int CornerRadius { get; set; }
    public Color FillColor { get; set; }
    public Color FillHover { get; set; }
    public Color FillDown  { get; set; }
    public Color BorderColor { get; set; }
    public int BorderSize { get; set; }

    private bool _hover;
    private bool _down;

    public RoundButton()
    {
        CornerRadius = 12;

        FillColor  = Color.FromArgb(0, 150, 255);
        FillHover  = Color.FromArgb(0, 165, 255);
        FillDown   = Color.FromArgb(0, 130, 220);

        BorderColor = Color.FromArgb(120, 170, 210);
        BorderSize  = 2;

        SetStyle(ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.UserPaint |
                 ControlStyles.ResizeRedraw, true);

        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        BackColor = Color.Transparent;
        ForeColor = Color.White;
        Cursor = Cursors.Hand;
    }

    protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { _hover = false; _down = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseDown(MouseEventArgs mevent) { _down = true; Invalidate(); base.OnMouseDown(mevent); }
    protected override void OnMouseUp(MouseEventArgs mevent) { _down = false; Invalidate(); base.OnMouseUp(mevent); }

    private GraphicsPath RoundedRect(Rectangle r, int radius)
    {
        int d = radius * 2;
        var path = new GraphicsPath();
        path.AddArc(r.X, r.Y, d, d, 180, 90);
        path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

        var rect = new Rectangle(0, 0, Width - 1, Height - 1);
        int rad = Math.Max(2, CornerRadius);

        Color fill = FillColor;
        if (!Enabled) fill = Color.FromArgb(210, 220, 230);
        else if (_down) fill = FillDown;
        else if (_hover) fill = FillHover;

        using (var gp = RoundedRect(rect, rad))
        {
            using (var br = (!_down && _hover && Enabled)
                ? (Brush)new LinearGradientBrush(rect, FillHover, FillColor, 90f)
                : (Brush)new SolidBrush(fill))
            {
                e.Graphics.FillPath(br, gp);
            }

            if (BorderSize > 0)
            {
                using (var pen = new Pen(BorderColor, BorderSize))
                {
                    pen.Alignment = PenAlignment.Inset;
                    e.Graphics.DrawPath(pen, gp);
                }
            }

            this.Region = new Region(gp);
        }

        TextRenderer.DrawText(
            e.Graphics,
            Text,
            Font,
            rect,
            Enabled ? ForeColor : Color.White,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis
        );
    }
}

/// <summary>
/// רשימת תצוגה מותאמת אישית לגופנים הכוללת ציור עצמאי (OwnerDraw).
/// </summary>
public class FontPreviewListView : ListView
{
    private class FontEntry
    {
        public string Path;
        public string DisplayName;
        public bool Checked;
        public bool Installed;
        public PrivateFontCollection Pfc;
        public Font PreviewFont;
        public string Meta;
    }

    private readonly Dictionary<ListViewItem, FontEntry> _map = new Dictionary<ListViewItem, FontEntry>();

    // הגדרות צבעים ועיצוב פנימיות
    private readonly Color _accent = Color.FromArgb(0, 150, 255);
    private readonly Color _accentSoft = Color.FromArgb(227, 247, 255);
    private readonly Color _rowAlt = Color.FromArgb(246, 252, 255);
    private readonly Color _grid = Color.FromArgb(220, 235, 245);
    private readonly Color _text = Color.FromArgb(20, 25, 35);
    private readonly Color _muted = Color.FromArgb(80, 95, 110);
    private readonly Color _title = Color.FromArgb(15, 76, 117);
    private readonly Color _meta  = Color.FromArgb(90, 110, 130);
    private readonly Color _good = Color.FromArgb(0, 140, 80);

    private string _sampleText = "אבגדהוזחטיכךלמםנןסעפףצץקרשת  | ABC abc 123 | ניקוד: שָׁלוֹם";
    private int _rowHeight = 74;
    private int _checkSize = 16;

    public FontPreviewListView()
    {
        this.View = View.Details;
        this.FullRowSelect = true;
        this.HideSelection = false;
        this.OwnerDraw = true;
        this.DoubleBuffered = true;
        this.BorderStyle = BorderStyle.FixedSingle;
        this.RightToLeft = RightToLeft.Yes;

        // שימוש ב-ImageList בלתי נראה כדי לכפות גובה שורה
        var il = new ImageList();
        il.ImageSize = new Size(1, _rowHeight);
        this.SmallImageList = il;

        this.Columns.Add("תצוגה מקדימה", 500, HorizontalAlignment.Left);
        this.Columns.Add("שם גופן", 400, HorizontalAlignment.Right);

        // ציור כותרות העמודות
        this.DrawColumnHeader += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var bg = new SolidBrush(Color.White))
                e.Graphics.FillRectangle(bg, e.Bounds);

            using (var p = new Pen(_grid, 1))
                e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);

            using (var b = new SolidBrush(_muted))
            {
                var fmt = new StringFormat();
                fmt.LineAlignment = StringAlignment.Center;
                fmt.Alignment = StringAlignment.Center;
                var r = new Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 20, e.Bounds.Height);
                e.Graphics.DrawString(e.Header.Text, this.Font, b, r, fmt);
            }
        };

        this.DrawItem += (s, e) => { };

        // ציור פריטי הרשימה (הגופנים עצמם)
        this.DrawSubItem += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            bool selected = (e.ItemState & ListViewItemStates.Selected) != 0;
            bool alt = (e.ItemIndex % 2) == 1;

            Color back = selected ? _accentSoft : (alt ? _rowAlt : Color.White);
            using (var bb = new SolidBrush(back))
                e.Graphics.FillRectangle(bb, e.Bounds);

            using (var p = new Pen(_grid, 1))
                e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);

            FontEntry fe;
            _map.TryGetValue(e.Item, out fe);

            if (e.ColumnIndex == 1) // עמודת פרטי הגופן
            {
                Rectangle col = e.Bounds;

                using (var pSep = new Pen(_grid, 1))
                {
                    pSep.Alignment = PenAlignment.Inset;
                    e.Graphics.DrawLine(pSep, col.Left, col.Top, col.Left, col.Bottom);
                }

                // תיבת סימון משמאל לשם הגופן
                int cx = col.Right - 12 - _checkSize;
                int cy = col.Y + (col.Height - _checkSize) / 2;
                var cb = new Rectangle(cx, cy, _checkSize, _checkSize);
                DrawRoundCheck(e.Graphics, cb, fe != null && fe.Checked);

                var textRect = new Rectangle(col.X + 10, col.Y + 6, col.Width - (10 + _checkSize + 18), col.Height - 12);

                string name = e.SubItem.Text;
                string status = (fe != null && fe.Installed) ? "מותקן" : "לא מותקן";

                using (var b1 = new SolidBrush(_title))
                using (var bMeta = new SolidBrush(_meta))
                using (var bStat = new SolidBrush((fe != null && fe.Installed) ? _good : _muted))
                {
                    var rTop = new Rectangle(textRect.X, textRect.Y, textRect.Width, (int)(textRect.Height * 0.45));
                    var rMeta = new Rectangle(textRect.X, textRect.Y + (int)(textRect.Height * 0.45), textRect.Width, textRect.Height - (int)(textRect.Height * 0.45));

                    // ציור שם הגופן
                    using (var fName = new Font(this.Font.FontFamily, this.Font.Size + 3.0f, FontStyle.Bold))
                    using (var fmtName = new StringFormat())
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(name, @"\p{IsHebrew}")) {
                            fmtName.FormatFlags |= StringFormatFlags.DirectionRightToLeft;
                            fmtName.Alignment = StringAlignment.Near; // ימין ב-RTL
                        } else {
                            fmtName.Alignment = StringAlignment.Far; // ימין ב-LTR
                        }
                        e.Graphics.DrawString(name, fName, b1, rTop, fmtName);
                    }

                    // ציור סטטוס מותקן באותה שורה אך בצד השני
                    using (var fStat = new Font(this.Font.FontFamily, 9.0f, FontStyle.Bold))
                    using (var fmtStat = new StringFormat())
                    {
                        fmtStat.FormatFlags |= StringFormatFlags.DirectionRightToLeft;
                        fmtStat.Alignment = StringAlignment.Far; // שמאל ב-RTL
                        e.Graphics.DrawString(status, fStat, bStat, rTop, fmtStat);
                    }

                    // ציור המידע הנוסף עם אפשרות גלישה (Wrap)
                    string meta = (fe != null) ? fe.Meta : null;
                    if (!string.IsNullOrWhiteSpace(meta))
                    {
                        using (var fMeta = new Font(this.Font.FontFamily, 8.5f, FontStyle.Regular))
                        using (var fmtMeta = new StringFormat())
                        {
                            if (System.Text.RegularExpressions.Regex.IsMatch(meta, @"\p{IsHebrew}")) {
                                fmtMeta.FormatFlags |= StringFormatFlags.DirectionRightToLeft;
                                fmtMeta.Alignment = StringAlignment.Near; 
                            } else {
                                fmtMeta.Alignment = StringAlignment.Far; 
                            }
                            fmtMeta.Trimming = StringTrimming.EllipsisWord;
                            e.Graphics.DrawString(meta, fMeta, bMeta, rMeta, fmtMeta);
                        }
                    }
                }
            }
            else // עמודת התצוגה המקדימה
            {
                Rectangle col = e.Bounds;
                var previewRect = new Rectangle(col.X + 10, col.Y + 6, col.Width - 20, col.Height - 12);

                var fmt = new StringFormat();
                fmt.LineAlignment = StringAlignment.Center;
                fmt.Alignment = StringAlignment.Far;
                fmt.FormatFlags |= StringFormatFlags.DirectionRightToLeft;

                Font f = (fe != null && fe.PreviewFont != null) ? fe.PreviewFont : this.Font;
                using (var b = new SolidBrush(_text))
                    e.Graphics.DrawString(e.SubItem.Text, f, b, previewRect, fmt);
            }
        };

        // טיפול בלחיצת עכבר על תיבת הסימון
        this.MouseDown += (s, e) =>
        {
            var hit = this.HitTest(e.Location);
            if (hit == null || hit.Item == null) return;

            var it = hit.Item;
            FontEntry fe;
            if (!_map.TryGetValue(it, out fe)) return;
            if (it.SubItems.Count < 2) return;

            Rectangle nameCol = it.SubItems[1].Bounds;
            int cx = nameCol.Right - 12 - _checkSize;
            int cy = nameCol.Y + (nameCol.Height - _checkSize) / 2;
            var cb = new Rectangle(cx, cy, _checkSize, _checkSize);

            if (cb.Contains(e.Location) || nameCol.Contains(e.Location))
            {
                fe.Checked = !fe.Checked;
                this.Invalidate(nameCol);
            }
        };

        this.ItemSelectionChanged += (s, e) => this.Invalidate();
    }

    private void DrawRoundCheck(Graphics g, Rectangle r, bool isChecked)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;

        using (var b = new SolidBrush(Color.White))
            g.FillEllipse(b, Rectangle.Inflate(r, -1, -1));

        using (var pen = new Pen(isChecked ? _accent : Color.FromArgb(170, 200, 220), 2))
            g.DrawEllipse(pen, r);

        if (isChecked)
        {
            var inner = Rectangle.Inflate(r, -4, -4);
            using (var b2 = new SolidBrush(_accent))
                g.FillEllipse(b2, inner);
        }
    }

    private bool IsInstalledByFamilyName(string familyName)
    {
        try
        {
            var installed = new InstalledFontCollection();
            foreach (var fam in installed.Families)
                if (string.Equals(fam.Name, familyName, StringComparison.InvariantCultureIgnoreCase))
                    return true;
        }
        catch { }
        return false;
    }

    private string PickDict(System.Collections.Generic.IDictionary<System.Globalization.CultureInfo, string> d)
    {
        try
        {
            if (d == null || d.Count == 0) return null;

            foreach (System.Collections.Generic.KeyValuePair<System.Globalization.CultureInfo, string> kv in d)
            {
                if (kv.Key != null)
                {
                    string n = kv.Key.Name ?? "";
                    if (n.StartsWith("he", StringComparison.OrdinalIgnoreCase) || n.StartsWith("en", StringComparison.OrdinalIgnoreCase))
                        return kv.Value;
                }
            }

            foreach (System.Collections.Generic.KeyValuePair<System.Globalization.CultureInfo, string> kv in d)
                return kv.Value;

            return null;
        }
        catch { return null; }
    }

    private string BuildMeta(string path)
    {
        try
        {
            var gt = new System.Windows.Media.GlyphTypeface(new Uri(path, UriKind.Absolute));

            string ver  = PickDict(gt.VersionStrings);
            string manu = PickDict(gt.ManufacturerNames);
            string des  = PickDict(gt.DesignerNames);

            string ext = System.IO.Path.GetExtension(path).ToLowerInvariant().TrimStart('.');
            string kind = (ext == "otf") ? "OTF" : (ext == "ttc" ? "TTC" : "TTF");

            string w = gt.Weight.ToString();
            string st = gt.Style.ToString();

            string meta = kind;
            if (!string.IsNullOrWhiteSpace(ver)) meta += " • גרסה: " + ver;
            meta += " • משקל: " + w + " • סגנון: " + st;

            if (!string.IsNullOrWhiteSpace(manu)) meta += " • מפרסם: " + manu;
            else if (!string.IsNullOrWhiteSpace(des)) meta += " • מעצב: " + des;

            return meta;
        }
        catch { return null; }
    }

    /// <summary>
    /// טוען את הגופנים לרשימה ויוצר עבורם תצוגה מקדימה.
    /// </summary>
    public void LoadFonts(string[] paths)
    {
        ClearAll();
        if (paths == null) return;

        foreach (var p in paths)
        {
            FontEntry entry = null;
            try
            {
                entry = new FontEntry();
                entry.Path = p;
                entry.Checked = false;

                entry.Pfc = new PrivateFontCollection();
                entry.Pfc.AddFontFile(p);

                string name = Path.GetFileNameWithoutExtension(p);
                if (entry.Pfc.Families != null && entry.Pfc.Families.Length > 0)
                {
                    name = entry.Pfc.Families[0].Name;
                    if (string.IsNullOrWhiteSpace(name))
                        name = Path.GetFileNameWithoutExtension(p);

                    entry.PreviewFont = new Font(entry.Pfc.Families[0], 13f, FontStyle.Regular, GraphicsUnit.Point);
                }
                entry.DisplayName = name;
                entry.Meta = BuildMeta(p);
                entry.Installed = IsInstalledByFamilyName(entry.DisplayName);

                var it = new ListViewItem(_sampleText);
                it.SubItems.Add(entry.DisplayName);

                this.Items.Add(it);
                _map[it] = entry;
            }
            catch
            {
                // במקרה של כישלון בקריאת הגופן, נציג אותו כשם קובץ בלבד
                var it = new ListViewItem(_sampleText);
                it.SubItems.Add(Path.GetFileNameWithoutExtension(p));
                this.Items.Add(it);

                if (entry == null)
                {
                    entry = new FontEntry();
                    entry.Path = p;
                    entry.DisplayName = Path.GetFileNameWithoutExtension(p);
                    entry.Checked = false;
                    entry.Installed = false;
                    entry.Meta = BuildMeta(p);
                }
                _map[it] = entry;
            }
        }
        AutoSizeColumns();
    }

    public string[] GetCheckedPaths()
    {
        var list = new List<string>();
        foreach (var kv in _map)
            if (kv.Value.Checked && !string.IsNullOrWhiteSpace(kv.Value.Path))
                list.Add(kv.Value.Path);
        return list.ToArray();
    }

    public void SetAllChecked(bool check)
    {
        foreach (var kv in _map)
            kv.Value.Checked = check;
        this.Invalidate();
    }

    public void AutoSizeColumns()
    {
        if (this.Columns.Count < 2) return;
        this.Columns[1].Width = 400;
        int w = Math.Max(320, this.ClientSize.Width - this.Columns[1].Width - 10);
        this.Columns[0].Width = w;
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        AutoSizeColumns();
    }

    /// <summary>
    /// מנקה משאבי גופנים כדי למנוע נעילות קבצים.
    /// </summary>
    public void ClearAll()
    {
        foreach (var kv in _map)
        {
            try
            {
                if (kv.Value.PreviewFont != null) kv.Value.PreviewFont.Dispose();
                if (kv.Value.Pfc != null) kv.Value.Pfc.Dispose();
            }
            catch { }
        }
        _map.Clear();
        this.Items.Clear();
    }
}
"@

Add-Type -TypeDefinition $csUI -ReferencedAssemblies 'System.Windows.Forms','System.Drawing','PresentationCore','WindowsBase' -Language CSharp
#endregion

#region Application State Management
# משתנים גלובליים לניהול נתיבי הגופנים שנבחרו
$scriptDir = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$pathsMap = @{}

function Add-Path([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }
    if (-not (Test-Path -LiteralPath $p)) { return $false }
    $k = $p.ToLowerInvariant()
    if ($pathsMap.ContainsKey($k)) { return $false }
    $pathsMap[$k] = $p
    return $true
}

function Clear-Paths { 
    $pathsMap.Clear() 
}

function Load-FolderFonts {
    Clear-Paths
    $files = Get-FontFilesInFolder -Folder $scriptDir
    foreach ($f in $files) { [void](Add-Path $f.FullName) }
    return $files.Count
}

function Get-CurrentPathsArray { 
    return ($pathsMap.Values | Sort-Object) 
}
#endregion

#region UI Construction & Theming

# Theme Configuration
$bgApp      = [System.Drawing.Color]::FromArgb(245, 250, 255)
$cardBg     = [System.Drawing.Color]::White
$accent     = [System.Drawing.Color]::FromArgb(0, 150, 255)
$accentSoft = [System.Drawing.Color]::FromArgb(227, 247, 255)
$border     = [System.Drawing.Color]::FromArgb(200, 225, 240)
$textMain   = [System.Drawing.Color]::FromArgb(15, 25, 35)
$textMuted  = [System.Drawing.Color]::FromArgb(80, 95, 110)

<#
.SYNOPSIS
פונקציית עזר ליצירת כפתור מעוגל מותאם אישית (RoundButton)
#>
function New-RoundBtn([string]$text, [int]$w, [int]$h, [bool]$primary=$false) {
    $b = New-Object RoundButton
    $b.Text = $text
    $b.Size = New-Object System.Drawing.Size($w, $h)
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 10, $(if($primary){[System.Drawing.FontStyle]::Bold}else{[System.Drawing.FontStyle]::Regular}))
    $b.CornerRadius = 12

    if ($primary) {
        $b.FillColor  = $accent
        $b.FillHover  = [System.Drawing.Color]::FromArgb(0, 165, 255)
        $b.FillDown   = [System.Drawing.Color]::FromArgb(0, 130, 220)
        $b.ForeColor  = [System.Drawing.Color]::White
        $b.BorderSize = 0
    } else {
        $b.FillColor   = [System.Drawing.Color]::White
        $b.FillHover   = [System.Drawing.Color]::FromArgb(245, 250, 255)
        $b.FillDown    = [System.Drawing.Color]::FromArgb(235, 245, 255)
        $b.ForeColor   = $textMain
        $b.BorderColor = [System.Drawing.Color]::FromArgb(160, 180, 210, 235)
        $b.BorderSize  = 1
    }
    return $b
}

# Form Setup
$form = New-Object System.Windows.Forms.Form
$form.Text = 'התקנת גופנים מהירה - FuntTop'
$form.RightToLeft = 'Yes'
$form.RightToLeftLayout = $true
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(990, 700)
$form.MinimumSize = New-Object System.Drawing.Size(700, 450)
$form.BackColor = $bgApp
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
$notifyIcon.Visible = $true

try {
    $icoPath = Join-Path $scriptDir 'app.ico'
    if (Test-Path -LiteralPath $icoPath) {
        $form.Icon = New-Object System.Drawing.Icon($icoPath)
    }
} catch { }

# Layout Panels
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.RowCount = 4
$root.ColumnCount = 1
$root.BackColor = $bgApp
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
$form.Controls.Add($root)

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Fill'
$header.BackColor = $accentSoft
$header.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
$root.Controls.Add($header, 0, 0)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Dock = 'Top'
$lblTitle.Height = 44
$lblTitle.Text = 'התקנת גופנים מהירה - FuntTop'
$lblTitle.TextAlign = 'MiddleCenter'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $textMain

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Dock = 'Top'
$lblPath.Height = 22
$lblPath.TextAlign = 'MiddleCenter'
$lblPath.Text = $scriptDir
$lblPath.ForeColor = $textMuted

$hdrStrip = New-Object System.Windows.Forms.TableLayoutPanel
$hdrStrip.Dock = 'Top'
$hdrStrip.Height = 24
$hdrStrip.ColumnCount = 2
$hdrStrip.RowCount = 1
$hdrStrip.BackColor = $accentSoft
[void]$hdrStrip.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$hdrStrip.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$lblAdmin = New-Object System.Windows.Forms.Label
$lblAdmin.Dock = 'Fill'
$lblAdmin.TextAlign = 'MiddleLeft'
$lblAdmin.Text = if (Test-IsAdmin) { 'מצב מנהל: כן' } else { 'מצב מנהל: לא' }
$lblAdmin.ForeColor = $textMuted

$lblFound = New-Object System.Windows.Forms.Label
$lblFound.Dock = 'Fill'
$lblFound.TextAlign = 'MiddleRight'
$lblFound.Text = 'נמצאו: 0'
$lblFound.ForeColor = $textMuted

$hdrStrip.Controls.Add($lblAdmin, 0, 0)
$hdrStrip.Controls.Add($lblFound, 1, 0)

$header.Controls.Add($hdrStrip)
$header.Controls.Add($lblPath)
$header.Controls.Add($lblTitle)

# ListView Container
$listPanel = New-Object System.Windows.Forms.Panel
$listPanel.Dock = 'Fill'
$listPanel.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
$listPanel.BackColor = $bgApp
$root.Controls.Add($listPanel, 0, 1)

$card = New-Object System.Windows.Forms.Panel
$card.Dock = 'Fill'
$card.BackColor = $cardBg
$card.Padding = New-Object System.Windows.Forms.Padding(12)
$card.BorderStyle = 'FixedSingle'
$listPanel.Controls.Add($card)

$lv = New-Object FontPreviewListView
$lv.Dock = 'Fill'
$card.Controls.Add($lv)

# Bottom Actions Panel
$buttonsPanel = New-Object System.Windows.Forms.Panel
$buttonsPanel.Dock = 'Fill'
$buttonsPanel.BackColor = $bgApp
$buttonsPanel.Padding = New-Object System.Windows.Forms.Padding(16, 6, 16, 8)
$root.Controls.Add($buttonsPanel, 0, 2)

$btnGrid = New-Object System.Windows.Forms.TableLayoutPanel
$btnGrid.Dock = 'Fill'
$btnGrid.BackColor = $bgApp
$btnGrid.RowCount = 3
$btnGrid.ColumnCount = 3
[void]$btnGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$btnGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$btnGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$btnGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$btnGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$btnGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$buttonsPanel.Controls.Add($btnGrid)

# Options & Checkboxes
$chkSkipInstalled = New-Object System.Windows.Forms.CheckBox
$chkSkipInstalled.Text = 'דלג על גופנים מותקנים ושגיאות בהתקנה'
$chkSkipInstalled.AutoSize = $true
$chkSkipInstalled.Checked = $true
$chkSkipInstalled.Anchor = 'None'
$chkSkipInstalled.ForeColor = $textMain
$chkSkipInstalled.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)

# Flow Layouts for Buttons
$flowTop = New-Object System.Windows.Forms.FlowLayoutPanel
$flowTop.AutoSize = $true
$flowTop.WrapContents = $false
$flowTop.FlowDirection = 'RightToLeft'
$flowTop.RightToLeft = 'Yes'
$flowTop.BackColor = $bgApp
$flowTop.Anchor = 'None'

$flowBottom = New-Object System.Windows.Forms.FlowLayoutPanel
$flowBottom.AutoSize = $true
$flowBottom.WrapContents = $false
$flowBottom.FlowDirection = 'RightToLeft'
$flowBottom.RightToLeft = 'Yes'
$flowBottom.BackColor = $bgApp
$flowBottom.Anchor = 'None'

$btnGrid.Controls.Add($chkSkipInstalled, 1, 0)
$btnGrid.Controls.Add($flowTop, 1, 1)
$btnGrid.Controls.Add($flowBottom, 1, 2)

# Buttons Initialization
$btnAddFonts        = New-RoundBtn 'בחר מיקום גופנים'     140 44 $false
$btnRefresh         = New-RoundBtn 'רענן רשימה'            96 44 $false
$btnSelectNone      = New-RoundBtn 'נקה הכל'              120 44 $false
$btnSelectAll       = New-RoundBtn 'סמן הכל'              120 44 $false

$btnInstallAll      = New-RoundBtn 'התקן הכל'             170 52 $false
$btnInstallSelected = New-RoundBtn 'התקן נבחרים'          220 52 $true

foreach ($c in @($btnAddFonts,$btnRefresh,$btnSelectNone,$btnSelectAll,$btnInstallAll,$btnInstallSelected)) {
    $c.Margin = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
}

$flowTop.Controls.Add($btnSelectAll)
$flowTop.Controls.Add($btnSelectNone)
$flowTop.Controls.Add($btnRefresh)
$flowTop.Controls.Add($btnAddFonts)

$flowBottom.Controls.Add($btnInstallSelected)
$flowBottom.Controls.Add($btnInstallAll)

# Footer Credit
$lblCredit = New-Object System.Windows.Forms.Label
$lblCredit.Dock = 'Fill'
$lblCredit.TextAlign = 'MiddleCenter'
$lblCredit.ForeColor = [System.Drawing.Color]::FromArgb(120, 135, 150)
$lblCredit.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$lblCredit.Text = 'פותח ע''י @מיכאלוש בסיוע AI'
$root.Controls.Add($lblCredit, 0, 3)
#endregion

#region Core Logic & Events

function Pick-FontsDialog {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'בחר קבצי גופנים'
    $dlg.Filter = 'קבצי גופנים (*.ttf;*.otf;*.ttc)|*.ttf;*.otf;*.ttc|כל הקבצים (*.*)|*.*'
    $dlg.Multiselect = $true
    $dlg.CheckFileExists = $true
    $dlg.RestoreDirectory = $true

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $added = 0
        foreach ($f in $dlg.FileNames) {
            if (Add-Path $f) { $added++ }
        }
        return $added
    }
    return 0
}

function Refresh-View {
    $arr = Get-CurrentPathsArray
    $lv.LoadFonts($arr)
    $lblFound.Text = "נמצאו: $($pathsMap.Count)"
}

<#
.SYNOPSIS
מנהל את תהליך ההתקנה, התראות למשתמש וניהול שגיאות.
#>
function Do-Install([string[]]$pathsToInstall) {
    if (-not $pathsToInstall -or $pathsToInstall.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('לא נבחרו גופנים להתקנה.', 'התקנת גופנים', 'OK', 'Information') | Out-Null
        return
    }

    # נעילת ממשק למניעת התנגשויות
    foreach ($c in @($btnInstallSelected,$btnInstallAll,$btnAddFonts,$btnRefresh,$btnSelectAll,$btnSelectNone,$chkSkipInstalled)) { $c.Enabled = $false }

    $ok = 0
    $skip = 0
    $fail = 0
    $failList = New-Object System.Collections.Generic.List[string]

    $lblFound.Text = "מתקין…"
    $form.Refresh()

    $skipErrors = $chkSkipInstalled.Checked
    $msgOptions = [System.Windows.Forms.MessageBoxOptions]'RightAlign, RtlReading'

    foreach ($p in $pathsToInstall) {
        $leaf = Split-Path -Leaf $p
        $fontName = [IO.Path]::GetFileNameWithoutExtension($leaf)
        $pfc = $null
        
        # ניסיון לחלץ את שם הגופן המדויק לצורך הודעת השגיאה למשתמש
        try {
            $pfc = New-Object System.Drawing.Text.PrivateFontCollection
            $pfc.AddFontFile($p)
            if ($pfc.Families.Count -gt 0) { $fontName = $pfc.Families[0].Name }
        } catch {
        } finally {
            if ($null -ne $pfc) { $pfc.Dispose() }
        }

        $force = $false

        try {
            # בדיקה האם הגופן מותקן והצגת התראה (אם לא הוגדר דילוג)
            if (Test-FontInstalledByFile $p) {
                if ($skipErrors) {
                    $skip++
                    continue
                } else {
                    $ans = [System.Windows.Forms.MessageBox]::Show(
                        "הגופן '$fontName' כבר מותקן. האם ברצונך להחליפו?", 
                        "התקנת גופן", 
                        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
                        [System.Windows.Forms.MessageBoxIcon]::Warning, 
                        [System.Windows.Forms.MessageBoxDefaultButton]::Button2, 
                        $msgOptions
                    )
                    
                    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
                        $force = $true
                    } else {
                        $skip++
                        continue
                    }
                }
            }

            Install-FontFile -Path $p -ForceOverwrite:$force
            $ok++
        } catch {
            $fail++
            $failList.Add("$leaf  ->  " + $_.Exception.Message)
            
            if (-not $skipErrors) {
                [System.Windows.Forms.MessageBox]::Show(
                    "שגיאה בהתקנת הגופן '$fontName':`n$($_.Exception.Message)", 
                    "שגיאת התקנה", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error, 
                    [System.Windows.Forms.MessageBoxDefaultButton]::Button1, 
                    $msgOptions
                ) | Out-Null
            }
        }
    }

    # שחרור נעילת הממשק
    foreach ($c in @($btnInstallSelected,$btnInstallAll,$btnAddFonts,$btnRefresh,$btnSelectAll,$btnSelectNone,$chkSkipInstalled)) { $c.Enabled = $true }

    # הצגת הודעת סיכום חכמה
    $summary =
        "הושלם.`r`n" +
        "הותקנו בהצלחה: $ok`r`n" +
        "דולגו (כבר מותקנים / בוטלו): $skip`r`n" +
        "שגיאות: $fail"

    if ($fail -gt 0) {
        $details = "`r`n`r`nפירוט שגיאות (ראשונים עד 25):`r`n" +
                   ($failList | Select-Object -First 25 | ForEach-Object { "• $_" } | Out-String)
        [System.Windows.Forms.MessageBox]::Show($summary + $details, 'התקנת גופנים', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, $msgOptions) | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show($summary, 'התקנת גופנים', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information, [System.Windows.Forms.MessageBoxDefaultButton]::Button1, $msgOptions) | Out-Null
    }

    try { [System.Media.SystemSounds]::Asterisk.Play() } catch { }
    try {
        if ($notifyIcon) {
            $notifyIcon.BalloonTipTitle = 'התקנת גופנים'
            $notifyIcon.BalloonTipText  = "הסתיים. הותקנו: $ok | דולגו: $skip | שגיאות: $fail"
            $notifyIcon.BalloonTipIcon  = if ($fail -gt 0) { [System.Windows.Forms.ToolTipIcon]::Warning } else { [System.Windows.Forms.ToolTipIcon]::Info }
            $notifyIcon.ShowBalloonTip(4000)
        }
    } catch { }

    Refresh-View
}

# חיבור אירועים לכפתורים
$btnSelectAll.Add_Click({ $lv.SetAllChecked($true) })
$btnSelectNone.Add_Click({ $lv.SetAllChecked($false) })
$btnRefresh.Add_Click({ Refresh-View })

$btnAddFonts.Add_Click({
    $added = Pick-FontsDialog
    Refresh-View
    if ($added -gt 0) { $lblFound.Text = "נמצאו: $($pathsMap.Count)  |  נוספו: $added" }
})

$btnInstallSelected.Add_Click({
    $selected = $lv.GetCheckedPaths()
    Do-Install $selected
})

$btnInstallAll.Add_Click({
    $all = Get-CurrentPathsArray
    Do-Install $all
})

# אתחול טעינת גופנים בעת פתיחת החלון
$form.Add_Shown({
    $cnt = Load-FolderFonts
    Refresh-View

    if ($cnt -eq 0) {
        $lblFound.Text = 'נמצאו: 0'
        $form.Refresh()
        $added = Pick-FontsDialog
        Refresh-View
        if ($added -gt 0) { 
            $lblFound.Text = "נמצאו: $($pathsMap.Count)  |  נוספו: $added" 
        } else { 
            $lblFound.Text = 'נמצאו: 0' 
        }
    } else {
        $lblFound.Text = "נמצאו: $cnt"
    }
})

# שחרור משאבים בסגירת החלון
$form.Add_FormClosing({
    try { $lv.ClearAll() } catch { }
    try { if ($notifyIcon) { $notifyIcon.Visible = $false; $notifyIcon.Dispose() } } catch { }
})

# הפעלת האפליקציה
[void]$form.ShowDialog()
#endregion
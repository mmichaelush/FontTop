#requires -version 1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Admin check (ללא העלאה אוטומטית)
# ----------------------------
function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p  = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

# ----------------------------
# Helpers
# ----------------------------
function Get-FontFilesInFolder([string]$Folder) {
    if (-not (Test-Path -LiteralPath $Folder)) { return @() }
    Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.ttf', '.otf', '.ttc') } |
        Sort-Object Name
}

function Install-FontFile {
    param([Parameter(Mandatory)] [string] $Path)

    $fontsFolder = 0x14  # CSIDL_FONTS
    $shell = New-Object -ComObject Shell.Application
    $dst   = $shell.NameSpace($fontsFolder)
    if ($null -eq $dst) { throw 'לא הצלחתי לגשת לתיקיית Fonts של Windows.' }

    $srcDir  = Split-Path -Parent $Path
    $srcName = Split-Path -Leaf $Path
    $srcNS   = $shell.NameSpace($srcDir)
    if ($null -eq $srcNS) { throw 'לא הצלחתי לגשת לתיקיית המקור.' }

    $item = $srcNS.ParseName($srcName)
    if ($null -eq $item) { throw "לא הצלחתי לקרוא את הקובץ: $srcName" }

    # silent + no confirm (אבל אם קיים זה עדיין יכול להקפיץ החלפה—לכן דולגים מראש)
    $dst.CopyHere($item, 0x10 -bor 0x04)
    Start-Sleep -Milliseconds 160
}

# בדיקה אם הגופן מותקן כבר (לדלג)
function Test-FontInstalledByFile([string]$fontPath) {
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
    } catch { }

    # fallback: לפי שם קובץ ב-Windows\Fonts
    try {
        $leaf = Split-Path -Leaf $fontPath
        $dst = Join-Path $env:WINDIR ("Fonts\" + $leaf)
        return (Test-Path -LiteralPath $dst)
    } catch { return $false }
}

# ----------------------------
# C# Controls: RoundButton + FontPreviewListView
# ----------------------------
$cs = @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Windows.Forms;

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
        using (var br = new SolidBrush(fill))
        {
            e.Graphics.FillPath(br, gp);

            if (BorderSize > 0)
            {
                using (var pen = new Pen(BorderColor, BorderSize))
                    e.Graphics.DrawPath(pen, gp);
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
    }

    private readonly Dictionary<ListViewItem, FontEntry> _map = new Dictionary<ListViewItem, FontEntry>();

    private readonly Color _accent = Color.FromArgb(0, 150, 255);
    private readonly Color _accentSoft = Color.FromArgb(227, 247, 255);
    private readonly Color _rowAlt = Color.FromArgb(246, 252, 255);
    private readonly Color _grid = Color.FromArgb(220, 235, 245);
    private readonly Color _text = Color.FromArgb(20, 25, 35);
    private readonly Color _muted = Color.FromArgb(80, 95, 110);
    private readonly Color _good = Color.FromArgb(0, 140, 80);

    private string _sampleText =
        "אבגדהוזחטיכךלמםנןסעפףצץקרשת  | ABC abc 123 | ניקוד: שָׁלוֹם";

    private int _rowHeight = 62;
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

        var il = new ImageList();
        il.ImageSize = new Size(1, _rowHeight);
        this.SmallImageList = il;

        // RTL ListView reverses visual order; add Preview then Font to get: Font on RIGHT, Preview on LEFT
        this.Columns.Add("תצוגה מקדימה", 560, HorizontalAlignment.Left); // LEFT
        this.Columns.Add("שם גופן", 300, HorizontalAlignment.Right);        // RIGHT

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
                fmt.Alignment = StringAlignment.Center; // כותרות ממורכזות

                var r = new Rectangle(e.Bounds.X + 10, e.Bounds.Y, e.Bounds.Width - 20, e.Bounds.Height);
                e.Graphics.DrawString(e.Header.Text, this.Font, b, r, fmt);
            }
        };

        this.DrawItem += (s, e) => { };

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

            if (e.ColumnIndex == 1)
            {
                Rectangle col = e.Bounds;

                int cx = col.Right - 12 - _checkSize;
                int cy = col.Y + (col.Height - _checkSize) / 2;
                var cb = new Rectangle(cx, cy, _checkSize, _checkSize);

                DrawRoundCheck(e.Graphics, cb, fe != null && fe.Checked);

                var textRect = new Rectangle(col.X + 10, col.Y + 6, col.Width - (10 + _checkSize + 18), col.Height - 12);

                string name = e.SubItem.Text;
                string status = (fe != null && fe.Installed) ? "מותקן" : "לא מותקן";

                using (var b1 = new SolidBrush(_text))
                using (var b2 = new SolidBrush((fe != null && fe.Installed) ? _good : _muted))
                {
                    var fmt = new StringFormat();
                    fmt.Alignment = StringAlignment.Far;
                    fmt.LineAlignment = StringAlignment.Near;

                    var rName = new Rectangle(textRect.X, textRect.Y, textRect.Width, (int)(textRect.Height * 0.62));
                    var rStat = new Rectangle(textRect.X, textRect.Y + (int)(textRect.Height * 0.58), textRect.Width, (int)(textRect.Height * 0.42));

                    e.Graphics.DrawString(name, this.Font, b1, rName, fmt);

                    using (var fSmall = new Font(this.Font.FontFamily, 8.5f, FontStyle.Regular))
                        e.Graphics.DrawString(status, fSmall, b2, rStat, fmt);
                }
            }
            else
            {
                Rectangle col = e.Bounds;
                var previewRect = new Rectangle(col.X + 10, col.Y + 6, col.Width - 20, col.Height - 12);

                var fmt = new StringFormat();
                fmt.LineAlignment = StringAlignment.Center;
                fmt.Alignment = StringAlignment.Near;

                Font f = (fe != null && fe.PreviewFont != null) ? fe.PreviewFont : this.Font;
                using (var b = new SolidBrush(_text))
                    e.Graphics.DrawString(e.SubItem.Text, f, b, previewRect, fmt);
            }
        };

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

                entry.Installed = IsInstalledByFamilyName(entry.DisplayName);

                var it = new ListViewItem(_sampleText);
                it.SubItems.Add(entry.DisplayName);

                this.Items.Add(it);
                _map[it] = entry;
            }
            catch
            {
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
        this.Columns[1].Width = 300;
        int w = Math.Max(320, this.ClientSize.Width - this.Columns[1].Width - 10);
        this.Columns[0].Width = w;
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        AutoSizeColumns();
    }

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

Add-Type -TypeDefinition $cs -ReferencedAssemblies 'System.Windows.Forms','System.Drawing' -Language CSharp

# ----------------------------
# Paths storage (PS 5.1 safe)
# ----------------------------
$scriptDir = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$pathsMap = @{}  # key=lowercased path, value=original path

function Add-Path([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return $false }
    if (-not (Test-Path -LiteralPath $p)) { return $false }
    $k = $p.ToLowerInvariant()
    if ($pathsMap.ContainsKey($k)) { return $false }
    $pathsMap[$k] = $p
    return $true
}

function Clear-Paths { $pathsMap.Clear() }

function Load-FolderFonts {
    Clear-Paths
    $files = Get-FontFilesInFolder -Folder $scriptDir
    foreach ($f in $files) { [void](Add-Path $f.FullName) }
    return $files.Count
}

function Get-CurrentPathsArray { return ($pathsMap.Values | Sort-Object) }

# ----------------------------
# Theme / UI constants
# ----------------------------
$bgApp      = [System.Drawing.Color]::FromArgb(245, 250, 255)  # תכלת בהיר כללי
$cardBg     = [System.Drawing.Color]::White
$accent     = [System.Drawing.Color]::FromArgb(0, 150, 255)
$accentSoft = [System.Drawing.Color]::FromArgb(227, 247, 255)
$border     = [System.Drawing.Color]::FromArgb(200, 225, 240)
$textMain   = [System.Drawing.Color]::FromArgb(15, 25, 35)
$textMuted  = [System.Drawing.Color]::FromArgb(80, 95, 110)

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
        $b.BorderColor = [System.Drawing.Color]::FromArgb(120, 170, 210)
        $b.BorderSize  = 2
    }
    return $b
}

# ----------------------------
# Form
# ----------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'התקנת גופנים מהירה'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(950, 700)
$form.MinimumSize = New-Object System.Drawing.Size(750, 500)
$form.BackColor = $bgApp
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.RightToLeft = 'Yes'
$form.RightToLeftLayout = $true

# Icon for title bar + taskbar
try {
    $icoPath = Join-Path $scriptDir 'app.ico'
    if (Test-Path -LiteralPath $icoPath) {
        $form.Icon = New-Object System.Drawing.Icon($icoPath)
    }
} catch { }

# Root layout: 4 rows (Header, List, Buttons, Credit)
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.RowCount = 4
$root.ColumnCount = 1
$root.BackColor = $bgApp
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 22)))
$form.Controls.Add($root)

# ----------------------------
# Header (כותרת מעל נתיב + שורת סטטוס)
# ----------------------------
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Fill'
$header.BackColor = $accentSoft
$header.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
$root.Controls.Add($header, 0, 0)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Dock = 'Top'
$lblTitle.Height = 44
$lblTitle.Text = 'התקנת גופנים מהירה'
$lblTitle.TextAlign = 'MiddleCenter'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $textMain

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Dock = 'Top'
$lblPath.Height = 22
$lblPath.TextAlign = 'MiddleCenter'
$lblPath.Text = $scriptDir
$lblPath.ForeColor = $textMuted

# סטטוס: מצב מנהל + נמצאו X
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

# חשוב: סדר Dock נכון (Top) => מוסיפים מלמטה למעלה
$header.Controls.Add($hdrStrip)
$header.Controls.Add($lblPath)
$header.Controls.Add($lblTitle)

# ----------------------------
# List area (כמו "כרטיס" לבן)
# ----------------------------
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

# ----------------------------
# Buttons: בדיוק 2 שורות, ממורכז
# ----------------------------
$buttonsPanel = New-Object System.Windows.Forms.Panel
$buttonsPanel.Dock = 'Fill'
$buttonsPanel.BackColor = $bgApp
$buttonsPanel.Padding = New-Object System.Windows.Forms.Padding(16, 6, 16, 8)
$root.Controls.Add($buttonsPanel, 0, 2)

$btnGrid = New-Object System.Windows.Forms.TableLayoutPanel
$btnGrid.Dock = 'Fill'
$btnGrid.BackColor = $bgApp
$btnGrid.RowCount = 2
$btnGrid.ColumnCount = 3
[void]$btnGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$btnGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$btnGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$btnGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$btnGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$buttonsPanel.Controls.Add($btnGrid)

$flowTop = New-Object System.Windows.Forms.FlowLayoutPanel
$flowTop.AutoSize = $true
$flowTop.WrapContents = $false      # <<< הכי חשוב: לא נשבר
$flowTop.FlowDirection = 'RightToLeft'
$flowTop.RightToLeft = 'Yes'
$flowTop.BackColor = $bgApp
$flowTop.Anchor = 'None'

$flowBottom = New-Object System.Windows.Forms.FlowLayoutPanel
$flowBottom.AutoSize = $true
$flowBottom.WrapContents = $false   # <<< הכי חשוב: לא נשבר
$flowBottom.FlowDirection = 'RightToLeft'
$flowBottom.RightToLeft = 'Yes'
$flowBottom.BackColor = $bgApp
$flowBottom.Anchor = 'None'

$btnGrid.Controls.Add($flowTop, 1, 0)
$btnGrid.Controls.Add($flowBottom, 1, 1)

# Buttons
$btnAddFonts        = New-RoundBtn 'בחר מיקום גופנים'     140 44 $false
$btnRefresh         = New-RoundBtn 'רענן רשימה'            96 44 $false
$btnSelectNone      = New-RoundBtn 'נקה הכל'        120 44 $false
$btnSelectAll       = New-RoundBtn 'סמן הכל'        120 44 $false

$btnInstallAll      = New-RoundBtn 'התקן הכל'       170 52 $false
$btnInstallSelected = New-RoundBtn 'התקן נבחרים'    220 52 $true

foreach ($c in @($btnAddFonts,$btnRefresh,$btnSelectNone,$btnSelectAll,$btnInstallAll,$btnInstallSelected)) {
    $c.Margin = New-Object System.Windows.Forms.Padding(10, 6, 10, 6)
}

# שורה 1
$flowTop.Controls.Add($btnSelectAll)
$flowTop.Controls.Add($btnSelectNone)
$flowTop.Controls.Add($btnRefresh)
$flowTop.Controls.Add($btnAddFonts)

# שורה 2
$flowBottom.Controls.Add($btnInstallSelected)
$flowBottom.Controls.Add($btnInstallAll)

# ----------------------------
# Credit
# ----------------------------
$lblCredit = New-Object System.Windows.Forms.Label
$lblCredit.Dock = 'Fill'
$lblCredit.TextAlign = 'MiddleCenter'
$lblCredit.ForeColor = [System.Drawing.Color]::FromArgb(120, 135, 150)
$lblCredit.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$lblCredit.Text = 'פותח ע''י @מיכאלוש בסיוע AI'
$root.Controls.Add($lblCredit, 0, 3)

# ----------------------------
# Dialog for picking fonts
# ----------------------------
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

# ----------------------------
# Install logic
# ----------------------------
function Do-Install([string[]]$pathsToInstall) {
    if (-not $pathsToInstall -or $pathsToInstall.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('לא נבחרו גופנים להתקנה.', 'התקנת גופנים', 'OK', 'Information') | Out-Null
        return
    }

    foreach ($c in @($btnInstallSelected,$btnInstallAll,$btnAddFonts,$btnRefresh,$btnSelectAll,$btnSelectNone)) { $c.Enabled = $false }

    $ok = 0
    $skip = 0
    $fail = 0
    $failList = New-Object System.Collections.Generic.List[string]

    $lblFound.Text = "מתקין…"
    $form.Refresh()

    foreach ($p in $pathsToInstall) {
        $leaf = Split-Path -Leaf $p
        try {
            if (Test-FontInstalledByFile $p) {
                $skip++
                continue
            }

            Install-FontFile -Path $p
            $ok++
        } catch {
            $fail++
            $failList.Add("$leaf  ->  " + $_.Exception.Message)
        }
    }

    foreach ($c in @($btnInstallSelected,$btnInstallAll,$btnAddFonts,$btnRefresh,$btnSelectAll,$btnSelectNone)) { $c.Enabled = $true }

    $summary =
        "הושלם.`r`n" +
        "הותקנו בהצלחה: $ok`r`n" +
        "דולגו (כבר מותקנים): $skip`r`n" +
        "שגיאות: $fail"

    if ($fail -gt 0) {
        $details = "`r`n`r`nפירוט שגיאות (ראשונים עד 25):`r`n" +
                   ($failList | Select-Object -First 25 | ForEach-Object { "• $_" } | Out-String)
        [System.Windows.Forms.MessageBox]::Show($summary + $details, 'התקנת גופנים', 'OK', 'Warning') | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show($summary, 'התקנת גופנים', 'OK', 'Information') | Out-Null
    }

    Refresh-View
}

# ----------------------------
# Events
# ----------------------------
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

$form.Add_Shown({
    $cnt = Load-FolderFonts
    Refresh-View

    if ($cnt -eq 0) {
        $lblFound.Text = 'נמצאו: 0'
        $form.Refresh()
        $added = Pick-FontsDialog
        Refresh-View
        if ($added -gt 0) { $lblFound.Text = "נמצאו: $($pathsMap.Count)  |  נוספו: $added" } else { $lblFound.Text = 'נמצאו: 0' }
    } else {
        $lblFound.Text = "נמצאו: $cnt"
    }
})

$form.Add_FormClosing({
    try { $lv.ClearAll() } catch { }
})

[void]$form.ShowDialog()

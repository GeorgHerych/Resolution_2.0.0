using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

// PdfiumViewer — перегляд PDF (потрібен рідний pdfium.dll під вашу архітектуру)
using PdfiumDoc = PdfiumViewer.PdfDocument;

// PdfSharp — фактичне вставлення зображення в PDF
using XGraphics = PdfSharp.Drawing.XGraphics;
using XImage = PdfSharp.Drawing.XImage;

#if NETFRAMEWORK || WINDOWS
using WordInterop = Microsoft.Office.Interop.Word;
#endif

namespace Resolution
{
    public partial class MainForm : Form
    {
        // ===== Локалізація місяців =====
        private static readonly string[] UA_MONTHS_GEN = {
            "", "січня","лютого","березня","квітня","травня","червня",
            "липня","серпня","вересня","жовтня","листопада","грудня"
        };
        private static readonly Dictionary<string, int> UA_MONTHS_REV =
            UA_MONTHS_GEN.Select((m, i) => (m, i)).Where(t => t.m != "")
                         .ToDictionary(t => t.m.ToLower(), t => t.i);

        // ===== Параметри тексту реєстраційного штампа =====
        private const float BASELINE_FACTOR = 0.70f;
        private static readonly Color TEXT_COLOR = Color.Black;
        private const float TEXT_STROKE_REG = 1.0f; // товщина обводки для реєстрації
        private const bool USE_BOLD_FONT_REG = false;

        // ===== Фіксовані поля реєстраційного штампа (ваші координати) =====
        private class FieldCfg { public string Type; public float X; public float Y; public float RelFontH; }
        private readonly Dictionary<string, FieldCfg> FIELDS_REG = new Dictionary<string, FieldCfg>
        {
            { "sheets", new FieldCfg{ Type="center",   X=0.175f, Y=0.110f, RelFontH=0.135f } },
            { "doc",    new FieldCfg{ Type="left_line",X=0.705f, Y=0.235f, RelFontH=0.145f } },
            { "day",    new FieldCfg{ Type="center",   X=0.125f, Y=0.580f, RelFontH=0.135f } },
            { "month",  new FieldCfg{ Type="left_line",X=0.225f, Y=0.610f, RelFontH=0.145f } },
            { "year",   new FieldCfg{ Type="left_line",X=0.778f, Y=0.600f, RelFontH=0.145f } },
        };

        // ===== Поля для "Резолюції" (відносні координати) =====
        // Малюємо все кодом на прозорому полотні розміром як stamp.png
        private readonly Dictionary<string, FieldCfg> FIELDS_RES = new Dictionary<string, FieldCfg>
        {
            { "title", new FieldCfg{ Type="center",    X=0.50f,  Y=0.08f,  RelFontH=0.085f } }, // рядок замість "НСЧ"
            // Пояснювальні надписи/галочки вгорі
            { "inorderLbl", new FieldCfg{ Type="left_line", X=0.05f, Y=0.24f, RelFontH=0.06f } },
            { "refuseLbl",  new FieldCfg{ Type="left_line", X=0.52f, Y=0.24f, RelFontH=0.06f } },
            // Текст під підписом командира (2 рядки)
            { "cmdr1", new FieldCfg{ Type="left_line", X=0.05f, Y=0.36f, RelFontH=0.070f } }, // "Командир військової частини…"
            { "rank",  new FieldCfg{ Type="left_line", X=0.05f, Y=0.50f, RelFontH=0.070f } }, // "підполковник"
        };

        // ===== Стан вільного позиціонування (для обох типів штампів) =====
        private bool _freePosition = false;
        private bool _dragging = false;
        private enum DragTarget { None, Reg, Res }
        private DragTarget _dragTarget = DragTarget.None;
        private float _posRegXPct = float.NaN, _posRegYPct = float.NaN;
        private float _posResXPct = float.NaN, _posResYPct = float.NaN;
        private int _dragOffsetX, _dragOffsetY;

        // Для хіт-тесту
        private int _lastPreviewW, _lastPreviewH;
        private int _lastStampRegW, _lastStampRegH, _lastStampRegX, _lastStampRegY;
        private int _lastStampResW, _lastStampResH, _lastStampResX, _lastStampResY;

        // ===== Стан документу =====
        private PdfiumDoc _pdfDoc;         // для попереднього перегляду
        private string _sourcePath;        // оригінальний шлях (PDF/DOC/DOCX)
        private string _previewPdfPath;    // якщо Word — тимчасовий PDF
        private string _tempDir;

        // ===== UI =====
        private TextBox tbFile;
        private Label lblSrcType;

        // перемикач типу штампа
        private RadioButton rbReg, rbRes;

        // -- реєстраційний
        private PictureBox pbStamp;
        private TextBox tbSheets, tbDocNo, tbDay, tbMonth, tbYear;
        private ComboBox cbRotate;
        private NumericUpDown nudWidthRatio, nudRightMm, nudBottomMm;
        private CheckBox cbFirstPageOnly, cbFreePos, cbDoubleStamp;
        private Button btnToday, btnSave;
        // -- резолюція
        private TextBox tbResTitle, tbResCmdr1, tbResRank;
        private CheckBox cbInOrder, cbRefuse;
        private NumericUpDown nudResPt, nudResLinePx;
        private CheckBox cbResThin;

        // перегляд
        private Panel pnlViewer;
        private PictureBox pbPreview;
        private NumericUpDown nudPage;
        private TrackBar tbZoom;

        // спліт (ліва — перегляд, справа — меню)
        private SplitContainer split;

        public MainForm()
        {
            BuildUi();
            Shown += (s, e) =>
            {
                // Після показу форми безпечно встановити початковий SplitterDistance
                split.SplitterDistance = Math.Max(split.Panel1MinSize, split.Width - Math.Max(split.Panel2MinSize, 420));
            };
            RenderAll();
        }

        // ============================ UI ============================
        private void BuildUi()
        {
            Text = "PDF/Word Stamp Tool (UA) — штамп і резолюція (перетягування мишею)";
            Width = 1120; Height = 780;

            split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterWidth = 6,
                Panel1MinSize = 0,
                Panel2MinSize = 0
            };
            Controls.Add(split);

            // ---- Ліва панель (перегляд з прокруткою)
            pnlViewer = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.DimGray };
            pbPreview = new PictureBox { SizeMode = PictureBoxSizeMode.Normal, BackColor = Color.Gray };
            pbPreview.MouseDown += PbPreview_MouseDown;
            pbPreview.MouseMove += PbPreview_MouseMove;
            pbPreview.MouseUp += PbPreview_MouseUp;
            pnlViewer.Controls.Add(pbPreview);
            split.Panel1.Controls.Add(pnlViewer);

            // ---- Права панель (меню стовпчиком)
            var column = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(6)
            };
            split.Panel2.Controls.Add(column);

            // Файл
            var grpFile = new GroupBox { Text = "Файл", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var flFile = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            flFile.Controls.Add(new Label { Text = "PDF / DOCX / DOC:", AutoSize = true, Margin = new Padding(6, 9, 6, 3) });
            tbFile = new TextBox { Width = 300 };
            var btnBrowse = new Button { Text = "Обрати…" };
            lblSrcType = new Label { Text = "—", AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(8, 9, 6, 3) };
            btnBrowse.Click += (s, e) => ChooseInput();
            flFile.Controls.Add(tbFile);
            flFile.Controls.Add(btnBrowse);
            flFile.Controls.Add(lblSrcType);
            grpFile.Controls.Add(flFile);
            column.Controls.Add(grpFile);

            // Тип штампа
            var grpType = new GroupBox { Text = "Тип штампа", AutoSize = true };
            var flType = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            rbReg = new RadioButton { Text = "Штамп реєстрації", Checked = true, AutoSize = true };
            rbRes = new RadioButton { Text = "Резолюція", AutoSize = true, Margin = new Padding(12, 3, 3, 3) };
            rbReg.CheckedChanged += (s, e) => { RenderAll(); };
            rbRes.CheckedChanged += (s, e) => { RenderAll(); };
            flType.Controls.Add(rbReg);
            flType.Controls.Add(rbRes);
            grpType.Controls.Add(flType);
            column.Controls.Add(grpType);

            // Прев’ю самого штампа (маленьке)
            var grpStampPrev = new GroupBox { Text = "Прев’ю штампа", AutoSize = true };
            pbStamp = new PictureBox { Width = 308, Height = 197, SizeMode = PictureBoxSizeMode.Zoom, Margin = new Padding(8) };
            grpStampPrev.Controls.Add(pbStamp);
            column.Controls.Add(grpStampPrev);

            // Поля реєстраційного штампа
            var grpReg = new GroupBox { Text = "Реєстраційний — поля", AutoSize = true };
            var flReg = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            flReg.Controls.Add(new Label { Text = "Аркушів “ ”:", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            tbSheets = new TextBox { Width = 60, Text = "1" };
            flReg.Controls.Add(tbSheets);

            flReg.Controls.Add(new Label { Text = "Вх. №:", AutoSize = true, Margin = new Padding(12, 8, 6, 3) });
            tbDocNo = new TextBox { Width = 140 };
            flReg.Controls.Add(tbDocNo);

            flReg.Controls.Add(new Label { Text = "Дата:", AutoSize = true, Margin = new Padding(12, 8, 6, 3) });
            var today = DateTime.Today;
            tbDay = new TextBox { Width = 36, Text = today.Day.ToString("00") };
            tbMonth = new TextBox { Width = 110, Text = UA_MONTHS_GEN[today.Month] };
            tbYear = new TextBox { Width = 40, Text = (today.Year % 100).ToString("00") };
            btnToday = new Button { Text = "Сьогодні" };
            btnToday.Click += (s, e) =>
            {
                var d = DateTime.Today;
                tbDay.Text = d.Day.ToString("00");
                tbMonth.Text = UA_MONTHS_GEN[d.Month];
                tbYear.Text = (d.Year % 100).ToString("00");
                RenderAll();
            };
            flReg.Controls.Add(tbDay); flReg.Controls.Add(tbMonth); flReg.Controls.Add(tbYear); flReg.Controls.Add(btnToday);
            grpReg.Controls.Add(flReg);
            column.Controls.Add(grpReg);

            // Поля резолюції
            var grpRes = new GroupBox { Text = "Резолюція — текст", AutoSize = true };
            var flRes = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            flRes.Controls.Add(new Label { Text = "Верхній рядок:", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            tbResTitle = new TextBox { Width = 300, Text = " " };
            flRes.Controls.Add(tbResTitle);

            flRes.Controls.Add(new Label { Text = "Командир/посада:", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            tbResCmdr1 = new TextBox { Width = 300, Text = "Командир військової частини A4844" };
            flRes.Controls.Add(tbResCmdr1);

            flRes.Controls.Add(new Label { Text = "Звання:", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            tbResRank = new TextBox { Width = 140, Text = "підполковник" };
            flRes.Controls.Add(tbResRank);

            cbInOrder = new CheckBox { Text = "В наказ ☑", Checked = false, Margin = new Padding(6, 6, 3, 3) };
            cbRefuse = new CheckBox { Text = "Відмова ☑", Checked = false, Margin = new Padding(6, 6, 3, 3) };
            flRes.Controls.Add(cbInOrder); flRes.Controls.Add(cbRefuse);

            grpRes.Controls.Add(flRes);
            column.Controls.Add(grpRes);

            // Формат резолюції
            var grpResFmt = new GroupBox { Text = "Резолюція — формат", AutoSize = true };
            var flResFmt = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            flResFmt.Controls.Add(new Label { Text = "Розмір (pt):", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            nudResPt = new NumericUpDown { Minimum = 10, Maximum = 40, Value = 20, DecimalPlaces = 0, Width = 70 };
            flResFmt.Controls.Add(nudResPt);

            flResFmt.Controls.Add(new Label { Text = "Міжряддя (px):", AutoSize = true, Margin = new Padding(12, 8, 6, 3) });
            nudResLinePx = new NumericUpDown { Minimum = -60, Maximum = 60, Value = -6, DecimalPlaces = 0, Width = 70 };
            flResFmt.Controls.Add(nudResLinePx);

            cbResThin = new CheckBox { Text = "Тонкий шрифт", Checked = true, Margin = new Padding(12, 6, 3, 3) };
            flResFmt.Controls.Add(cbResThin);
            grpResFmt.Controls.Add(flResFmt);
            column.Controls.Add(grpResFmt);

            // Розміщення / поворот / вільне позиціонування
            var grpPlace = new GroupBox { Text = "Розміщення на сторінці", AutoSize = true };
            var flPlace = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            flPlace.Controls.Add(new Label { Text = "Ширина (% стор.):", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            nudWidthRatio = new NumericUpDown { DecimalPlaces = 2, Increment = 0.01M, Minimum = 0.10M, Maximum = 0.80M, Value = 0.45M, Width = 80 };
            flPlace.Controls.Add(nudWidthRatio);

            flPlace.Controls.Add(new Label { Text = "Відступ справа (мм):", AutoSize = true, Margin = new Padding(10, 8, 6, 3) });
            nudRightMm = new NumericUpDown { DecimalPlaces = 1, Increment = 0.5M, Minimum = 0, Maximum = 50, Value = 10M, Width = 80 };
            flPlace.Controls.Add(nudRightMm);

            flPlace.Controls.Add(new Label { Text = "Відступ знизу (мм):", AutoSize = true, Margin = new Padding(10, 8, 6, 3) });
            nudBottomMm = new NumericUpDown { DecimalPlaces = 1, Increment = 0.5M, Minimum = 0, Maximum = 50, Value = 10M, Width = 80 };
            flPlace.Controls.Add(nudBottomMm);

            cbFirstPageOnly = new CheckBox { Text = "Лише 1-ша сторінка", Checked = false, Margin = new Padding(12, 6, 3, 3) };
            flPlace.Controls.Add(cbFirstPageOnly);

            cbDoubleStamp = new CheckBox { Text = "Два штампи", Checked = false, Margin = new Padding(12, 6, 3, 3) };
            flPlace.Controls.Add(cbDoubleStamp);

            flPlace.Controls.Add(new Label { Text = "Поворот (°):", AutoSize = true, Margin = new Padding(10, 8, 6, 3) });
            cbRotate = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70 };
            cbRotate.Items.AddRange(new object[] { 0, 90, 180, 270 });
            cbRotate.SelectedIndex = 0;
            flPlace.Controls.Add(cbRotate);

            cbFreePos = new CheckBox { Text = "Вільне позиціонування (мишею)", Checked = false, Margin = new Padding(12, 6, 3, 3) };
            cbFreePos.CheckedChanged += (s, e) => { _freePosition = cbFreePos.Checked; RenderAll(); };
            flPlace.Controls.Add(cbFreePos);

            grpPlace.Controls.Add(flPlace);
            column.Controls.Add(grpPlace);

            // Сторінка/зум
            var grpPg = new GroupBox { Text = "Перегляд", AutoSize = true };
            var flPg = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            flPg.Controls.Add(new Label { Text = "Сторінка:", AutoSize = true, Margin = new Padding(6, 8, 6, 3) });
            nudPage = new NumericUpDown { Minimum = 1, Maximum = 1, Value = 1, Width = 80 };
            flPg.Controls.Add(nudPage);
            flPg.Controls.Add(new Label { Text = "Зум (%):", AutoSize = true, Margin = new Padding(12, 8, 6, 3) });
            tbZoom = new TrackBar { Minimum = 60, Maximum = 250, TickFrequency = 10, Value = 120, Width = 280 };
            flPg.Controls.Add(tbZoom);
            grpPg.Controls.Add(flPg);
            column.Controls.Add(grpPg);

            // Зберегти
            var flSave = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            btnSave = new Button { Text = "Вставити та зберегти…" };
            btnSave.Click += (s, e) => ProcessExport();
            flSave.Controls.Add(btnSave);
            column.Controls.Add(flSave);

            // реагування на зміни
            foreach (var tb in new[] { tbSheets, tbDocNo, tbDay, tbMonth, tbYear, tbResTitle, tbResCmdr1, tbResRank })
                tb.TextChanged += (s, e) => RenderAll();
            cbInOrder.CheckedChanged += (s, e) => RenderAll();
            cbRefuse.CheckedChanged += (s, e) => RenderAll();
            nudResPt.ValueChanged += (s, e) => RenderAll();
            nudResLinePx.ValueChanged += (s, e) => RenderAll();
            cbResThin.CheckedChanged += (s, e) => RenderAll();

            nudWidthRatio.ValueChanged += (s, e) => RenderAll();
            nudRightMm.ValueChanged += (s, e) => RenderAll();
            nudBottomMm.ValueChanged += (s, e) => RenderAll();
            cbFirstPageOnly.CheckedChanged += (s, e) => RenderAll();
            cbDoubleStamp.CheckedChanged += (s, e) => RenderAll();
            cbRotate.SelectedIndexChanged += (s, e) => RenderAll();
            tbZoom.Scroll += (s, e) => RenderAll();
            nudPage.ValueChanged += (s, e) => RenderAll();
        }

        // ============================ Utils ============================
        private static string AppDir()
            => Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? Environment.CurrentDirectory;

        private static float MmToPt(float mm) => (float)(mm * 72.0 / 25.4);

        private static Bitmap LoadPng(string name)
        {
            string p = Path.Combine(AppDir(), name);
            if (!File.Exists(p)) throw new FileNotFoundException($"Не знайдено '{name}' поруч із .exe");
            return (Bitmap)System.Drawing.Image.FromFile(p);
        }

        private static Bitmap LoadStampPng() => LoadPng("stamp.png");

        private static int SafeInt(string s, int fallback)
            => int.TryParse(s.Trim(), out var v) ? v : fallback;

        private static SizeF MeasureString(Graphics g, Font f, string text)
        {
            var path = new GraphicsPath();
            path.AddString(text, f.FontFamily, (int)f.Style, f.Size, PointF.Empty, StringFormat.GenericTypographic);
            var b = path.GetBounds();
            return new SizeF(b.Width, b.Height);
        }

        private static void DrawTextWithStroke(Graphics g, string text, Font font, Color color, float strokeW, PointF pt)
        {
            var path = new GraphicsPath();
            path.AddString(text, font.FontFamily, (int)font.Style, font.Size, pt, StringFormat.GenericTypographic);
            if (strokeW > 0f)
                using (var pen = new Pen(color, strokeW) { LineJoin = LineJoin.Round })
                    g.DrawPath(pen, path);
            var brush = new SolidBrush(color);
            g.FillPath(brush, path);
        }

        private static void DrawFieldCenter(Graphics g, Size imgSize, Font f, FieldCfg cfg, string text, float strokeW)
        {
            var size = MeasureString(g, f, text);
            float x = imgSize.Width * cfg.X - size.Width / 2f;
            float y = imgSize.Height * cfg.Y - size.Height / 2f;
            DrawTextWithStroke(g, text, f, TEXT_COLOR, strokeW, new PointF(x, y));
        }
        private static void DrawFieldLeftLine(Graphics g, Size imgSize, Font f, FieldCfg cfg, string text, float strokeW)
        {
            float x = imgSize.Width * cfg.X;
            float y = imgSize.Height * cfg.Y - f.Size * BASELINE_FACTOR; // f.Size в пікселях
            DrawTextWithStroke(g, text, f, TEXT_COLOR, strokeW, new PointF(x, y));
        }

        private static Font TimesPxFromPt(Graphics g, float pt, bool bold = false)
        {
            float px = pt * g.DpiY / 72f;
            return new Font("Times New Roman", px, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
        }

        private Bitmap RotateBitmapCW(Bitmap src, int cw)
        {
            if (cw % 360 == 0) return (Bitmap)src.Clone();

            var pts = new[] { new PointF(0, 0), new PointF(src.Width, 0), new PointF(0, src.Height), new PointF(src.Width, src.Height) };
            var m = new Matrix();
            m.Translate(src.Width / 2f, src.Height / 2f);
            m.Rotate(-cw);
            m.Translate(-src.Width / 2f, -src.Height / 2f);
            var tp = pts.Select(p => { var arr = new[] { p }; m.TransformPoints(arr); return arr[0]; }).ToArray();
            var minX = tp.Min(p => p.X); var maxX = tp.Max(p => p.X);
            var minY = tp.Min(p => p.Y); var maxY = tp.Max(p => p.Y);
            int outW = (int)Math.Ceiling(maxX - minX);
            int outH = (int)Math.Ceiling(maxY - minY);

            var outBmp = new Bitmap(outW, outH, PixelFormat.Format32bppArgb);
            outBmp.SetResolution(src.HorizontalResolution, src.VerticalResolution);
            using (var g = Graphics.FromImage(outBmp))
            {
                g.Clear(Color.Transparent);
                g.TranslateTransform(outW / 2f, outH / 2f);
                g.RotateTransform(-cw);
                g.TranslateTransform(-src.Width / 2f, -src.Height / 2f);
                g.DrawImage(src, 0, 0);
            }
            return outBmp;
        }

        private Size GetBaseStampSize()
        {
            try { var bmp = LoadStampPng(); return bmp.Size; }
            catch { return new Size(900, 560); }
        }

        // ============================ Побудова штампів ============================
        private Bitmap BuildRegistrationStamp()
        {
            var baseStamp = LoadStampPng();
            var img = new Bitmap(baseStamp.Width, baseStamp.Height, PixelFormat.Format32bppArgb);
            img.SetResolution(baseStamp.HorizontalResolution, baseStamp.VerticalResolution);

            using (var g = Graphics.FromImage(img))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                g.DrawImage(baseStamp, 0, 0, baseStamp.Width, baseStamp.Height);

                string sheets = tbSheets.Text.Trim();
                string docNo = tbDocNo.Text.Trim();

                // дата
                int monthNum;
                if (!int.TryParse(tbMonth.Text.Trim(), out monthNum))
                    UA_MONTHS_REV.TryGetValue(tbMonth.Text.Trim().ToLower(), out monthNum);
                if (monthNum < 1 || monthNum > 12) monthNum = DateTime.Today.Month;

                int day = SafeInt(tbDay.Text, DateTime.Today.Day);
                int yTail = SafeInt(tbYear.Text, DateTime.Today.Year % 100);
                int year = (yTail <= 99) ? 2000 + yTail : yTail;
                var date = new DateTime(year, monthNum, Math.Min(day, DateTime.DaysInMonth(year, monthNum)));

                string dayStr = date.Day.ToString("00");
                string monthStr = UA_MONTHS_GEN[date.Month];
                string yearTail = (date.Year % 100).ToString("00");

                // шрифти (пікселі)
                var fonts = FIELDS_REG.ToDictionary(
                    kv => kv.Key,
                    kv =>
                    {
                        float h = kv.Value.RelFontH * baseStamp.Height;
                        var style = USE_BOLD_FONT_REG ? FontStyle.Bold : FontStyle.Regular;
                        try { return new Font("DejaVu Sans", h, style, GraphicsUnit.Pixel); } catch { }
                        try { return new Font("Arial", h, style, GraphicsUnit.Pixel); } catch { }
                        return new Font(FontFamily.GenericSansSerif, h, style, GraphicsUnit.Pixel);
                    });

                DrawFieldCenter(g, baseStamp.Size, fonts["sheets"], FIELDS_REG["sheets"], sheets, TEXT_STROKE_REG);
                DrawFieldLeftLine(g, baseStamp.Size, fonts["doc"], FIELDS_REG["doc"], docNo, TEXT_STROKE_REG);
                DrawFieldCenter(g, baseStamp.Size, fonts["day"], FIELDS_REG["day"], dayStr, TEXT_STROKE_REG);
                DrawFieldLeftLine(g, baseStamp.Size, fonts["month"], FIELDS_REG["month"], monthStr, TEXT_STROKE_REG);
                DrawFieldLeftLine(g, baseStamp.Size, fonts["year"], FIELDS_REG["year"], yearTail, TEXT_STROKE_REG);
            }
            baseStamp.Dispose();

            int deg = (int)cbRotate.SelectedItem;
            var rotated = RotateBitmapCW(img, deg);
            img.Dispose();
            return rotated;
        }

        private Bitmap BuildResolutionStamp()
        {
            // полотно за розміром базового штампа
            var baseSize = GetBaseStampSize();
            var img = new Bitmap(baseSize.Width, baseSize.Height, PixelFormat.Format32bppArgb);

            using (var g = Graphics.FromImage(img))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                // Тонка верхня риска (декор)
                using (var pen = new Pen(Color.Black, 2f))
                {
                    g.DrawLine(pen, 0.04f * baseSize.Width, 0.14f * baseSize.Height, 0.96f * baseSize.Width, 0.14f * baseSize.Height);
                }

                float pt = (float)nudResPt.Value;
                float lineDeltaPx = (float)nudResLinePx.Value;
                float strokeW = cbResThin.Checked ? 0f : 0.8f;

                // Конвертація в пікселі для правильної «посадки на лінію»
                var fTitle = TimesPxFromPt(g, pt);
                var fLine = TimesPxFromPt(g, pt);

                // верхній рядок (замість "НСЧ")
                DrawFieldCenter(g, baseSize, fTitle, FIELDS_RES["title"], tbResTitle.Text.Trim(), strokeW);

                // Підписи "В наказ" / "Відмова" + порожні рамки
                using (var pen = new Pen(Color.Black, 1.4f))
                using (var fSmall = TimesPxFromPt(g, Math.Max(12, pt - 4)))
                {
                    DrawFieldLeftLine(g, baseSize, fSmall, FIELDS_RES["inorderLbl"], "В наказ", strokeW);
                    DrawFieldLeftLine(g, baseSize, fSmall, FIELDS_RES["refuseLbl"], "Відмова", strokeW);

                    // Рамочки праворуч від написів
                    var boxW = 0.11f * baseSize.Width;
                    var boxH = 0.12f * baseSize.Height;
                    var yBox = baseSize.Height * 0.17f;
                    var xBox1 = baseSize.Width * 0.28f;
                    var xBox2 = baseSize.Width * 0.86f - boxW;

                    g.DrawRectangle(pen, xBox1, yBox, boxW, boxH);
                    g.DrawRectangle(pen, xBox2, yBox, boxW, boxH);

                    if (cbInOrder.Checked)
                        g.FillRectangle(Brushes.Black, xBox1 + 4, yBox + 4, boxW - 8, boxH - 8);
                    if (cbRefuse.Checked)
                        g.FillRectangle(Brushes.Black, xBox2 + 4, yBox + 4, boxW - 8, boxH - 8);
                }

                // Міжряддя: коригуємо Y у відносних координатах через пікселі
                float dRel = lineDeltaPx / baseSize.Height;

                var cfgCmdr = FIELDS_RES["cmdr1"];
                var cfgRank = new FieldCfg { Type = "left_line", X = FIELDS_RES["rank"].X, Y = FIELDS_RES["rank"].Y + dRel, RelFontH = FIELDS_RES["rank"].RelFontH };

                DrawFieldLeftLine(g, baseSize, fLine, cfgCmdr, tbResCmdr1.Text.Trim(), strokeW);
                DrawFieldLeftLine(g, baseSize, fLine, cfgRank, tbResRank.Text.Trim(), strokeW);
            }

            int deg = (int)cbRotate.SelectedItem;
            var rotated = RotateBitmapCW(img, deg);
            img.Dispose();
            return rotated;
        }

        private Bitmap BuildCurrentStamp()
            => rbReg.Checked ? BuildRegistrationStamp() : BuildResolutionStamp();

        private void RenderStampPreview()
        {
            try
            {
                var bmp = BuildCurrentStamp();
                pbStamp.Image?.Dispose();
                pbStamp.Image = (Bitmap)bmp.Clone();
                bmp.Dispose();
            }
            catch (Exception ex)
            {
                pbStamp.Image?.Dispose();
                pbStamp.Image = null;
                var img = new Bitmap(308, 197);
                var g = Graphics.FromImage(img);
                g.Clear(Color.White);
                g.DrawString(ex.Message, Font, Brushes.Red, new RectangleF(0, 0, img.Width, img.Height));
                pbStamp.Image = img;
            }
        }

        // ============================ Рендер перегляду ============================
        private void RenderAll()
        {
            RenderStampPreview();

            if (_pdfDoc == null)
            {
                pbPreview.Image?.Dispose();
                pbPreview.Image = null;
                return;
            }

            int pageIndex = Math.Max(0, Math.Min((int)nudPage.Value - 1, _pdfDoc.PageCount - 1));
            var pageSizePt = _pdfDoc.PageSizes[pageIndex];
            float zoom = tbZoom.Value / 100f;

            int bmpW = Math.Max(50, (int)(pageSizePt.Width * zoom));
            int bmpH = Math.Max(50, (int)(pageSizePt.Height * zoom));

            var rendered = _pdfDoc.Render(pageIndex, bmpW, bmpH, 96, 96, PdfiumViewer.PdfRenderFlags.Annotations);
            var baseBmp = new Bitmap(rendered);

            var regBmp = BuildRegistrationStamp();
            var resBmp = BuildResolutionStamp();

            int targetWpx = (int)(pageSizePt.Width * (float)nudWidthRatio.Value * zoom);
            int targetHpx = Math.Max(1, (int)(targetWpx * (regBmp.Height / (float)regBmp.Width)));
            int mrPx = (int)(MmToPt((float)nudRightMm.Value) * zoom);
            int mbPx = (int)(MmToPt((float)nudBottomMm.Value) * zoom);
            int gapPx = (int)(MmToPt(5f) * zoom);

            int xReg = 0, yReg = 0, xRes = 0, yRes = 0;
            if (cbDoubleStamp.Checked)
            {
                if (_freePosition)
                {
                    if (float.IsNaN(_posRegXPct) || float.IsNaN(_posRegYPct))
                    {
                        int defX = baseBmp.Width - mrPx - targetWpx;
                        int defY = baseBmp.Height - mbPx - targetHpx;
                        _posRegXPct = defX / (float)baseBmp.Width;
                        _posRegYPct = defY / (float)baseBmp.Height;
                    }
                    if (float.IsNaN(_posResXPct) || float.IsNaN(_posResYPct))
                    {
                        int defX = (int)(_posRegXPct * baseBmp.Width);
                        int defY = (int)(_posRegYPct * baseBmp.Height) - targetHpx - gapPx;
                        if (defY < 0) defY = (int)(_posRegYPct * baseBmp.Height);
                        _posResXPct = defX / (float)baseBmp.Width;
                        _posResYPct = defY / (float)baseBmp.Height;
                    }
                    xReg = (int)(_posRegXPct * baseBmp.Width);
                    yReg = (int)(_posRegYPct * baseBmp.Height);
                    xRes = (int)(_posResXPct * baseBmp.Width);
                    yRes = (int)(_posResYPct * baseBmp.Height);
                }
                else
                {
                    xReg = baseBmp.Width - mrPx - targetWpx;
                    yReg = baseBmp.Height - mbPx - targetHpx;
                    xRes = xReg;
                    yRes = yReg - targetHpx - gapPx;
                    if (yRes < 0) yRes = yReg;
                    _posRegXPct = xReg / (float)baseBmp.Width;
                    _posRegYPct = yReg / (float)baseBmp.Height;
                    _posResXPct = xRes / (float)baseBmp.Width;
                    _posResYPct = yRes / (float)baseBmp.Height;
                }
            }
            else if (rbReg.Checked)
            {
                if (_freePosition && (float.IsNaN(_posRegXPct) || float.IsNaN(_posRegYPct)))
                {
                    int defX = baseBmp.Width - mrPx - targetWpx;
                    int defY = baseBmp.Height - mbPx - targetHpx;
                    _posRegXPct = defX / (float)baseBmp.Width;
                    _posRegYPct = defY / (float)baseBmp.Height;
                }
                if (_freePosition)
                {
                    xReg = (int)(_posRegXPct * baseBmp.Width);
                    yReg = (int)(_posRegYPct * baseBmp.Height);
                }
                else
                {
                    xReg = baseBmp.Width - mrPx - targetWpx;
                    yReg = baseBmp.Height - mbPx - targetHpx;
                    _posRegXPct = xReg / (float)baseBmp.Width;
                    _posRegYPct = yReg / (float)baseBmp.Height;
                }
            }
            else // резолюція одна
            {
                if (_freePosition && (float.IsNaN(_posResXPct) || float.IsNaN(_posResYPct)))
                {
                    int defX = baseBmp.Width - mrPx - targetWpx;
                    int defY = baseBmp.Height - mbPx - targetHpx;
                    _posResXPct = defX / (float)baseBmp.Width;
                    _posResYPct = defY / (float)baseBmp.Height;
                }
                if (_freePosition)
                {
                    xRes = (int)(_posResXPct * baseBmp.Width);
                    yRes = (int)(_posResYPct * baseBmp.Height);
                }
                else
                {
                    xRes = baseBmp.Width - mrPx - targetWpx;
                    yRes = baseBmp.Height - mbPx - targetHpx;
                    _posResXPct = xRes / (float)baseBmp.Width;
                    _posResYPct = yRes / (float)baseBmp.Height;
                }
            }

            using (var g = Graphics.FromImage(baseBmp))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                if (cbDoubleStamp.Checked)
                {
                    g.DrawImage(regBmp, new Rectangle(xReg, yReg, targetWpx, targetHpx));
                    g.DrawImage(resBmp, new Rectangle(xRes, yRes, targetWpx, targetHpx));
                }
                else if (rbReg.Checked)
                {
                    g.DrawImage(regBmp, new Rectangle(xReg, yReg, targetWpx, targetHpx));
                }
                else
                {
                    g.DrawImage(resBmp, new Rectangle(xRes, yRes, targetWpx, targetHpx));
                }
            }

            _lastPreviewW = baseBmp.Width; _lastPreviewH = baseBmp.Height;
            if (cbDoubleStamp.Checked || rbReg.Checked)
            {
                _lastStampRegW = targetWpx; _lastStampRegH = targetHpx;
                _lastStampRegX = xReg; _lastStampRegY = yReg;
            }
            else
            {
                _lastStampRegW = _lastStampRegH = 0;
            }
            if (cbDoubleStamp.Checked || rbRes.Checked)
            {
                _lastStampResW = targetWpx; _lastStampResH = targetHpx;
                _lastStampResX = xRes; _lastStampResY = yRes;
            }
            else
            {
                _lastStampResW = _lastStampResH = 0;
            }

            pbPreview.Image?.Dispose();
            pbPreview.Image = (Bitmap)baseBmp.Clone();
            pbPreview.Size = baseBmp.Size;

            regBmp.Dispose();
            resBmp.Dispose();
            baseBmp.Dispose();
        }

        // ============================ Перетягування мишею ============================
        private void PbPreview_MouseDown(object sender, MouseEventArgs e)
        {
            if (!_freePosition || _pdfDoc == null || pbPreview.Image == null) return;

            _dragTarget = DragTarget.None;
            if (cbDoubleStamp.Checked)
            {
                var rectReg = new Rectangle(_lastStampRegX, _lastStampRegY, _lastStampRegW, _lastStampRegH);
                var rectRes = new Rectangle(_lastStampResX, _lastStampResY, _lastStampResW, _lastStampResH);
                if (rectReg.Contains(e.Location))
                {
                    _dragTarget = DragTarget.Reg;
                    _dragging = true;
                    _dragOffsetX = e.X - _lastStampRegX;
                    _dragOffsetY = e.Y - _lastStampRegY;
                }
                else if (rectRes.Contains(e.Location))
                {
                    _dragTarget = DragTarget.Res;
                    _dragging = true;
                    _dragOffsetX = e.X - _lastStampResX;
                    _dragOffsetY = e.Y - _lastStampResY;
                }
            }
            else if (rbReg.Checked)
            {
                var rect = new Rectangle(_lastStampRegX, _lastStampRegY, _lastStampRegW, _lastStampRegH);
                if (rect.Contains(e.Location))
                {
                    _dragTarget = DragTarget.Reg;
                    _dragging = true;
                    _dragOffsetX = e.X - _lastStampRegX;
                    _dragOffsetY = e.Y - _lastStampRegY;
                }
            }
            else
            {
                var rect = new Rectangle(_lastStampResX, _lastStampResY, _lastStampResW, _lastStampResH);
                if (rect.Contains(e.Location))
                {
                    _dragTarget = DragTarget.Res;
                    _dragging = true;
                    _dragOffsetX = e.X - _lastStampResX;
                    _dragOffsetY = e.Y - _lastStampResY;
                }
            }
        }
        private void PbPreview_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging || !_freePosition || pbPreview.Image == null) return;
            if (_dragTarget == DragTarget.None) return;

            int w = _dragTarget == DragTarget.Reg ? _lastStampRegW : _lastStampResW;
            int h = _dragTarget == DragTarget.Reg ? _lastStampRegH : _lastStampResH;

            int x = e.X - _dragOffsetX;
            int y = e.Y - _dragOffsetY;

            x = Math.Max(0, Math.Min(_lastPreviewW - w, x));
            y = Math.Max(0, Math.Min(_lastPreviewH - h, y));

            if (_dragTarget == DragTarget.Reg)
            {
                _posRegXPct = (float)x / _lastPreviewW;
                _posRegYPct = (float)y / _lastPreviewH;
            }
            else if (_dragTarget == DragTarget.Res)
            {
                _posResXPct = (float)x / _lastPreviewW;
                _posResYPct = (float)y / _lastPreviewH;
            }

            RenderAll();
        }
        private void PbPreview_MouseUp(object sender, MouseEventArgs e)
        {
            _dragging = false;
            _dragTarget = DragTarget.None;
        }

        // ============================ Робота з файлами ============================
        private void ChooseInput()
        {
            var ofd = new OpenFileDialog
            {
                Title = "Оберіть файл",
                Filter = "Supported|*.pdf;*.doc;*.docx|PDF|*.pdf|Word|*.doc;*.docx"
            };
            if (ofd.ShowDialog(this) != DialogResult.OK) return;

            _sourcePath = ofd.FileName;
            tbFile.Text = _sourcePath;
            CleanupTemp();

            try
            {
                string ext = Path.GetExtension(_sourcePath).ToLowerInvariant();
                if (ext == ".doc" || ext == ".docx")
                {
                    _tempDir = Path.Combine(Path.GetTempPath(), "stamp_preview_" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(_tempDir);
                    _previewPdfPath = Path.Combine(_tempDir, "preview.pdf");
                    ConvertWordToPdf(_sourcePath, _previewPdfPath);
                    lblSrcType.Text = "Тип: Word → PDF (перетворено)";
                }
                else
                {
                    _previewPdfPath = _sourcePath;
                    lblSrcType.Text = "Тип: PDF";
                }

                _pdfDoc?.Dispose();
                _pdfDoc = PdfiumViewer.PdfDocument.Load(_previewPdfPath);
                nudPage.Maximum = _pdfDoc.PageCount;
                nudPage.Value = 1;

                // скидаємо вільну позицію
                _posRegXPct = _posRegYPct = _posResXPct = _posResYPct = float.NaN;

                RenderAll();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Не вдалося підготувати документ:\n" + ex.Message, "Помилка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _pdfDoc = null;
                _previewPdfPath = null;
                lblSrcType.Text = "—";
            }
        }

        private void ProcessExport()
        {
            if (string.IsNullOrEmpty(_previewPdfPath))
            {
                MessageBox.Show(this, "Спочатку оберіть PDF або Word-файл.", "Помилка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            byte[] pngMain;
            byte[] pngSecond = null;
            if (cbDoubleStamp.Checked)
            {
                var bmpReg = BuildRegistrationStamp();
                var bmpRes = BuildResolutionStamp();
                using (var ms = new MemoryStream())
                {
                    bmpReg.Save(ms, ImageFormat.Png);
                    pngMain = ms.ToArray();
                }
                using (var ms = new MemoryStream())
                {
                    bmpRes.Save(ms, ImageFormat.Png);
                    pngSecond = ms.ToArray();
                }
                bmpReg.Dispose();
                bmpRes.Dispose();
            }
            else
            {
                var bmp = BuildCurrentStamp();
                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Png);
                    pngMain = ms.ToArray();
                }
                bmp.Dispose();
            }

            string baseDir = Path.GetDirectoryName(_sourcePath) ?? Environment.CurrentDirectory;
            string suggested = (!string.IsNullOrWhiteSpace(tbDocNo.Text) ? tbDocNo.Text.Trim()
                : Path.GetFileNameWithoutExtension(_sourcePath) + "_stamped") + ".pdf";

            var sfd = new SaveFileDialog
            {
                Title = "Зберегти як (PDF)",
                Filter = "PDF files|*.pdf",
                InitialDirectory = baseDir,
                FileName = suggested
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                float posX = cbDoubleStamp.Checked ? _posRegXPct : (rbReg.Checked ? _posRegXPct : _posResXPct);
                float posY = cbDoubleStamp.Checked ? _posRegYPct : (rbReg.Checked ? _posRegYPct : _posResYPct);
                InsertStampIntoPdf(
                    _previewPdfPath, sfd.FileName, pngMain,
                    (float)nudWidthRatio.Value, (float)nudRightMm.Value, (float)nudBottomMm.Value,
                    cbFirstPageOnly.Checked, _freePosition, posX, posY, false, pngSecond, _posResXPct, _posResYPct
                );
                MessageBox.Show(this, "Файл збережено:\n" + sfd.FileName, "Готово",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Не вдалося вставити штамп:\n" + ex.Message, "Помилка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void InsertStampIntoPdf(
            string srcPdf,
            string outPdf,
            byte[] pngStamp,
            float widthRatio,
            float rightMm,
            float bottomMm,
            bool firstPageOnly,
            bool freePos,
            float posXPct,
            float posYPct,
            bool doubleStamp,
            byte[] pngStamp2 = null,
            float posXPct2 = float.NaN,
            float posYPct2 = float.NaN)
        {
            using (var input = PdfSharp.Pdf.IO.PdfReader.Open(srcPdf, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify))
            using (var ms = new MemoryStream(pngStamp, writable: false))
            using (var ximg = XImage.FromStream(ms))
            {
                for (int i = 0; i < input.PageCount; i++)
                {
                    if (firstPageOnly && i > 0) break;

                    var page = input.Pages[i];
                    using (var gfx = XGraphics.FromPdfPage(page))
                    {
                        double pageW = page.Width;
                        double pageH = page.Height;

                        double targetW = pageW * widthRatio;
                        double targetH = targetW * (ximg.PixelHeight / (double)ximg.PixelWidth);

                        double x, y;
                        if (freePos && !float.IsNaN(posXPct) && !float.IsNaN(posYPct))
                        {
                            x = posXPct * pageW;
                            y = posYPct * pageH;
                        }
                        else
                        {
                            double mr = rightMm * 72.0 / 25.4;
                            double mb = bottomMm * 72.0 / 25.4;
                            x = pageW - mr - targetW;
                            y = pageH - mb - targetH;
                        }

                        var rect = new PdfSharp.Drawing.XRect(x, y, targetW, targetH);
                        gfx.DrawImage(ximg, rect);
                        if (doubleStamp && pngStamp2 == null)
                        {
                            double gap = MmToPt(5f);
                            double y2 = y - targetH - gap;
                            if (y2 < 0) y2 = y;
                            var rect2 = new PdfSharp.Drawing.XRect(x, y2, targetW, targetH);
                            gfx.DrawImage(ximg, rect2);
                        }
                        if (pngStamp2 != null)
                        {
                            using (var ms2 = new MemoryStream(pngStamp2, writable: false))
                            using (var ximg2 = XImage.FromStream(ms2))
                            {
                                double x2, y2;
                                if (freePos && !float.IsNaN(posXPct2) && !float.IsNaN(posYPct2))
                                {
                                    x2 = posXPct2 * pageW;
                                    y2 = posYPct2 * pageH;
                                }
                                else
                                {
                                    double gap = MmToPt(5f);
                                    x2 = x;
                                    y2 = y - targetH - gap;
                                    if (y2 < 0) y2 = y;
                                }
                                var rect2 = new PdfSharp.Drawing.XRect(x2, y2, targetW, targetH);
                                gfx.DrawImage(ximg2, rect2);
                            }
                        }
                    }
                }
                input.Save(outPdf);
            }
        }

        private static void ConvertWordToPdf(string src, string dst)
        {

#if NETFRAMEWORK || WINDOWS
            try
            {
                var word = new WordInterop.Application { Visible = false };
                try { word.DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone; } catch { }
                var doc = word.Documents.Open(src, ReadOnly: true, Visible: false);
                doc.ExportAsFixedFormat(dst, WordInterop.WdExportFormat.wdExportFormatPDF);
                doc.Close(false);
                word.Quit();
                if (File.Exists(dst)) return;
            }
            catch { }
#endif
            string[] candidates =
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),    "LibreOffice", "program", "soffice.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),"LibreOffice", "program", "soffice.exe")
            };
            string soffice = candidates.FirstOrDefault(File.Exists) ?? "soffice";
            try
            {
                var outdir = Path.GetDirectoryName(dst) ?? Environment.CurrentDirectory;
                Directory.CreateDirectory(outdir);
                var psi = new ProcessStartInfo
                {
                    FileName = soffice,
                    Arguments = $"--headless --convert-to pdf --outdir \"{outdir}\" \"{src}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                var p = Process.Start(psi);
                p.WaitForExit();
                var produced = Path.Combine(outdir, Path.GetFileNameWithoutExtension(src) + ".pdf");
                if (File.Exists(produced))
                {
                    if (!dst.Equals(produced, StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(dst)) File.Delete(dst);
                        File.Move(produced, dst);
                    }
                    return;
                }
            }
            catch { }

            throw new InvalidOperationException("Не вдалося конвертувати DOC/DOCX у PDF (Word або LibreOffice).");
        }

        private void CleanupTemp()
        {
            try
            {
                if (!string.IsNullOrEmpty(_tempDir) && Directory.Exists(_tempDir))
                    Directory.Delete(_tempDir, true);
            }
            catch { }
            _tempDir = null;
            _previewPdfPath = null;
            _pdfDoc?.Dispose();
            _pdfDoc = null;
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            CleanupTemp();
            base.OnClosing(e);
        }
    }
}


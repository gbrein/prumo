// Prumo Painel para Windows: add-in COM do PowerPoint com um painel lateral fixo (Custom Task Pane).
// Sem VSTO e sem PIAs: as interfaces do Office são declaradas aqui e o resto vai por ligação tardia (dynamic).
// Cada botão roda a macro PrumoP_<id> do suplemento Prumo.ppam (Application.Run). As caixas de pergunta
// e os avisos são os do próprio VBA, que no Windows abrem na janela certa.
// Os botões vêm de panel.tsv (recurso embutido), gerado do painel do Mac por build/gen_panel.py.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: ComVisible(false)]

namespace Prumo
{
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode,
                                      [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C033E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)] void CTPFactoryAvailable([In, MarshalAs(UnmanagedType.Interface)] object factory);
    }

    // O add-in. O VBA chega aqui por COMAddIns("Prumo.Panel").Object.Show (botão Painel do ribbon).
    [ComVisible(true), Guid("6B0B3E4C-6C1E-4C0B-9A57-0E5C2A1D7F01"), ProgId("Prumo.Panel")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, ICustomTaskPaneConsumer
    {
        internal dynamic App;
        dynamic factory;
        readonly Dictionary<string, dynamic> panes = new Dictionary<string, dynamic>();
        bool visible = true;   // o X do painel esconde em todas as janelas; o botão Painel traz de volta
        string lastKey;
        Timer timer;

        // Idioma: o da interface do Office (msoLanguageIDUI). Português fica em português; qualquer outro, inglês.
        internal static bool EN;
        internal static string L(string pt, string en) { return EN ? en : pt; }

        static bool DetectEN(object app)
        {
            try
            {
                object ls = app.GetType().InvokeMember("LanguageSettings", BindingFlags.GetProperty, null, app, null);
                int id = (int)ls.GetType().InvokeMember("LanguageID", BindingFlags.GetProperty, null, ls, new object[] { 2 });
                return (id & 0x3FF) != 0x16;  // 0x16 = português (Brasil e Portugal)
            }
            catch { return System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName != "pt"; }
        }

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            Trace("OnConnection " + connectMode);
            try
            {
                App = application;
                EN = DetectEN(application);
                object inst = addInInst;  // COMAddIn.Object: o VBA chama Show por ele
                inst.GetType().InvokeMember("Object", BindingFlags.SetProperty, null, inst, new object[] { this });
                // shortcut: sem eventos COM do PowerPoint, um timer acha janelas novas; trocar por WindowActivate se pesar
                timer = new Timer { Interval = 700 };
                timer.Tick += (s, e) => Tick();
                timer.Start();
            }
            catch (Exception ex) { Log(ex); }
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            if (timer != null) timer.Dispose();
            foreach (var p in panes.Values) { try { p.Delete(); } catch { } }
            panes.Clear();
            factory = null;
            App = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void CTPFactoryAvailable(object f)
        {
            Trace("CTPFactoryAvailable");
            factory = f;
            try { Tick(); } catch (Exception ex) { Log(ex); }
        }

        public void Show()
        {
            visible = true;
            foreach (var p in panes.Values) { try { p.Visible = true; } catch { } }
            Tick();
        }

        // Janela do documento ativo: "arquivo|título" (o título diferencia as janelas :2 da mesma apresentação).
        string ActiveKey(out object window)
        {
            window = null;
            try
            {
                if (App.Windows.Count == 0) return null;
                window = App.ActiveWindow;
                dynamic w = window;
                return (string)w.Presentation.FullName + "|" + (string)w.Caption;
            }
            catch (Exception ex) { Trace("sem janela: " + ex.Message); return null; }
        }

        void Tick()
        {
            if (factory == null || App == null) return;
            // Painel fechado pelo X em alguma janela: some de todas (como no Mac).
            var dead = new List<string>();
            foreach (var kv in panes)
            {
                try { if (visible && !(bool)kv.Value.Visible) visible = false; }
                catch { dead.Add(kv.Key); }  // janela fechada: o Office já desfez o painel
            }
            foreach (var k in dead) panes.Remove(k);
            if (!visible) { foreach (var p in panes.Values) { try { p.Visible = false; } catch { } } }

            object window;
            string key = ActiveKey(out window);
            if (key == null) return;
            dynamic pane;
            if (!panes.TryGetValue(key, out pane))
            {
                if (!visible) return;
                try
                {
                    pane = factory.CreateCTP("Prumo.PanelControl", "Prumo", window);
                    pane.DockPosition = 2;  // msoCTPDockPositionRight
                    pane.Width = 250;
                    ((PanelControl)pane.ContentControl).Init(this);
                    pane.Visible = true;
                    panes[key] = pane;
                    Trace("painel criado: " + key);
                }
                catch (Exception ex) { Log(ex); return; }
            }
            if (key != lastKey)
            {
                lastKey = key;
                try { ((PanelControl)pane.ContentControl).RefreshColors(); } catch { }
            }
        }

        // Roda a ação do Prumo; erro de chamada vira texto para o rodapé do painel.
        internal string Run(string id)
        {
            try
            {
                object app = App;  // Run tem ParamArray: chamada direta por IDispatch, com um argumento só
                app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, new object[] { "PrumoP_" + id });
                return null;
            }
            catch (Exception ex)
            {
                var inner = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                return L("Falhou: ", "Failed: ") + inner.Message;
            }
        }

        // Cores do tema da apresentação ativa (Escuro 1, Claro 1, Escuro 2, Claro 2, Ênfase 1-6).
        internal List<Color> ThemeColors()
        {
            var list = new List<Color>();
            try
            {
                dynamic scheme = App.ActivePresentation.SlideMaster.Theme.ThemeColorScheme;
                for (int i = 1; i <= 10; i++) list.Add(ColorTranslator.FromOle((int)scheme.Colors(i).RGB));
            }
            catch { }
            return list;
        }

        // Diagnóstico: com o arquivo %TEMP%\prumo_trace, registra os passos em %TEMP%\prumo_panel.log.
        internal static void Trace(string m) { if (File.Exists(Path.Combine(Path.GetTempPath(), "prumo_trace"))) Log(m); }
        internal static void Log(object ex)
        {
            try { File.AppendAllText(Path.Combine(Path.GetTempPath(), "prumo_panel.log"), DateTime.Now + " " + ex + "\r\n"); } catch { }
        }
    }

    sealed class Item
    {
        public string Id, Tip, Title, Colors, Icon, Tint;
        public List<KeyValuePair<string, string>> Menu = new List<KeyValuePair<string, string>>();
    }

    sealed class Section
    {
        public string Title;
        public bool ColorBar;
        public List<Item> Items = new List<Item>();
    }

    // Converte o IPictureDisp do GetImageMso em Image (método protegido do AxHost).
    sealed class Pic : AxHost
    {
        Pic() : base("") { }
        public static Image From(object p) { return GetPictureFromIPicture(p); }
    }

    // O conteúdo do painel. O Office cria pelo ProgId; o Connect chama Init em seguida.
    [ComVisible(true), Guid("6B0B3E4C-6C1E-4C0B-9A57-0E5C2A1D7F02"), ProgId("Prumo.PanelControl")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class PanelControl : UserControl
    {
        static string[] ColorNames
        {
            get
            {
                return Connect.EN
                    ? new[] { "Dark 1", "Light 1", "Dark 2", "Light 2", "Accent 1", "Accent 2", "Accent 3", "Accent 4", "Accent 5", "Accent 6" }
                    : new[] { "Escuro 1", "Claro 1", "Escuro 2", "Claro 2", "Ênfase 1", "Ênfase 2", "Ênfase 3", "Ênfase 4", "Ênfase 5", "Ênfase 6" };
            }
        }
        Connect owner;
        readonly ToolTip tips = new ToolTip();
        readonly Label status = new Label { AutoSize = true, ForeColor = Color.Firebrick };
        FlowLayoutPanel stack, colorBar;
        int btn;  // lado do botão, em pixels da tela

        internal void Init(Connect c)
        {
            owner = c;
            Font = new Font("Segoe UI", 9f);
            // sem BackColor/ForeColor: o painel herda as cores do tema do Office (claro ou escuro)
            btn = (int)Math.Round(Font.Height * 1.9);
            stack = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false,
                AutoScroll = true, Padding = new Padding(8, 4, 4, 8),
            };
            foreach (var s in LoadSections())
            {
                stack.Controls.Add(new Label
                {
                    Text = s.Title, AutoSize = true, Font = new Font(Font, FontStyle.Bold),
                    Margin = new Padding(0, 10, 0, 2),
                });
                var row = NewRow();
                foreach (var it in s.Items) AddControl(row, it);
                stack.Controls.Add(row);
                if (s.ColorBar) { colorBar = NewRow(); stack.Controls.Add(colorBar); }
            }
            status.Margin = new Padding(0, 10, 0, 0);
            stack.Controls.Add(status);
            Controls.Add(stack);
            Resize += (s, e) => FitRows();
            FitRows();
            RefreshColors();
        }

        FlowLayoutPanel NewRow()
        {
            return new FlowLayoutPanel { AutoSize = true, WrapContents = true, Margin = new Padding(0) };
        }

        // As linhas quebram na largura do painel.
        void FitRows()
        {
            if (stack == null) return;
            int w = Math.Max(btn * 2, stack.ClientSize.Width - stack.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth);
            foreach (Control c in stack.Controls)
                if (c is FlowLayoutPanel) c.MaximumSize = new Size(w, 0);
            status.MaximumSize = new Size(w, 0);
        }

        static List<Section> LoadSections()
        {
            var list = new List<Section>();
            using (var st = typeof(PanelControl).Assembly.GetManifestResourceStream("panel.tsv"))
            using (var rd = new StreamReader(st, System.Text.Encoding.UTF8))
            {
                string line;
                while ((line = rd.ReadLine()) != null)
                {
                    var f = line.Split('\t');
                    // o texto em inglês vem na última coluna
                    if (f[0] == "S") list.Add(new Section { Title = Connect.L(f[1], f[3]), ColorBar = f[2] == "1" });
                    else if (f[0] == "I") list[list.Count - 1].Items.Add(new Item { Id = f[1], Tip = Connect.L(f[2], f[7]), Title = f[3], Colors = f[4], Icon = f[5], Tint = f[6] });
                    else if (f[0] == "M") { var s = list[list.Count - 1]; s.Items[s.Items.Count - 1].Menu.Add(new KeyValuePair<string, string>(f[1], Connect.L(f[2], f[3]))); }
                }
            }
            return list;
        }

        Button NewButton(string tip, int width)
        {
            var b = new Button
            {
                Size = new Size(width, btn), FlatStyle = FlatStyle.Flat, Margin = new Padding(0, 0, 2, 2),
                AccessibleName = tip, TabStop = true, UseMnemonic = false,
            };
            b.FlatAppearance.BorderSize = 0;
            tips.SetToolTip(b, tip);
            return b;
        }

        void AddControl(FlowLayoutPanel row, Item it)
        {
            var b = NewButton(it.Tip, btn);
            if (it.Title != "") { b.Text = it.Title; b.Font = new Font(Font, FontStyle.Bold); }
            else if (it.Icon.StartsWith("txt:")) { b.Text = it.Icon.Substring(4); b.Font = new Font("Segoe UI Symbol", 11f); }
            else
            {
                try { b.Image = Pic.From(owner.App.CommandBars.GetImageMso(it.Icon, 16, 16)); }
                catch { b.Text = it.Id.Substring(0, 3); b.Font = new Font(Font.FontFamily, 7f); }  // ícone ausente nesta versão do Office
            }
            if (it.Tint != "") b.ForeColor = ColorTranslator.FromHtml(it.Tint);
            b.Click += (s, e) => Run(it.Id);
            if (it.Menu.Count == 0 && it.Colors == "") { row.Controls.Add(b); return; }

            // Botão + setinha de menu, como os botões divididos do Office (os dois juntos na mesma linha).
            var pair = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 2, 2) };
            b.Margin = new Padding(0);
            var arrow = NewButton(Connect.L("Mais opções: ", "More options: ") + it.Tip, btn / 2);
            arrow.Text = "▾";
            arrow.Margin = new Padding(0);
            arrow.Click += (s, e) => ShowMenu(arrow, it);
            pair.Controls.Add(b);
            pair.Controls.Add(arrow);
            row.Controls.Add(pair);
        }

        void ShowMenu(Control at, Item it)
        {
            var m = new ContextMenuStrip();
            if (it.Colors != "")
            {
                var colors = owner.ThemeColors();
                for (int i = 0; i < ColorNames.Length; i++)
                {
                    string id = it.Colors + "T" + (i + 1);
                    var mi = m.Items.Add(ColorNames[i], i < colors.Count ? Swatch(colors[i], 14) : null);
                    mi.Click += (s, e) => Run(id);
                }
                if (it.Colors != "Font")
                {
                    m.Items.Add(new ToolStripSeparator());
                    string none = it.Colors + "None";
                    m.Items.Add(it.Colors == "Fill" ? Connect.L("Sem preenchimento", "No fill") : Connect.L("Sem linha", "No line")).Click += (s, e) => Run(none);
                }
            }
            else
            {
                foreach (var kv in it.Menu) { string id = kv.Key; m.Items.Add(kv.Value).Click += (s, e) => Run(id); }
            }
            m.Show(at, new Point(0, at.Height));
        }

        static Bitmap Swatch(Color c, int side)
        {
            var bmp = new Bitmap(side, side);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(c);
                g.DrawRectangle(Pens.Gray, 0, 0, side - 1, side - 1);
            }
            return bmp;
        }

        // Barra de cores: um clique pinta o preenchimento com a cor do tema.
        internal void RefreshColors()
        {
            if (colorBar == null) return;
            colorBar.Controls.Clear();
            var colors = owner.ThemeColors();
            for (int i = 0; i < colors.Count; i++)
            {
                string id = "FillT" + (i + 1);
                var b = NewButton(Connect.L("Preencher com a cor " + (i + 1) + " do tema", "Fill with theme color " + (i + 1)), btn * 2 / 3);
                b.Height = btn * 2 / 3;
                b.Image = Swatch(colors[i], btn / 2);
                b.Click += (s, e) => Run(id);
                colorBar.Controls.Add(b);
            }
            var again = NewButton(Connect.L("Atualizar as cores (depois de trocar de apresentação)", "Refresh the colors (after switching presentations)"), btn * 2 / 3);
            again.Height = btn * 2 / 3;
            again.Text = "↻";
            again.Click += (s, e) => RefreshColors();
            colorBar.Controls.Add(again);
        }

        void Run(string id)
        {
            status.Text = "";
            var err = owner.Run(id);
            if (err != null) status.Text = err;
        }
    }
}

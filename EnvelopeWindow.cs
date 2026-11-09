using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;

internal sealed class EnvelopeWindow : Form
{
    readonly EventWaitHandle reopen;
    readonly TextBox tenant = new TextBox(), client = new TextBox();
    readonly Button connect = new Button { Text = "Sign in" };
    readonly Button open = new Button { Text = "Open Inbox" };
    readonly Button preview = new Button { Text = "Preview envelope" };
    readonly CheckBox enabled = new CheckBox { Text = "Show the yellow envelope for new Inbox mail" };
    readonly Label status = new Label();
    readonly NotifyIcon tray = new NotifyIcon();
    readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 500 };
    readonly CancellationTokenSource closing = new CancellationTokenSource();
    readonly Icon envelope = DrawEnvelope();
    Mailbox box;
    Form helpWindow;
    SignIn auth;
    string account;
    DateTime next;
    bool busy, navigating, quit, forceRefresh;
    int failures;

    public EnvelopeWindow(EventWaitHandle reopen)
    {
        this.reopen = reopen;
        Text = "Outlook Envelope";
        Font = new Font("Segoe UI", 10);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(495, 423);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        Icon = envelope;
        Add(new Label { Text = "Microsoft 365 work account", Font = new Font(Font, FontStyle.Bold) }, 18, 14, 450, 27);
        Add(new Label { Text = "Use the IDs from your organization’s Outlook Envelope app registration." }, 18, 45, 455, 40);
        Add(new Label { Text = "Directory (tenant) ID" }, 18, 90, 150, 25);
        Add(tenant, 178, 86, 295, 28);
        Add(new Label { Text = "Application (client) ID" }, 18, 125, 155, 25);
        Add(client, 178, 121, 295, 28);
        tenant.Text = Settings.Get("Tenant") as string ?? "";
        client.Text = Settings.Get("Client") as string ?? "";
        account = Settings.Get("Username") as string;
        Add(enabled, 18, 164, 455, 27);
        enabled.Checked = !Equals(Settings.Get("Enabled"), 0);
        enabled.CheckedChanged += (s, e) =>
        {
            Settings.Set("Enabled", enabled.Checked ? 1 : 0);
            if (enabled.Checked) { box = new Mailbox(); next = DateTime.UtcNow; }
            tray.Visible = false;
            preview.Text = "Preview envelope";
            status.Text = enabled.Checked ? "Envelope enabled." : "Envelope hidden until you enable it again.";
        };
        Add(status, 18, 203, 457, 66);
        status.Text = "Sign in to start watching your Inbox.";
        Add(preview, 18, 273, 160, 34);
        var help = new Button { Text = "Help" };
        Add(help, 369, 273, 104, 34);
        help.Click += (s, e) => ShowHelp();
        preview.Click += (s, e) =>
        {
            tray.Visible = !tray.Visible;
            preview.Text = tray.Visible ? "End preview" : "Preview envelope";
            status.Text = tray.Visible ? "Previewing the envelope. Look beside the clock or in the hidden icons menu. Double-click the envelope to open Inbox." : "Preview ended.";
        };
        Add(connect, 18, 321, 90, 34);
        Add(open, 117, 321, 110, 34);
        var exit = new Button { Text = "Exit" };
        Add(exit, 369, 321, 104, 34);
        connect.Click += async (s, e) => await Start(true);
        open.Click += async (s, e) => await OpenInbox();
        exit.Click += (s, e) => ExitApp();
        var menu = new ContextMenuStrip();
        menu.Items.Add("Hide Envelope", null, (s, e) => { enabled.Checked = false; tray.Visible = false; preview.Text = "Preview envelope"; });
        menu.Items.Add("Settings", null, (s, e) => Reveal());
        menu.Items.Add("Exit", null, (s, e) => ExitApp());
        tray.ContextMenuStrip = menu;
        tray.Icon = envelope;
        tray.Text = "New mail in your work Inbox";
        tray.MouseDoubleClick += async (s, e) => { if (e.Button == MouseButtons.Left) await OpenInbox(); };
        timer.Tick += async (s, e) =>
        {
            if (reopen.WaitOne(0)) Reveal();
            if (!busy && enabled.Checked && auth != null && DateTime.UtcNow >= next) await Poll();
        };
        Shown += async (s, e) =>
        {
            timer.Start();
            if (Settings.Get("Account") != null) await Start(false);
        };
        FormClosing += (s, e) =>
        {
            if (!quit && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                var answer = MessageBox.Show(this,
                    "Outlook Envelope will continue running in the background. To stop the app, choose Exit or end its process in Task Manager.\n\nClose this window and continue in the background?",
                    Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (answer == DialogResult.Yes) Hide();
            }
            else { closing.Cancel(); timer.Stop(); tray.Visible = false; }
        };
    }

    void Add(Control c, int x, int y, int width, int height)
    { c.SetBounds(x, y, width, height); Controls.Add(c); }
    void Reveal() { Show(); WindowState = FormWindowState.Normal; Activate(); }
    void ShowHelp()
    {
        if (helpWindow == null || helpWindow.IsDisposed)
        {
            using (var reader = new StreamReader(typeof(EnvelopeWindow).Assembly.GetManifestResourceStream("Setup.txt")))
            {
                var instructions = new RichTextBox
                {
                    Dock = DockStyle.Fill, ReadOnly = true, DetectUrls = false,
                    BorderStyle = BorderStyle.None, BackColor = SystemColors.Window,
                    ScrollBars = RichTextBoxScrollBars.Vertical, Text = reader.ReadToEnd(),
                    AccessibleName = "Setup instructions"
                };
                helpWindow = new Form
                {
                    Text = "Outlook Envelope Help", Font = Font, Icon = envelope,
                    Padding = new Padding(14), BackColor = SystemColors.Window,
                    ClientSize = new Size(680, 480), StartPosition = FormStartPosition.CenterParent
                };
                helpWindow.Controls.Add(instructions);
            }
            helpWindow.Show(this);
        }
        if (helpWindow.WindowState == FormWindowState.Minimized) helpWindow.WindowState = FormWindowState.Normal;
        helpWindow.Activate();
    }
    void ExitApp() { quit = true; Close(); }
    async Task OpenInbox()
    {
        if (navigating) return;
        navigating = true;
        try { await Inbox.Open(account); }
        catch (Exception e) { if (!quit) { status.Text = e.Message; Reveal(); } }
        finally { navigating = false; }
    }
    async Task Start(bool interactive)
    {
        if (busy) return;
        if (!Guid.TryParse(tenant.Text.Trim(), out var tid) || !Guid.TryParse(client.Text.Trim(), out var cid))
        { status.Text = "Enter both IDs from the app registration, then choose Sign in."; return; }
        busy = true;
        connect.Enabled = false;
        try
        {
            var session = new SignIn(tid, cid);
            status.Text = interactive ? "Finish signing in using your browser." : "Connecting to Microsoft 365…";
            var token = await session.Get(interactive, false, closing.Token);
            auth = session;
            preview.Enabled = false;
            Settings.Set("Tenant", tid.ToString());
            Settings.Set("Client", cid.ToString());
            Settings.Set("Account", token.Account.HomeAccountId.Identifier);
            account = token.Account.Username;
            Settings.Set("Username", account);
            box = new Mailbox();
            tray.Visible = false;
            forceRefresh = false;
            failures = 0;
            next = DateTime.UtcNow;
            status.Text = enabled.Checked ? "Checking your Inbox…" : "Envelope hidden. Check the box above to enable it.";
            if (!interactive && enabled.Checked) Hide();
        }
        catch (Exception e) { status.Text = "Sign-in requires attention. " + e.Message; }
        finally { busy = false; connect.Enabled = true; }
    }
    async Task Poll()
    {
        busy = true;
        connect.Enabled = false;
        var current = box;
        try
        {
            var token = await auth.Get(false, forceRefresh, closing.Token);
            await current.Poll(token.AccessToken, closing.Token);
            if (current != box || !enabled.Checked) return;
            tray.Visible = current.State.Visible;
            next = DateTime.UtcNow.AddSeconds(20);
            status.Text = "Watching " + account + ". Last checked at " + DateTime.Now.ToShortTimeString() +
                ". Click X to close this window. Run the EXE again to reopen it.";
            forceRefresh = false;
            failures = 0;
        }
        catch (ResyncRequired) { box = new Mailbox(); tray.Visible = false; }
        catch (MsalUiRequiredException) { NeedsSignIn("Microsoft 365 needs you to sign in again."); }
        catch (MailError e) when (e.Code == HttpStatusCode.Forbidden || (e.Code == HttpStatusCode.Unauthorized && forceRefresh))
        { NeedsSignIn(e.Message + " Check the app’s Mail.ReadBasic permission and sign in again."); }
        catch (Exception e)
        {
            var service = e as MailError;
            forceRefresh = service?.Code == HttpStatusCode.Unauthorized;
            failures = Math.Min(failures + 1, 5);
            int seconds = Math.Max(service?.RetrySeconds ?? 0, 10 << failures);
            next = DateTime.UtcNow.AddSeconds(seconds);
            status.Text = e.Message + " Retrying at " + next.ToLocalTime().ToShortTimeString() + ".";
        }
        finally { busy = false; connect.Enabled = true; }
    }
    void NeedsSignIn(string message)
    { auth = null; tray.Visible = false; preview.Enabled = true; preview.Text = "Preview envelope"; status.Text = message; Reveal(); }
    [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr handle);
    internal static Icon DrawEnvelope()
    {
        using (var bmp = new Bitmap(32, 32))
        using (var g = Graphics.FromImage(bmp))
        using (var edge = new Pen(Color.FromArgb(104, 66, 9), 1.7f))
        using (var fold = new Pen(Color.FromArgb(145, 91, 12), 1.6f))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.FillRectangle(Brushes.Gold, 1, 5, 30, 22);
            g.DrawRectangle(edge, 1, 5, 30, 22);
            g.DrawLines(fold, new[] { new Point(2, 26), new Point(13, 16), new Point(19, 16), new Point(30, 26) });
            g.FillPolygon(Brushes.Khaki, new[] { new Point(2, 6), new Point(30, 6), new Point(16, 18) });
            g.DrawLines(fold, new[] { new Point(2, 6), new Point(16, 18), new Point(30, 6) });
            var handle = bmp.GetHicon();
            try { using (var icon = Icon.FromHandle(handle)) return (Icon)icon.Clone(); }
            finally { DestroyIcon(handle); }
        }
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing) { timer.Dispose(); tray.ContextMenuStrip.Dispose(); tray.Dispose(); envelope.Dispose(); closing.Dispose(); }
        base.Dispose(disposing);
    }
}

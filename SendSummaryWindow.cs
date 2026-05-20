using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Outlook = Microsoft.Office.Interop.Outlook;
using WpfBorder = System.Windows.Controls.Border;
using WpfThickness = System.Windows.Thickness;

namespace Marketing_Meetings_Summary;

public class SendSummaryWindow : Window
{
    // ── public result ─────────────────────────────────────────────────────
    public bool ShouldSend { get; private set; }

    // ── data passed in ────────────────────────────────────────────────────
    private readonly string        _orderNumber;
    private readonly string        _customerName;
    private readonly string        _projectType;
    private readonly List<string>  _participants;
    private readonly string        _attachmentPath;   // already-saved .doc path (may be empty)

    // ── editable controls ─────────────────────────────────────────────────
    private TextBox  _txtTo        = new();
    private TextBox  _txtSubject   = new();
    private TextBox  _txtBody      = new();
    private TextBlock _lblAttach   = new();

    // ── stored attachment chosen in the dialog ────────────────────────────
    public string ResolvedAttachmentPath { get; private set; } = "";

    public SendSummaryWindow(
        string        orderNumber,
        string        customerName,
        string        projectType,
        List<string>  participants,
        string        attachmentPath)
    {
        _orderNumber    = orderNumber;
        _customerName   = customerName;
        _projectType    = projectType;
        _participants   = participants;
        _attachmentPath = attachmentPath;

        Title                  = "Send Summary via Outlook";
        Width                  = 680;
        Height                 = 640;
        MinWidth               = 580;
        MinHeight              = 540;
        WindowStartupLocation  = WindowStartupLocation.CenterOwner;
        ResizeMode             = ResizeMode.CanResize;
        Background             = new SolidColorBrush(Color.FromRgb(248, 250, 252)); // #F8FAFC

        BuildUI();
    }

    // ──────────────────────────────────────────────────────────────────────
    //  UI Construction
    // ──────────────────────────────────────────────────────────────────────
    private void BuildUI()
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // header
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // body
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // footer

        // ── Header bar ────────────────────────────────────────────────────
        var header = new WpfBorder
        {
            Background      = new SolidColorBrush(Color.FromRgb(37, 99, 235)),   // #2563EB
            Padding         = new WpfThickness(24, 16, 24, 16),
            CornerRadius    = new CornerRadius(0)
        };

        var headerContent = new StackPanel { Orientation = Orientation.Horizontal };
        headerContent.Children.Add(new TextBlock
        {
            Text                = "✉️",
            FontSize            = 22,
            VerticalAlignment   = VerticalAlignment.Center,
            Margin              = new WpfThickness(0, 0, 12, 0)
        });
        var headerTextStack = new StackPanel();
        headerTextStack.Children.Add(new TextBlock
        {
            Text       = "Send Summary via Outlook",
            FontSize   = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.White
        });
        headerTextStack.Children.Add(new TextBlock
        {
            Text       = "Review and send the meeting summary to all participants",
            FontSize   = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(191, 219, 254)), // #BFDBFE
            Margin     = new WpfThickness(0, 2, 0, 0)
        });
        headerContent.Children.Add(headerTextStack);
        header.Child = headerContent;
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // ── Scrollable body ───────────────────────────────────────────────
        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Padding                     = new WpfThickness(0)
        };
        var body = new StackPanel { Margin = new WpfThickness(24, 20, 24, 8) };
        scroll.Content = body;
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        // ── To: ───────────────────────────────────────────────────────────
        body.Children.Add(MakeLabel("To:"));
        _txtTo = MakeTextBox(string.Join("; ", _participants), false);
        _txtTo.FontSize = 12;
        _txtTo.Height = 36;
        body.Children.Add(_txtTo);

        // ── Subject: ──────────────────────────────────────────────────────
        body.Children.Add(MakeLabel("Subject:"));
        _txtSubject = MakeTextBox(BuildSubject(), false);
        _txtSubject.Height = 36;
        body.Children.Add(_txtSubject);

        // ── Attachment: ───────────────────────────────────────────────────
        body.Children.Add(MakeLabel("Attachment:"));
        var attachRow = new Grid();
        attachRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        attachRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var attachBox = new WpfBorder
        {
            BorderBrush     = new SolidColorBrush(Color.FromRgb(209, 213, 219)),
            BorderThickness = new WpfThickness(1),
            CornerRadius    = new CornerRadius(8),
            Background      = new SolidColorBrush(Color.FromRgb(249, 250, 251)),
            Padding         = new WpfThickness(10, 0, 10, 0),
            Margin          = new WpfThickness(0, 0, 8, 0),
            Height          = 36
        };

        string attachDisplay = string.IsNullOrEmpty(_attachmentPath)
            ? "⚠️  No file selected — click Browse to attach the summary document"
            : $"📎  {Path.GetFileName(_attachmentPath)}";

        _lblAttach = new TextBlock
        {
            Text                = attachDisplay,
            FontSize            = 12,
            Foreground          = string.IsNullOrEmpty(_attachmentPath)
                                    ? new SolidColorBrush(Color.FromRgb(180, 83, 9))
                                    : new SolidColorBrush(Color.FromRgb(30, 41, 59)),
            VerticalAlignment   = VerticalAlignment.Center,
            TextTrimming        = TextTrimming.CharacterEllipsis
        };
        attachBox.Child = _lblAttach;
        Grid.SetColumn(attachBox, 0);
        attachRow.Children.Add(attachBox);

        var browseBtn = MakeActionButton("📁  Browse…", "#374151");
        browseBtn.Width = 110;
        Grid.SetColumn(browseBtn, 1);
        browseBtn.Click += BrowseAttachment_Click;
        attachRow.Children.Add(browseBtn);

        body.Children.Add(attachRow);

        // ── Divider ───────────────────────────────────────────────────────
        body.Children.Add(new WpfBorder
        {
            Height          = 1,
            Background      = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
            Margin          = new WpfThickness(0, 4, 0, 16)
        });

        // ── Body label + note ─────────────────────────────────────────────
        var bodyLabelRow = new Grid();
        bodyLabelRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        bodyLabelRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var bodyLabel = MakeLabel("Email Body:");
        bodyLabel.Margin = new WpfThickness(0);
        Grid.SetColumn(bodyLabel, 0);
        bodyLabelRow.Children.Add(bodyLabel);
        var editNote = new TextBlock
        {
            Text              = "✏️ Editable",
            FontSize          = 10,
            Foreground        = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(editNote, 1);
        bodyLabelRow.Children.Add(editNote);
        body.Children.Add(bodyLabelRow);

        // ── Email body preview ────────────────────────────────────────────
        var bodyBorder = new WpfBorder
        {
            BorderBrush     = new SolidColorBrush(Color.FromRgb(209, 213, 219)),
            BorderThickness = new WpfThickness(1),
            CornerRadius    = new CornerRadius(8),
            Background      = Brushes.White,
            Margin          = new WpfThickness(0, 6, 0, 0),
            Effect          = new DropShadowEffect
            {
                ShadowDepth = 0,
                BlurRadius  = 6,
                Opacity     = 0.06,
                Color       = Colors.Black
            }
        };

        _txtBody = new TextBox
        {
            Text                        = BuildBody(),
            FontSize                    = 13,
            FontFamily                  = new FontFamily("Segoe UI"),
            Foreground                  = new SolidColorBrush(Color.FromRgb(30, 41, 59)),
            Background                  = Brushes.Transparent,
            BorderThickness             = new WpfThickness(0),
            AcceptsReturn               = true,
            TextWrapping                = TextWrapping.Wrap,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Padding                     = new WpfThickness(14, 12, 14, 12),
            MinHeight                   = 220,
            VerticalContentAlignment    = VerticalAlignment.Top
        };
        bodyBorder.Child = _txtBody;
        body.Children.Add(bodyBorder);

        // ── Footer ────────────────────────────────────────────────────────
        var footer = new WpfBorder
        {
            Background      = new SolidColorBrush(Color.FromRgb(248, 250, 252)),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(0, 1, 0, 0),
            Padding         = new WpfThickness(20, 8, 20, 8)
        };

        var footerGrid = new Grid();
        footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        // Left note
        var footerNote = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        footerNote.Children.Add(new TextBlock
        {
            Text       = "ℹ️  ",
            FontSize   = 12,
            Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139))
        });
        footerNote.Children.Add(new TextBlock
        {
            Text       = "Outlook will open with this draft ready to review before sending.",
            FontSize   = 11,
            Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
            VerticalAlignment = VerticalAlignment.Center
        });
        Grid.SetColumn(footerNote, 0);
        footerGrid.Children.Add(footerNote);

        // Right buttons
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(btnRow, 1);

        var cancelBtn = MakeActionButton("Cancel", "#6B7280");
        cancelBtn.Margin = new WpfThickness(0, 0, 8, 0);
        cancelBtn.Height = 28;
        cancelBtn.Click += (s, e) => { ShouldSend = false; Close(); };
        btnRow.Children.Add(cancelBtn);

        var sendBtn = MakeActionButton("✉️  Open in Outlook", "#2563EB");
        sendBtn.Height = 28;
        sendBtn.Click += SendButton_Click;
        btnRow.Children.Add(sendBtn);

        footerGrid.Children.Add(btnRow);
        footer.Child = footerGrid;
        Grid.SetRow(footer, 2);
        root.Children.Add(footer);

        // Pre-fill resolved attachment
        ResolvedAttachmentPath = _attachmentPath;

        Content = root;
    }

    // ──────────────────────────────────────────────────────────────────────
    //  Helpers
    // ──────────────────────────────────────────────────────────────────────
    private static TextBlock MakeLabel(string text) => new TextBlock
    {
        Text       = text,
        FontSize   = 12,
        FontWeight = FontWeights.SemiBold,
        Foreground = new SolidColorBrush(Color.FromRgb(71, 85, 105)),
        Margin     = new WpfThickness(0, 10, 0, 5)
    };

    private static TextBox MakeTextBox(string text, bool multiline)
    {
        var tb = new TextBox
        {
            Text                     = text,
            FontSize                 = 12,
            Padding                  = new WpfThickness(10, multiline ? 8 : 0, 10, multiline ? 8 : 0),
            BorderBrush              = new SolidColorBrush(Color.FromRgb(209, 213, 219)),
            BorderThickness          = new WpfThickness(1),
            Background               = Brushes.White,
            Foreground               = new SolidColorBrush(Color.FromRgb(30, 41, 59)),
            AcceptsReturn            = multiline,
            TextWrapping             = multiline ? TextWrapping.Wrap : TextWrapping.NoWrap,
            VerticalContentAlignment = VerticalAlignment.Center
        };

        // Rounded corners via template
        var tmpl  = new System.Windows.Controls.ControlTemplate(typeof(TextBox));
        var bFact = new FrameworkElementFactory(typeof(WpfBorder));
        bFact.SetValue(WpfBorder.BackgroundProperty,      new TemplateBindingExtension(BackgroundProperty));
        bFact.SetValue(WpfBorder.BorderBrushProperty,     new TemplateBindingExtension(BorderBrushProperty));
        bFact.SetValue(WpfBorder.BorderThicknessProperty, new TemplateBindingExtension(BorderThicknessProperty));
        bFact.SetValue(WpfBorder.CornerRadiusProperty,    new CornerRadius(8));
        var svFact = new FrameworkElementFactory(typeof(ScrollViewer));
        svFact.Name = "PART_ContentHost";
        svFact.SetValue(ScrollViewer.VerticalAlignmentProperty,   VerticalAlignment.Center);
        svFact.SetValue(ScrollViewer.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
        svFact.SetValue(ScrollViewer.MarginProperty, new WpfThickness(10, 0, 10, 0));
        bFact.AppendChild(svFact);
        tmpl.VisualTree = bFact;
        tb.Template = tmpl;

        return tb;
    }

    private static Button MakeActionButton(string label, string hexColor)
    {
        var baseColor  = (Color)ColorConverter.ConvertFromString(hexColor);
        var hoverColor = Color.FromRgb(
            (byte)Math.Max(0, baseColor.R - 20),
            (byte)Math.Max(0, baseColor.G - 20),
            (byte)Math.Max(0, baseColor.B - 20));
        var pressColor = Color.FromRgb(
            (byte)Math.Max(0, baseColor.R - 40),
            (byte)Math.Max(0, baseColor.G - 40),
            (byte)Math.Max(0, baseColor.B - 40));

        var btn = new Button
        {
            BorderThickness = new WpfThickness(0),
            Cursor          = System.Windows.Input.Cursors.Hand,
            Background      = new SolidColorBrush(baseColor),
            Padding         = new WpfThickness(0)
        };

        // Use an explicit TextBlock as Content — bypasses all inheritance issues
        var txt = new TextBlock
        {
            Text       = label,
            FontSize   = 11,
            FontWeight = FontWeights.Medium,
            Foreground = Brushes.White,
            Padding    = new WpfThickness(12, 5, 12, 5),
            VerticalAlignment   = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
        };
        btn.Content = txt;

        var tmpl  = new System.Windows.Controls.ControlTemplate(typeof(Button));
        var bFact = new FrameworkElementFactory(typeof(WpfBorder));
        bFact.Name = "bd";
        bFact.SetValue(WpfBorder.BackgroundProperty,   new TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        bFact.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(6));
        var cFact = new FrameworkElementFactory(typeof(ContentPresenter));
        cFact.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        cFact.SetValue(ContentPresenter.VerticalAlignmentProperty,   VerticalAlignment.Center);
        bFact.AppendChild(cFact);
        tmpl.VisualTree = bFact;

        var hoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
        hoverTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(hoverColor), "bd"));
        tmpl.Triggers.Add(hoverTrigger);

        var pressTrigger = new Trigger { Property = Button.IsPressedProperty, Value = true };
        pressTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(pressColor), "bd"));
        tmpl.Triggers.Add(pressTrigger);

        btn.Template = tmpl;
        return btn;
    }

    // ──────────────────────────────────────────────────────────────────────
    //  Content builders
    // ──────────────────────────────────────────────────────────────────────
    private string BuildSubject() =>
        $"Marketing Kickoff Meeting Summary – Order {_orderNumber}" +
        (string.IsNullOrWhiteSpace(_customerName) ? "" : $" | {_customerName}");

    private string BuildBody()
    {
        string date        = DateTime.Now.ToString("MMMM dd, yyyy", CultureInfo.InvariantCulture);
        string order       = string.IsNullOrWhiteSpace(_orderNumber) ? "[Order Number]" : _orderNumber;
        string customer    = string.IsNullOrWhiteSpace(_customerName) ? "[Customer]" : _customerName;
        string projectType = string.IsNullOrWhiteSpace(_projectType) ? "" : $" ({_projectType})";
        string greeting    = _participants.Count > 0
            ? $"Dear {_participants[0].Split(' ')[0]},"
            : "Dear Team,";

        return
$@"{greeting}

Please find attached the Marketing Kickoff Meeting Summary for Order {order} – {customer}{projectType}, prepared on {date}.

This document covers the key topics and decisions discussed during our kickoff meeting, including:

  •  Order overview and project scope
  •  System configuration and selected components
  •  Target specifications
  •  Open questions and action items
  •  Marketing and PM notes

Please review the summary at your earliest convenience. Should you have any questions, require clarification, or need to discuss any of the items outlined, do not hesitate to reach out.

We look forward to a successful collaboration on this project.

Best regards,
CI Systems – Marketing Team";
    }

    // ──────────────────────────────────────────────────────────────────────
    //  Event handlers
    // ──────────────────────────────────────────────────────────────────────
    private void BrowseAttachment_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title  = "Select Summary Document",
            Filter = "Word Document (*.doc;*.docx)|*.doc;*.docx|All Files (*.*)|*.*",
            FileName = string.IsNullOrEmpty(_attachmentPath)
                ? $"Kickoff {_orderNumber}"
                : Path.GetFileName(_attachmentPath)
        };

        if (!string.IsNullOrEmpty(_attachmentPath) && File.Exists(_attachmentPath))
            dlg.InitialDirectory = Path.GetDirectoryName(_attachmentPath);

        if (dlg.ShowDialog(this) == true)
        {
            ResolvedAttachmentPath = dlg.FileName;
            _lblAttach.Text        = $"📎  {Path.GetFileName(dlg.FileName)}";
            _lblAttach.Foreground  = new SolidColorBrush(Color.FromRgb(30, 41, 59));
        }
    }

    private void SendButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_txtTo.Text))
        {
            MessageBox.Show("Please add at least one recipient in the To: field.",
                "No Recipients", MessageBoxButton.OK, MessageBoxImage.Warning);
            _txtTo.Focus();
            return;
        }

        try
        {
            SendViaOutlook(
                _txtTo.Text,
                _txtSubject.Text,
                _txtBody.Text,
                ResolvedAttachmentPath);

            ShouldSend = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Could not open Outlook.\n\nPlease make sure Outlook is installed.\n\nDetails: {ex.Message}",
                "Outlook Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // GetActiveObject was removed from Marshal in .NET 5+ — call oleaut32 directly.
    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    // Keep static refs so GC doesn't release COM objects before the user clicks Send.
    private static Outlook.Application? _outlookApp;
    private static Outlook.MailItem?    _pendingMail;

    private static void SendViaOutlook(string to, string subject, string body, string attachmentPath)
    {
        // Get or create the Outlook Application COM object.
        // For classic Outlook (2016/2019/2021): GetActiveObject returns the running instance.
        // For new Outlook (Microsoft 365): GetActiveObject fails; Activator.CreateInstance
        //   routes through the COM registration and attaches to the running new Outlook.
        Outlook.Application? app = null;

        // Try ROT first
        try
        {
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType != null)
            {
                int hr = GetActiveObject(outlookType.GUID, IntPtr.Zero, out object runningObj);
                if (hr == 0) app = (Outlook.Application)runningObj;
            }
        }
        catch { }

        // Fallback: create via registered COM server (attaches to running instance)
        if (app == null)
        {
            try { app = new Outlook.Application(); }
            catch { }
        }

        if (app == null)
            throw new InvalidOperationException("Could not connect to Outlook. Make sure Outlook is installed and running.");

        _outlookApp = app;

        Outlook.MailItem mail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
        _pendingMail = mail;

        foreach (var addr in to.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string a = addr.Trim();
            if (!string.IsNullOrEmpty(a))
                mail.Recipients.Add(a);
        }
        mail.Recipients.ResolveAll();

        mail.Subject  = subject;
        mail.HTMLBody = BuildHtmlBody(body);

        if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
            mail.Attachments.Add(attachmentPath, Outlook.OlAttachmentType.olByValue, 1,
                                 Path.GetFileName(attachmentPath));

        mail.Display(false);   // opens compose window; user reviews and clicks Send in Outlook
    }

    // Wraps the email body in LTR HTML so Outlook never auto-flips to RTL
    // even when the text contains Hebrew characters (e.g. in the date).
    private static string BuildHtmlBody(string plainBody)
    {
        string escaped = System.Security.SecurityElement.Escape(plainBody)
                             .Replace("\r\n", "<br>")
                             .Replace("\n",   "<br>");
        return "<html><head><meta charset=\"UTF-8\"></head>" +
               "<body dir=\"ltr\" style=\"direction:ltr;text-align:left;" +
               "font-family:Calibri,Arial,sans-serif;font-size:11pt\">" +
               escaped +
               "</body></html>";
    }
}

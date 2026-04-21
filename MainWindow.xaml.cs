using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WpfCheckBox = System.Windows.Controls.CheckBox;
using WpfBorder = System.Windows.Controls.Border;
using WpfThickness = System.Windows.Thickness;

namespace Marketing_Meetings_Summary;

public class TargetItem : INotifyPropertyChanged
{
    private string _type = "";
    private string _qty = "";
    private string _details = "";

    public string Type
    {
        get => _type;
        set { _type = value; OnPropertyChanged(nameof(Type)); }
    }

    public string Qty
    {
        get => _qty;
        set { _qty = value; OnPropertyChanged(nameof(Qty)); }
    }

    public string Details
    {
        get => _details;
        set { _details = value; OnPropertyChanged(nameof(Details)); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public class QuestionItem : INotifyPropertyChanged
{
    private string _text = "";

    public string Text
    {
        get => _text;
        set { _text = value; OnPropertyChanged(nameof(Text)); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public partial class MainWindow : Window
{
    private readonly string[] _configItems = new[]
    {
        "Motorized Target Wheel",
        "Motorized Source Stage",
        "Motorized Focus Stage",
        "Integrating Sphere",
        "Blackbody / SR800N",
        "PC + CTE",
        "Frame Grabber",
        "Installation + Training",
        "Rackmount / PDU",
        "Optical Table"
    };

    private readonly Dictionary<string, WpfCheckBox> _configCheckBoxes = new();
    private readonly ObservableCollection<TargetItem> _targets = new();
    private readonly ObservableCollection<QuestionItem> _questions = new();

    public MainWindow()
    {
        InitializeComponent();
        InitializeConfigItems();
        InitializeTargets();
        InitializeQuestions();
    }

    private void InitializeConfigItems()
    {
        foreach (var item in _configItems)
        {
            var checkBox = new WpfCheckBox
            {
                Content = item,
                Margin = new WpfThickness(0, 0, 8, 8),
                FontSize = 13,
                Padding = new WpfThickness(12, 8, 12, 8)
            };

            var border = new WpfBorder
            {
                Background = System.Windows.Media.Brushes.White,
                CornerRadius = new CornerRadius(12),
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(226, 232, 240)),
                BorderThickness = new WpfThickness(1),
                Padding = new WpfThickness(12, 8, 12, 8),
                Margin = new WpfThickness(0, 0, 8, 8),
                Child = checkBox
            };

            _configCheckBoxes[item] = checkBox;
            configItemsControl.Items.Add(border);
        }
    }

    private void InitializeTargets()
    {
        var defaultTargets = new[]
        {
            "4Bar", "Pin Hole", "Square", "Cross", "Step", "USAF", "Boresight", "LOS Alignment"
        };

        foreach (var type in defaultTargets)
        {
            var target = new TargetItem { Type = type };
            _targets.Add(target);
        }

        targetsItemsControl.ItemsSource = _targets;
    }

    private void InitializeQuestions()
    {
        _questions.Add(new QuestionItem { Text = "Define all target sizes" });
        _questions.Add(new QuestionItem { Text = "Confirm frame grabber compatibility" });
        _questions.Add(new QuestionItem { Text = "Confirm rack / table / logistics details" });

        questionsItemsControl.ItemsSource = _questions;
    }

    private void AddTarget_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var newTarget = new TargetItem();
            _targets.Add(newTarget);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding target: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void RemoveTarget_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is TargetItem target)
            {
                _targets.Remove(target);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error removing target: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AddQuestion_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _questions.Add(new QuestionItem { Text = "" });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding question: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void btnGenerateDocument_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtOrderNumber.Text))
        {
            MessageBox.Show("Please enter an Order Number.", "Required Field", 
                MessageBoxButton.OK, MessageBoxImage.Warning);
            mainTabs.SelectedIndex = 0;
            txtOrderNumber.Focus();
            return;
        }

        SaveFileDialog saveDialog = new SaveFileDialog
        {
            Filter = "Word Documents (*.docx)|*.docx",
            DefaultExt = "docx",
            FileName = $"Kickoff {txtOrderNumber.Text}.docx"
        };

        if (saveDialog.ShowDialog() == true)
        {
            try
            {
                GenerateWordDocument(saveDialog.FileName);
                MessageBox.Show($"Document successfully created!\n\nSaved to:\n{saveDialog.FileName}", 
                    "Success", MessageBoxButton.OK, MessageBoxImage.Information);

                if (MessageBox.Show("Would you like to open the document now?", "Open Document", 
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating document:\n\n{ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void GenerateWordDocument(string filePath)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, 
            WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // Header
            string headerText = $"Order {txtOrderNumber.Text}";
            if (!string.IsNullOrWhiteSpace(txtCustomerName.Text))
                headerText += $" | Customer: {txtCustomerName.Text}";
            if (!string.IsNullOrWhiteSpace(txtFinalCustomer.Text))
                headerText += $"\nFinal Customer: {txtFinalCustomer.Text}";
            AddParagraph(body, headerText, 24, true);

            // System Type
            if (!string.IsNullOrWhiteSpace(txtProjectType.Text))
                AddParagraph(body, txtProjectType.Text, 24, true);

            // Dates
            string datesText = "";
            if (dpDeliveryDate.SelectedDate.HasValue)
                datesText = $"Delivery Date: {dpDeliveryDate.SelectedDate.Value:dd.MM.yyyy}";
            if (dpDesignDueDate.SelectedDate.HasValue)
            {
                if (!string.IsNullOrEmpty(datesText)) datesText += "  |  ";
                datesText += $"Design Due Date: {dpDesignDueDate.SelectedDate.Value:dd.MM.yyyy}";
            }
            if (!string.IsNullOrEmpty(datesText))
                AddParagraph(body, datesText, 24, false);

            // Configuration
            var configLines = _configCheckBoxes
                .Where(kv => kv.Value.IsChecked == true)
                .Select(kv => kv.Key);

            if (!string.IsNullOrWhiteSpace(txtCustomConfig.Text))
            {
                var customLines = txtCustomConfig.Text
                    .Split('\n', StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => line.Trim());
                configLines = configLines.Concat(customLines);
            }

            foreach (var line in configLines)
                AddParagraph(body, line, 24, false);

            // Targets
            var targetLines = _targets
                .Where(t => !string.IsNullOrWhiteSpace(t.Qty) || !string.IsNullOrWhiteSpace(t.Details))
                .Select(t => $"{t.Qty} {t.Type}{(string.IsNullOrWhiteSpace(t.Details) ? "" : $" | {t.Details}")}".Trim());

            if (targetLines.Any())
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "Targets: " + string.Join(" | ", targetLines), 24, false);
            }

            // Paka
            if (!string.IsNullOrWhiteSpace(txtPakaNumber.Text))
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, $"Paka : {txtPakaNumber.Text}", 28, true, true);
            }

            // Reference
            if (!string.IsNullOrWhiteSpace(txtReferenceOrder.Text))
                AddParagraph(body, $"Reference order: {txtReferenceOrder.Text}", 24, false);

            // Questions
            var questions = _questions.Where(q => !string.IsNullOrWhiteSpace(q.Text)).ToList();
            if (questions.Any())
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "Questions:", 32, true, true, "EE0000");
                foreach (var question in questions)
                    AddBulletPoint(body, question.Text, 24);
            }

            // Summary
            if (!string.IsNullOrWhiteSpace(txtMeetingSummary.Text))
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "Meeting Summary:", 24, true);
                AddParagraph(body, txtMeetingSummary.Text, 24, false);
            }

            // Decisions
            if (!string.IsNullOrWhiteSpace(txtDecisions.Text))
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "Key Decisions:", 24, true);
                AddParagraph(body, txtDecisions.Text, 24, false);
            }

            // Risks
            if (!string.IsNullOrWhiteSpace(txtRisks.Text))
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "Risks / Special Requirements:", 24, true);
                AddParagraph(body, txtRisks.Text, 24, false);
            }

            mainPart.Document.Save();
        }
    }

    private void AddParagraph(Body body, string text, int fontSize = 24, bool bold = false, 
        bool underline = false, string color = "000000")
    {
        Paragraph para = body.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());

        RunProperties runProperties = new RunProperties();
        runProperties.AppendChild(new FontSize { Val = fontSize.ToString() });
        runProperties.AppendChild(new FontSizeComplexScript { Val = fontSize.ToString() });

        if (bold)
        {
            runProperties.AppendChild(new Bold());
            runProperties.AppendChild(new BoldComplexScript());
        }

        if (underline)
            runProperties.AppendChild(new Underline { Val = UnderlineValues.Single });

        if (color != "000000")
            runProperties.AppendChild(new Color { Val = color });

        run.AppendChild(runProperties);
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private void AddBulletPoint(Body body, string text, int fontSize = 24)
    {
        Paragraph para = body.AppendChild(new Paragraph());

        Run bulletRun = para.AppendChild(new Run());
        RunProperties bulletRunProps = new RunProperties();
        bulletRunProps.AppendChild(new FontSize { Val = fontSize.ToString() });
        bulletRun.AppendChild(bulletRunProps);
        bulletRun.AppendChild(new Text("• ") { Space = SpaceProcessingModeValues.Preserve });

        Run textRun = para.AppendChild(new Run());
        RunProperties runProperties = new RunProperties();
        runProperties.AppendChild(new FontSize { Val = fontSize.ToString() });
        runProperties.AppendChild(new FontSizeComplexScript { Val = fontSize.ToString() });
        textRun.AppendChild(runProperties);
        textRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }
}

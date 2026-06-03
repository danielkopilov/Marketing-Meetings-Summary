using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Xml.Linq;
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

public class BoldFontWeightConverter : System.Windows.Data.IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
        value is true ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal;
    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
        throw new NotImplementedException();
}

public class BoldBackgroundConverter : System.Windows.Data.IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
        value is true
            ? (object)new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(186, 230, 253))
            : System.Windows.Media.Brushes.Transparent;
    public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) =>
        throw new NotImplementedException();
}

public class QuestionItem : INotifyPropertyChanged
{
    private string _text = "";
    private int _number;
    private bool _isBold;

    public string Text
    {
        get => _text;
        set { _text = value; OnPropertyChanged(nameof(Text)); }
    }

    public int Number
    {
        get => _number;
        set { _number = value; OnPropertyChanged(nameof(Number)); }
    }

    public bool IsBold
    {
        get => _isBold;
        set { _isBold = value; OnPropertyChanged(nameof(IsBold)); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public class TemplateData
{
    public string OrderNumber { get; set; } = "";
    public string CustomerName { get; set; } = "";
    public string FinalCustomer { get; set; } = "";
    public string ProjectType { get; set; } = "";   // legacy alias – mirrors SystemType
    public string SystemType { get; set; } = "";
    public string SystemVariant { get; set; } = "";
    public string SystemAperture { get; set; } = "";
    public string PakaNumber { get; set; } = "";
    public string ReferenceOrder { get; set; } = "";
    public string DeliveryDate { get; set; } = "";
    public string DesignDueDate { get; set; } = "";
    public List<string> Participants { get; set; } = new();
    public string SellingPrice { get; set; } = "";
    public string MaterialCost { get; set; } = "";
    public string ProjectHours { get; set; } = "";
    public string Penalties { get; set; } = "";
    public bool DORated { get; set; }
    public Dictionary<string, bool> ConfigCheckBoxes { get; set; } = new();
    public Dictionary<string, string> ComponentNotes { get; set; } = new();
    public string BBType { get; set; } = "";
    public string BBSize { get; set; } = "";
    public string ISAperture { get; set; } = "";
    public string BacklightType { get; set; } = "";
    public string MaxWeight { get; set; } = "";
    public string FiniteDistance { get; set; } = "";
    public string Vrs1 { get; set; } = "";
    public string Vrs2 { get; set; } = "";
    public string Vrs3 { get; set; } = "";
    public string Vrs4 { get; set; } = "";
    public string GimbalSize { get; set; } = "";
    public string GimbalLoadCapacity { get; set; } = "";
    public string GimbalAccuracy { get; set; } = "";
    public bool GimbalJoystick { get; set; }
    public bool LosHalogen { get; set; }
    public bool SourceStageManual { get; set; }
    public bool XyStageManual { get; set; }
    public string FrameGrabber1 { get; set; } = "";
    public string FrameGrabber2 { get; set; } = "";
    public string FrameGrabber3 { get; set; } = "";
    public string FrameGrabber4 { get; set; } = "";
    public bool RackmountMonitorArm { get; set; }
    public string RackmountHeight { get; set; } = "";
    public string OpticalTableWidth { get; set; } = "";
    public string OpticalTableLength { get; set; } = "";
    public string OpticalTableHeight { get; set; } = "";
    public bool OpticalTableActive { get; set; }
    public List<TargetItem> Targets { get; set; } = new();
    public List<string> PmQuestions { get; set; } = new();
    public List<string> MarketingQuestions { get; set; } = new();
    public List<string> MarketingNotes { get; set; } = new();
    public List<string> PmNotes { get; set; } = new();
}

public partial class MainWindow : Window
{
    // Configuration items organized by sections
    private readonly string[] _radiationSourceItems = new[]
    {
        "B.B",
        "I.S",
        "Backlight",
        "LOS Laser",
        "QTH Lamp"
    };

    private readonly string[] _systemComponentsItems = new[]
    {
        "Source Stage",
        "Rackmount",
        "Manual Choke",
        "CTE",
        "Device Center",
        "LOS alignment target",
        "XY Stage",
        "Power Meter",
        "Energy Meter",
        "NewPort Stage",
        "Focus Stage",
        "VRS",
        "Gimbal",
        "Frame Grabbers",
        "Optical Table"
    };

    private readonly Dictionary<string, WpfCheckBox> _configCheckBoxes = new();
    private readonly Dictionary<string, string> _componentNotes = new();
    private System.Windows.Threading.DispatcherTimer? _autoSaveTimer;
    private System.Windows.Threading.DispatcherTimer? _countdownTimer;
    private int _countdownSeconds = 60;
    private string? _lastExportedPath;
    private System.Windows.Controls.ComboBox? _bbTypeComboBox;
    private System.Windows.Controls.ComboBox? _bbSizeComboBox;
    private System.Windows.Controls.ComboBox? _isExitApertureComboBox;
    private System.Windows.Controls.ComboBox? _backlightTypeComboBox;
    private System.Windows.Controls.TextBox? _maxWeightTextBox;
    private System.Windows.Controls.TextBox? _finiteDistance1TextBox;
    private System.Windows.Controls.TextBox? _finiteDistance2TextBox;
    private System.Windows.Controls.TextBox? _finiteDistance3TextBox;
    private System.Windows.Controls.ComboBox? _vrsComboBox1;
    private System.Windows.Controls.ComboBox? _vrsComboBox2;
    private System.Windows.Controls.ComboBox? _vrsComboBox3;
    private System.Windows.Controls.ComboBox? _vrsComboBox4;
    private System.Windows.Controls.TextBox? _gimbalSizeTextBox;
    private System.Windows.Controls.TextBox? _gimbalLoadCapacityTextBox;
    private WpfCheckBox? _gimbalJoystickCheckBox;
    private System.Windows.Controls.ComboBox? _gimbalAccuracyComboBox;
    private System.Windows.Controls.ComboBox? _frameGrabbersComboBox;
    private System.Windows.Controls.ComboBox? _frameGrabbersComboBox2;
    private System.Windows.Controls.ComboBox? _frameGrabbersComboBox3;
    private System.Windows.Controls.ComboBox? _frameGrabbersComboBox4;
    private WpfCheckBox? _losHalogenCheckBox;
    private WpfCheckBox? _sourceStageManualCheckBox;
    private WpfCheckBox? _xyStageManualCheckBox;
    private WpfCheckBox? _rackmountMonitorArmCheckBox;
    private System.Windows.Controls.TextBox? _rackmountHeightTextBox;
    private System.Windows.Controls.TextBlock? _lblRackmountHeight, _lblRackmountU;
    private System.Windows.Controls.TextBox? _opticalTableWidthTextBox;
    private System.Windows.Controls.TextBox? _opticalTableLengthTextBox;
    private System.Windows.Controls.TextBox? _opticalTableHeightTextBox;
    private WpfCheckBox? _opticalTableActiveCheckBox;
    private System.Windows.Controls.TextBlock? _lblOTWidth, _lblOTLength, _lblOTHeight;
    // Inline label references for enable/disable
    private System.Windows.Controls.TextBlock? _lblBBType, _lblBBSize;
    private System.Windows.Controls.TextBlock? _lblISAperture;
    private System.Windows.Controls.TextBlock? _lblBacklightType;
    private System.Windows.Controls.TextBlock? _lblMaxWeight, _lblKG;
    private System.Windows.Controls.TextBlock? _lblFiniteDistance, _lblM;
    private System.Windows.Controls.TextBlock? _lblGimbalSize, _lblInches, _lblLoadCapacity, _lblKGGimbal, _lblAccuracy;
    private readonly ObservableCollection<TargetItem> _targets = new();
    private readonly ObservableCollection<QuestionItem> _questions = new();
    private readonly ObservableCollection<QuestionItem> _marketingQuestions = new();
    private readonly ObservableCollection<QuestionItem> _marketingNotes = new();
    private readonly ObservableCollection<QuestionItem> _pmNotes = new();

    // Form controls
    private System.Windows.Controls.TextBox txtOrderNumber = new();
    private System.Windows.Controls.TextBox txtCustomerName = new();
    private System.Windows.Controls.TextBox txtContactPerson = new();
    private System.Windows.Controls.TextBox txtFinalCustomer = new();
    private System.Windows.Controls.ComboBox cmbSystemType    = new();
    private System.Windows.Controls.ComboBox cmbSystemVariant = new();
    private System.Windows.Controls.ComboBox cmbSystemAperture = new();
    private System.Windows.Controls.TextBox txtPakaNumber = new();
    private System.Windows.Controls.DatePicker dpDeliveryDate = new();
    private System.Windows.Controls.DatePicker dpDesignDueDate = new();
    private System.Windows.Controls.WrapPanel? participantsPanel;
    private System.Windows.Controls.TextBox? txtParticipantsInput;
    private System.Windows.Controls.TextBox txtReferenceOrder = new();
    private List<string> selectedParticipants = new();

    // Marketing Overview controls
    private System.Windows.Controls.TextBox txtSellingPrice = new();
    private System.Windows.Controls.TextBox txtMaterialCost = new();
    private System.Windows.Controls.TextBox txtMondayCRM = new();
    private System.Windows.Controls.TextBox txtProjectHours = new();
    private System.Windows.Controls.TextBox txtPenalties = new();
    private WpfCheckBox chkDORated = new();

    // Helper property to get participants as a string
    private string GetParticipantsText()
    {
        return string.Join("; ", selectedParticipants);
    }
    private System.Windows.Controls.ItemsControl configItemsControl = new();
    private System.Windows.Controls.TextBox txtCustomConfig = new();
    private System.Windows.Controls.ItemsControl targetsItemsControl = new();
    private System.Windows.Controls.ItemsControl questionsItemsControl = new();
    private System.Windows.Controls.ItemsControl marketingQuestionsItemsControl = new();
    private System.Windows.Controls.TextBox txtActions = new();
    private System.Windows.Controls.ItemsControl marketingNotesItemsControl = new();
    private System.Windows.Controls.ItemsControl pmNotesItemsControl = new();

    public MainWindow()
    {
        try
        {
            InitializeComponent();
            InitializeConfigItems();
            InitializeTargets();
            InitializeQuestions();
            InitializeMarketingQuestions();
            InitializeNotes();
            StartAutoSaveTimer();

            // Load Overview section after window is loaded
            this.Loaded += MainWindow_Loaded;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error initializing window: {ex.Message}\n\n{ex.StackTrace}", 
                "Initialization Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        try
        {
            LoadSection(0);
            UpdateBlockDiagramButtonState();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading Overview section: {ex.Message}\n\n{ex.StackTrace}", 
                "Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void NavigateToSection(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button button) return;

        int sectionIndex = int.Parse(button.Tag.ToString() ?? "0");
        LoadSection(sectionIndex);

        // Update menu button styles
        UpdateMenuSelection(button);
    }

    private void UpdateMenuSelection(System.Windows.Controls.Button selectedButton)
    {
        // Find all menu buttons by searching the visual tree
        var sideMenu = this.FindName("btnOverview") as System.Windows.Controls.Button;
        if (sideMenu?.Parent is System.Windows.Controls.Panel panel)
        {
            foreach (var child in panel.Children)
            {
                if (child is System.Windows.Controls.Button btn)
                {
                    btn.Background = System.Windows.Media.Brushes.Transparent;
                    btn.Foreground = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(100, 116, 139)); // #64748B
                }
            }
        }

        // Highlight selected button
        selectedButton.Background = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(59, 130, 246)); // #3B82F6
        selectedButton.Foreground = System.Windows.Media.Brushes.White;
    }

    private void UpdateBlockDiagramButtonState()
    {
        var btn = this.FindName("btnBlockDiagram") as System.Windows.Controls.Button;
        if (btn == null) return;
        bool anyChecked = _configCheckBoxes.Values.Any(cb => cb.IsChecked == true);
        btn.IsEnabled = anyChecked;
        btn.Opacity = anyChecked ? 1.0 : 0.4;
    }

    private void LoadSection(int sectionIndex)
    {
        try
        {
            var panel = this.FindName("contentPanel") as System.Windows.Controls.Panel;
            if (panel == null)
            {
                MessageBox.Show("Error: contentPanel not found in XAML!", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            panel.Children.Clear();

            switch (sectionIndex)
            {
                case 0: // Overview
                    LoadOverviewSection(panel);
                    break;
                case 1: // Configuration
                    LoadConfigurationSection(panel);
                    break;
                case 2: // Targets
                    LoadTargetsSection(panel);
                    break;
                case 3: // Actions
                    LoadActionsSection(panel);
                    break;
                case 4: // Block Diagram
                    LoadBlockDiagramSection(panel);
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error in LoadSection: {ex.Message}\n\n{ex.StackTrace}", 
                "Section Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadOverviewSection(System.Windows.Controls.Panel panel)
    {
        // ===== ORDER OVERVIEW CARD =====
        var orderCard = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(1),
            Padding = new WpfThickness(16),
            Margin = new WpfThickness(0, 0, 0, 12)
        };

        var orderStackPanel = new System.Windows.Controls.StackPanel();
        orderCard.Child = orderStackPanel;

        // Title
        var orderTitle = new System.Windows.Controls.TextBlock
        {
            Text = "Order Overview",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        orderStackPanel.Children.Add(orderTitle);

        // Create form grid with 3 columns
        var orderFormGrid = new System.Windows.Controls.Grid();
        orderFormGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        orderFormGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        orderFormGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        orderStackPanel.Children.Add(orderFormGrid);

        int row = 0;

        // Participants FIRST - spans full width
        AddParticipantsFieldWithAutocomplete(orderFormGrid, "Participants:", participantsPanel, txtParticipantsInput, row++, 0, 3);

        // Order Number - spans 1 column (1/3 width)
        AddFormField(orderFormGrid, "Order Number:", txtOrderNumber, row, 0, 1);

        // Customer Name - spans 1 column (1/3 width)
        AddFormField(orderFormGrid, "Customer Name:", txtCustomerName, row, 1, 1);

        // Customer/Agent Contact Person - spans 1 column (1/3 width)
        AddFormField(orderFormGrid, "Customer/Agent Contact Person:", txtContactPerson, row++, 2, 1);

        // Agent and Project Type side by side
        AddFormField(orderFormGrid, "Territory:", txtFinalCustomer, row, 0, 1);

        // System Type — three inline combo boxes spanning 2 columns, wrapped in rounded border
        {
            orderFormGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            var st = new System.Windows.Controls.StackPanel { Margin = new WpfThickness(0, 0, 0, 8) };
            st.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "System Type:",
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Medium,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(71, 85, 105)),
                Margin = new WpfThickness(0, 0, 0, 4)
            });

            // Rounded border wrapping all three combos — matches text-field style
            var outerBorder = new WpfBorder
            {
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)),
                BorderThickness = new WpfThickness(1),
                CornerRadius = new CornerRadius(6),
                Background = System.Windows.Media.Brushes.White,
                Padding = new WpfThickness(6, 2, 6, 2)
            };

            var typeRow = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            };

            // ── Type combo ───────────────────────────────────────────────────────
            bool firstLoad = cmbSystemType.Items.Count == 0;
            if (firstLoad)
                cmbSystemType = MakeCompactComboBox(95, "METS", "ILET", "WFOV", "CFT");
            typeRow.Children.Add(new System.Windows.Controls.TextBlock { Text = "Type:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 6, 0), FontSize = 12, Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(71, 85, 105)) });
            if (cmbSystemType.Parent is System.Windows.Controls.Panel p1) p1.Children.Remove(cmbSystemType);
            typeRow.Children.Add(cmbSystemType);

            // separator
            typeRow.Children.Add(new System.Windows.Controls.Border { Width = 1, Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)), Margin = new WpfThickness(14, 2, 14, 2) });

            // ── Variant combo ────────────────────────────────────────────────────
            if (firstLoad)
                cmbSystemVariant = MakeCompactComboBox(75);
            typeRow.Children.Add(new System.Windows.Controls.TextBlock { Text = "Variant:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 6, 0), FontSize = 12, Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(71, 85, 105)) });
            if (cmbSystemVariant.Parent is System.Windows.Controls.Panel p2) p2.Children.Remove(cmbSystemVariant);
            typeRow.Children.Add(cmbSystemVariant);

            // separator
            typeRow.Children.Add(new System.Windows.Controls.Border { Width = 1, Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)), Margin = new WpfThickness(14, 2, 14, 2) });

            // ── Aperture combo ───────────────────────────────────────────────────
            if (firstLoad)
                cmbSystemAperture = MakeCompactComboBox(75, "8\"", "10\"", "12\"", "14\"", "16\"", "19\"", "21\"");
            typeRow.Children.Add(new System.Windows.Controls.TextBlock { Text = "Aperture:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 6, 0), FontSize = 12, Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(71, 85, 105)) });
            if (cmbSystemAperture.Parent is System.Windows.Controls.Panel p3) p3.Children.Remove(cmbSystemAperture);
            typeRow.Children.Add(cmbSystemAperture);

            outerBorder.Child = typeRow;
            st.Children.Add(outerBorder);

            System.Windows.Controls.Grid.SetRow(st, row);
            System.Windows.Controls.Grid.SetColumn(st, 1);
            System.Windows.Controls.Grid.SetColumnSpan(st, 2);
            orderFormGrid.Children.Add(st);

            // Wire up SelectionChanged only on first load
            if (firstLoad)
                cmbSystemType.SelectionChanged += (s, e) => RefreshVariantOptions();
            row++;
        }

        // Paka Number, Delivery Date, Design Due Date (all side by side)
        AddFormField(orderFormGrid, "Paka Number:", txtPakaNumber, row, 0, 1);
        AddFormField(orderFormGrid, "Delivery Date:", dpDeliveryDate, row, 1, 1);
        AddFormField(orderFormGrid, "Design Due Date:", dpDesignDueDate, row++, 2, 1);

        // Reference Order - spans full width
        AddFormField(orderFormGrid, "Reference Order:", txtReferenceOrder, row++, 0, 3);

        panel.Children.Add(orderCard);

        // ===== MARKETING OVERVIEW CARD =====
        var marketingCard = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(1),
            Padding = new WpfThickness(16),
            Margin = new WpfThickness(0, 0, 0, 12)
        };

        var marketingStackPanel = new System.Windows.Controls.StackPanel();
        marketingCard.Child = marketingStackPanel;

        // Marketing Overview title
        var marketingTitle = new System.Windows.Controls.TextBlock
        {
            Text = "Marketing Overview",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        marketingStackPanel.Children.Add(marketingTitle);

        // Create form grid with 3 columns
        var marketingFormGrid = new System.Windows.Controls.Grid();
        marketingFormGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        marketingFormGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        marketingFormGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        marketingStackPanel.Children.Add(marketingFormGrid);

        int marketingRow = 0;

        // Selling Price, Material Cost and Monday # (CRM)
        AddFormField(marketingFormGrid, "Selling Price:", txtSellingPrice, marketingRow, 0, 1);
        AddFormField(marketingFormGrid, "Material Cost:", txtMaterialCost, marketingRow, 1, 1);
        AddFormField(marketingFormGrid, "Monday # (CRM):", txtMondayCRM, marketingRow++, 2, 1);

        // Project Hours and Penalties
        AddFormField(marketingFormGrid, "Project Hours:", txtProjectHours, marketingRow, 0, 1);
        AddFormField(marketingFormGrid, "Penalties:", txtPenalties, marketingRow++, 1, 1);

        // D.O rated checkbox
        AddCheckBoxField(marketingFormGrid, "D.O rated", chkDORated, marketingRow++, 0, 1);

        panel.Children.Add(marketingCard);
    }

    private void AddCheckBoxField(System.Windows.Controls.Grid grid, string labelText, 
        WpfCheckBox checkBox, int row, int column, int columnSpan)
    {
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var fieldStack = new System.Windows.Controls.StackPanel
        {
            Margin = new WpfThickness(0, 0, column < 2 ? 8 : 0, 8)
        };

        // Remove from any previous parent
        if (checkBox.Parent is System.Windows.Controls.Panel oldPanel)
        {
            oldPanel.Children.Remove(checkBox);
        }

        checkBox.Content = labelText;
        checkBox.FontSize = 13;
        checkBox.FontWeight = System.Windows.FontWeights.Medium;
        checkBox.Foreground = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(71, 85, 105));
        checkBox.Margin = new WpfThickness(0, 4, 0, 0);

        fieldStack.Children.Add(checkBox);

        System.Windows.Controls.Grid.SetRow(fieldStack, row);
        System.Windows.Controls.Grid.SetColumn(fieldStack, column);
        System.Windows.Controls.Grid.SetColumnSpan(fieldStack, columnSpan);
        grid.Children.Add(fieldStack);
    }

    private void AddFormField(System.Windows.Controls.Grid grid, string labelText, 
        System.Windows.Controls.Control control, int row, int column, int columnSpan)
    {
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var fieldStack = new System.Windows.Controls.StackPanel
        {
            Margin = new WpfThickness(0, 0, column < 2 ? 8 : 0, 8)
        };

        var label = new System.Windows.Controls.TextBlock
        {
            Text = labelText,
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        fieldStack.Children.Add(label);

        // Apply modern style to control
        ApplyModernControlStyle(control);

        SafeAddChild(fieldStack, control);

        System.Windows.Controls.Grid.SetRow(fieldStack, row);
        System.Windows.Controls.Grid.SetColumn(fieldStack, column);
        System.Windows.Controls.Grid.SetColumnSpan(fieldStack, columnSpan);
        grid.Children.Add(fieldStack);
    }

    private void ApplyModernControlStyle(System.Windows.Controls.Control control)
    {
        control.FontSize = 13;
        control.Padding = new WpfThickness(8, 4, 8, 4); // Reduced vertical padding from 6 to 4
        control.BorderBrush = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(226, 232, 240));
        control.BorderThickness = new WpfThickness(1);
        control.Margin = new WpfThickness(0, 0, 0, 0);

        if (control is System.Windows.Controls.TextBox textBox)
        {
            // Apply rounded corner template manually
            var template = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.TextBox));
            var borderFactory = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
            borderFactory.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
            borderFactory.SetValue(WpfBorder.BorderBrushProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderBrushProperty));
            borderFactory.SetValue(WpfBorder.BorderThicknessProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderThicknessProperty));
            borderFactory.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(6));

            var scrollFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ScrollViewer));
            scrollFactory.Name = "PART_ContentHost";
            scrollFactory.SetValue(System.Windows.Controls.ScrollViewer.MarginProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.PaddingProperty));

            borderFactory.AppendChild(scrollFactory);
            template.VisualTree = borderFactory;
            textBox.Template = template;
        }
        else if (control is System.Windows.Controls.DatePicker datePicker)
        {
            datePicker.Style = (System.Windows.Style)FindResource("RoundedDatePicker");
        }
    }

    private void ApplyModernButtonStyle(System.Windows.Controls.Button button)
    {
        button.Background = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(59, 130, 246));
        button.Foreground = System.Windows.Media.Brushes.White;
        button.BorderThickness = new WpfThickness(0);
        button.Padding = new WpfThickness(14, 8, 14, 8);
        button.FontSize = 14;
        button.FontWeight = System.Windows.FontWeights.SemiBold;
        button.Cursor = System.Windows.Input.Cursors.Hand;

        var template = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var borderFactory = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        borderFactory.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        borderFactory.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(16));
        borderFactory.SetValue(WpfBorder.PaddingProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.PaddingProperty));

        var contentFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
        contentFactory.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        contentFactory.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);

        borderFactory.AppendChild(contentFactory);
        template.VisualTree = borderFactory;
        button.Template = template;
    }

    private void SafeAddChild(System.Windows.Controls.Panel parent, System.Windows.UIElement child)
    {
        // Remove from previous parent if it has one
        if (child is System.Windows.FrameworkElement element && element.Parent is System.Windows.Controls.Panel oldParent)
        {
            oldParent.Children.Remove(child);
        }
        parent.Children.Add(child);
    }

    private System.Windows.Controls.Primitives.Popup? _participantsPopup;
    private System.Windows.Controls.ListBox? _participantsListBox;
    private List<string> _outlookContacts = new();
    private int _lastAtPosition = -1;

    private void AddParticipantsFieldWithAutocomplete(System.Windows.Controls.Grid grid, string labelText, 
        System.Windows.Controls.WrapPanel? chipsPanel, System.Windows.Controls.TextBox? inputBox, int row, int column, int columnSpan)
    {
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        var fieldStack = new System.Windows.Controls.StackPanel
        {
            Margin = new WpfThickness(0, 0, column < 2 ? 8 : 0, 8)
        };

        var label = new System.Windows.Controls.TextBlock
        {
            Text = labelText,
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        fieldStack.Children.Add(label);

        // Container with border that looks like a textbox
        var containerBorder = new WpfBorder
        {
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(6),
            Background = System.Windows.Media.Brushes.White,
            Padding = new WpfThickness(4, 0, 4, 0),
            Height = 32,
            MaxHeight = 32
        };

        // Create new WrapPanel each time
        participantsPanel = new System.Windows.Controls.WrapPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        // Re-add all existing participant chips
        foreach (var participant in selectedParticipants)
        {
            AddParticipantChipInternal(participantsPanel, participant);
        }

        // Create new Input textbox each time
        txtParticipantsInput = new System.Windows.Controls.TextBox
        {
            BorderThickness = new WpfThickness(0),
            Background = System.Windows.Media.Brushes.Transparent,
            MinWidth = 100,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Padding = new WpfThickness(4, 0, 4, 0), // Reduced from 4,6,4,6
            FontSize = 13
        };
        txtParticipantsInput.TextChanged += ParticipantsTextBox_TextChanged;
        txtParticipantsInput.PreviewKeyDown += ParticipantsTextBox_PreviewKeyDown;

        // Add input box to the panel
        participantsPanel.Children.Add(txtParticipantsInput);

        containerBorder.Child = participantsPanel;
        fieldStack.Children.Add(containerBorder);

        // Create popup for autocomplete suggestions
        _participantsPopup = new System.Windows.Controls.Primitives.Popup
        {
            PlacementTarget = containerBorder,
            Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom,
            Width = 300,
            MaxHeight = 200,
            StaysOpen = false,
            AllowsTransparency = true
        };

        // Create ListBox for contact suggestions
        _participantsListBox = new System.Windows.Controls.ListBox
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(59, 130, 246)),
            BorderThickness = new WpfThickness(2),
            FontSize = 13,
            Padding = new WpfThickness(0)
        };
        _participantsListBox.MouseLeftButtonUp += ParticipantsListBox_MouseClick;

        var popupBorder = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(59, 130, 246)),
            BorderThickness = new WpfThickness(2),
            CornerRadius = new CornerRadius(6),
            MaxHeight = 200,
            Child = _participantsListBox,
            Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                ShadowDepth = 2,
                BlurRadius = 8,
                Opacity = 0.3
            }
        };

        _participantsPopup.Child = popupBorder;

        System.Windows.Controls.Grid.SetRow(fieldStack, row);
        System.Windows.Controls.Grid.SetColumn(fieldStack, column);
        System.Windows.Controls.Grid.SetColumnSpan(fieldStack, columnSpan);
        grid.Children.Add(fieldStack);

        // Load Outlook contacts in background
        LoadOutlookContactsAsync();
    }

    private void AddParticipantChip(string name)
    {
        if (selectedParticipants.Contains(name)) return;

        selectedParticipants.Add(name);

        if (participantsPanel != null)
        {
            AddParticipantChipInternal(participantsPanel, name);
        }
    }

    private void AddParticipantChipInternal(System.Windows.Controls.WrapPanel panel, string name)
    {
        // Create chip/pill
        var chipBorder = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(225, 239, 254)), // Light blue background
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(147, 197, 253)),
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(12),
            Padding = new WpfThickness(8, 4, 4, 4),
            Margin = new WpfThickness(2),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        var chipPanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal
        };

        // Name text
        var nameText = new System.Windows.Controls.TextBlock
        {
            Text = name,
            FontSize = 13,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 64, 175)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Margin = new WpfThickness(0, 0, 4, 0)
        };
        chipPanel.Children.Add(nameText);

        // Remove X — plain TextBlock so no Button ControlTemplate can hide it
        var removeX = new System.Windows.Controls.TextBlock
        {
            Text = "✕",
            FontSize = 13,
            FontWeight = System.Windows.FontWeights.Bold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(220, 38, 38)),
            Cursor = System.Windows.Input.Cursors.Hand,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Margin = new WpfThickness(4, 0, 2, 0),
            Tag = name,
            ToolTip = $"Remove {name}"
        };
        removeX.MouseEnter += (s, e) =>
            removeX.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(153, 27, 27));
        removeX.MouseLeave += (s, e) =>
            removeX.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(220, 38, 38));
        removeX.MouseLeftButtonUp += RemoveParticipantChip_Click;
        chipPanel.Children.Add(removeX);

        chipBorder.Child = chipPanel;

        // Insert chip before the input textbox
        if (txtParticipantsInput != null)
        {
            int inputIndex = panel.Children.IndexOf(txtParticipantsInput);
            if (inputIndex >= 0)
            {
                panel.Children.Insert(inputIndex, chipBorder);
            }
            else
            {
                panel.Children.Add(chipBorder);
            }
        }
        else
        {
            panel.Children.Add(chipBorder);
        }
    }

    private void RemoveParticipantChip_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not System.Windows.Controls.TextBlock removeX) return;
        if (removeX.Tag is not string name) return;

        selectedParticipants.Remove(name);

        // Walk up: TextBlock → StackPanel → Border
        if (removeX.Parent is System.Windows.Controls.Panel chipPanel &&
            chipPanel.Parent is WpfBorder chipBorder &&
            participantsPanel != null)
        {
            participantsPanel.Children.Remove(chipBorder);
        }

        txtParticipantsInput?.Focus();
    }

    private async void LoadOutlookContactsAsync()
    {
        await Task.Run(() =>
        {
            try
            {
                Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null) return;

                dynamic outlookApp = Activator.CreateInstance(outlookType)!;
                dynamic outlookNamespace = outlookApp.GetNamespace("MAPI");

                // Get the Contacts folder
                dynamic contactsFolder = outlookNamespace.GetDefaultFolder(10); // 10 = olFolderContacts
                dynamic contacts = contactsFolder.Items;

                var contactList = new List<string>();

                foreach (dynamic contact in contacts)
                {
                    try
                    {
                        string? name = contact.FullName?.ToString();
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            contactList.Add(name);
                        }
                    }
                    catch { /* Skip invalid contacts */ }
                }

                // Also try to get Global Address List
                try
                {
                    dynamic addressLists = outlookNamespace.AddressLists;
                    foreach (dynamic addressList in addressLists)
                    {
                        if (addressList.Name == "Global Address List")
                        {
                            dynamic entries = addressList.AddressEntries;
                            foreach (dynamic entry in entries)
                            {
                                try
                                {
                                    string? name = entry.Name?.ToString();
                                    if (!string.IsNullOrWhiteSpace(name) && !contactList.Contains(name))
                                    {
                                        contactList.Add(name);
                                    }
                                }
                                catch { /* Skip invalid entries */ }
                            }
                            break;
                        }
                    }
                }
                catch { /* GAL not available */ }

                // Cleanup
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNamespace);
                outlookApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);

                // Update UI on main thread
                Dispatcher.Invoke(() =>
                {
                    _outlookContacts = contactList.OrderBy(c => c).ToList();
                });
            }
            catch
            {
                // Silently fail if Outlook is not available
            }
        });
    }

    private void ParticipantsTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        if (sender is not System.Windows.Controls.TextBox textBox) return;
        if (_participantsPopup == null || _participantsListBox == null) return;

        string text = textBox.Text;
        int caretIndex = textBox.CaretIndex;

        // Check if user just typed '@'
        if (caretIndex > 0 && caretIndex <= text.Length && text[caretIndex - 1] == '@')
        {
            _lastAtPosition = caretIndex - 1;
            ShowContactSuggestions("");
            return;
        }

        // If we're in autocomplete mode (after @)
        if (_lastAtPosition >= 0 && caretIndex > _lastAtPosition)
        {
            // Get the text after the @ symbol
            int searchStart = _lastAtPosition + 1;
            int searchLength = caretIndex - searchStart;

            if (searchStart <= text.Length && searchLength >= 0)
            {
                string searchText = searchStart + searchLength <= text.Length 
                    ? text.Substring(searchStart, searchLength) 
                    : "";
                ShowContactSuggestions(searchText);
            }
        }
        else
        {
            // Close popup if we're not in autocomplete mode
            _participantsPopup.IsOpen = false;
            _lastAtPosition = -1;
        }
    }

    private void ShowContactSuggestions(string searchText)
    {
        if (_participantsListBox == null || _participantsPopup == null) return;

        // Filter contacts based on search text
        var filteredContacts = string.IsNullOrEmpty(searchText)
            ? _outlookContacts.Take(10).ToList()
            : _outlookContacts
                .Where(c => c.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                .Take(10)
                .ToList();

        _participantsListBox.Items.Clear();

        if (filteredContacts.Count > 0)
        {
            foreach (var contact in filteredContacts)
            {
                _participantsListBox.Items.Add(contact);
            }

            _participantsPopup.IsOpen = true;
            _participantsListBox.SelectedIndex = 0;
        }
        else
        {
            _participantsPopup.IsOpen = false;
        }
    }

    private void ParticipantsTextBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (_participantsPopup?.IsOpen != true || _participantsListBox == null) return;

        // Handle arrow keys and Enter for navigation and selection
        if (e.Key == System.Windows.Input.Key.Down)
        {
            e.Handled = true;
            if (_participantsListBox.SelectedIndex < _participantsListBox.Items.Count - 1)
            {
                _participantsListBox.SelectedIndex++;
                _participantsListBox.ScrollIntoView(_participantsListBox.SelectedItem);
            }
        }
        else if (e.Key == System.Windows.Input.Key.Up)
        {
            e.Handled = true;
            if (_participantsListBox.SelectedIndex > 0)
            {
                _participantsListBox.SelectedIndex--;
                _participantsListBox.ScrollIntoView(_participantsListBox.SelectedItem);
            }
        }
        else if (e.Key == System.Windows.Input.Key.Enter || e.Key == System.Windows.Input.Key.Tab)
        {
            e.Handled = true;
            if (_participantsListBox.SelectedItem != null)
            {
                InsertSelectedContact(_participantsListBox.SelectedItem.ToString()!);
            }
        }
        else if (e.Key == System.Windows.Input.Key.Escape)
        {
            e.Handled = true;
            _participantsPopup.IsOpen = false;
            _lastAtPosition = -1;
        }
    }

    private void ParticipantsListBox_MouseClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender is not System.Windows.Controls.ListBox listBox) return;
        if (listBox.SelectedItem == null) return;

        InsertSelectedContact(listBox.SelectedItem.ToString()!);
    }

    private void InsertSelectedContact(string contactName)
    {
        // Add the contact as a chip
        AddParticipantChip(contactName);

        // Clear the input textbox
        if (txtParticipantsInput != null)
        {
            txtParticipantsInput.Text = "";
        }
        _lastAtPosition = -1;

        // Close popup
        if (_participantsPopup != null)
        {
            _participantsPopup.IsOpen = false;
        }

        // Focus back to input textbox
        txtParticipantsInput?.Focus();
    }

    private void LoadConfigurationSection(System.Windows.Controls.Panel panel)
    {
        // ── Save existing inline-control values before rebuilding the UI ──────
        var savedBBType         = _bbTypeComboBox?.SelectedItem?.ToString();
        var savedBBSize         = _bbSizeComboBox?.SelectedItem?.ToString();
        var savedISAperture     = _isExitApertureComboBox?.SelectedItem?.ToString();
        var savedBacklightType  = _backlightTypeComboBox?.SelectedItem?.ToString();
        var savedMaxWeight      = _maxWeightTextBox?.Text;
        var savedFiniteDist     = _finiteDistance1TextBox?.Text;
        var savedVrs1           = _vrsComboBox1?.SelectedItem?.ToString();
        var savedVrs2           = _vrsComboBox2?.SelectedItem?.ToString();
        var savedVrs3           = _vrsComboBox3?.SelectedItem?.ToString();
        var savedVrs4           = _vrsComboBox4?.SelectedItem?.ToString();
        var savedGimbalSize     = _gimbalSizeTextBox?.Text;
        var savedGimbalLC       = _gimbalLoadCapacityTextBox?.Text;
        var savedGimbalAccuracy = _gimbalAccuracyComboBox?.SelectedItem?.ToString();
        var savedGimbalJoystick = _gimbalJoystickCheckBox?.IsChecked ?? false;
        var savedLosHalogen     = _losHalogenCheckBox?.IsChecked ?? false;
        var savedSourceStageManual = _sourceStageManualCheckBox?.IsChecked ?? false;
        var savedXyStageManual     = _xyStageManualCheckBox?.IsChecked ?? false;
        var savedFG1            = _frameGrabbersComboBox?.SelectedItem?.ToString();
        var savedFG2            = _frameGrabbersComboBox2?.SelectedItem?.ToString();
        var savedFG3            = _frameGrabbersComboBox3?.SelectedItem?.ToString();
        var savedFG4            = _frameGrabbersComboBox4?.SelectedItem?.ToString();
        var savedRackmountMonitorArm = _rackmountMonitorArmCheckBox?.IsChecked ?? false;
        var savedRackmountHeight    = _rackmountHeightTextBox?.Text;
        var savedOTWidth            = _opticalTableWidthTextBox?.Text;
        var savedOTLength           = _opticalTableLengthTextBox?.Text;
        var savedOTHeight           = _opticalTableHeightTextBox?.Text;
        var savedOTActive           = _opticalTableActiveCheckBox?.IsChecked ?? false;

        // Main Configuration Card - matching the image exactly
        var card = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderThickness = new WpfThickness(0),
            Padding = new WpfThickness(16, 8, 32, 32), // Reduced top padding from 12 to 8
            Margin = new WpfThickness(0, 0, 0, 0)
        };

        // Soft shadow
        card.Effect = new System.Windows.Media.Effects.DropShadowEffect
        {
            ShadowDepth = 0,
            BlurRadius = 15,
            Opacity = 0.08,
            Color = System.Windows.Media.Color.FromRgb(0, 0, 0)
        };

        var stackPanel = new System.Windows.Controls.StackPanel();
        card.Child = stackPanel;

        // Title - Configuration (less bold, closer to top)
        var title = new System.Windows.Controls.TextBlock
        {
            Text = "Configuration",
            FontSize = 24,
            FontWeight = System.Windows.FontWeights.Normal,
            Foreground = System.Windows.Media.Brushes.Black,
            Margin = new WpfThickness(0, 0, 0, 8)
        };
        stackPanel.Children.Add(title);

        // Subtitle
        var subtitle = new System.Windows.Controls.TextBlock
        {
            Text = "Select the relevant options for this meeting.",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.Normal,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(128, 128, 128)),
            Margin = new WpfThickness(0, 0, 0, 16) // Reduced bottom margin from 32 to 16
        };
        stackPanel.Children.Add(subtitle);

        // ===== RADIATION SOURCE SECTION =====
        var radiationSourceSection = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White, // White background for the whole section
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(229, 231, 235)), // Light border around entire section
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new WpfThickness(0),
            Margin = new WpfThickness(0, 0, 0, 24), // No left margin - starts at edge
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch, // Stretch to full width
            MaxWidth = 920 // Limit width but start from left edge
        };

        var radiationSourcePanel = new System.Windows.Controls.StackPanel();
        radiationSourceSection.Child = radiationSourcePanel;

        // Title with icon - lighter blue background
        var radiationTitleContainer = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(239, 246, 255)), // Lighter blue background (#EFF6FF)
            Padding = new WpfThickness(16, 8, 16, 8), // Reduced vertical padding for narrower box
            CornerRadius = new CornerRadius(8, 8, 0, 0) // Rounded top corners only
        };

        var radiationTitlePanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center // Center everything vertically
        };

        var radiationIcon = new System.Windows.Controls.TextBlock
        {
            Text = "☢️",
            FontSize = 16,
            Margin = new WpfThickness(0, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        radiationTitlePanel.Children.Add(radiationIcon);

        var radiationSourceTitle = new System.Windows.Controls.TextBlock
        {
            Text = "Radiation Source",
            FontSize = 15,
            FontWeight = System.Windows.FontWeights.Bold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)), // Blue color
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Margin = new WpfThickness(0, -1, 0, 0) // Move text up slightly (less than before)
        };
        radiationTitlePanel.Children.Add(radiationSourceTitle);

        radiationTitleContainer.Child = radiationTitlePanel;
        radiationSourcePanel.Children.Add(radiationTitleContainer);

        // Options container with white background and padding
        var optionsContainer = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            Padding = new WpfThickness(16, 16, 16, 16)
        };

        // Grid layout: Col0=checkbox(100), Col1=label(110), Col2=combo(140), Col3=label(60), Col4=combo(100)
        var radiationGrid = new System.Windows.Controls.Grid();
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(110) }); // checkbox
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(110) }); // label1
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(140) }); // combo1
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(60)  }); // label2
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(110) }); // combo2

        int radRowIdx = 0;

        // B.B, I.S, Backlight each get their own row (they have inline combo controls).
        // LOS Laser + QTH Lamp share one row side-by-side.
        var simpleRadItems = new List<string>();

        foreach (var item in _radiationSourceItems)
        {
            bool isSimple = item is "LOS Laser" or "QTH Lamp";
            if (isSimple) { simpleRadItems.Add(item); continue; }

            radiationGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

            var checkBox = _configCheckBoxes[item];
            if (checkBox.Parent is WpfBorder ob) ob.Child = null;
            else if (checkBox.Parent is System.Windows.Controls.Panel op) op.Children.Remove(checkBox);

            checkBox.Margin = new WpfThickness(0);
            checkBox.FontSize = 14;
            checkBox.FontWeight = System.Windows.FontWeights.Normal;
            checkBox.Foreground = System.Windows.Media.Brushes.Black;
            checkBox.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            checkBox.VerticalContentAlignment = System.Windows.VerticalAlignment.Center;
            if (checkBox.Content is System.Windows.Controls.TextBlock cbtb) { cbtb.FontSize = 14; cbtb.FontWeight = System.Windows.FontWeights.Normal; }

            var cbWrapper = WrapCheckBoxWithNoteButton(checkBox, item);
            cbWrapper.Margin = new WpfThickness(0, 0, 0, 14);
            cbWrapper.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            System.Windows.Controls.Grid.SetRow(cbWrapper, radRowIdx);
            System.Windows.Controls.Grid.SetColumn(cbWrapper, 0);
            radiationGrid.Children.Add(cbWrapper);

            if (item == "B.B")
            {
                var lbl1 = MakeComboLabel("Type:");
                lbl1.Margin = new WpfThickness(0, 0, 6, 14);
                System.Windows.Controls.Grid.SetRow(lbl1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(lbl1, 1);
                radiationGrid.Children.Add(lbl1);
                _lblBBType = lbl1;

                _bbTypeComboBox = MakeWhiteComboBox(138, "RR", "STD", "SR200N-33");
                _bbTypeComboBox.SelectedIndex = -1;
                var w1 = MakeComboWrapper(_bbTypeComboBox, 138);
                w1.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w1, 2);
                radiationGrid.Children.Add(w1);

                var lbl2 = MakeComboLabel("Size:");
                lbl2.Margin = new WpfThickness(8, 0, 6, 14);
                System.Windows.Controls.Grid.SetRow(lbl2, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(lbl2, 3);
                radiationGrid.Children.Add(lbl2);
                _lblBBSize = lbl2;

                _bbSizeComboBox = MakeWhiteComboBox(108, "1D", "2D", "4D", "8D", "12D");
                var w2 = MakeComboWrapper(_bbSizeComboBox, 108);
                w2.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w2, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w2, 4);
                radiationGrid.Children.Add(w2);

                // Initially disabled
                SetInlineEnabled(false, lbl1, lbl2);
                _bbTypeComboBox.IsEnabled = false; _bbSizeComboBox.IsEnabled = false;
                var bbCb = _configCheckBoxes["B.B"];
                bbCb.Checked   += (s,e) => { SetInlineEnabled(true,  lbl1, lbl2); _bbTypeComboBox.IsEnabled = true; _bbSizeComboBox.IsEnabled = true; };
                bbCb.Unchecked += (s,e) => { SetInlineEnabled(false, lbl1, lbl2); _bbTypeComboBox.IsEnabled = false; _bbSizeComboBox.IsEnabled = false; };
            }
            else if (item == "I.S")
            {
                var lbl1 = MakeComboLabel("Exit Aperture:");
                lbl1.Margin = new WpfThickness(0, 0, 6, 14);
                System.Windows.Controls.Grid.SetRow(lbl1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(lbl1, 1);
                radiationGrid.Children.Add(lbl1);
                _lblISAperture = lbl1;

                _isExitApertureComboBox = MakeWhiteComboBox(138, "2\"", "3\"", "4\"", "5\"");
                var w1 = MakeComboWrapper(_isExitApertureComboBox, 138);
                w1.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w1, 2);
                radiationGrid.Children.Add(w1);

                SetInlineEnabled(false, lbl1);
                _isExitApertureComboBox.IsEnabled = false;
                var isCb = _configCheckBoxes["I.S"];
                isCb.Checked   += (s,e) => { SetInlineEnabled(true,  lbl1); _isExitApertureComboBox.IsEnabled = true; };
                isCb.Unchecked += (s,e) => { SetInlineEnabled(false, lbl1); _isExitApertureComboBox.IsEnabled = false; };
            }
            else if (item == "Backlight")
            {
                var lbl1 = MakeComboLabel("Type:");
                lbl1.Margin = new WpfThickness(0, 0, 6, 14);
                System.Windows.Controls.Grid.SetRow(lbl1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(lbl1, 1);
                radiationGrid.Children.Add(lbl1);
                _lblBacklightType = lbl1;

                _backlightTypeComboBox = MakeWhiteComboBox(138, "LED", "Fiber Optic");
                var w1 = MakeComboWrapper(_backlightTypeComboBox, 138);
                w1.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w1, 2);
                radiationGrid.Children.Add(w1);

                SetInlineEnabled(false, lbl1);
                _backlightTypeComboBox.IsEnabled = false;
                var blCb = _configCheckBoxes["Backlight"];
                blCb.Checked   += (s,e) => { SetInlineEnabled(true,  lbl1); _backlightTypeComboBox.IsEnabled = true; };
                blCb.Unchecked += (s,e) => { SetInlineEnabled(false, lbl1); _backlightTypeComboBox.IsEnabled = false; };
            }

            radRowIdx++;
        }

        // LOS Laser + QTH Lamp on one shared row
        if (simpleRadItems.Count > 0)
        {
            radiationGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int colOffset = 0;
            foreach (var si in simpleRadItems)
            {
                var cb = _configCheckBoxes[si];
                if (cb.Parent is WpfBorder ob2) ob2.Child = null;
                else if (cb.Parent is System.Windows.Controls.Panel op2) op2.Children.Remove(cb);

                cb.Margin = new WpfThickness(0);
                cb.FontSize = 14;
                cb.FontWeight = System.Windows.FontWeights.Normal;
                cb.Foreground = System.Windows.Media.Brushes.Black;
                cb.VerticalAlignment = System.Windows.VerticalAlignment.Center;
                cb.VerticalContentAlignment = System.Windows.VerticalAlignment.Center;
                if (cb.Content is System.Windows.Controls.TextBlock stb) { stb.FontSize = 14; stb.FontWeight = System.Windows.FontWeights.Normal; }

                var siWrapper = WrapCheckBoxWithNoteButton(cb, si);
                siWrapper.Margin = new WpfThickness(0, 10, 24, 14);
                siWrapper.VerticalAlignment = System.Windows.VerticalAlignment.Center;
                System.Windows.Controls.Grid.SetRow(siWrapper, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(siWrapper, colOffset);
                // span 2 for LOS Laser (col 0) so the note icon isn't clipped; span 3 for QTH Lamp
                System.Windows.Controls.Grid.SetColumnSpan(siWrapper, colOffset == 0 ? 2 : 3);
                radiationGrid.Children.Add(siWrapper);
                colOffset = 2; // QTH Lamp starts at col 2, spans remaining
            }
        }

        optionsContainer.Child = radiationGrid;
        radiationSourcePanel.Children.Add(optionsContainer);
        stackPanel.Children.Add(radiationSourceSection);

        // ===== SYSTEM COMPONENTS SECTION =====
        var systemComponentsSection = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White, // White background for the whole section
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(229, 231, 235)), // Light border around entire section
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new WpfThickness(0),
            Margin = new WpfThickness(0, 0, 0, 24),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch, // Stretch to full width
            MaxWidth = 920 // Limit width but start from left edge
        };

        var systemComponentsPanel = new System.Windows.Controls.StackPanel();
        systemComponentsSection.Child = systemComponentsPanel;

        // Title with icon - lighter blue background
        var systemTitleContainer = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(239, 246, 255)), // Lighter blue background (#EFF6FF)
            Padding = new WpfThickness(16, 8, 16, 8), // Reduced vertical padding for narrower box
            CornerRadius = new CornerRadius(8, 8, 0, 0) // Rounded top corners only
        };

        var systemTitlePanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center // Center everything vertically
        };

        var systemIcon = new System.Windows.Controls.TextBlock
        {
            Text = "📦",
            FontSize = 16,
            Margin = new WpfThickness(0, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        systemTitlePanel.Children.Add(systemIcon);

        var systemComponentsTitle = new System.Windows.Controls.TextBlock
        {
            Text = "System Components",
            FontSize = 15,
            FontWeight = System.Windows.FontWeights.Bold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)), // Blue color
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Margin = new WpfThickness(0, -1, 0, 0) // Move text up slightly to align with icon
        };
        systemTitlePanel.Children.Add(systemComponentsTitle);

        systemTitleContainer.Child = systemTitlePanel;
        systemComponentsPanel.Children.Add(systemTitleContainer);

        // Options container with white background and padding
        var systemOptionsContainer = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            Padding = new WpfThickness(16, 16, 16, 16)
        };

        // Helper: detach a checkbox from any previous parent
        void DetachCheckBox(WpfCheckBox cb)
        {
            if (cb.Parent is WpfBorder ob) ob.Child = null;
            else if (cb.Parent is System.Windows.Controls.Panel op) op.Children.Remove(cb);
        }

        // Helper: style a checkbox for System Components (not bold)
        void StyleSysCheckBox(WpfCheckBox cb)
        {
            cb.Margin = new WpfThickness(0);
            cb.FontSize = 14;
            cb.FontWeight = System.Windows.FontWeights.Normal;
            cb.Foreground = System.Windows.Media.Brushes.Black;
            cb.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            cb.VerticalContentAlignment = System.Windows.VerticalAlignment.Center;
            if (cb.Content is System.Windows.Controls.TextBlock cbtb2) { cbtb2.FontSize = 14; cbtb2.FontWeight = System.Windows.FontWeights.Normal; }
        }

        // Helper: create a small label
        System.Windows.Controls.TextBlock SysLabel(string text) =>
            new System.Windows.Controls.TextBlock
            {
                Text = text,
                FontSize = 14,
                FontWeight = System.Windows.FontWeights.Normal,
                Foreground = System.Windows.Media.Brushes.Black,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new WpfThickness(6, 0, 4, 0)
            };

        // Helper: create a small textbox
        System.Windows.Controls.TextBox SysTextBox(double width) =>
            new System.Windows.Controls.TextBox
            {
                Width = width,
                Height = 26,
                FontSize = 14,
                Padding = new WpfThickness(4, 1, 4, 1),
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Center
            };

        // Helper: inline combobox
        System.Windows.Controls.ComboBox SysComboBox(double width) =>
            new System.Windows.Controls.ComboBox
            {
                Width = width,
                Height = 26,
                FontSize = 14,
                FontWeight = System.Windows.FontWeights.Normal,
                Padding = new WpfThickness(4, 1, 4, 1),
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new WpfThickness(4, 0, 0, 0)
            };

        // Grid layout:
        // col0 = 190px  (left checkboxes: Source Stage, CTE, XY Stage + inline rows — fixed so spacing equals col1/col2)
        // col1 = 190px  (rows 0-2: Device Center/Power Meter  |  rows 3/4/6: inline dim label)
        // col2 = 190px  (rows 0-2: Manual Choke/LOS alignment target/Energy Meter  |  rows 3/4/6: inline textbox)
        // col3 = Auto   (inline unit: [KG], [m], [Inches] — empty in rows 0-2)
        var componentsGrid = new System.Windows.Controls.Grid();
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(190) });  // col0
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(190) });  // col1
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(190) });  // col2
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });       // col3 unit

        // Rows 0-2: 8 plain checkboxes — Rackmount is placed separately below Optical Table
        // Layout: row0: Source Stage(col0), Device Center(col1), Manual Choke(col2)
        //         row1: CTE(col0), Power Meter(col1), LOS alignment target(col2)
        //         row2: XY Stage(col0), Energy Meter(col1)
        string[] top8      = { "Source Stage", "Device Center", "Manual Choke",
                                "CTE",          "Power Meter",   "LOS alignment target",
                                "XY Stage",     "Energy Meter" };
        int[]    top8Cols  = { 0, 1, 2,  0, 1, 2,  0, 1 };
        int[]    top8Rows  = { 0, 0, 0,  1, 1, 1,  2, 2 };

        // Ensure 3 row definitions for rows 0-2
        componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

        for (int i = 0; i < top8.Length; i++)
        {
            var cb = _configCheckBoxes[top8[i]];
            DetachCheckBox(cb);
            StyleSysCheckBox(cb);
            cb.Margin = new WpfThickness(0, 0, 0, 0);

            var cbw = WrapCheckBoxWithNoteButton(cb, top8[i]);

            if (top8[i] == "Source Stage")
            {
                _sourceStageManualCheckBox = new WpfCheckBox
                {
                    Content = new System.Windows.Controls.TextBlock
                    {
                        Text = "Manual",
                        FontSize = 11,
                        FontWeight = System.Windows.FontWeights.Normal,
                        Foreground = System.Windows.Media.Brushes.Black,
                        Margin = new WpfThickness(2, 0, 0, 0)
                    },
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new WpfThickness(6, 0, 0, 0),
                    Padding = new WpfThickness(0),
                    LayoutTransform = new System.Windows.Media.ScaleTransform(0.8, 0.8)
                };
                _sourceStageManualCheckBox.IsEnabled = false;
                var ssRow = new System.Windows.Controls.StackPanel
                {
                    Orientation = System.Windows.Controls.Orientation.Horizontal,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new WpfThickness(0, 0, 0, 14)
                };
                ssRow.Children.Add(cbw);
                ssRow.Children.Add(_sourceStageManualCheckBox);
                System.Windows.Controls.Grid.SetRow(ssRow, top8Rows[i]);
                System.Windows.Controls.Grid.SetColumn(ssRow, top8Cols[i]);
                componentsGrid.Children.Add(ssRow);
                var ssCb = _configCheckBoxes["Source Stage"];
                ssCb.Checked   += (s, e) => { _sourceStageManualCheckBox.IsEnabled = true; };
                ssCb.Unchecked += (s, e) => { _sourceStageManualCheckBox.IsEnabled = false; };
            }
            else if (top8[i] == "XY Stage")
            {
                _xyStageManualCheckBox = new WpfCheckBox
                {
                    Content = new System.Windows.Controls.TextBlock
                    {
                        Text = "Manual",
                        FontSize = 11,
                        FontWeight = System.Windows.FontWeights.Normal,
                        Foreground = System.Windows.Media.Brushes.Black,
                        Margin = new WpfThickness(2, 0, 0, 0)
                    },
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new WpfThickness(6, 0, 0, 0),
                    Padding = new WpfThickness(0),
                    LayoutTransform = new System.Windows.Media.ScaleTransform(0.8, 0.8)
                };
                _xyStageManualCheckBox.IsEnabled = false;
                var xyRow = new System.Windows.Controls.StackPanel
                {
                    Orientation = System.Windows.Controls.Orientation.Horizontal,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new WpfThickness(0, 0, 0, 14)
                };
                xyRow.Children.Add(cbw);
                xyRow.Children.Add(_xyStageManualCheckBox);
                System.Windows.Controls.Grid.SetRow(xyRow, top8Rows[i]);
                System.Windows.Controls.Grid.SetColumn(xyRow, top8Cols[i]);
                componentsGrid.Children.Add(xyRow);
                var xyCb = _configCheckBoxes["XY Stage"];
                xyCb.Checked   += (s, e) => { _xyStageManualCheckBox.IsEnabled = true; };
                xyCb.Unchecked += (s, e) => { _xyStageManualCheckBox.IsEnabled = false; };
            }
            else if (top8[i] == "LOS alignment target")
            {
                // Wrap the note-button wrapper + Halogen checkbox in a horizontal StackPanel
                _losHalogenCheckBox = new WpfCheckBox
                {
                    Content = new System.Windows.Controls.TextBlock
                    {
                        Text = "+Halogen",
                        FontSize = 11,
                        FontWeight = System.Windows.FontWeights.Normal,
                        Foreground = System.Windows.Media.Brushes.Black,
                        Margin = new WpfThickness(2, 0, 0, 0)
                    },
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new WpfThickness(6, 0, 0, 0),
                    Padding = new WpfThickness(0),
                    LayoutTransform = new System.Windows.Media.ScaleTransform(0.8, 0.8)
                };
                var losRow = new System.Windows.Controls.StackPanel
                {
                    Orientation = System.Windows.Controls.Orientation.Horizontal,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new WpfThickness(0, 0, 0, 14)
                };
                losRow.Children.Add(cbw);
                losRow.Children.Add(_losHalogenCheckBox);
                System.Windows.Controls.Grid.SetRow(losRow, top8Rows[i]);
                System.Windows.Controls.Grid.SetColumn(losRow, top8Cols[i]);
                System.Windows.Controls.Grid.SetColumnSpan(losRow, 2);
                componentsGrid.Children.Add(losRow);
            }
            else
            {
                cbw.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(cbw, top8Rows[i]);
                System.Windows.Controls.Grid.SetColumn(cbw, top8Cols[i]);
                componentsGrid.Children.Add(cbw);
            }
        }

        // Dimmed style helpers for inline labels/units
        var dimColor = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(160, 160, 160));
        const double dimFontSize = 12;

        System.Windows.Controls.TextBlock DimLabel(string text) =>
            new System.Windows.Controls.TextBlock
            {
                Text = text,
                FontSize = dimFontSize,
                FontWeight = System.Windows.FontWeights.Normal,
                Foreground = dimColor,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new WpfThickness(4, 0, 2, 14)
            };

        System.Windows.Controls.TextBox DimTextBox(double width) =>
            new System.Windows.Controls.TextBox
            {
                Width = width,
                Height = 24,
                FontSize = dimFontSize,
                Foreground = dimColor,
                Padding = new WpfThickness(3, 1, 3, 1),
                Margin = new WpfThickness(2, 0, 0, 14),
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Center
            };

        // ── Row 3: NewPort Stage | Max Weight: [___] [KG] ─────────────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 3;

            var cbNP = _configCheckBoxes["NewPort Stage"];
            DetachCheckBox(cbNP); StyleSysCheckBox(cbNP);
            cbNP.Margin = new WpfThickness(0, 0, 0, 0);

            var lblMW = DimLabel("Max Weight:"); lblMW.Margin = new WpfThickness(20, 0, 4, 0);
            _maxWeightTextBox = DimTextBox(55); _maxWeightTextBox.Margin = new WpfThickness(0);
            var lblKG = DimLabel("[KG]"); lblKG.Margin = new WpfThickness(4, 0, 0, 0);
            _lblMaxWeight = lblMW; _lblKG = lblKG;

            var mwRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbNPWrapper = WrapCheckBoxWithNoteButton(cbNP, "NewPort Stage");
            mwRow.Children.Add(cbNPWrapper); mwRow.Children.Add(lblMW); mwRow.Children.Add(_maxWeightTextBox); mwRow.Children.Add(lblKG);
            System.Windows.Controls.Grid.SetRow(mwRow, r); System.Windows.Controls.Grid.SetColumn(mwRow, 0); System.Windows.Controls.Grid.SetColumnSpan(mwRow, 4);
            componentsGrid.Children.Add(mwRow);

            SetInlineEnabled(false, lblMW, lblKG);
            _maxWeightTextBox.IsEnabled = false;
            cbNP.Checked   += (s, e) => { SetInlineEnabled(true,  lblMW, lblKG); _maxWeightTextBox.IsEnabled = true;  _maxWeightTextBox.Foreground = System.Windows.Media.Brushes.Black; };
            cbNP.Unchecked += (s, e) => { SetInlineEnabled(false, lblMW, lblKG); _maxWeightTextBox.IsEnabled = false; _maxWeightTextBox.Foreground = _greyBrush; };
        }

        // ── Row 4: Focus Stage | Finite Distance: [___] [m] ──────────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 4;

            var cbFS = _configCheckBoxes["Focus Stage"];
            DetachCheckBox(cbFS); StyleSysCheckBox(cbFS);
            cbFS.Margin = new WpfThickness(0, 0, 0, 0);

            var lblFD = DimLabel("Finite Distance:"); lblFD.Margin = new WpfThickness(20, 0, 4, 0);
            _finiteDistance1TextBox = DimTextBox(55); _finiteDistance1TextBox.Margin = new WpfThickness(0);
            _finiteDistance2TextBox = null; _finiteDistance3TextBox = null;
            var lblM = DimLabel("[m]"); lblM.Margin = new WpfThickness(4, 0, 0, 0);
            _lblFiniteDistance = lblFD; _lblM = lblM;

            var fdRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbFSWrapper = WrapCheckBoxWithNoteButton(cbFS, "Focus Stage");
            fdRow.Children.Add(cbFSWrapper); fdRow.Children.Add(lblFD); fdRow.Children.Add(_finiteDistance1TextBox); fdRow.Children.Add(lblM);
            System.Windows.Controls.Grid.SetRow(fdRow, r); System.Windows.Controls.Grid.SetColumn(fdRow, 0); System.Windows.Controls.Grid.SetColumnSpan(fdRow, 4);
            componentsGrid.Children.Add(fdRow);

            SetInlineEnabled(false, lblFD, lblM);
            _finiteDistance1TextBox.IsEnabled = false;
            cbFS.Checked   += (s, e) => { SetInlineEnabled(true,  lblFD, lblM); _finiteDistance1TextBox.IsEnabled = true;  _finiteDistance1TextBox.Foreground = System.Windows.Media.Brushes.Black; };
            cbFS.Unchecked += (s, e) => { SetInlineEnabled(false, lblFD, lblM); _finiteDistance1TextBox.IsEnabled = false; _finiteDistance1TextBox.Foreground = _greyBrush; };
        }

        // ── Row 5: VRS | 4 wavelength dropdowns ──────────────────────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 5;

            var cbVRS = _configCheckBoxes["VRS"];
            DetachCheckBox(cbVRS); StyleSysCheckBox(cbVRS);
            cbVRS.Margin = new WpfThickness(0, 0, 0, 0);
            var cbVRSWrapper = WrapCheckBoxWithNoteButton(cbVRS, "VRS");

            string[] vrsOptions = { "NA", "1550 [nm]", "1570 [nm]", "1064 [nm]", "1540 [nm]" };
            _vrsComboBox1 = MakeWhiteComboBox(110, vrsOptions); _vrsComboBox1.SelectedIndex = 0;
            _vrsComboBox2 = MakeWhiteComboBox(110, vrsOptions); _vrsComboBox2.SelectedIndex = 0;
            _vrsComboBox3 = MakeWhiteComboBox(110, vrsOptions); _vrsComboBox3.SelectedIndex = 0;
            _vrsComboBox4 = MakeWhiteComboBox(110, vrsOptions); _vrsComboBox4.SelectedIndex = 0;
            _vrsComboBox2.Margin = new WpfThickness(6, 0, 0, 0);
            _vrsComboBox3.Margin = new WpfThickness(6, 0, 0, 0);
            _vrsComboBox4.Margin = new WpfThickness(6, 0, 0, 0);

            var vrsRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            vrsRow.Children.Add(cbVRSWrapper);
            vrsRow.Children.Add(MakeComboWrapper(_vrsComboBox1, 110));
            vrsRow.Children.Add(MakeComboWrapper(_vrsComboBox2, 110));
            vrsRow.Children.Add(MakeComboWrapper(_vrsComboBox3, 110));
            vrsRow.Children.Add(MakeComboWrapper(_vrsComboBox4, 110));
            System.Windows.Controls.Grid.SetRow(vrsRow, r); System.Windows.Controls.Grid.SetColumn(vrsRow, 0); System.Windows.Controls.Grid.SetColumnSpan(vrsRow, 4);
            componentsGrid.Children.Add(vrsRow);

            _vrsComboBox1.IsEnabled = false; _vrsComboBox2.IsEnabled = false;
            _vrsComboBox3.IsEnabled = false; _vrsComboBox4.IsEnabled = false;
            cbVRS.Checked   += (s, e) => { _vrsComboBox1.IsEnabled = true;  _vrsComboBox2.IsEnabled = true;  _vrsComboBox3.IsEnabled = true;  _vrsComboBox4.IsEnabled = true; };
            cbVRS.Unchecked += (s, e) => { _vrsComboBox1.IsEnabled = false; _vrsComboBox2.IsEnabled = false; _vrsComboBox3.IsEnabled = false; _vrsComboBox4.IsEnabled = false; };
        }

        // ── Row 6: Gimbal | +Joystick | Size:[__][Inches] | Load Capacity:[__][KG] | Accuracy:[v] ──
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 6;

            var cbGimbal = _configCheckBoxes["Gimbal"];
            DetachCheckBox(cbGimbal); StyleSysCheckBox(cbGimbal);
            cbGimbal.Margin = new WpfThickness(0, 0, 0, 0);

            // +Joystick small checkbox (right after note icon, same style as +Halogen)
            _gimbalJoystickCheckBox = new WpfCheckBox
            {
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = "+Joystick",
                    FontSize = 11,
                    FontWeight = System.Windows.FontWeights.Normal,
                    Foreground = System.Windows.Media.Brushes.Black,
                    Margin = new WpfThickness(2, 0, 0, 0)
                },
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new WpfThickness(8, 0, 0, 0),
                Padding = new WpfThickness(0),
                LayoutTransform = new System.Windows.Media.ScaleTransform(0.8, 0.8)
            };

            // Size: [__] [Inches]
            var lblSize = DimLabel("Size:"); lblSize.Margin = new WpfThickness(16, 0, 4, 0);
            _gimbalSizeTextBox = DimTextBox(55); _gimbalSizeTextBox.Margin = new WpfThickness(0);
            var lblInches = DimLabel("[Inches]"); lblInches.Margin = new WpfThickness(4, 0, 0, 0);
            _lblGimbalSize = lblSize; _lblInches = lblInches;

            // Load Capacity: [__] [KG]
            var lblLC = DimLabel("Load Capacity:"); lblLC.Margin = new WpfThickness(16, 0, 4, 0);
            _gimbalLoadCapacityTextBox = DimTextBox(55); _gimbalLoadCapacityTextBox.Margin = new WpfThickness(0);
            var lblKGG = DimLabel("[KG]"); lblKGG.Margin = new WpfThickness(4, 0, 0, 0);
            _lblLoadCapacity = lblLC; _lblKGGimbal = lblKGG;

            // Accuracy: [v]
            var lblAcc = DimLabel("Accuracy:"); lblAcc.Margin = new WpfThickness(16, 0, 4, 0);
            _lblAccuracy = lblAcc;
            _gimbalAccuracyComboBox = MakeWhiteComboBox(148, "Standard Accuracy", "High Accuracy");
            _gimbalAccuracyComboBox.SelectedIndex = -1;

            var szRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbGimbalWrapper = WrapCheckBoxWithNoteButton(cbGimbal, "Gimbal");
            szRow.Children.Add(cbGimbalWrapper);
            szRow.Children.Add(_gimbalJoystickCheckBox);
            szRow.Children.Add(lblSize);
            szRow.Children.Add(_gimbalSizeTextBox);
            szRow.Children.Add(lblInches);
            szRow.Children.Add(lblLC);
            szRow.Children.Add(_gimbalLoadCapacityTextBox);
            szRow.Children.Add(lblKGG);
            szRow.Children.Add(lblAcc);
            szRow.Children.Add(MakeComboWrapper(_gimbalAccuracyComboBox, 148));
            System.Windows.Controls.Grid.SetRow(szRow, r); System.Windows.Controls.Grid.SetColumn(szRow, 0); System.Windows.Controls.Grid.SetColumnSpan(szRow, 4);
            componentsGrid.Children.Add(szRow);

            SetInlineEnabled(false, lblSize, lblInches, lblLC, lblKGG, lblAcc);
            _gimbalSizeTextBox.IsEnabled = false;
            _gimbalLoadCapacityTextBox.IsEnabled = false;
            _gimbalJoystickCheckBox.IsEnabled = false;
            _gimbalAccuracyComboBox.IsEnabled = false;
            cbGimbal.Checked   += (s, e) => { SetInlineEnabled(true,  lblSize, lblInches, lblLC, lblKGG, lblAcc); _gimbalSizeTextBox.IsEnabled = true;  _gimbalSizeTextBox.Foreground = System.Windows.Media.Brushes.Black; _gimbalLoadCapacityTextBox.IsEnabled = true; _gimbalLoadCapacityTextBox.Foreground = System.Windows.Media.Brushes.Black; _gimbalJoystickCheckBox.IsEnabled = true; _gimbalAccuracyComboBox.IsEnabled = true; };
            cbGimbal.Unchecked += (s, e) => { SetInlineEnabled(false, lblSize, lblInches, lblLC, lblKGG, lblAcc); _gimbalSizeTextBox.IsEnabled = false; _gimbalSizeTextBox.Foreground = _greyBrush; _gimbalLoadCapacityTextBox.IsEnabled = false; _gimbalLoadCapacityTextBox.Foreground = _greyBrush; _gimbalJoystickCheckBox.IsEnabled = false; _gimbalAccuracyComboBox.IsEnabled = false; };
        }

        // ── Row 7: Frame Grabbers | up to 4 dropdowns ────────────────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 7;

            var cbFG = _configCheckBoxes["Frame Grabbers"];
            DetachCheckBox(cbFG); StyleSysCheckBox(cbFG);
            cbFG.Margin = new WpfThickness(0, 0, 0, 0);

            string[] fgOptions = { "Analog", "Camera Link", "GigE Vision", "CoaxPress", "HD-SDI" };
            _frameGrabbersComboBox  = MakeWhiteComboBox(120, fgOptions);
            _frameGrabbersComboBox2 = MakeWhiteComboBox(120, fgOptions);
            _frameGrabbersComboBox3 = MakeWhiteComboBox(120, fgOptions);
            _frameGrabbersComboBox4 = MakeWhiteComboBox(120, fgOptions);

            var fgRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbFGWrapper = WrapCheckBoxWithNoteButton(cbFG, "Frame Grabbers");
            fgRow.Children.Add(cbFGWrapper);
            fgRow.Children.Add(MakeComboWrapper(_frameGrabbersComboBox,  120));
            _frameGrabbersComboBox2.Margin = new WpfThickness(6, 0, 0, 0);
            fgRow.Children.Add(MakeComboWrapper(_frameGrabbersComboBox2, 120));
            _frameGrabbersComboBox3.Margin = new WpfThickness(6, 0, 0, 0);
            fgRow.Children.Add(MakeComboWrapper(_frameGrabbersComboBox3, 120));
            _frameGrabbersComboBox4.Margin = new WpfThickness(6, 0, 0, 0);
            fgRow.Children.Add(MakeComboWrapper(_frameGrabbersComboBox4, 120));
            System.Windows.Controls.Grid.SetRow(fgRow, r); System.Windows.Controls.Grid.SetColumn(fgRow, 0); System.Windows.Controls.Grid.SetColumnSpan(fgRow, 4);
            componentsGrid.Children.Add(fgRow);

            _frameGrabbersComboBox.IsEnabled  = false; _frameGrabbersComboBox2.IsEnabled = false;
            _frameGrabbersComboBox3.IsEnabled = false; _frameGrabbersComboBox4.IsEnabled = false;
            cbFG.Checked   += (s, e) => { _frameGrabbersComboBox.IsEnabled  = true;  _frameGrabbersComboBox2.IsEnabled = true;  _frameGrabbersComboBox3.IsEnabled = true;  _frameGrabbersComboBox4.IsEnabled = true; };
            cbFG.Unchecked += (s, e) => { _frameGrabbersComboBox.IsEnabled  = false; _frameGrabbersComboBox2.IsEnabled = false; _frameGrabbersComboBox3.IsEnabled = false; _frameGrabbersComboBox4.IsEnabled = false; };
        }

        // ── Row 8: Optical Table | Active checkbox | Width:[__] Length:[__] Height:[__] ──
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 8;

            var cbOT = _configCheckBoxes["Optical Table"];
            DetachCheckBox(cbOT); StyleSysCheckBox(cbOT);
            cbOT.Margin = new WpfThickness(0, 0, 0, 0);

            _opticalTableActiveCheckBox = new WpfCheckBox
            {
                Content = new System.Windows.Controls.TextBlock { Text = "Active", FontSize = 11, FontWeight = System.Windows.FontWeights.Normal, Foreground = System.Windows.Media.Brushes.Black, Margin = new WpfThickness(2, 0, 0, 0) },
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new WpfThickness(8, 0, 0, 0),
                Padding = new WpfThickness(0),
                LayoutTransform = new System.Windows.Media.ScaleTransform(0.8, 0.8)
            };

            var lblOTW = DimLabel("Width:");  lblOTW.Margin  = new WpfThickness(16, 0, 4, 0);
            var lblOTL = DimLabel("Length:"); lblOTL.Margin  = new WpfThickness(10, 0, 4, 0);
            var lblOTH = DimLabel("Height:"); lblOTH.Margin  = new WpfThickness(10, 0, 4, 0);
            _lblOTWidth = lblOTW; _lblOTLength = lblOTL; _lblOTHeight = lblOTH;

            _opticalTableWidthTextBox  = DimTextBox(60); _opticalTableWidthTextBox.Margin  = new WpfThickness(0);
            _opticalTableLengthTextBox = DimTextBox(60); _opticalTableLengthTextBox.Margin = new WpfThickness(0);
            _opticalTableHeightTextBox = DimTextBox(60); _opticalTableHeightTextBox.Margin = new WpfThickness(0);

            var otRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbOTWrapper = WrapCheckBoxWithNoteButton(cbOT, "Optical Table");
            otRow.Children.Add(cbOTWrapper);
            otRow.Children.Add(_opticalTableActiveCheckBox);
            otRow.Children.Add(lblOTW); otRow.Children.Add(_opticalTableWidthTextBox);
            otRow.Children.Add(lblOTL); otRow.Children.Add(_opticalTableLengthTextBox);
            otRow.Children.Add(lblOTH); otRow.Children.Add(_opticalTableHeightTextBox);
            System.Windows.Controls.Grid.SetRow(otRow, r); System.Windows.Controls.Grid.SetColumn(otRow, 0); System.Windows.Controls.Grid.SetColumnSpan(otRow, 4);
            componentsGrid.Children.Add(otRow);

            SetInlineEnabled(false, lblOTW, lblOTL, lblOTH);
            _opticalTableActiveCheckBox.IsEnabled  = false;
            _opticalTableWidthTextBox.IsEnabled    = false;
            _opticalTableLengthTextBox.IsEnabled   = false;
            _opticalTableHeightTextBox.IsEnabled   = false;
            cbOT.Checked   += (s, e) => { SetInlineEnabled(true,  lblOTW, lblOTL, lblOTH); _opticalTableActiveCheckBox.IsEnabled = true; _opticalTableWidthTextBox.IsEnabled  = true; _opticalTableWidthTextBox.Foreground  = System.Windows.Media.Brushes.Black; _opticalTableLengthTextBox.IsEnabled = true; _opticalTableLengthTextBox.Foreground = System.Windows.Media.Brushes.Black; _opticalTableHeightTextBox.IsEnabled = true; _opticalTableHeightTextBox.Foreground = System.Windows.Media.Brushes.Black; };
            cbOT.Unchecked += (s, e) => { SetInlineEnabled(false, lblOTW, lblOTL, lblOTH); _opticalTableActiveCheckBox.IsEnabled = false; _opticalTableWidthTextBox.IsEnabled  = false; _opticalTableWidthTextBox.Foreground  = _greyBrush; _opticalTableLengthTextBox.IsEnabled = false; _opticalTableLengthTextBox.Foreground = _greyBrush; _opticalTableHeightTextBox.IsEnabled = false; _opticalTableHeightTextBox.Foreground = _greyBrush; };
        }

        // ── Row 9: Rackmount | +Monitor Arm | Height: [__] [U] ───────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 9;

            var cbRM = _configCheckBoxes["Rackmount"];
            DetachCheckBox(cbRM); StyleSysCheckBox(cbRM);
            cbRM.Margin = new WpfThickness(0, 0, 0, 0);
            var cbRMWrapper = WrapCheckBoxWithNoteButton(cbRM, "Rackmount");

            _rackmountMonitorArmCheckBox = new WpfCheckBox
            {
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = "+Monitor Arm",
                    FontSize = 11,
                    FontWeight = System.Windows.FontWeights.Normal,
                    Foreground = System.Windows.Media.Brushes.Black,
                    Margin = new WpfThickness(2, 0, 0, 0)
                },
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new WpfThickness(6, 0, 0, 0),
                Padding = new WpfThickness(0),
                LayoutTransform = new System.Windows.Media.ScaleTransform(0.8, 0.8)
            };
            var lblRH = new System.Windows.Controls.TextBlock { Text = "Height:", FontSize = 12, FontWeight = System.Windows.FontWeights.Normal, Foreground = _greyBrush, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(10, 0, 3, 0) };
            _rackmountHeightTextBox = new System.Windows.Controls.TextBox { Width = 45, Height = 24, FontSize = 12, Foreground = _greyBrush, Padding = new WpfThickness(3, 1, 3, 1), VerticalAlignment = System.Windows.VerticalAlignment.Center, VerticalContentAlignment = System.Windows.VerticalAlignment.Center };
            var lblRU = new System.Windows.Controls.TextBlock { Text = "[U]", FontSize = 12, FontWeight = System.Windows.FontWeights.Normal, Foreground = _greyBrush, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(3, 0, 0, 0) };
            _lblRackmountHeight = lblRH; _lblRackmountU = lblRU;

            var rmRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            rmRow.Children.Add(cbRMWrapper);
            rmRow.Children.Add(_rackmountMonitorArmCheckBox);
            rmRow.Children.Add(lblRH);
            rmRow.Children.Add(_rackmountHeightTextBox);
            rmRow.Children.Add(lblRU);
            System.Windows.Controls.Grid.SetRow(rmRow, r);
            System.Windows.Controls.Grid.SetColumn(rmRow, 0);
            System.Windows.Controls.Grid.SetColumnSpan(rmRow, 4);
            componentsGrid.Children.Add(rmRow);

            _rackmountMonitorArmCheckBox.IsEnabled = false;
            _rackmountHeightTextBox.IsEnabled = false;
            cbRM.Checked   += (s, e) => { _rackmountMonitorArmCheckBox.IsEnabled = true;  _rackmountHeightTextBox.IsEnabled = true;  _rackmountHeightTextBox.Foreground = System.Windows.Media.Brushes.Black; SetInlineEnabled(true,  lblRH, lblRU); };
            cbRM.Unchecked += (s, e) => { _rackmountMonitorArmCheckBox.IsEnabled = false; _rackmountHeightTextBox.IsEnabled = false; _rackmountHeightTextBox.Foreground = _greyBrush;                         SetInlineEnabled(false, lblRH, lblRU); };
        }

        systemOptionsContainer.Child = componentsGrid;
        systemComponentsPanel.Children.Add(systemOptionsContainer);
        stackPanel.Children.Add(systemComponentsSection);

        // ── Restore saved inline-control values and sync enabled state ────────
        RestoreComboSelection(_bbTypeComboBox,         savedBBType);
        RestoreComboSelection(_bbSizeComboBox,         savedBBSize);
        RestoreComboSelection(_isExitApertureComboBox, savedISAperture);
        RestoreComboSelection(_backlightTypeComboBox,  savedBacklightType);
        if (_maxWeightTextBox      != null && savedMaxWeight   != null) _maxWeightTextBox.Text      = savedMaxWeight;
        if (_finiteDistance1TextBox != null && savedFiniteDist  != null) _finiteDistance1TextBox.Text = savedFiniteDist;
        RestoreComboSelection(_vrsComboBox1, savedVrs1);
        RestoreComboSelection(_vrsComboBox2, savedVrs2);
        RestoreComboSelection(_vrsComboBox3, savedVrs3);
        RestoreComboSelection(_vrsComboBox4, savedVrs4);
        if (_gimbalSizeTextBox          != null && savedGimbalSize != null) _gimbalSizeTextBox.Text           = savedGimbalSize;
        if (_gimbalLoadCapacityTextBox  != null && savedGimbalLC   != null) _gimbalLoadCapacityTextBox.Text   = savedGimbalLC;
        RestoreComboSelection(_gimbalAccuracyComboBox, savedGimbalAccuracy);
        if (_gimbalJoystickCheckBox != null) _gimbalJoystickCheckBox.IsChecked = savedGimbalJoystick;
        if (_losHalogenCheckBox        != null) _losHalogenCheckBox.IsChecked        = savedLosHalogen;
        if (_sourceStageManualCheckBox  != null) _sourceStageManualCheckBox.IsChecked = savedSourceStageManual;
        if (_xyStageManualCheckBox      != null) _xyStageManualCheckBox.IsChecked     = savedXyStageManual;
        RestoreComboSelection(_frameGrabbersComboBox,  savedFG1);
        RestoreComboSelection(_frameGrabbersComboBox2, savedFG2);
        RestoreComboSelection(_frameGrabbersComboBox3, savedFG3);
        RestoreComboSelection(_frameGrabbersComboBox4, savedFG4);
        if (_rackmountMonitorArmCheckBox != null) _rackmountMonitorArmCheckBox.IsChecked = savedRackmountMonitorArm;
        if (_rackmountHeightTextBox      != null && savedRackmountHeight != null) _rackmountHeightTextBox.Text = savedRackmountHeight;
        if (_opticalTableWidthTextBox    != null && savedOTWidth  != null) _opticalTableWidthTextBox.Text  = savedOTWidth;
        if (_opticalTableLengthTextBox   != null && savedOTLength != null) _opticalTableLengthTextBox.Text = savedOTLength;
        if (_opticalTableHeightTextBox   != null && savedOTHeight != null) _opticalTableHeightTextBox.Text = savedOTHeight;
        if (_opticalTableActiveCheckBox  != null) _opticalTableActiveCheckBox.IsChecked = savedOTActive;

        // Sync enabled state for all inline controls to match current checkbox states
        SyncInlineControlStates();

        // Apply any pending template config (inline combos/textboxes loaded from template)
        ApplyPendingTemplateConfig();

        panel.Children.Add(card);
    }

    private void LoadTargetsSection(System.Windows.Controls.Panel panel)
    {
        // Main Targets Card - matching Configuration design
        var card = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderThickness = new WpfThickness(0),
            Padding = new WpfThickness(16, 8, 32, 32),
            Margin = new WpfThickness(0, 0, 0, 0)
        };

        card.Effect = new System.Windows.Media.Effects.DropShadowEffect
        {
            ShadowDepth = 0,
            BlurRadius = 15,
            Opacity = 0.08,
            Color = System.Windows.Media.Color.FromRgb(0, 0, 0)
        };

        var stackPanel = new System.Windows.Controls.StackPanel();
        card.Child = stackPanel;

        // Title - "Targets List"
        var title = new System.Windows.Controls.TextBlock
        {
            Text = "Targets List",
            FontSize = 24,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = System.Windows.Media.Brushes.Black,
            Margin = new WpfThickness(0, 0, 0, 8)
        };
        stackPanel.Children.Add(title);

        // Subtitle
        var subtitle = new System.Windows.Controls.TextBlock
        {
            Text = "Add and manage target specifications for this project.",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.Normal,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(128, 128, 128)),
            Margin = new WpfThickness(0, 0, 0, 16)
        };
        stackPanel.Children.Add(subtitle);

        // Targets Section with light blue title background (matching Configuration style)
        var targetsSection = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(229, 231, 235)),
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new WpfThickness(0),
            Margin = new WpfThickness(0, 0, 0, 24),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch,
            MaxWidth = 920
        };

        var targetsPanel = new System.Windows.Controls.StackPanel();
        targetsSection.Child = targetsPanel;

        // Title container with light blue background and Add Target button on the right
        var titleContainer = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(239, 246, 255)), // Light blue background
            Padding = new WpfThickness(16, 3, 16, 3), // Minimal vertical padding - much less than Radiation Source's 8px
            CornerRadius = new CornerRadius(8, 8, 0, 0)
        };

        var titleGrid = new System.Windows.Controls.Grid();
        titleGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        titleGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });

        var titlePanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        var icon = new System.Windows.Controls.TextBlock
        {
            Text = "🎯",
            FontSize = 16, // Same size as Radiation Source icon
            Margin = new WpfThickness(0, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        titlePanel.Children.Add(icon);

        var sectionTitle = new System.Windows.Controls.TextBlock
        {
            Text = "Target Specifications",
            FontSize = 15, // Same size as Radiation Source text
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)), // Blue color
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Margin = new WpfThickness(0, -1, 0, 0) // Same alignment adjustment as Radiation Source
        };
        titlePanel.Children.Add(sectionTitle);

        System.Windows.Controls.Grid.SetColumn(titlePanel, 0);
        titleGrid.Children.Add(titlePanel);

        // Add Target Button (moved to title bar)
        var addButton = new System.Windows.Controls.Button
        {
            Content = "+ Add Target",
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Padding = new WpfThickness(10, 1, 10, 1), // Extremely minimal padding
            FontSize = 10, // Smaller font
            FontWeight = System.Windows.FontWeights.Medium,
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)),
            Foreground = System.Windows.Media.Brushes.White,
            BorderThickness = new WpfThickness(0),
            Cursor = System.Windows.Input.Cursors.Hand,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        var addButtonTemplate = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var addBorderFactory = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        addBorderFactory.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        addBorderFactory.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(6));
        addBorderFactory.SetValue(WpfBorder.PaddingProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.PaddingProperty));

        var addContentFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
        addContentFactory.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        addContentFactory.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);

        addBorderFactory.AppendChild(addContentFactory);
        addButtonTemplate.VisualTree = addBorderFactory;
        addButton.Template = addButtonTemplate;

        System.Windows.Controls.Grid.SetColumn(addButton, 1);
        titleGrid.Children.Add(addButton);

        titleContainer.Child = titleGrid;
        targetsPanel.Children.Add(titleContainer);

        // Content container with white background
        var contentContainer = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            Padding = new WpfThickness(16, 2, 16, 16)
        };

        var contentStack = new System.Windows.Controls.StackPanel();
        contentContainer.Child = contentStack;

        // Column Headers (#, Target Name, Size, Unit, Actions)
        var headersGrid = new System.Windows.Controls.Grid
        {
            Margin = new WpfThickness(0, 0, 0, 2)
        };
        headersGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(28) });  // #
        headersGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) }); // Target Name
        headersGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(150) }); // Size
        headersGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(150) }); // Unit
        headersGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(80) });  // Actions

        var numHeader = new System.Windows.Controls.TextBlock
        {
            Text = "#",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        };
        System.Windows.Controls.Grid.SetColumn(numHeader, 0);
        headersGrid.Children.Add(numHeader);

        var nameHeader = new System.Windows.Controls.TextBlock
        {
            Text = "Target Name",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)),
            Margin = new WpfThickness(4, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        };
        System.Windows.Controls.Grid.SetColumn(nameHeader, 1);
        headersGrid.Children.Add(nameHeader);

        var sizeHeader = new System.Windows.Controls.TextBlock
        {
            Text = "Size",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)),
            Margin = new WpfThickness(8, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        };
        System.Windows.Controls.Grid.SetColumn(sizeHeader, 2);
        headersGrid.Children.Add(sizeHeader);

        var unitHeader = new System.Windows.Controls.TextBlock
        {
            Text = "Unit",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)),
            Margin = new WpfThickness(8, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        };
        System.Windows.Controls.Grid.SetColumn(unitHeader, 3);
        headersGrid.Children.Add(unitHeader);

        var actionsHeader = new System.Windows.Controls.TextBlock
        {
            Text = "Actions",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)),
            Margin = new WpfThickness(8, 0, 0, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center
        };
        System.Windows.Controls.Grid.SetColumn(actionsHeader, 4);
        headersGrid.Children.Add(actionsHeader);

        contentStack.Children.Add(headersGrid);

        // Initialize with 3 empty targets (only if targets list is empty)
        if (_targets.Count == 0)
        {
            _targets.Add(new TargetItem { Type = "", Qty = "", Details = "mm" });
            _targets.Add(new TargetItem { Type = "", Qty = "", Details = "mm" });
            _targets.Add(new TargetItem { Type = "", Qty = "", Details = "mm" });
        }

        // Create target rows
        for (int i = 0; i < _targets.Count; i++)
        {
            var targetRow = CreateTargetRow(_targets[i], contentStack, i + 1);
            contentStack.Children.Add(targetRow);
        }

        addButton.Click += (s, e) => AddNewTarget(contentStack);

        targetsPanel.Children.Add(contentContainer);
        stackPanel.Children.Add(targetsSection);

        panel.Children.Add(card);
    }

    private WpfBorder CreateTargetRow(TargetItem target, System.Windows.Controls.Panel container, int rowNumber)
    {
        // Single bordered row matching Configuration section style
        var rowBorder = new WpfBorder
        {
            BorderThickness = new WpfThickness(0),
            Background = System.Windows.Media.Brushes.Transparent,
            Padding = new WpfThickness(0, 0, 0, 0),
            Margin = new WpfThickness(0, 0, 0, 1)
        };

        var grid = new System.Windows.Controls.Grid();
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(28) });  // #
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) }); // Target Name
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(150) }); // Size
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(150) }); // Unit
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(80) });  // Actions

        // Row number label
        var numberLabel = new System.Windows.Controls.TextBlock
        {
            Text = rowNumber.ToString(),
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            Tag = "RowNumber"
        };
        System.Windows.Controls.Grid.SetColumn(numberLabel, 0);
        grid.Children.Add(numberLabel);

        // Target Name TextBox (bordered)
        var nameTextBox = new System.Windows.Controls.TextBox
        {
            Text = target.Type,
            FontSize = 13,
            Height = 32,
            Padding = new WpfThickness(8, 0, 8, 0),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(209, 213, 219)),
            Background = System.Windows.Media.Brushes.White,
            BorderThickness = new WpfThickness(1),
            Margin = new WpfThickness(0, 0, 8, 0),
            VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        };
        ApplyRoundedTextBoxTemplate(nameTextBox);

        nameTextBox.TextChanged += (s, e) => target.Type = nameTextBox.Text;

        System.Windows.Controls.Grid.SetColumn(nameTextBox, 1);
        grid.Children.Add(nameTextBox);

        // Size TextBox (bordered)
        var sizeTextBox = new System.Windows.Controls.TextBox
        {
            Text = target.Qty,
            FontSize = 13,
            Height = 32,
            Padding = new WpfThickness(8, 0, 8, 0),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(209, 213, 219)),
            Background = System.Windows.Media.Brushes.White,
            BorderThickness = new WpfThickness(1),
            Margin = new WpfThickness(0, 0, 8, 0),
            VerticalContentAlignment = System.Windows.VerticalAlignment.Center
        };
        ApplyRoundedTextBoxTemplate(sizeTextBox);

        sizeTextBox.TextChanged += (s, e) => target.Qty = sizeTextBox.Text;

        System.Windows.Controls.Grid.SetColumn(sizeTextBox, 2);
        grid.Children.Add(sizeTextBox);

        // Unit ComboBox (borderless, inside the row border)
        // Wrap ComboBox in a rounded Border for the visual style
        var comboWrapper = new WpfBorder
        {
            CornerRadius = new CornerRadius(6),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(209, 213, 219)),
            BorderThickness = new WpfThickness(1),
            Background = System.Windows.Media.Brushes.White,
            Margin = new WpfThickness(0, 0, 8, 0),
            Height = 32
        };

        var unitComboBox = new System.Windows.Controls.ComboBox
        {
            FontSize = 13,
            BorderThickness = new WpfThickness(0),
            Background = System.Windows.Media.Brushes.Transparent,
            VerticalContentAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center,
            Padding = new WpfThickness(6, 0, 2, 0)
        };
        comboWrapper.Child = unitComboBox;
        unitComboBox.Items.Add("mm");
        unitComboBox.Items.Add("mRad");
        unitComboBox.Items.Add("cy/mRad");
        unitComboBox.Items.Add("Na");
        unitComboBox.SelectedItem = target.Details ?? "mm";
        unitComboBox.SelectionChanged += (s, e) => target.Details = unitComboBox.SelectedItem?.ToString() ?? "mm";

        System.Windows.Controls.Grid.SetColumn(comboWrapper, 3);
        grid.Children.Add(comboWrapper);

        // Delete Button (trash icon, gray color)
        var deleteButton = new System.Windows.Controls.Button
        {
            Content = "🗑️",
            FontSize = 16,
            Width = 32,
            Height = 32,
            Background = System.Windows.Media.Brushes.Transparent,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)), // Gray color
            BorderThickness = new WpfThickness(0),
            Cursor = System.Windows.Input.Cursors.Hand,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        };

        var delButtonTemplate = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var delBorderFactory = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        delBorderFactory.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        delBorderFactory.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(4));

        var delContentFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
        delContentFactory.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        delContentFactory.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);

        delBorderFactory.AppendChild(delContentFactory);
        delButtonTemplate.VisualTree = delBorderFactory;
        deleteButton.Template = delButtonTemplate;

        deleteButton.MouseEnter += (s, e) =>
        {
            deleteButton.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(239, 68, 68)); // Red on hover
        };

        deleteButton.MouseLeave += (s, e) =>
        {
            deleteButton.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(107, 114, 128)); // Gray default
        };

        deleteButton.Click += (s, e) => RemoveTargetRow(target, rowBorder, container);

        System.Windows.Controls.Grid.SetColumn(deleteButton, 4);
        grid.Children.Add(deleteButton);

        rowBorder.Child = grid;
        return rowBorder;
    }

    private void AddNewTarget(System.Windows.Controls.Panel container)
    {
        var newTarget = new TargetItem { Type = "", Qty = "", Details = "mm" };
        _targets.Add(newTarget);

        var targetRow = CreateTargetRow(newTarget, container, _targets.Count);
        container.Children.Add(targetRow);
    }

    private void RemoveTargetRow(TargetItem target, WpfBorder rowBorder, System.Windows.Controls.Panel container)
    {
        _targets.Remove(target);
        container.Children.Remove(rowBorder);

        // Renumber remaining rows (skip header rows — only WpfBorder children are data rows)
        int num = 1;
        foreach (var child in container.Children)
        {
            if (child is WpfBorder rb && rb.Child is System.Windows.Controls.Grid g)
            {
                foreach (var gc in g.Children)
                {
                    if (gc is System.Windows.Controls.TextBlock tb && tb.Tag is string tag && tag == "RowNumber")
                    {
                        tb.Text = num.ToString();
                        break;
                    }
                }
                num++;
            }
        }
    }

    private static readonly System.Windows.Media.SolidColorBrush _greyBrush =
        new(System.Windows.Media.Color.FromRgb(160, 160, 160));

    private static void SetInlineEnabled(bool enabled, params System.Windows.Controls.TextBlock[] labels)
    {
        var color = enabled
            ? System.Windows.Media.Brushes.Black
            : _greyBrush;
        foreach (var l in labels) l.Foreground = color;
    }

    private static void RestoreComboSelection(System.Windows.Controls.ComboBox? combo, string? value)
    {
        if (combo == null || value == null) return;
        for (int i = 0; i < combo.Items.Count; i++)
        {
            if (combo.Items[i]?.ToString() == value)
            {
                combo.SelectedIndex = i;
                return;
            }
        }
    }

    /// <summary>
    /// After rebuilding the configuration UI, re-apply enabled/foreground state
    /// to all inline controls based on the current checkbox states.
    /// </summary>
    private void SyncInlineControlStates()
    {
        bool bb = _configCheckBoxes.TryGetValue("B.B", out var bbCb) && bbCb.IsChecked == true;
        SetInlineEnabled(bb, _lblBBType!, _lblBBSize!);
        if (_bbTypeComboBox  != null) _bbTypeComboBox.IsEnabled  = bb;
        if (_bbSizeComboBox  != null) _bbSizeComboBox.IsEnabled  = bb;

        bool isS = _configCheckBoxes.TryGetValue("I.S", out var isCb) && isCb.IsChecked == true;
        SetInlineEnabled(isS, _lblISAperture!);
        if (_isExitApertureComboBox != null) _isExitApertureComboBox.IsEnabled = isS;

        bool bl = _configCheckBoxes.TryGetValue("Backlight", out var blCb) && blCb.IsChecked == true;
        SetInlineEnabled(bl, _lblBacklightType!);
        if (_backlightTypeComboBox != null) _backlightTypeComboBox.IsEnabled = bl;

        bool np = _configCheckBoxes.TryGetValue("NewPort Stage", out var npCb) && npCb.IsChecked == true;
        SetInlineEnabled(np, _lblMaxWeight!, _lblKG!);
        if (_maxWeightTextBox != null) { _maxWeightTextBox.IsEnabled = np; _maxWeightTextBox.Foreground = np ? System.Windows.Media.Brushes.Black : _greyBrush; }

        bool fs = _configCheckBoxes.TryGetValue("Focus Stage", out var fsCb) && fsCb.IsChecked == true;
        SetInlineEnabled(fs, _lblFiniteDistance!, _lblM!);
        if (_finiteDistance1TextBox != null) { _finiteDistance1TextBox.IsEnabled = fs; _finiteDistance1TextBox.Foreground = fs ? System.Windows.Media.Brushes.Black : _greyBrush; }

        bool vrs = _configCheckBoxes.TryGetValue("VRS", out var vrsCb) && vrsCb.IsChecked == true;
        if (_vrsComboBox1 != null) _vrsComboBox1.IsEnabled = vrs;
        if (_vrsComboBox2 != null) _vrsComboBox2.IsEnabled = vrs;
        if (_vrsComboBox3 != null) _vrsComboBox3.IsEnabled = vrs;
        if (_vrsComboBox4 != null) _vrsComboBox4.IsEnabled = vrs;

        bool gim = _configCheckBoxes.TryGetValue("Gimbal", out var gimCb) && gimCb.IsChecked == true;
        SetInlineEnabled(gim, _lblGimbalSize!, _lblInches!, _lblLoadCapacity!, _lblKGGimbal!, _lblAccuracy!);
        if (_gimbalSizeTextBox         != null) { _gimbalSizeTextBox.IsEnabled         = gim; _gimbalSizeTextBox.Foreground         = gim ? System.Windows.Media.Brushes.Black : _greyBrush; }
        if (_gimbalLoadCapacityTextBox != null) { _gimbalLoadCapacityTextBox.IsEnabled = gim; _gimbalLoadCapacityTextBox.Foreground = gim ? System.Windows.Media.Brushes.Black : _greyBrush; }
        if (_gimbalJoystickCheckBox    != null) _gimbalJoystickCheckBox.IsEnabled      = gim;
        if (_gimbalAccuracyComboBox    != null) _gimbalAccuracyComboBox.IsEnabled      = gim;

        bool fg = _configCheckBoxes.TryGetValue("Frame Grabbers", out var fgCb) && fgCb.IsChecked == true;
        if (_frameGrabbersComboBox  != null) _frameGrabbersComboBox.IsEnabled  = fg;
        if (_frameGrabbersComboBox2 != null) _frameGrabbersComboBox2.IsEnabled = fg;
        if (_frameGrabbersComboBox3 != null) _frameGrabbersComboBox3.IsEnabled = fg;
        if (_frameGrabbersComboBox4 != null) _frameGrabbersComboBox4.IsEnabled = fg;

        bool rm = _configCheckBoxes.TryGetValue("Rackmount", out var rmCb2) && rmCb2.IsChecked == true;
        if (_sourceStageManualCheckBox != null) _sourceStageManualCheckBox.IsEnabled = _configCheckBoxes.TryGetValue("Source Stage", out var ssCb2) && ssCb2.IsChecked == true;
        if (_xyStageManualCheckBox     != null) _xyStageManualCheckBox.IsEnabled     = _configCheckBoxes.TryGetValue("XY Stage",     out var xyCb2) && xyCb2.IsChecked == true;
        if (_rackmountMonitorArmCheckBox != null) _rackmountMonitorArmCheckBox.IsEnabled = rm;
        if (_rackmountHeightTextBox      != null) { _rackmountHeightTextBox.IsEnabled = rm; _rackmountHeightTextBox.Foreground = rm ? System.Windows.Media.Brushes.Black : _greyBrush; }
        SetInlineEnabled(rm, _lblRackmountHeight!, _lblRackmountU!);

        bool ot = _configCheckBoxes.TryGetValue("Optical Table", out var otCb) && otCb.IsChecked == true;
        if (_opticalTableActiveCheckBox  != null) _opticalTableActiveCheckBox.IsEnabled  = ot;
        if (_opticalTableWidthTextBox    != null) { _opticalTableWidthTextBox.IsEnabled  = ot; _opticalTableWidthTextBox.Foreground  = ot ? System.Windows.Media.Brushes.Black : _greyBrush; }
        if (_opticalTableLengthTextBox   != null) { _opticalTableLengthTextBox.IsEnabled = ot; _opticalTableLengthTextBox.Foreground = ot ? System.Windows.Media.Brushes.Black : _greyBrush; }
        if (_opticalTableHeightTextBox   != null) { _opticalTableHeightTextBox.IsEnabled = ot; _opticalTableHeightTextBox.Foreground = ot ? System.Windows.Media.Brushes.Black : _greyBrush; }
        SetInlineEnabled(ot, _lblOTWidth!, _lblOTLength!, _lblOTHeight!);
    }

    private System.Windows.Controls.TextBlock MakeComboLabel(string text) =>
        new System.Windows.Controls.TextBlock
        {
            Text = text,
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.Normal,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Margin = new WpfThickness(16, 0, 6, 0)
        };

    private System.Windows.Controls.ComboBox MakeWhiteComboBox(int width, params string[] items)
    {
        var cb = new System.Windows.Controls.ComboBox
        {
            Width = width,
            Style = (System.Windows.Style)FindResource("WhiteComboBox")
        };
        foreach (var item in items)
            cb.Items.Add(item);
        return cb;
    }

    private System.Windows.Controls.ComboBox MakeCompactComboBox(int width, params string[] items)
    {
        var cb = new System.Windows.Controls.ComboBox
        {
            Width = width,
            Style = (System.Windows.Style)FindResource("CompactRoundedComboBox")
        };
        foreach (var item in items)
            cb.Items.Add(item);
        return cb;
    }

    private System.Windows.Controls.ComboBox MakeComboWrapper(System.Windows.Controls.ComboBox cb, int width)
    {
        cb.Width = width;
        return cb;
    }

    private System.Windows.Controls.ControlTemplate CreateRoundedTextBoxControlTemplate()
    {
        var template = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.TextBox));
        var borderFact = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        borderFact.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        borderFact.SetValue(WpfBorder.BorderBrushProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderBrushProperty));
        borderFact.SetValue(WpfBorder.BorderThicknessProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderThicknessProperty));
        borderFact.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(6));
        borderFact.SetValue(WpfBorder.PaddingProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.PaddingProperty));
        var scrollFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ScrollViewer));
        scrollFact.Name = "PART_ContentHost";
        scrollFact.SetValue(System.Windows.Controls.ScrollViewer.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        borderFact.AppendChild(scrollFact);
        template.VisualTree = borderFact;
        return template;
    }

    private void ApplyRoundedTextBoxTemplate(System.Windows.Controls.TextBox textBox)
    {
        var template = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.TextBox));
        var borderFact = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        borderFact.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        borderFact.SetValue(WpfBorder.BorderBrushProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderBrushProperty));
        borderFact.SetValue(WpfBorder.BorderThicknessProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderThicknessProperty));
        borderFact.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(6));
        borderFact.SetValue(WpfBorder.PaddingProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.PaddingProperty));
        var scrollFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ScrollViewer));
        scrollFact.Name = "PART_ContentHost";
        scrollFact.SetValue(System.Windows.Controls.ScrollViewer.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        borderFact.AppendChild(scrollFact);
        template.VisualTree = borderFact;
        textBox.Template = template;
    }


    // ── Block Diagram drag state ──────────────────────────────────────────
    private WpfBorder? _draggedBlock;
    private System.Windows.Point _dragOffset;

    private void LoadBlockDiagramSection(System.Windows.Controls.Panel panel)
    {
        var card = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderThickness = new WpfThickness(0),
            Padding = new WpfThickness(16, 12, 16, 24)
        };
        card.Effect = new System.Windows.Media.Effects.DropShadowEffect
        { ShadowDepth = 0, BlurRadius = 15, Opacity = 0.08, Color = System.Windows.Media.Color.FromRgb(0, 0, 0) };

        var outer = new System.Windows.Controls.StackPanel();
        card.Child = outer;

        // Title
        outer.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Block Diagram",
            FontSize = 22, FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = System.Windows.Media.Brushes.Black,
            Margin = new WpfThickness(0, 0, 0, 4)
        });
        outer.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Auto-generated from selected components. Drag blocks to reposition. Edit text by clicking a block label.",
            FontSize = 12, Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100,116,139)),
            Margin = new WpfThickness(0, 0, 0, 16), TextWrapping = System.Windows.TextWrapping.Wrap
        });

        // Canvas
        var canvas = new System.Windows.Controls.Canvas
        {
            Width = 750, Height = 460,
            Background = System.Windows.Media.Brushes.White,
            ClipToBounds = true
        };
        var canvasBorder = new WpfBorder
        {
            BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226,232,240)),
            BorderThickness = new WpfThickness(1), CornerRadius = new CornerRadius(6),
            Child = canvas, HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        };
        outer.Children.Add(canvasBorder);

        // ── Read selected components ────────────────────────────────────────
        bool hasRackmount    = IsChecked("Rackmount");
        bool hasBB           = IsChecked("B.B");
        bool hasIS           = IsChecked("I.S");
        bool hasBacklight    = IsChecked("Backlight");
        bool hasSourceStage  = IsChecked("Source Stage");
        bool hasPowerMeter   = IsChecked("Power Meter");
        bool hasEnergyMeter  = IsChecked("Energy Meter");
        bool hasOptTable     = IsChecked("Optical Table");
        bool hasCTE          = IsChecked("CTE");
        bool hasDevCenter    = IsChecked("Device Center");
        bool hasXYStage      = IsChecked("XY Stage");
        bool hasMonitorArm   = _rackmountMonitorArmCheckBox?.IsChecked == true;
        bool hasLOS          = IsChecked("LOS Laser");
        bool hasQTH          = IsChecked("QTH Lamp");
        bool hasNewport      = IsChecked("NewPort Stage");
        bool isXYManual      = _xyStageManualCheckBox?.IsChecked == true;
        bool hasLOSTarget    = IsChecked("LOS alignment target");
        bool hasFocus        = IsChecked("Focus Stage");
        bool hasGimbal       = IsChecked("Gimbal");

        // System type label
        string sysType = string.Join(" ", new[]
        {
            cmbSystemType.SelectedItem?.ToString() ?? "",
            cmbSystemVariant.SelectedItem?.ToString() ?? "",
            cmbSystemAperture.SelectedItem?.ToString() ?? ""
        }.Where(s => !string.IsNullOrEmpty(s)));
        if (string.IsNullOrEmpty(sysType)) sysType = "METS";

        // B.B label
        string bbLabel = "B.B";
        if (hasBB)
        {
            var t = _bbTypeComboBox?.SelectedItem?.ToString() ?? "";
            var s2 = _bbSizeComboBox?.SelectedItem?.ToString() ?? "";
            if (!string.IsNullOrEmpty(t) || !string.IsNullOrEmpty(s2))
                bbLabel = $"B.B {s2} {t}".Trim();
        }
        // I.S label
        string isLabel = "I.S";
        if (hasIS)
        {
            var ap = _isExitApertureComboBox?.SelectedItem?.ToString() ?? "";
            if (!string.IsNullOrEmpty(ap)) isLabel = $"I.S {ap}";
        }
        // Backlight label
        string blLabel = "Backlight";
        if (hasBacklight)
        {
            var bt = _backlightTypeComboBox?.SelectedItem?.ToString() ?? "";
            if (!string.IsNullOrEmpty(bt)) blLabel = $"Backlight\n{bt}";
        }

        // Rackmount height label
        string rackLabel = "Rackmount";
        if (hasRackmount)
        {
            var rh = _rackmountHeightTextBox?.Text.Trim() ?? "";
            if (!string.IsNullOrEmpty(rh)) rackLabel = $"Rackmount {rh}U";
        }

        // Optical Table label
        string otLabel = "Optical Table";
        if (hasOptTable)
        {
            var w = _opticalTableWidthTextBox?.Text.Trim() ?? "";
            var l = _opticalTableLengthTextBox?.Text.Trim() ?? "";
            var h = _opticalTableHeightTextBox?.Text.Trim() ?? "";
            var dims = string.Join(" x ", new[]{w,l,h}.Where(x=>!string.IsNullOrEmpty(x)));
            if (!string.IsNullOrEmpty(dims)) otLabel = $"Optical Table\n{dims} [mm]";
        }

        // Helper: make a draggable block
        WpfBorder MakeBlock(string text, double x, double y, double w, double h,
            bool bold = false, bool filled = false, double fontSize = 11,
            System.Windows.Media.Brush? textColor = null)
        {
            var tb = new System.Windows.Controls.TextBox
            {
                Text = text,
                FontSize = fontSize,
                FontWeight = bold ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal,
                Foreground = textColor ?? System.Windows.Media.Brushes.Black,
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new WpfThickness(0),
                TextWrapping = System.Windows.TextWrapping.Wrap,
                TextAlignment = System.Windows.TextAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                AcceptsReturn = true,
                IsReadOnly = false,
                Padding = new WpfThickness(2)
            };
            var border = new WpfBorder
            {
                Width = w, Height = h,
                BorderBrush = System.Windows.Media.Brushes.Black,
                BorderThickness = new WpfThickness(1.5),
                CornerRadius = new CornerRadius(0),
                Background = filled
                    ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220,235,255))
                    : System.Windows.Media.Brushes.White,
                Child = tb,
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            System.Windows.Controls.Canvas.SetLeft(border, x);
            System.Windows.Controls.Canvas.SetTop(border, y);

            border.MouseLeftButtonDown += (s, e) =>
            {
                _draggedBlock = border;
                _dragOffset = e.GetPosition(canvas);
                _dragOffset.X -= System.Windows.Controls.Canvas.GetLeft(border);
                _dragOffset.Y -= System.Windows.Controls.Canvas.GetTop(border);
                border.CaptureMouse();
                e.Handled = true;
            };
            border.MouseMove += (s, e) =>
            {
                if (_draggedBlock == border && border.IsMouseCaptured)
                {
                    var pos = e.GetPosition(canvas);
                    double nx = pos.X - _dragOffset.X;
                    double ny = pos.Y - _dragOffset.Y;
                    nx = Math.Max(0, Math.Min(canvas.Width - border.Width, nx));
                    ny = Math.Max(0, Math.Min(canvas.Height - border.Height, ny));
                    System.Windows.Controls.Canvas.SetLeft(border, nx);
                    System.Windows.Controls.Canvas.SetTop(border, ny);
                }
            };
            border.MouseLeftButtonUp += (s, e) =>
            {
                if (_draggedBlock == border) { _draggedBlock = null; border.ReleaseMouseCapture(); }
            };
            canvas.Children.Add(border);
            return border;
        }

        // Helper: plain text label (no border)
        void MakeLabel(string text, double x, double y, double fontSize = 10, bool bold = false)
        {
            var tb = new System.Windows.Controls.TextBlock
            {
                Text = text, FontSize = fontSize,
                FontWeight = bold ? System.Windows.FontWeights.Bold : System.Windows.FontWeights.Normal,
                Foreground = System.Windows.Media.Brushes.Black,
                TextAlignment = System.Windows.TextAlignment.Center,
                TextWrapping = System.Windows.TextWrapping.Wrap
            };
            System.Windows.Controls.Canvas.SetLeft(tb, x);
            System.Windows.Controls.Canvas.SetTop(tb, y);
            canvas.Children.Add(tb);
        }

        // Helper: draw arrow line on canvas
        void DrawArrow(double x1, double y1, double x2, double y2, string midLabel = "")
        {
            var green = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 160, 0));
            var line = new System.Windows.Shapes.Line
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Stroke = green, StrokeThickness = 1.5
            };
            canvas.Children.Add(line);

            // Arrow heads (both ends = double-headed)
            double dx = x2 - x1, dy = y2 - y1;
            double len = Math.Sqrt(dx*dx + dy*dy);
            if (len < 1) return;
            double ux = dx/len, uy = dy/len;
            double aLen = 10, aW = 5;

            System.Windows.Shapes.Polygon MakeHead(double ax, double ay, double dirX, double dirY)
            {
                return new System.Windows.Shapes.Polygon
                {
                    Fill = green,
                    Points = new System.Windows.Media.PointCollection
                    {
                        new System.Windows.Point(ax, ay),
                        new System.Windows.Point(ax - dirX*aLen + dirY*aW, ay - dirY*aLen - dirX*aW),
                        new System.Windows.Point(ax - dirX*aLen - dirY*aW, ay - dirY*aLen + dirX*aW)
                    }
                };
            }
            canvas.Children.Add(MakeHead(x2, y2, ux, uy));
            canvas.Children.Add(MakeHead(x1, y1, -ux, -uy));

            if (!string.IsNullOrEmpty(midLabel))
            {
                double mx = (x1+x2)/2, my = (y1+y2)/2;
                var lbl = new System.Windows.Controls.TextBlock
                {
                    Text = midLabel, FontSize = 10, FontWeight = System.Windows.FontWeights.Bold,
                    Foreground = System.Windows.Media.Brushes.Black,
                    Background = System.Windows.Media.Brushes.White
                };
                System.Windows.Controls.Canvas.SetLeft(lbl, mx - 18);
                System.Windows.Controls.Canvas.SetTop(lbl, my - 22);  // above the arrow
                canvas.Children.Add(lbl);
            }
        }

        // ── LAYOUT CONSTANTS ───────────────────────────────────────────────
        const double canvasW = 750;
        double topPad = 60;   // vertical padding from top

        // Fixed sizes
        double srcStageW = 90,  srcStageH = 330;
        double narrowW   = 22,  narrowH   = 120;
        double collBodyW = 320, collBodyH = 310;
        double rackW     = 100, rackH     = 210;
        double rackGap   = 60;  // space between rackmount right edge and source stage left edge
        double meterW    = 68,  meterH    = 52;
        double meterGap  = 16;  // space between collimator right edge and meter

        // ── Pre-compute required height and scale down to fit canvas ─────
        {
            double neededH = topPad + collBodyH + 20; // base: topPad + collimator + bottom pad
            if (hasOptTable) neededH += 22 + 50 + 12; // gap + table height + pad

            const double availH = 455;
            if (neededH > availH)
            {
                double s = availH / neededH;
                srcStageW *= s; srcStageH *= s;
                narrowW   *= s; narrowH   *= s;
                collBodyW *= s; collBodyH *= s;
                rackW     *= s; rackH     *= s;
                rackGap   *= s;
                meterW    *= s; meterH    *= s;
                meterGap  *= s;
                topPad    *= s;
            }
        }

        // ── Compute total diagram width to centre it ─────────────────────
        double diagramW = srcStageW + narrowW + collBodyW;
        if (hasRackmount)   diagramW += rackW + rackGap;
        if (hasPowerMeter || hasEnergyMeter) diagramW += meterW + meterGap;

        double offsetX = Math.Max(10, (canvasW - diagramW) / 2.0);

        // ── Absolute X positions derived from offsetX ────────────────────
        double rackX, srcStageX, narrowX, collBodyX, meterX;
        if (hasRackmount)
        {
            rackX      = offsetX;
            srcStageX  = rackX + rackW + rackGap;
        }
        else
        {
            rackX     = offsetX;          // unused but kept for arrow calc
            srcStageX = offsetX;
        }
        narrowX    = srcStageX + srcStageW;
        collBodyX  = narrowX   + narrowW;
        meterX     = collBodyX + collBodyW + meterGap;

        // ── Vertical positions ───────────────────────────────────────────
        double collBodyY = topPad;
        double srcStageY = collBodyY - 15;
        double narrowY   = collBodyY + 30;
        double rackY     = collBodyY + 5;
        double meterY    = collBodyY;

        // Arrow between rack and source stage
        double arrowStartX = rackX + rackW + 6;
        double arrowEndX   = srcStageX - 8;
        double arrowY      = rackY + rackH / 2;

        // ── Monitor on Arm (far left of rack) ──────────────────────────────
        if (hasMonitorArm)
        {
            MakeBlock("Monitor\non Arm", rackX - 65, rackY + 45, 58, 42, fontSize: 9);
            var ln = new System.Windows.Shapes.Line
            { X1 = rackX - 7, Y1 = rackY + 66, X2 = rackX, Y2 = rackY + 66,
              Stroke = System.Windows.Media.Brushes.Black, StrokeThickness = 1.5 };
            canvas.Children.Add(ln);
        }

        // ── RACKMOUNT block ─────────────────────────────────────────────────
        if (hasRackmount)
        {
            // "Rackmount Size" label above the block — same style as "Motorized Source Stage"
            var rackAboveLabel = new System.Windows.Controls.TextBlock
            {
                Text = rackLabel,
                FontSize = 11, FontWeight = System.Windows.FontWeights.Bold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 38, 38)),
                TextAlignment = System.Windows.TextAlignment.Left,
                TextWrapping = System.Windows.TextWrapping.Wrap
            };
            System.Windows.Controls.Canvas.SetLeft(rackAboveLabel, rackX);
            System.Windows.Controls.Canvas.SetTop(rackAboveLabel, rackY - 28);
            canvas.Children.Add(rackAboveLabel);

            // Build slot list: PDU → controllers → Monitor → Computer (no CTE)
            var rackSlots = new List<(string text, bool isRed)>();
            rackSlots.Add(("PDU", true));
            if (hasBB)  rackSlots.Add(("SR800N", true));
            if (hasIS)  rackSlots.Add(("SR300N", true));
            rackSlots.Add(("Monitor", true));
            rackSlots.Add(("Computer", true));

            // Height logic:
            // - Minimum height is collBodyH (matches collimator height)
            // - Required height = slots * (minSlotH + gap) + padding
            // - If required > min, use required; otherwise use min and stretch slots evenly
            double minSlotH = 38, slotGap2 = 4;
            double padding2 = 8;
            double requiredH = rackSlots.Count * (minSlotH + slotGap2) + padding2;
            double rackDynH = Math.Max(collBodyH, requiredH);

            // Distribute available space evenly across slots
            double totalSlotSpace = rackDynH - padding2 - (rackSlots.Count - 1) * slotGap2;
            double slotH2 = totalSlotSpace / rackSlots.Count;

            var rackBorder = new WpfBorder
            {
                Width = rackW, Height = rackDynH,
                BorderBrush = System.Windows.Media.Brushes.Black,
                BorderThickness = new WpfThickness(2.5),
                Background = System.Windows.Media.Brushes.White,
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            var rackStack = new System.Windows.Controls.StackPanel
            {
                Margin = new WpfThickness(0, 4, 0, 4)
            };
            rackBorder.Child = rackStack;

            foreach (var (slotText, slotRed) in rackSlots)
            {
                var ib = new WpfBorder
                {
                    BorderBrush = System.Windows.Media.Brushes.Black,
                    BorderThickness = new WpfThickness(1.5),
                    Margin = new WpfThickness(5, 2, 5, 2),
                    Height = slotH2
                };
                ib.Child = new System.Windows.Controls.TextBox
                {
                    Text = slotText,
                    FontSize = 10,
                    FontWeight = System.Windows.FontWeights.Bold,
                    Foreground = slotRed
                        ? System.Windows.Media.Brushes.Red
                        : System.Windows.Media.Brushes.Black,
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderThickness = new WpfThickness(0),
                    TextAlignment = System.Windows.TextAlignment.Center,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    VerticalContentAlignment = System.Windows.VerticalAlignment.Center,
                    IsReadOnly = false
                };
                rackStack.Children.Add(ib);
            }

            System.Windows.Controls.Canvas.SetLeft(rackBorder, rackX);
            System.Windows.Controls.Canvas.SetTop(rackBorder, rackY);
            rackBorder.MouseLeftButtonDown += (s, e) =>
            {
                _draggedBlock = rackBorder; _dragOffset = e.GetPosition(canvas);
                _dragOffset.X -= System.Windows.Controls.Canvas.GetLeft(rackBorder);
                _dragOffset.Y -= System.Windows.Controls.Canvas.GetTop(rackBorder);
                rackBorder.CaptureMouse(); e.Handled = true;
            };
            rackBorder.MouseMove += (s, e) =>
            {
                if (_draggedBlock == rackBorder && rackBorder.IsMouseCaptured)
                {
                    var pos = e.GetPosition(canvas);
                    System.Windows.Controls.Canvas.SetLeft(rackBorder, pos.X - _dragOffset.X);
                    System.Windows.Controls.Canvas.SetTop(rackBorder, pos.Y - _dragOffset.Y);
                }
            };
            rackBorder.MouseLeftButtonUp += (s, e) =>
            { if (_draggedBlock == rackBorder) { _draggedBlock = null; rackBorder.ReleaseMouseCapture(); } };
            canvas.Children.Add(rackBorder);

            // Wheels drawn OUTSIDE the border, centered below the rack block
            double wheelY = rackY + rackDynH + 3;
            double wheelSize = 16;
            double wheelSpacing = rackW * 0.28;
            foreach (var wx in new[] { rackX + wheelSpacing, rackX + rackW - wheelSpacing - wheelSize })
            {
                var wheel = new System.Windows.Shapes.Ellipse
                {
                    Width = wheelSize, Height = wheelSize,
                    Stroke = System.Windows.Media.Brushes.Black,
                    StrokeThickness = 2,
                    Fill = System.Windows.Media.Brushes.White
                };
                System.Windows.Controls.Canvas.SetLeft(wheel, wx);
                System.Windows.Controls.Canvas.SetTop(wheel, wheelY);
                canvas.Children.Add(wheel);
            }

            DrawArrow(arrowStartX, arrowY, arrowEndX, arrowY, "4 Meter");
        }

        // ── COLLIMATOR DEFAULT LAYOUT ───────────────────────────────────────
        // Source Stage block (left standalone)
        {
            var srcBorder = new WpfBorder
            {
                Width = srcStageW, Height = srcStageH,
                BorderBrush = System.Windows.Media.Brushes.Black,
                BorderThickness = new WpfThickness(2.5),
                Background = System.Windows.Media.Brushes.White,
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            var srcTb = new System.Windows.Controls.TextBox
            {
                Text = "", FontSize = 11, BorderThickness = new WpfThickness(0),
                Background = System.Windows.Media.Brushes.Transparent,
                TextAlignment = System.Windows.TextAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Center,
                TextWrapping = System.Windows.TextWrapping.Wrap, IsReadOnly = false,
                Padding = new WpfThickness(2)
            };
            srcBorder.Child = srcTb;
            System.Windows.Controls.Canvas.SetLeft(srcBorder, srcStageX);
            System.Windows.Controls.Canvas.SetTop(srcBorder, srcStageY);
            srcBorder.MouseLeftButtonDown += (s, e) =>
            {
                _draggedBlock = srcBorder; _dragOffset = e.GetPosition(canvas);
                _dragOffset.X -= System.Windows.Controls.Canvas.GetLeft(srcBorder);
                _dragOffset.Y -= System.Windows.Controls.Canvas.GetTop(srcBorder);
                srcBorder.CaptureMouse(); e.Handled = true;
            };
            srcBorder.MouseMove += (s, e) =>
            {
                if (_draggedBlock == srcBorder && srcBorder.IsMouseCaptured)
                {
                    var pos = e.GetPosition(canvas);
                    System.Windows.Controls.Canvas.SetLeft(srcBorder, pos.X - _dragOffset.X);
                    System.Windows.Controls.Canvas.SetTop(srcBorder, pos.Y - _dragOffset.Y);
                }
            };
            srcBorder.MouseLeftButtonUp += (s, e) =>
            { if (_draggedBlock == srcBorder) { _draggedBlock = null; srcBorder.ReleaseMouseCapture(); } };
            canvas.Children.Add(srcBorder);

            // "Motorized Source Stage" label above — red, bold, matching the reference
            var srcLabel = new System.Windows.Controls.TextBlock
            {
                Text = "Motorized\nSource Stage",
                FontSize = 10, FontWeight = System.Windows.FontWeights.Bold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(200, 0, 0)),
                TextAlignment = System.Windows.TextAlignment.Center,
                TextWrapping = System.Windows.TextWrapping.Wrap
            };
            System.Windows.Controls.Canvas.SetLeft(srcLabel, srcStageX);
            System.Windows.Controls.Canvas.SetTop(srcLabel, srcStageY - 44);
            canvas.Children.Add(srcLabel);
        }

        // Narrow vertical connector bar (between source stage and collimator body)
        {
            var narrowBorder = new WpfBorder
            {
                Width = narrowW, Height = narrowH,
                BorderBrush = System.Windows.Media.Brushes.Black,
                BorderThickness = new WpfThickness(2),
                Background = System.Windows.Media.Brushes.White,
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            System.Windows.Controls.Canvas.SetLeft(narrowBorder, narrowX);
            System.Windows.Controls.Canvas.SetTop(narrowBorder, narrowY);
            narrowBorder.MouseLeftButtonDown += (s, e) =>
            {
                _draggedBlock = narrowBorder; _dragOffset = e.GetPosition(canvas);
                _dragOffset.X -= System.Windows.Controls.Canvas.GetLeft(narrowBorder);
                _dragOffset.Y -= System.Windows.Controls.Canvas.GetTop(narrowBorder);
                narrowBorder.CaptureMouse(); e.Handled = true;
            };
            narrowBorder.MouseMove += (s, e) =>
            {
                if (_draggedBlock == narrowBorder && narrowBorder.IsMouseCaptured)
                {
                    var pos = e.GetPosition(canvas);
                    System.Windows.Controls.Canvas.SetLeft(narrowBorder, pos.X - _dragOffset.X);
                    System.Windows.Controls.Canvas.SetTop(narrowBorder, pos.Y - _dragOffset.Y);
                }
            };
            narrowBorder.MouseLeftButtonUp += (s, e) =>
            { if (_draggedBlock == narrowBorder) { _draggedBlock = null; narrowBorder.ReleaseMouseCapture(); } };
            canvas.Children.Add(narrowBorder);

            // "Targets Wheel" label — right of the narrow bar, above it, never overlapping
            double twCenterX = narrowX + narrowW / 2.0;  // arrow stem X (center of narrow bar)
            var twLabel = new System.Windows.Controls.TextBlock
            {
                Text = "Targets Wheel",
                FontSize = 10, FontWeight = System.Windows.FontWeights.Normal,
                Foreground = System.Windows.Media.Brushes.Black,
                TextAlignment = System.Windows.TextAlignment.Left,
                TextWrapping = System.Windows.TextWrapping.NoWrap
            };
            // Place label starting from the narrow bar center, extending right — fully clear of Source Stage
            System.Windows.Controls.Canvas.SetLeft(twLabel, twCenterX + 4);
            System.Windows.Controls.Canvas.SetTop(twLabel,  narrowY - 58);
            canvas.Children.Add(twLabel);

            // Green down arrow: stem starts high above narrow bar, tip stops above block top
            double arrowStemX   = twCenterX;
            double arrowLineTop = narrowY - 44;   // well above the block
            double arrowLineBot = narrowY - 14;   // tip stops clear of block top
            var twLine = new System.Windows.Shapes.Line
            {
                X1 = arrowStemX, Y1 = arrowLineTop, X2 = arrowStemX, Y2 = arrowLineBot,
                Stroke = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 160, 0)),
                StrokeThickness = 2.5
            };
            canvas.Children.Add(twLine);
            // Arrowhead pointing down
            canvas.Children.Add(new System.Windows.Shapes.Polygon
            {
                Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 160, 0)),
                Points = new System.Windows.Media.PointCollection
                {
                    new System.Windows.Point(arrowStemX,     arrowLineBot + 8),
                    new System.Windows.Point(arrowStemX - 6, arrowLineBot),
                    new System.Windows.Point(arrowStemX + 6, arrowLineBot)
                }
            });
        }

        // Large rounded Collimator body (main block)
        {
            var collBorder = new WpfBorder
            {
                Width = collBodyW, Height = collBodyH,
                BorderBrush = System.Windows.Media.Brushes.Black,
                BorderThickness = new WpfThickness(3),
                CornerRadius = new CornerRadius(30),
                Background = System.Windows.Media.Brushes.White,
                Cursor = System.Windows.Input.Cursors.SizeAll
            };
            var collTb = new System.Windows.Controls.TextBox
            {
                Text = sysType, FontSize = 32, FontWeight = System.Windows.FontWeights.Bold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(210, 0, 0)),
                BorderThickness = new WpfThickness(0),
                Background = System.Windows.Media.Brushes.Transparent,
                TextAlignment = System.Windows.TextAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Top,
                VerticalContentAlignment = System.Windows.VerticalAlignment.Top,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                TextWrapping = System.Windows.TextWrapping.Wrap, IsReadOnly = false,
                Padding = new WpfThickness(12, 18, 12, 0),
                Width = collBodyW - 6, Height = collBodyH - 6
            };
            collBorder.Child = collTb;
            System.Windows.Controls.Canvas.SetLeft(collBorder, collBodyX);
            System.Windows.Controls.Canvas.SetTop(collBorder, collBodyY);
            collBorder.MouseLeftButtonDown += (s, e) =>
            {
                _draggedBlock = collBorder; _dragOffset = e.GetPosition(canvas);
                _dragOffset.X -= System.Windows.Controls.Canvas.GetLeft(collBorder);
                _dragOffset.Y -= System.Windows.Controls.Canvas.GetTop(collBorder);
                collBorder.CaptureMouse(); e.Handled = true;
            };
            collBorder.MouseMove += (s, e) =>
            {
                if (_draggedBlock == collBorder && collBorder.IsMouseCaptured)
                {
                    var pos = e.GetPosition(canvas);
                    System.Windows.Controls.Canvas.SetLeft(collBorder, pos.X - _dragOffset.X);
                    System.Windows.Controls.Canvas.SetTop(collBorder, pos.Y - _dragOffset.Y);
                }
            };
            collBorder.MouseLeftButtonUp += (s, e) =>
            { if (_draggedBlock == collBorder) { _draggedBlock = null; collBorder.ReleaseMouseCapture(); } };
            canvas.Children.Add(collBorder);
        }

        // ── B.B block inside Source Stage + SR800N Controller (no Rackmount) ──
        if (hasBB)
        {
            // B.B sub-block at the top of the Source Stage — red bold text
            MakeBlock(bbLabel, srcStageX + 4, srcStageY + 6, srcStageW - 8, 44, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);

            // SR800N Controller — bottom-aligned with collimator bottom, left of source stage
            if (!hasRackmount)
            {
                double ctrlW = 90, ctrlH = 50;
                double ctrlX = srcStageX - ctrlW - 20;
                double ctrlY = collBodyY + collBodyH - ctrlH;   // aligned to collimator bottom
                MakeBlock("SR800N\nController", ctrlX, ctrlY, ctrlW, ctrlH, bold: true, fontSize: 10, textColor: System.Windows.Media.Brushes.Red);
            }
        }

        // ── I.S block inside Source Stage + SR300N Controller (no Rackmount) ──
        if (hasIS)
        {
            double isOffsetY = hasBB ? 56 : 6;
            MakeBlock(isLabel, srcStageX + 4, srcStageY + isOffsetY, srcStageW - 8, 44, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);

            if (!hasRackmount)
            {
                double ctrlW = 90, ctrlH = 50;
                double ctrlX = srcStageX - ctrlW - 20;
                double ctrlY = collBodyY + collBodyH - ctrlH - (hasBB ? 58 : 0);  // stack above SR800N if both
                MakeBlock("SR300N\nController", ctrlX, ctrlY, ctrlW, ctrlH, bold: true, fontSize: 10, textColor: System.Windows.Media.Brushes.Red);
            }
        }

        // ── Backlight, LOS Laser, QTH Lamp — sub-blocks in Source Stage, no controllers ──
        {
            // Count how many slots are already used by B.B and I.S
            int slotIndex = (hasBB ? 1 : 0) + (hasIS ? 1 : 0);
            const double slotH = 44, slotGap = 6;

            if (hasBacklight)
            {
                double bly = srcStageY + 6 + slotIndex * (slotH + slotGap);
                MakeBlock(blLabel, srcStageX + 4, bly, srcStageW - 8, slotH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
                slotIndex++;
            }
            if (hasLOS)
            {
                double losy = srcStageY + 6 + slotIndex * (slotH + slotGap);
                MakeBlock("LOS Laser", srcStageX + 4, losy, srcStageW - 8, slotH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
                slotIndex++;
            }
            if (hasQTH)
            {
                double qthy = srcStageY + 6 + slotIndex * (slotH + slotGap);
                MakeBlock("QTH Lamp", srcStageX + 4, qthy, srcStageW - 8, slotH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
            }
        }

        // ── LOS + XY Stage — both INSIDE the collimator body, stacked at bottom ──
        {
            double losW = 80, losH = 46;
            double xyBarH2 = 20, xyVH2 = 50, xyBarW2 = collBodyW * 0.55, xyVW2 = 20;
            double xyBarX2 = collBodyX + (collBodyW - xyBarW2) / 2;
            double xyVX2   = xyBarX2 + (xyBarW2 - xyVW2) / 2;

            // LOS always anchors 12px above the collimator bottom
            double losX = collBodyX + (collBodyW - losW) / 2;
            double losY = collBodyY + collBodyH - losH - 12;

            // XY cross sits above LOS (if both), otherwise also 12px above collimator bottom
            double xyBarY2, xyVY2;
            if (hasLOSTarget && hasXYStage)
            {
                xyBarY2 = losY - xyVH2 - 10 + (xyVH2 - xyBarH2) / 2;
                xyVY2   = losY - xyVH2 - 10;
            }
            else
            {
                xyBarY2 = collBodyY + collBodyH - 12 - xyBarH2;
                xyVY2   = xyBarY2 - (xyVH2 - xyBarH2) / 2;
            }

            if (hasXYStage)
            {
                string xyLabel = isXYManual ? "Manual XY Stage" : "Motorized XY Stage";
                var xyH2 = new WpfBorder { Width = xyBarW2, Height = xyBarH2,
                    BorderBrush = System.Windows.Media.Brushes.Black, BorderThickness = new WpfThickness(1.5),
                    Background = System.Windows.Media.Brushes.White };
                System.Windows.Controls.Canvas.SetLeft(xyH2, xyBarX2);
                System.Windows.Controls.Canvas.SetTop(xyH2, xyBarY2);
                canvas.Children.Add(xyH2);

                var xyV2 = new WpfBorder { Width = xyVW2, Height = xyVH2,
                    BorderBrush = System.Windows.Media.Brushes.Black, BorderThickness = new WpfThickness(1.5),
                    Background = System.Windows.Media.Brushes.White };
                System.Windows.Controls.Canvas.SetLeft(xyV2, xyVX2);
                System.Windows.Controls.Canvas.SetTop(xyV2, xyVY2);
                canvas.Children.Add(xyV2);

                // Label above the horizontal bar, right portion
                var xyLbl2 = new System.Windows.Controls.TextBlock
                {
                    Text = xyLabel, FontSize = 9, FontWeight = System.Windows.FontWeights.Bold,
                    Foreground = System.Windows.Media.Brushes.Red,
                    TextAlignment = System.Windows.TextAlignment.Left,
                    TextWrapping = System.Windows.TextWrapping.Wrap
                };
                double lblX = xyVX2 + xyVW2 + 4;
                double lblW = (xyBarX2 + xyBarW2) - lblX - 2;
                xyLbl2.Width = Math.Max(lblW, 70); // ensure enough room for "Motorized XY Stage"
                System.Windows.Controls.Canvas.SetLeft(xyLbl2, lblX);
                System.Windows.Controls.Canvas.SetTop(xyLbl2, xyBarY2 - 38);
                canvas.Children.Add(xyLbl2);
            }

            if (hasLOSTarget)
            {
                MakeBlock("LOS CCD\nLOS LED", losX, losY, losW, losH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
            }
        }

        // ── Focus Stage block — left side INSIDE collimator, vertically centered ──
        if (hasFocus)
        {
            double fsW = collBodyW * 0.30, fsH = 36;
            double fsX = collBodyX + 8;
            double fsY = collBodyY + collBodyH * 0.38;
            MakeBlock("Focus Stage", fsX, fsY, fsW, fsH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
        }

        // ── Newport Stage + Gimbal blocks — small, outside collimator, bottom-right ──
        {
            double npW = 90, npH = 46;
            double npX = collBodyX + collBodyW + 16;
            int rightSlot = 0;
            if (hasGimbal)
            {
                double gY = collBodyY + collBodyH - npH * (rightSlot + 1) - 8 * rightSlot;
                MakeBlock("Gimbal", npX, gY, npW, npH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
                rightSlot++;
            }
            if (hasNewport)
            {
                double nY = collBodyY + collBodyH - npH * (rightSlot + 1) - 8 * rightSlot;
                MakeBlock("Newport Stage", npX, nY, npW, npH, bold: true, fontSize: 9, textColor: System.Windows.Media.Brushes.Red);
            }
        }

        // ── OPTICAL TABLE block ─────────────────────────────────────────────
        if (hasOptTable)
        {
            double otX = srcStageX - 4;
            double otBaseY = collBodyY + collBodyH + 22;
            MakeBlock(otLabel, otX, otBaseY, collBodyX + collBodyW - otX + 4, 50, bold: true, filled: false, fontSize: 10, textColor: System.Windows.Media.Brushes.Red);
        }

        // ── POWER METER / ENERGY METER ──────────────────────────────────────
        double curMeterY = meterY;
        if (hasPowerMeter)
        {
            MakeBlock("Power\nMeter", meterX, curMeterY, meterW, meterH, fontSize: 9);
            curMeterY += meterH + 8;
        }
        if (hasEnergyMeter)
        {
            MakeBlock("Energy\nMeter", meterX, curMeterY, meterW, meterH, fontSize: 9);
        }

        panel.Children.Add(card);
    }

    private bool IsChecked(string key) =>
        _configCheckBoxes.TryGetValue(key, out var cb) && cb.IsChecked == true;

    private void LoadActionsSection(System.Windows.Controls.Panel panel)
    {
        // Create card container with compact styling
        var card = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(1),
            Padding = new WpfThickness(16),
            Margin = new WpfThickness(0, 0, 0, 12)
        };

        var stackPanel = new System.Windows.Controls.StackPanel();
        card.Child = stackPanel;

        // Title - compact
        var title = new System.Windows.Controls.TextBlock
        {
            Text = "Questions/Notes",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        stackPanel.Children.Add(title);

        // PM Questions/Notes — outer border wrapping header + content (matches Targets style)
        var pmOuterBorder = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(229, 231, 235)),
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(8),
            Margin = new WpfThickness(0, 0, 0, 16)
        };
        var pmInnerStack = new System.Windows.Controls.StackPanel();
        pmOuterBorder.Child = pmInnerStack;
        pmInnerStack.Children.Add(MakeSectionHeader("📋", "PM Questions/Notes", "+ Add Question/Note", AddQuestion_Click));

        // Questions ItemsControl
        questionsItemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0)
        };

        // Numbered row: "1. [TextBox] [B] [🗑]"
        var dataTemplate = new System.Windows.DataTemplate();
        var rowFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.DockPanel));
        rowFactory.SetValue(System.Windows.Controls.DockPanel.MarginProperty, new WpfThickness(0, 0, 0, 6));

        var numFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBlock));
        numFactory.SetBinding(System.Windows.Controls.TextBlock.TextProperty,
            new System.Windows.Data.Binding("Number") { StringFormat = "{0}." });
        numFactory.SetValue(System.Windows.Controls.TextBlock.FontSizeProperty, 13.0);
        numFactory.SetValue(System.Windows.Controls.TextBlock.FontWeightProperty, System.Windows.FontWeights.SemiBold);
        numFactory.SetValue(System.Windows.Controls.TextBlock.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        numFactory.SetValue(System.Windows.Controls.TextBlock.MarginProperty, new WpfThickness(0, 0, 4, 0));
        numFactory.SetValue(System.Windows.Controls.TextBlock.WidthProperty, 20.0);
        numFactory.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Left);
        rowFactory.AppendChild(numFactory);

        var delFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.Button));
        delFactory.SetValue(System.Windows.Controls.Button.ContentProperty, "🗑");
        delFactory.SetValue(System.Windows.Controls.Button.FontSizeProperty, 13.0);
        delFactory.SetValue(System.Windows.Controls.Button.ForegroundProperty, System.Windows.Media.Brushes.Black);
        delFactory.SetValue(System.Windows.Controls.Button.BackgroundProperty, System.Windows.Media.Brushes.Transparent);
        delFactory.SetValue(System.Windows.Controls.Button.BorderThicknessProperty, new WpfThickness(0));
        delFactory.SetValue(System.Windows.Controls.Button.CursorProperty, System.Windows.Input.Cursors.Hand);
        delFactory.SetValue(System.Windows.Controls.Button.PaddingProperty, new WpfThickness(4, 0, 0, 0));
        delFactory.SetValue(System.Windows.Controls.Button.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        delFactory.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Right);
        delFactory.SetBinding(System.Windows.Controls.Button.TagProperty, new System.Windows.Data.Binding());
        delFactory.AddHandler(System.Windows.Controls.Button.ClickEvent, new RoutedEventHandler(DeletePmQuestion_Click));
        rowFactory.AppendChild(delFactory);

        var boldFactory = MakeBoldButtonFactory(ToggleBold_Click);
        rowFactory.AppendChild(boldFactory);

        var tbFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        tbFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty,
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
        tbFactory.SetBinding(System.Windows.Controls.TextBox.FontWeightProperty,
            new System.Windows.Data.Binding("IsBold") { Converter = new BoldFontWeightConverter() });
        tbFactory.SetValue(System.Windows.Controls.TextBox.HeightProperty, 28.0);
        tbFactory.SetValue(System.Windows.Controls.TextBox.PaddingProperty, new WpfThickness(8, 2, 8, 2));
        tbFactory.SetValue(System.Windows.Controls.TextBox.VerticalContentAlignmentProperty, System.Windows.VerticalAlignment.Center);
        tbFactory.SetValue(System.Windows.Controls.TextBox.FontSizeProperty, 13.0);
        tbFactory.SetValue(System.Windows.Controls.TextBox.BorderBrushProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)));
        tbFactory.SetValue(System.Windows.Controls.TextBox.BorderThicknessProperty, new WpfThickness(1));
        tbFactory.SetValue(System.Windows.Controls.Control.TemplateProperty, CreateRoundedTextBoxControlTemplate());
        rowFactory.AppendChild(tbFactory);

        dataTemplate.VisualTree = rowFactory;
        questionsItemsControl.ItemTemplate = dataTemplate;
        questionsItemsControl.ItemsSource = _questions;

        var pmContentBorder = new WpfBorder
        {
            Padding = new WpfThickness(16, 6, 16, 2)
        };
        var pmContentStack = new System.Windows.Controls.StackPanel();
        pmContentBorder.Child = pmContentStack;
        pmContentStack.Children.Add(questionsItemsControl);
        pmInnerStack.Children.Add(pmContentBorder);
        stackPanel.Children.Add(pmOuterBorder);

        // Marketing Questions/Notes — outer border wrapping header + content (matches Targets style)
        var mktOuterBorder = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(229, 231, 235)),
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(8),
            Margin = new WpfThickness(0, 0, 0, 16)
        };
        var mktInnerStack = new System.Windows.Controls.StackPanel();
        mktOuterBorder.Child = mktInnerStack;
        mktInnerStack.Children.Add(MakeSectionHeader("📣", "Marketing Questions/Notes", "+ Add Question/Note", AddMarketingQuestion_Click));

        // Marketing Questions ItemsControl
        marketingQuestionsItemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0)
        };

        // Numbered row: "1. [TextBox] [B] [🗑]"
        var marketingDataTemplate = new System.Windows.DataTemplate();
        var mRowFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.DockPanel));
        mRowFactory.SetValue(System.Windows.Controls.DockPanel.MarginProperty, new WpfThickness(0, 0, 0, 6));

        var mNumFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBlock));
        mNumFactory.SetBinding(System.Windows.Controls.TextBlock.TextProperty,
            new System.Windows.Data.Binding("Number") { StringFormat = "{0}." });
        mNumFactory.SetValue(System.Windows.Controls.TextBlock.FontSizeProperty, 13.0);
        mNumFactory.SetValue(System.Windows.Controls.TextBlock.FontWeightProperty, System.Windows.FontWeights.SemiBold);
        mNumFactory.SetValue(System.Windows.Controls.TextBlock.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        mNumFactory.SetValue(System.Windows.Controls.TextBlock.MarginProperty, new WpfThickness(0, 0, 4, 0));
        mNumFactory.SetValue(System.Windows.Controls.TextBlock.WidthProperty, 20.0);
        mNumFactory.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Left);
        mRowFactory.AppendChild(mNumFactory);

        var mDelFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.Button));
        mDelFactory.SetValue(System.Windows.Controls.Button.ContentProperty, "🗑");
        mDelFactory.SetValue(System.Windows.Controls.Button.FontSizeProperty, 13.0);
        mDelFactory.SetValue(System.Windows.Controls.Button.ForegroundProperty, System.Windows.Media.Brushes.Black);
        mDelFactory.SetValue(System.Windows.Controls.Button.BackgroundProperty, System.Windows.Media.Brushes.Transparent);
        mDelFactory.SetValue(System.Windows.Controls.Button.BorderThicknessProperty, new WpfThickness(0));
        mDelFactory.SetValue(System.Windows.Controls.Button.CursorProperty, System.Windows.Input.Cursors.Hand);
        mDelFactory.SetValue(System.Windows.Controls.Button.PaddingProperty, new WpfThickness(4, 0, 0, 0));
        mDelFactory.SetValue(System.Windows.Controls.Button.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        mDelFactory.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Right);
        mDelFactory.SetBinding(System.Windows.Controls.Button.TagProperty, new System.Windows.Data.Binding());
        mDelFactory.AddHandler(System.Windows.Controls.Button.ClickEvent, new RoutedEventHandler(DeleteMarketingQuestion_Click));
        mRowFactory.AppendChild(mDelFactory);

        var mBoldFactory = MakeBoldButtonFactory(ToggleBold_Click);
        mRowFactory.AppendChild(mBoldFactory);

        var mTbFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        mTbFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty,
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
        mTbFactory.SetBinding(System.Windows.Controls.TextBox.FontWeightProperty,
            new System.Windows.Data.Binding("IsBold") { Converter = new BoldFontWeightConverter() });
        mTbFactory.SetValue(System.Windows.Controls.TextBox.HeightProperty, 28.0);
        mTbFactory.SetValue(System.Windows.Controls.TextBox.PaddingProperty, new WpfThickness(8, 2, 8, 2));
        mTbFactory.SetValue(System.Windows.Controls.TextBox.VerticalContentAlignmentProperty, System.Windows.VerticalAlignment.Center);
        mTbFactory.SetValue(System.Windows.Controls.TextBox.FontSizeProperty, 13.0);
        mTbFactory.SetValue(System.Windows.Controls.TextBox.BorderBrushProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)));
        mTbFactory.SetValue(System.Windows.Controls.TextBox.BorderThicknessProperty, new WpfThickness(1));
        mTbFactory.SetValue(System.Windows.Controls.Control.TemplateProperty, CreateRoundedTextBoxControlTemplate());
        mRowFactory.AppendChild(mTbFactory);

        marketingDataTemplate.VisualTree = mRowFactory;
        marketingQuestionsItemsControl.ItemTemplate = marketingDataTemplate;
        marketingQuestionsItemsControl.ItemsSource = _marketingQuestions;

        var mktContentBorder = new WpfBorder
        {
            Padding = new WpfThickness(16, 6, 16, 2)
        };
        var mktContentStack = new System.Windows.Controls.StackPanel();
        mktContentBorder.Child = mktContentStack;
        mktContentStack.Children.Add(marketingQuestionsItemsControl);
        mktInnerStack.Children.Add(mktContentBorder);
        stackPanel.Children.Add(mktOuterBorder);

        panel.Children.Add(card);
    }

    private void LoadNotesSection(System.Windows.Controls.Panel panel)
    {
        var card = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            CornerRadius = new CornerRadius(12),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(1),
            Padding = new WpfThickness(16),
            Margin = new WpfThickness(0, 0, 0, 12)
        };

        var stackPanel = new System.Windows.Controls.StackPanel();
        card.Child = stackPanel;

        // Title
        stackPanel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Notes",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        });

        // ── Marketing Notes ──────────────────────────────────────
        stackPanel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "Marketing Notes:",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(51, 65, 85)),
            Margin = new WpfThickness(0, 0, 0, 8)
        });

        AddNotesSubSection(stackPanel, _marketingNotes, ref marketingNotesItemsControl,
            AddMarketingNote_Click, DeleteMarketingNote_Click);

        // ── PM Notes ─────────────────────────────────────────────
        stackPanel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "PM Notes:",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(51, 65, 85)),
            Margin = new WpfThickness(0, 0, 0, 8)
        });

        AddNotesSubSection(stackPanel, _pmNotes, ref pmNotesItemsControl,
            AddPmNote_Click, DeletePmNote_Click);

        panel.Children.Add(card);
    }

    // ── Creates a Targets-style section header bar (light-blue bg, title left, button right) ──
    private WpfBorder MakeSectionHeader(string icon, string title, string buttonLabel, RoutedEventHandler buttonClick)
    {
        var container = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(239, 246, 255)),
            Padding = new WpfThickness(16, 1, 16, 1),
            CornerRadius = new CornerRadius(8, 8, 0, 0)
        };

        var grid = new System.Windows.Controls.Grid();
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });

        var titlePanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        titlePanel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = icon,
            FontSize = 16,
            Margin = new WpfThickness(0, 0, 8, 0),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        });
        titlePanel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = title,
            FontSize = 15,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)),
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        });
        System.Windows.Controls.Grid.SetColumn(titlePanel, 0);
        grid.Children.Add(titlePanel);

        var btn = new System.Windows.Controls.Button
        {
            Content = buttonLabel,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
            Padding = new WpfThickness(6, 2, 6, 2),
            FontSize = 7,
            FontWeight = System.Windows.FontWeights.Medium,
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)),
            Foreground = System.Windows.Media.Brushes.White,
            BorderThickness = new WpfThickness(0),
            Cursor = System.Windows.Input.Cursors.Hand,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
        var btnTemplate = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var btnBorder = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        btnBorder.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        btnBorder.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(6));
        btnBorder.SetValue(WpfBorder.PaddingProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.PaddingProperty));
        var btnContent = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
        btnContent.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        btnContent.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        btnBorder.AppendChild(btnContent);
        btnTemplate.VisualTree = btnBorder;
        btn.Template = btnTemplate;
        btn.Click += buttonClick;
        System.Windows.Controls.Grid.SetColumn(btn, 1);
        grid.Children.Add(btn);

        container.Child = grid;
        return container;
    }

    // ── Creates a reusable Bold-toggle button factory ─────────────────────────
    private System.Windows.FrameworkElementFactory MakeBoldButtonFactory(RoutedEventHandler clickHandler)
    {
        var factory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.Button));
        factory.SetValue(System.Windows.Controls.Button.ContentProperty, "B");
        factory.SetValue(System.Windows.Controls.Button.FontSizeProperty, 13.0);
        factory.SetValue(System.Windows.Controls.Button.FontWeightProperty, System.Windows.FontWeights.Bold);
        factory.SetValue(System.Windows.Controls.Button.ForegroundProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)));
        factory.SetValue(System.Windows.Controls.Button.WidthProperty, 28.0);
        factory.SetValue(System.Windows.Controls.Button.HeightProperty, 28.0);
        factory.SetValue(System.Windows.Controls.Button.CursorProperty, System.Windows.Input.Cursors.Hand);
        factory.SetValue(System.Windows.Controls.Button.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        factory.SetValue(System.Windows.Controls.Button.MarginProperty, new WpfThickness(0, 0, 0, 0));
        factory.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Right);
        factory.SetBinding(System.Windows.Controls.Button.TagProperty, new System.Windows.Data.Binding());

        // Custom ControlTemplate so WPF chrome doesn't swallow the "B" text or colours
        var ct = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var borderFact = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        borderFact.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(3));
        borderFact.SetValue(WpfBorder.BorderThicknessProperty, new WpfThickness(1));
        borderFact.SetValue(WpfBorder.PaddingProperty, new WpfThickness(0));
        borderFact.SetValue(WpfBorder.BorderBrushProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(203, 213, 225)));

        // Background: light-blue when IsBold=true, transparent otherwise
        // Bind via TemplatedParent (the Button) → its Tag (the QuestionItem) → IsBold
        borderFact.SetBinding(WpfBorder.BackgroundProperty,
            new System.Windows.Data.Binding("Tag.IsBold")
            {
                RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent),
                Converter = new BoldBackgroundConverter()
            });

        var cp = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBlock));
        cp.SetValue(System.Windows.Controls.TextBlock.TextProperty, "B");
        cp.SetValue(System.Windows.Controls.TextBlock.FontSizeProperty, 16.0);
        cp.SetValue(System.Windows.Controls.TextBlock.FontWeightProperty, System.Windows.FontWeights.Black);
        cp.SetValue(System.Windows.Controls.TextBlock.ForegroundProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 41, 59)));
        cp.SetValue(System.Windows.Controls.TextBlock.HorizontalAlignmentProperty,
            System.Windows.HorizontalAlignment.Center);
        cp.SetValue(System.Windows.Controls.TextBlock.VerticalAlignmentProperty,
            System.Windows.VerticalAlignment.Center);
        cp.SetValue(System.Windows.Controls.TextBlock.LineHeightProperty, 16.0);
        cp.SetValue(System.Windows.Controls.TextBlock.LineStackingStrategyProperty,
            System.Windows.LineStackingStrategy.BlockLineHeight);
        borderFact.AppendChild(cp);
        ct.VisualTree = borderFact;
        factory.SetValue(System.Windows.Controls.Control.TemplateProperty, ct);

        factory.AddHandler(System.Windows.Controls.Button.ClickEvent, clickHandler);
        return factory;
    }

    private void AddNotesSubSection(
        System.Windows.Controls.StackPanel parent,
        ObservableCollection<QuestionItem> collection,
        ref System.Windows.Controls.ItemsControl itemsControl,
        RoutedEventHandler addHandler,
        RoutedEventHandler deleteHandler)
    {
        // Add Note button
        var addBtn = new System.Windows.Controls.Button
        {
            Content = "+ Add Note",
            Margin = new WpfThickness(0, 0, 0, 8),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
            FontSize = 12
        };
        ApplyModernButtonStyle(addBtn);
        addBtn.Padding = new WpfThickness(10, 2, 10, 2);
        addBtn.Click += addHandler;
        parent.Children.Add(addBtn);

        // ItemsControl with numbered rows
        itemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0, 0, 0, 16)
        };

        var template = new System.Windows.DataTemplate();
        var rowFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.DockPanel));
        rowFact.SetValue(System.Windows.Controls.DockPanel.MarginProperty, new WpfThickness(0, 0, 0, 6));

        var numFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBlock));
        numFact.SetBinding(System.Windows.Controls.TextBlock.TextProperty,
            new System.Windows.Data.Binding("Number") { StringFormat = "{0}." });
        numFact.SetValue(System.Windows.Controls.TextBlock.FontSizeProperty, 13.0);
        numFact.SetValue(System.Windows.Controls.TextBlock.FontWeightProperty, System.Windows.FontWeights.SemiBold);
        numFact.SetValue(System.Windows.Controls.TextBlock.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        numFact.SetValue(System.Windows.Controls.TextBlock.MarginProperty, new WpfThickness(0, 0, 4, 0));
        numFact.SetValue(System.Windows.Controls.TextBlock.WidthProperty, 20.0);
        numFact.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Left);
        rowFact.AppendChild(numFact);

        var delFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.Button));
        delFact.SetValue(System.Windows.Controls.Button.ContentProperty, "🗑");
        delFact.SetValue(System.Windows.Controls.Button.FontSizeProperty, 13.0);
        delFact.SetValue(System.Windows.Controls.Button.ForegroundProperty, System.Windows.Media.Brushes.Black);
        delFact.SetValue(System.Windows.Controls.Button.BackgroundProperty, System.Windows.Media.Brushes.Transparent);
        delFact.SetValue(System.Windows.Controls.Button.BorderThicknessProperty, new WpfThickness(0));
        delFact.SetValue(System.Windows.Controls.Button.CursorProperty, System.Windows.Input.Cursors.Hand);
        delFact.SetValue(System.Windows.Controls.Button.PaddingProperty, new WpfThickness(4, 0, 0, 0));
        delFact.SetValue(System.Windows.Controls.Button.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        delFact.SetValue(System.Windows.Controls.DockPanel.DockProperty, System.Windows.Controls.Dock.Right);
        delFact.SetBinding(System.Windows.Controls.Button.TagProperty, new System.Windows.Data.Binding());
        delFact.AddHandler(System.Windows.Controls.Button.ClickEvent, deleteHandler);
        rowFact.AppendChild(delFact);

        var boldFact = MakeBoldButtonFactory(ToggleBold_Click);
        rowFact.AppendChild(boldFact);

        var tbFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        tbFact.SetBinding(System.Windows.Controls.TextBox.TextProperty,
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
        tbFact.SetBinding(System.Windows.Controls.TextBox.FontWeightProperty,
            new System.Windows.Data.Binding("IsBold") { Converter = new BoldFontWeightConverter() });
        tbFact.SetValue(System.Windows.Controls.TextBox.HeightProperty, 28.0);
        tbFact.SetValue(System.Windows.Controls.TextBox.PaddingProperty, new WpfThickness(8, 2, 8, 2));
        tbFact.SetValue(System.Windows.Controls.TextBox.VerticalContentAlignmentProperty, System.Windows.VerticalAlignment.Center);
        tbFact.SetValue(System.Windows.Controls.TextBox.FontSizeProperty, 13.0);
        tbFact.SetValue(System.Windows.Controls.TextBox.BorderBrushProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)));
        tbFact.SetValue(System.Windows.Controls.TextBox.BorderThicknessProperty, new WpfThickness(1));
        tbFact.SetValue(System.Windows.Controls.Control.TemplateProperty, CreateRoundedTextBoxControlTemplate());
        rowFact.AppendChild(tbFact);

        template.VisualTree = rowFact;
        itemsControl.ItemTemplate = template;
        itemsControl.ItemsSource = collection;

        parent.Children.Add(itemsControl);
    }

    // ── Bold toggle handler (shared by all sections) ──────────────────────
    private void ToggleBold_Click(object s, RoutedEventArgs e)
    {
        if (s is System.Windows.Controls.Button b && b.Tag is QuestionItem item)
            item.IsBold = !item.IsBold;
    }

    // ── Marketing Notes handlers ──────────────────────────────────
    private void AddMarketingNote_Click(object s, RoutedEventArgs e) =>
        _marketingNotes.Add(new QuestionItem { Text = "", Number = _marketingNotes.Count + 1 });
    private void DeleteMarketingNote_Click(object s, RoutedEventArgs e)
    {
        if (s is System.Windows.Controls.Button b && b.Tag is QuestionItem item)
        { _marketingNotes.Remove(item); RenumberCollection(_marketingNotes); }
    }

    // ── PM Notes handlers ────────────────────────────────────────
    private void AddPmNote_Click(object s, RoutedEventArgs e) =>
        _pmNotes.Add(new QuestionItem { Text = "", Number = _pmNotes.Count + 1 });
    private void DeletePmNote_Click(object s, RoutedEventArgs e)
    {
        if (s is System.Windows.Controls.Button b && b.Tag is QuestionItem item)
        { _pmNotes.Remove(item); RenumberCollection(_pmNotes); }
    }

    private void InitializeConfigItems()
    {
        // Initialize Radiation Source items
        foreach (var item in _radiationSourceItems)
        {
            var checkBox = new WpfCheckBox
            {
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = item,
                    FontSize = 14,
                    FontWeight = System.Windows.FontWeights.Normal,
                    Margin = new WpfThickness(0)
                },
                FontSize = 14,
                FontWeight = System.Windows.FontWeights.Normal
            };
            _configCheckBoxes[item] = checkBox;
        }

        // Initialize System Components items
        foreach (var item in _systemComponentsItems)
        {
            var checkBox = new WpfCheckBox
            {
                Content = new System.Windows.Controls.TextBlock
                {
                    Text = item,
                    FontSize = 14,
                    FontWeight = System.Windows.FontWeights.Normal,
                    Margin = new WpfThickness(0)
                },
                FontSize = 14,
                FontWeight = System.Windows.FontWeights.Normal
            };
            _configCheckBoxes[item] = checkBox;
        }

        // Wire all checkboxes to update the Block Diagram button state
        foreach (var cb in _configCheckBoxes.Values)
        {
            cb.Checked   += (s, e) => UpdateBlockDiagramButtonState();
            cb.Unchecked += (s, e) => UpdateBlockDiagramButtonState();
        }
    }

    // Creates a WPF Path that draws a speech bubble icon.
    private static System.Windows.Shapes.Path MakeSpeechBubbleIcon(System.Windows.Media.Brush fill)
    {
        // A simple rounded speech bubble with a small tail at the bottom-left.
        var geometry = System.Windows.Media.Geometry.Parse(
            "M 1,1 Q 1,0 2,0 L 12,0 Q 13,0 13,1 L 13,8 Q 13,9 12,9 L 5,9 L 2,12 L 2,9 L 2,9 Q 1,9 1,8 Z");
        return new System.Windows.Shapes.Path
        {
            Data = geometry,
            Fill = fill,
            Width = 14,
            Height = 13,
            Stretch = System.Windows.Media.Stretch.Uniform,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };
    }

    // Wraps a checkbox in a horizontal StackPanel together with a speech-bubble note button.
    private System.Windows.Controls.StackPanel WrapCheckBoxWithNoteButton(WpfCheckBox checkBox, string itemKey)
    {
        var wrapper = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            VerticalAlignment = System.Windows.VerticalAlignment.Center
        };

        wrapper.Children.Add(checkBox);

        bool hasNote = _componentNotes.TryGetValue(itemKey, out var existingNote) && !string.IsNullOrWhiteSpace(existingNote);
        var iconFill = hasNote
            ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(37, 99, 235))
            : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(160, 174, 192));

        var noteBtn = new System.Windows.Controls.Button
        {
            Content = MakeSpeechBubbleIcon(iconFill),
            Width = 22,
            Height = 22,
            Background = System.Windows.Media.Brushes.Transparent,
            BorderThickness = new WpfThickness(0),
            Cursor = System.Windows.Input.Cursors.Hand,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Padding = new WpfThickness(0),
            Margin = new WpfThickness(4, 0, 0, 0),
            Tag = itemKey,
            ToolTip = hasNote ? existingNote : "Add note"
        };

        var noteBtnTemplate = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
        var noteBorderFact = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        noteBorderFact.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
        var noteContentFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
        noteContentFact.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
        noteContentFact.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
        noteBorderFact.AppendChild(noteContentFact);
        noteBtnTemplate.VisualTree = noteBorderFact;
        noteBtn.Template = noteBtnTemplate;

        noteBtn.Click += NoteButton_Click;
        wrapper.Children.Add(noteBtn);

        return wrapper;
    }

    private void NoteButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button btn) return;
        if (btn.Tag is not string itemKey) return;

        // Build popup
        var popup = new System.Windows.Controls.Primitives.Popup
        {
            PlacementTarget = btn,
            Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom,
            StaysOpen = false,
            AllowsTransparency = true
        };

        // Outer border — light gray border, white background, rounded, with shadow (matches screenshot)
        var popupBorder = new WpfBorder
        {
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(180, 190, 205)),
            BorderThickness = new WpfThickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new WpfThickness(0),
            Width = 260,
            Effect = new System.Windows.Media.Effects.DropShadowEffect
            {
                ShadowDepth = 2,
                BlurRadius = 8,
                Opacity = 0.18,
                Color = System.Windows.Media.Color.FromRgb(0, 0, 0)
            }
        };

        var outerStack = new System.Windows.Controls.StackPanel();

        // Title bar — light blue-gray background with label (matches screenshot header)
        var titleBar = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(220, 230, 242)),
            Padding = new WpfThickness(10, 4, 10, 4),
            CornerRadius = new CornerRadius(5, 5, 0, 0)
        };
        var titleText = new System.Windows.Controls.TextBlock
        {
            Text = $"Notes for {itemKey}",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59))
        };
        titleBar.Child = titleText;
        outerStack.Children.Add(titleBar);

        // Body area
        var bodyStack = new System.Windows.Controls.StackPanel
        {
            Margin = new WpfThickness(10, 8, 10, 6)
        };

        _componentNotes.TryGetValue(itemKey, out var currentNote);
        var noteTextBox = new System.Windows.Controls.TextBox
        {
            Text = currentNote ?? "",
            FontSize = 13,
            FontWeight = System.Windows.FontWeights.Normal,
            Height = 70,
            Padding = new WpfThickness(6, 4, 6, 4),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(180, 190, 205)),
            BorderThickness = new WpfThickness(1),
            AcceptsReturn = true,
            TextWrapping = System.Windows.TextWrapping.Wrap,
            VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
            Margin = new WpfThickness(0, 0, 0, 10)
        };
        ApplyRoundedTextBoxTemplate(noteTextBox);
        bodyStack.Children.Add(noteTextBox);

        // Buttons row — Save & Cancel aligned to the right
        var buttonsRow = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Right
        };

        // Helper: create a flat dialog button matching the screenshot style
        System.Windows.Controls.Button MakeDialogButton(string label)
        {
            var b = new System.Windows.Controls.Button
            {
                Content = label,
                FontSize = 12,
                FontWeight = System.Windows.FontWeights.Normal,
                Padding = new WpfThickness(14, 4, 14, 4),
                Background = System.Windows.Media.Brushes.White,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(30, 41, 59)),
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(180, 190, 205)),
                BorderThickness = new WpfThickness(1),
                Cursor = System.Windows.Input.Cursors.Hand
            };
            var tmpl = new System.Windows.Controls.ControlTemplate(typeof(System.Windows.Controls.Button));
            var bf = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
            bf.SetValue(WpfBorder.BackgroundProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BackgroundProperty));
            bf.SetValue(WpfBorder.BorderBrushProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderBrushProperty));
            bf.SetValue(WpfBorder.BorderThicknessProperty, new System.Windows.TemplateBindingExtension(System.Windows.Controls.Control.BorderThicknessProperty));
            bf.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(4));
            var cf = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ContentPresenter));
            cf.SetValue(System.Windows.Controls.ContentPresenter.HorizontalAlignmentProperty, System.Windows.HorizontalAlignment.Center);
            cf.SetValue(System.Windows.Controls.ContentPresenter.VerticalAlignmentProperty, System.Windows.VerticalAlignment.Center);
            cf.SetValue(System.Windows.Controls.ContentPresenter.MarginProperty, new WpfThickness(0, -6, 0, 0));
            bf.AppendChild(cf);
            tmpl.VisualTree = bf;
            b.Template = tmpl;
            b.VerticalContentAlignment = System.Windows.VerticalAlignment.Center;
            return b;
        }

        var saveBtn   = MakeDialogButton("Save");
        saveBtn.Width = 62;
        saveBtn.Height = 26;
        saveBtn.Padding = new WpfThickness(0, 0, 0, 0);
        var cancelBtn = MakeDialogButton("Cancel");
        cancelBtn.Width = 62;
        cancelBtn.Height = 26;
        cancelBtn.Padding = new WpfThickness(0, 0, 0, 0);
        cancelBtn.Margin = new WpfThickness(6, 0, 0, 0);

        saveBtn.Click += (s2, e2) =>
        {
            _componentNotes[itemKey] = noteTextBox.Text;
            bool saved = !string.IsNullOrWhiteSpace(noteTextBox.Text);
            btn.ToolTip = saved ? noteTextBox.Text : "Add note";
            if (btn.Content is System.Windows.Shapes.Path iconPath)
            {
                iconPath.Fill = saved
                    ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(37, 99, 235))
                    : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(160, 174, 192));
            }
            popup.IsOpen = false;
        };

        cancelBtn.Click += (s2, e2) => { popup.IsOpen = false; };

        buttonsRow.Children.Add(saveBtn);
        buttonsRow.Children.Add(cancelBtn);
        bodyStack.Children.Add(buttonsRow);

        outerStack.Children.Add(bodyStack);
        popupBorder.Child = outerStack;
        popup.Child = popupBorder;
        popup.IsOpen = true;

        noteTextBox.Focus();
        noteTextBox.SelectAll();
    }

    private void InitializeTargets()
    {
        // Targets will be initialized with 3 empty items in LoadTargetsSection
        // Keep collection empty here
    }

    private void InitializeQuestions()
    {
        _questions.Add(new QuestionItem { Text = "", Number = 1 });
    }

    private void InitializeMarketingQuestions()
    {
        _marketingQuestions.Add(new QuestionItem { Text = "", Number = 1 });
    }

    private void InitializeNotes()
    {
        // No default notes — PM/Marketing notes start empty; only questions seed one row
    }

    private void RenumberCollection(ObservableCollection<QuestionItem> collection)
    {
        for (int i = 0; i < collection.Count; i++)
            collection[i].Number = i + 1;
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
            _questions.Add(new QuestionItem { Text = "", Number = _questions.Count + 1 });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding question: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void AddMarketingQuestion_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            _marketingQuestions.Add(new QuestionItem { Text = "", Number = _marketingQuestions.Count + 1 });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error adding marketing question: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void DeletePmQuestion_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is QuestionItem item)
            {
                _questions.Remove(item);
                RenumberCollection(_questions);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error deleting question: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void DeleteMarketingQuestion_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is QuestionItem item)
            {
                _marketingQuestions.Remove(item);
                RenumberCollection(_marketingQuestions);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error deleting marketing question: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void btnGenerateDocument_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtOrderNumber.Text))
        {
            MessageBox.Show("Please enter an Order Number.", "Required Field",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            LoadSection(0);
            txtOrderNumber.Focus();
            return;
        }

        SaveFileDialog saveDialog = new SaveFileDialog
        {
            Filter = "Word Document (*.doc)|*.doc",
            DefaultExt = "doc",
            FileName = $"Kickoff {txtOrderNumber.Text}.doc"
        };

        if (saveDialog.ShowDialog() == true)
        {
            try
            {
                GenerateFromTemplate(saveDialog.FileName);
                _lastExportedPath = saveDialog.FileName;
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

    private void GenerateFromTemplate(string outputPath)
    {
        string templatePath = System.IO.Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Marketing_Meeting_Summary_Template.xml");

        if (!System.IO.File.Exists(templatePath))
            throw new System.IO.FileNotFoundException($"Template not found at:\n{templatePath}");

        XDocument doc;
        using (var fs = new System.IO.FileStream(templatePath, System.IO.FileMode.Open,
                   System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
            doc = XDocument.Load(fs, LoadOptions.PreserveWhitespace);

        XNamespace w   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        var wordPart = doc.Root!
            .Elements(pkg + "part")
            .FirstOrDefault(p => ((string?)p.Attribute(pkg + "name")) == "/word/document.xml")
            ?? throw new InvalidOperationException("Cannot find /word/document.xml in template.");

        var body = wordPart
            .Element(pkg + "xmlData")!
            .Element(w + "document")!
            .Element(w + "body")!;

        // Collect all searchable containers: body + all header/footer parts
        var allContainers = new List<XElement> { body };
        foreach (var part in doc.Root!.Elements(pkg + "part"))
        {
            var partName = (string?)part.Attribute(pkg + "name") ?? "";
            if (partName.StartsWith("/word/header") || partName.StartsWith("/word/footer"))
            {
                var hdrBody = part.Element(pkg + "xmlData")?.Elements().FirstOrDefault();
                if (hdrBody != null) allContainers.Add(hdrBody);
            }
        }

        // ── Local helpers ─────────────────────────────────────────────────────
        bool IsRedRun(XElement r)
        {
            var color = r.Element(w + "rPr")?.Element(w + "color");
            return ((string?)color?.Attribute(w + "val"))
                   ?.Equals("EE0000", StringComparison.OrdinalIgnoreCase) == true;
        }

        string RunText(XElement r) =>
            string.Concat(r.Elements(w + "t").Select(t => (string?)t ?? ""));

        void MakeBlackRun(XElement r, string value)
        {
            var rPr = r.Element(w + "rPr");
            if (rPr != null)
            {
                rPr.Elements(w + "b").Remove();
                rPr.Elements(w + "bCs").Remove();
                rPr.Elements(w + "color").Remove();
            }
            r.Elements(w + "t").Remove();
            var t = new XElement(w + "t", value);
            if (value.StartsWith(" ") || value.EndsWith(" "))
                t.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            r.Add(t);
        }

        void ReplacePlaceholder(string placeholder, string value)
        {
            foreach (var container in allContainers)
            {
                var run = container.Descendants(w + "r")
                    .FirstOrDefault(r => IsRedRun(r) && RunText(r) == placeholder);
                if (run != null) { MakeBlackRun(run, value); break; }
            }
        }

        void InjectAfterLabelRun(string labelFragment, string value)
        {
            foreach (var container in allContainers)
            {
                var run = container.Descendants(w + "r")
                    .FirstOrDefault(r => RunText(r).Contains(labelFragment));
                if (run == null) continue;
                var extra = new XElement(w + "r",
                    new XElement(w + "rPr",
                        new XElement(w + "sz",   new XAttribute(w + "val", "22")),
                        new XElement(w + "szCs", new XAttribute(w + "val", "22"))),
                    new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"), value));
                run.AddAfterSelf(extra);
                break;
            }
        }

        XElement MakeBulletPara(string text)
        {
            return new XElement(w + "p",
                new XElement(w + "pPr",
                    new XElement(w + "spacing", new XAttribute(w + "after", "80"))),
                new XElement(w + "r",
                    new XElement(w + "rPr",
                        new XElement(w + "sz",   new XAttribute(w + "val", "22")),
                        new XElement(w + "szCs", new XAttribute(w + "val", "22"))),
                    new XElement(w + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        text)));
        }

        void ReplaceMarkerTableRows(string markerFragment, IEnumerable<(string text, bool bold)> items_in)
        {
            // Find the paragraph inside the table cell that contains the marker
            var markerPara = body.Descendants(w + "p").FirstOrDefault(p =>
                string.Concat(p.Descendants(w + "t").Select(t => (string?)t ?? ""))
                      .Contains(markerFragment));
            if (markerPara == null) return;

            // Walk up to find the table row that owns this paragraph
            var templateRow = markerPara.Ancestors(w + "tr").FirstOrDefault();
            if (templateRow == null) return;

            var items = items_in.Where(s => !string.IsNullOrWhiteSpace(s.text)).ToList();

            if (items.Count == 0)
            {
                // Clear the marker text, leave the row with "1" and empty content
                markerPara.Elements(w + "r").Remove();
                return;
            }

            // Fill the first (template) row with item #1
            var cells = templateRow.Elements(w + "tc").ToList();
            // Update number cell to "1"
            if (cells.Count > 0)
            {
                var numPara = cells[0].Element(w + "p");
                numPara?.Elements(w + "r").Remove();
                numPara?.Add(MakeDataRun("1"));
            }
            // Replace marker paragraph in content cell with item text
            markerPara.Elements(w + "r").Remove();
            markerPara.Add(items[0].bold ? MakeBoldDataRun(items[0].text) : MakeDataRun(items[0].text));

            // Clone the template row for each additional item
            var insertAfter = templateRow;
            for (int i = 1; i < items.Count; i++)
            {
                var clonedRow = new XElement(templateRow);
                var clonedCells = clonedRow.Elements(w + "tc").ToList();

                // Set row number
                if (clonedCells.Count > 0)
                {
                    var numPara = clonedCells[0].Element(w + "p");
                    numPara?.Elements(w + "r").Remove();
                    numPara?.Add(MakeDataRun((i + 1).ToString()));
                }
                // Set content
                if (clonedCells.Count > 1)
                {
                    var contentPara = clonedCells[1].Element(w + "p");
                    contentPara?.Elements(w + "r").Remove();
                    contentPara?.Add(items[i].bold ? MakeBoldDataRun(items[i].text) : MakeDataRun(items[i].text));
                }

                insertAfter.AddAfterSelf(clonedRow);
                insertAfter = clonedRow;
            }
        }

        // ── Simple field replacements ─────────────────────────────────────────
        ReplacePlaceholder("Order Number",   txtOrderNumber.Text);
        ReplacePlaceholder("Customer Name",  txtCustomerName.Text);
        ReplacePlaceholder("Territory",      txtFinalCustomer.Text);
        ReplacePlaceholder("Paka Number",    txtPakaNumber.Text);
        var systemTypeStr = string.Join(" ", new[]
        {
            cmbSystemType.SelectedItem?.ToString() ?? "",
            cmbSystemVariant.SelectedItem?.ToString() ?? "",
            cmbSystemAperture.SelectedItem?.ToString() ?? ""
        }.Where(s => !string.IsNullOrEmpty(s)));
        ReplacePlaceholder("Project Type",   systemTypeStr);
        ReplacePlaceholder("Project Hours",  txtProjectHours.Text);
        ReplacePlaceholder("Selling Price",  txtSellingPrice.Text);
        ReplacePlaceholder("Material Cost",  txtMaterialCost.Text);
        // The Monday # (CRM) placeholder is split across two red runs in the template.
        // Find consecutive red runs whose combined text equals "Monday # (CRM):" and replace them.
        {
            string mondayValue = txtMondayCRM.Text;
            bool mondayReplaced = false;
            foreach (var container in allContainers)
            {
                var redRuns = container.Descendants(w + "r").Where(r => IsRedRun(r)).ToList();
                for (int ri = 0; ri < redRuns.Count - 1; ri++)
                {
                    string t1 = RunText(redRuns[ri]);
                    string t2 = RunText(redRuns[ri + 1]);
                    string combined = t1 + t2;
                    if (combined == "Monday # (CRM):" || combined == "Monday # (CRM)")
                    {
                        MakeBlackRun(redRuns[ri], mondayValue);
                        redRuns[ri + 1].Remove();
                        mondayReplaced = true;
                        break;
                    }
                }
                if (mondayReplaced) break;
            }
            if (!mondayReplaced)
                ReplacePlaceholder("Monday # (CRM)", mondayValue);
        }
        ReplacePlaceholder("Penalties",      txtPenalties.Text);
        ReplacePlaceholder("Reference Order", txtReferenceOrder.Text);
        InjectAfterLabelRun("D.O Rated:", chkDORated.IsChecked == true ? "Yes" : "No");

        ReplacePlaceholder("Delivery Date",
            dpDeliveryDate.SelectedDate.HasValue
                ? dpDeliveryDate.SelectedDate.Value.ToString("dd.MM.yyyy") : "");
        ReplacePlaceholder("Design Due Date",
            dpDesignDueDate.SelectedDate.HasValue
                ? dpDesignDueDate.SelectedDate.Value.ToString("dd.MM.yyyy") : "");

        // ── Targets table cell ────────────────────────────────────────────────
        var targetItems = _targets
            .Where(t => !string.IsNullOrWhiteSpace(t.Qty) || !string.IsNullOrWhiteSpace(t.Type))
            .ToList();

        var targetNameRun = body.Descendants(w + "r")
            .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Target Name");
        if (targetNameRun != null)
        {
            var templateRow = targetNameRun.Ancestors(w + "tr").First();
            // Save a pristine copy BEFORE modifying the first row, so clones get clean placeholders
            var pristineRow = new XElement(templateRow);
            var sizeUnitRun = templateRow.Descendants(w + "r")
                .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Size");
            var unitRun = templateRow.Descendants(w + "r")
                .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Size Unit");

            if (targetItems.Count == 0)
            {
                MakeBlackRun(targetNameRun, "");
                if (sizeUnitRun != null) MakeBlackRun(sizeUnitRun, "");
                if (unitRun != null) MakeBlackRun(unitRun, "");
            }
            else
            {
                MakeBlackRun(targetNameRun, targetItems[0].Type);
                // Center-align the Target cell paragraph
                var targetPara = targetNameRun.Parent;
                if (targetPara != null)
                {
                    var pPr = targetPara.Element(w + "pPr") ?? new XElement(w + "pPr");
                    if (targetPara.Element(w + "pPr") == null) targetPara.AddFirst(pPr);
                    pPr.Elements(w + "jc").Remove();
                    pPr.Add(new XElement(w + "jc", new XAttribute(w + "val", "center")));
                }
                if (sizeUnitRun != null)
                    MakeBlackRun(sizeUnitRun, targetItems[0].Qty);
                if (unitRun != null)
                    MakeBlackRun(unitRun, targetItems[0].Details);

                var insertAfter = templateRow;
                for (int i = 1; i < targetItems.Count; i++)
                {
                    // Clone from the pristine (unmodified) template row
                    var clonedRow = new XElement(pristineRow);
                    var cloneNameRun = clonedRow.Descendants(w + "r")
                        .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Target Name");
                    var cloneSizeRun = clonedRow.Descendants(w + "r")
                        .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Size");
                    var cloneUnitRun = clonedRow.Descendants(w + "r")
                        .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Size Unit");
                    if (cloneNameRun != null)
                    {
                        MakeBlackRun(cloneNameRun, targetItems[i].Type);
                        var cloneTargetPara = cloneNameRun.Parent;
                        if (cloneTargetPara != null)
                        {
                            var pPr = cloneTargetPara.Element(w + "pPr") ?? new XElement(w + "pPr");
                            if (cloneTargetPara.Element(w + "pPr") == null) cloneTargetPara.AddFirst(pPr);
                            pPr.Elements(w + "jc").Remove();
                            pPr.Add(new XElement(w + "jc", new XAttribute(w + "val", "center")));
                        }
                    }
                    if (cloneSizeRun != null)
                        MakeBlackRun(cloneSizeRun, targetItems[i].Qty);
                    if (cloneUnitRun != null)
                        MakeBlackRun(cloneUnitRun, targetItems[i].Details);
                    insertAfter.AddAfterSelf(clonedRow);
                    insertAfter = clonedRow;
                }
            }
        }

        // ── Configuration table — one row per checked component ──────────────
        // Build list of (componentText, noteText) for every checked item
        var configItems = _configCheckBoxes
            .Where(kv => kv.Value.IsChecked == true)
            .Select(kv =>
            {
                string config;
                if (kv.Key == "B.B")
                {
                    var parts = new List<string>();
                    var bt = _bbTypeComboBox?.SelectedItem?.ToString() ?? "";
                    var bs = _bbSizeComboBox?.SelectedItem?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(bt)) parts.Add($"Type: {bt}");
                    if (!string.IsNullOrEmpty(bs)) parts.Add($"Size: {bs}");
                    config = string.Join(", ", parts);
                }
                else if (kv.Key == "I.S")
                {
                    var ap = _isExitApertureComboBox?.SelectedItem?.ToString() ?? "";
                    config = !string.IsNullOrEmpty(ap) ? $"Exit Aperture: {ap}" : "";
                }
                else if (kv.Key == "Backlight")
                {
                    var blt = _backlightTypeComboBox?.SelectedItem?.ToString() ?? "";
                    config = !string.IsNullOrEmpty(blt) ? $"Type: {blt}" : "";
                }
                else if (kv.Key == "Frame Grabbers")
                {
                    var fgParts = new List<string>();
                    foreach (var fgCb in new[] { _frameGrabbersComboBox, _frameGrabbersComboBox2, _frameGrabbersComboBox3, _frameGrabbersComboBox4 })
                    {
                        var v = fgCb?.SelectedItem?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(v)) fgParts.Add(v);
                    }
                    config = string.Join(", ", fgParts);
                }
                else if (kv.Key == "Gimbal")
                {
                    var parts = new List<string>();
                    var sz = _gimbalSizeTextBox?.Text.Trim() ?? "";
                    if (!string.IsNullOrEmpty(sz)) parts.Add($"Size: {sz} Inches");
                    if (_gimbalJoystickCheckBox?.IsChecked == true) parts.Add("+Joystick");
                    var lc = _gimbalLoadCapacityTextBox?.Text.Trim() ?? "";
                    if (!string.IsNullOrEmpty(lc)) parts.Add($"Load Capacity: {lc} KG");
                    var acc = _gimbalAccuracyComboBox?.SelectedItem?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(acc)) parts.Add($"Accuracy: {acc}");
                    config = string.Join(", ", parts);
                }
                else if (kv.Key == "LOS alignment target")
                {
                    config = (_losHalogenCheckBox?.IsChecked == true) ? "+Halogen" : "";
                }
                else if (kv.Key == "Rackmount")
                {
                    var parts = new List<string>();
                    if (_rackmountMonitorArmCheckBox?.IsChecked == true) parts.Add("+Monitor Arm");
                    var rh = _rackmountHeightTextBox?.Text.Trim() ?? "";
                    if (!string.IsNullOrEmpty(rh)) parts.Add($"Height: {rh} U");
                    config = string.Join(", ", parts);
                }
                else if (kv.Key == "Optical Table")
                {
                    var parts = new List<string>();
                    if (_opticalTableActiveCheckBox?.IsChecked == true) parts.Add("Active");
                    var w = _opticalTableWidthTextBox?.Text.Trim() ?? "";
                    var l = _opticalTableLengthTextBox?.Text.Trim() ?? "";
                    var h = _opticalTableHeightTextBox?.Text.Trim() ?? "";
                    if (!string.IsNullOrEmpty(w)) parts.Add($"Width: {w}");
                    if (!string.IsNullOrEmpty(l)) parts.Add($"Length: {l}");
                    if (!string.IsNullOrEmpty(h)) parts.Add($"Height: {h}");
                    config = string.Join(", ", parts);
                }
                else if (kv.Key == "VRS")
                {
                    var vrsParts = new List<string>();
                    foreach (var vrsCb in new[] { _vrsComboBox1, _vrsComboBox2, _vrsComboBox3, _vrsComboBox4 })
                    {
                        var v = vrsCb?.SelectedItem?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(v) && v != "NA") vrsParts.Add(v);
                    }
                    config = string.Join(", ", vrsParts);
                }
                else
                {
                    config = "";
                }
                _componentNotes.TryGetValue(kv.Key, out var note);
                return (name: kv.Key, config, note: note ?? "");
            }).ToList();

        // Add any custom config lines (no note)
        if (!string.IsNullOrWhiteSpace(txtCustomConfig.Text))
            configItems.AddRange(txtCustomConfig.Text
                .Split('\n', StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim()).Where(l => l.Length > 0)
                .Select(l => (name: l, config: "", note: "")));

        // Helper: build a plain run with 11pt text
        XElement MakeDataRun(string value) =>
            new XElement(w + "r",
                new XElement(w + "rPr",
                    new XElement(w + "sz",   new XAttribute(w + "val", "22")),
                    new XElement(w + "szCs", new XAttribute(w + "val", "22"))),
                new XElement(w + "t",
                    value.StartsWith(" ") || value.EndsWith(" ")
                        ? new XAttribute(XNamespace.Xml + "space", "preserve") : null!,
                    value));

        // Helper: build a bold run with 11pt text
        XElement MakeBoldDataRun(string value) =>
            new XElement(w + "r",
                new XElement(w + "rPr",
                    new XElement(w + "b"),
                    new XElement(w + "bCs"),
                    new XElement(w + "sz",   new XAttribute(w + "val", "22")),
                    new XElement(w + "szCs", new XAttribute(w + "val", "22"))),
                new XElement(w + "t",
                    value.StartsWith(" ") || value.EndsWith(" ")
                        ? new XAttribute(XNamespace.Xml + "space", "preserve") : null!,
                    value));

        // Helper: add bold-name + regular-options runs into a paragraph element.
        // Splits "Name (options)" → bold "Name" + regular " (options)".
        // If no parenthesis, the whole text is bold.
        void AddComponentRuns(XElement para, string component)
        {
            int paren = component.IndexOf('(');
            if (paren > 0)
            {
                para.Add(MakeBoldDataRun(component[..paren].TrimEnd()));
                para.Add(MakeDataRun(" " + component[paren..]));
            }
            else
            {
                para.Add(MakeBoldDataRun(component));
            }
        }

        // Helper: clone a table cell's tcPr and fill it with plain (non-bold) text — used for Configuration and Notes columns
        XElement CloneConfigCell(XElement sourceCell, string text) => CloneNoteCell(sourceCell, text);

        // Helper: clone a table cell's tcPr and fill it with plain (non-bold) text — used for Notes column
        XElement CloneNoteCell(XElement sourceCell, string text)
        {
            var tcPr = sourceCell.Element(w + "tcPr");
            var cell = new XElement(w + "tc");
            if (tcPr != null) cell.Add(new XElement(tcPr));
            var para = new XElement(w + "p",
                new XElement(w + "pPr",
                    new XElement(w + "rPr",
                        new XElement(w + "sz",   new XAttribute(w + "val", "22")),
                        new XElement(w + "szCs", new XAttribute(w + "val", "22")))));
            if (!string.IsNullOrEmpty(text)) para.Add(MakeDataRun(text));
            cell.Add(para);
            return cell;
        }

        // Helper: clone a table cell's tcPr and fill it with a bold run — used for Component column
        XElement CloneCell(XElement sourceCell, string text)
        {
            var tcPr = sourceCell.Element(w + "tcPr");
            var cell = new XElement(w + "tc");
            if (tcPr != null) cell.Add(new XElement(tcPr));
            var para = new XElement(w + "p",
                new XElement(w + "pPr",
                    new XElement(w + "rPr",
                        new XElement(w + "sz",   new XAttribute(w + "val", "22")),
                        new XElement(w + "szCs", new XAttribute(w + "val", "22")))));
            if (!string.IsNullOrEmpty(text)) para.Add(MakeBoldDataRun(text));
            cell.Add(para);
            return cell;
        }

        var configRun = body.Descendants(w + "r")
            .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Place components from here!");
        if (configRun != null)
        {
            var templateRow = configRun.Ancestors(w + "tr").First();
            var cells = templateRow.Elements(w + "tc").ToList(); // [0]=Component, [1]=Configuration, [2]=Notes

            if (configItems.Count == 0)
            {
                MakeBlackRun(configRun, "");
            }
            else
            {
                // Fill the first (template) row:
                // cells[0] = Component (bold name only)
                // cells[1] = Configuration (options, plain)
                // cells[2] = Notes (plain)
                var configPara = configRun.Parent; // <w:p>
                configPara?.Elements(w + "r").Remove();
                if (configPara != null) configPara.Add(MakeBoldDataRun(configItems[0].name));

                // Fill configuration cell of first row (cells[1])
                if (cells.Count >= 2)
                {
                    var configCell1Para = cells[1].Element(w + "p");
                    configCell1Para?.Elements(w + "r").Remove();
                    if (!string.IsNullOrWhiteSpace(configItems[0].config))
                        configCell1Para?.Add(MakeDataRun(configItems[0].config));
                }

                // Fill notes cell of first row (cells[2])
                if (cells.Count >= 3)
                {
                    var notesPara = cells[2].Element(w + "p");
                    notesPara?.Elements(w + "r").Remove();
                    var notesPPr = notesPara?.Element(w + "pPr");
                    notesPPr?.Element(w + "rPr")?.Elements(w + "b").Remove();
                    notesPPr?.Element(w + "rPr")?.Elements(w + "bCs").Remove();
                    if (!string.IsNullOrWhiteSpace(configItems[0].note))
                        notesPara?.Add(MakeDataRun(configItems[0].note));
                }

                // Add one new row per remaining component
                var insertAfter = templateRow;
                for (int i = 1; i < configItems.Count; i++)
                {
                    var newRow = new XElement(w + "tr");
                    var trPr = templateRow.Element(w + "trPr");
                    if (trPr != null) newRow.Add(new XElement(trPr));

                    var compCell = cells.Count > 0
                        ? CloneCell(cells[0], configItems[i].name)
                        : new XElement(w + "tc", new XElement(w + "p", MakeBoldDataRun(configItems[i].name)));

                    var configCell = cells.Count > 1
                        ? CloneConfigCell(cells[1], configItems[i].config)
                        : new XElement(w + "tc", new XElement(w + "p",
                            string.IsNullOrEmpty(configItems[i].config) ? null : MakeDataRun(configItems[i].config)));

                    var notesCell = cells.Count > 2
                        ? CloneNoteCell(cells[2], configItems[i].note)
                        : new XElement(w + "tc", new XElement(w + "p",
                            string.IsNullOrEmpty(configItems[i].note) ? null : MakeDataRun(configItems[i].note)));

                    newRow.Add(compCell);
                    newRow.Add(configCell);
                    newRow.Add(notesCell);

                    insertAfter.AddAfterSelf(newRow);
                    insertAfter = newRow;
                }
            }
        }

        // Convert leftover lone red space / separator runs to black (they carry spacing between fields)
        foreach (var container in allContainers)
        foreach (var r in container.Descendants(w + "r").Where(r => IsRedRun(r)).ToList())
        {
            var rPr = r.Element(w + "rPr");
            if (rPr != null)
            {
                rPr.Elements(w + "b").Remove();
                rPr.Elements(w + "bCs").Remove();
                rPr.Elements(w + "color").Remove();
            }
        }

        // ── Section paragraph replacements ────────────────────────────────────
        ReplaceMarkerTableRows("Start inserting PM questions",        _questions.Select(q => (q.Text, q.IsBold)));
        ReplaceMarkerTableRows("Start inserting Marketing questions", _marketingQuestions.Select(q => (q.Text, q.IsBold)));
        ReplaceMarkerTableRows("Start inserting Marketing notes",     _marketingNotes.Select(n => (n.Text, n.IsBold)));
        ReplaceMarkerTableRows("Start inserting PM notes",           _pmNotes.Select(n => (n.Text, n.IsBold)));

        // ── Save ──────────────────────────────────────────────────────────────
        var settings = new System.Xml.XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: true),
            Indent = false
        };
        using var writer = System.Xml.XmlWriter.Create(outputPath, settings);
        doc.Save(writer);
    }

    private void GenerateWordDocument_PLACEHOLDER(string filePath)
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
            var sysTypePh = string.Join(" ", new[] { cmbSystemType.SelectedItem?.ToString() ?? "", cmbSystemVariant.SelectedItem?.ToString() ?? "", cmbSystemAperture.SelectedItem?.ToString() ?? "" }.Where(s => !string.IsNullOrEmpty(s)));
            if (!string.IsNullOrWhiteSpace(sysTypePh))
                AddParagraph(body, sysTypePh, 24, true);

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
                .Select(kv =>
                {
                    if (kv.Key == "B.B")
                    {
                        var type = _bbTypeComboBox?.SelectedItem?.ToString() ?? "";
                        var size = _bbSizeComboBox?.SelectedItem?.ToString() ?? "";
                        var extras = new List<string>();
                        if (!string.IsNullOrEmpty(type)) extras.Add($"Type: {type}");
                        if (!string.IsNullOrEmpty(size)) extras.Add($"Size: {size}");
                        return extras.Count > 0 ? $"B.B ({string.Join(", ", extras)})" : "B.B";
                    }
                    if (kv.Key == "I.S")
                    {
                        var aperture = _isExitApertureComboBox?.SelectedItem?.ToString() ?? "";
                        return !string.IsNullOrEmpty(aperture) ? $"I.S (Exit Aperture: {aperture})" : "I.S";
                    }
                    if (kv.Key == "Backlight")
                    {
                        var btype = _backlightTypeComboBox?.SelectedItem?.ToString() ?? "";
                        return !string.IsNullOrEmpty(btype) ? $"Backlight (Type: {btype})" : "Backlight";
                    }
                    if (kv.Key == "Frame Grabbers")
                    {
                        var fgParts = new List<string>();
                        foreach (var fgCb in new[] { _frameGrabbersComboBox, _frameGrabbersComboBox2, _frameGrabbersComboBox3, _frameGrabbersComboBox4 })
                        {
                            var v = fgCb?.SelectedItem?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(v)) fgParts.Add(v);
                        }
                        return fgParts.Count > 0 ? $"Frame Grabbers ({string.Join(", ", fgParts)})" : "Frame Grabbers";
                    }
                    if (kv.Key == "Gimbal")
                    {
                        var parts = new List<string>();
                        var sz = _gimbalSizeTextBox?.Text.Trim() ?? "";
                        if (!string.IsNullOrEmpty(sz)) parts.Add($"Size: {sz} Inches");
                        if (_gimbalJoystickCheckBox?.IsChecked == true) parts.Add("+Joystick");
                        var lc = _gimbalLoadCapacityTextBox?.Text.Trim() ?? "";
                        if (!string.IsNullOrEmpty(lc)) parts.Add($"Load Capacity: {lc} KG");
                        var acc = _gimbalAccuracyComboBox?.SelectedItem?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(acc)) parts.Add($"Accuracy: {acc}");
                        return parts.Count > 0 ? $"Gimbal ({string.Join(", ", parts)})" : "Gimbal";
                    }
                    if (kv.Key == "LOS alignment target")
                        return (_losHalogenCheckBox?.IsChecked == true)
                            ? "LOS alignment target + Halogen"
                            : "LOS alignment target";
                    if (kv.Key == "VRS")
                    {
                        var vrsParts = new List<string>();
                        foreach (var vrsCb in new[] { _vrsComboBox1, _vrsComboBox2, _vrsComboBox3, _vrsComboBox4 })
                        {
                            var v = vrsCb?.SelectedItem?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(v) && v != "NA") vrsParts.Add(v);
                        }
                        return vrsParts.Count > 0 ? $"VRS ({string.Join(", ", vrsParts)})" : "VRS";
                    }
                    return kv.Key;
                });

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

            // Marketing Notes
            var mktNotes = _marketingNotes.Where(n => !string.IsNullOrWhiteSpace(n.Text)).ToList();
            if (mktNotes.Any())
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "Marketing Notes:", 24, true);
                foreach (var n in mktNotes) AddBulletPoint(body, n.Text, 24);
            }

            // PM Notes
            var pmNotesList = _pmNotes.Where(n => !string.IsNullOrWhiteSpace(n.Text)).ToList();
            if (pmNotesList.Any())
            {
                AddParagraph(body, "", 24, false);
                AddParagraph(body, "PM Notes:", 24, true);
                foreach (var n in pmNotesList) AddBulletPoint(body, n.Text, 24);
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

    // ── Template System ──────────────────────────────────────────────────────

    private static string GetTemplatesFolder()
    {
        string folder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
        System.IO.Directory.CreateDirectory(folder);
        return folder;
    }

    private string GetTemplateName()
    {
        string order    = txtOrderNumber.Text.Trim();
        string customer = txtCustomerName.Text.Trim();
        if (string.IsNullOrEmpty(order) && string.IsNullOrEmpty(customer))
            return "";
        return $"{order} {customer}".Trim();
    }

    private TemplateData CollectTemplateData()
    {
        var data = new TemplateData
        {
            OrderNumber        = txtOrderNumber.Text,
            CustomerName       = txtCustomerName.Text,
            FinalCustomer      = txtFinalCustomer.Text,
            SystemType         = cmbSystemType.SelectedItem?.ToString() ?? "",
            SystemVariant      = cmbSystemVariant.SelectedItem?.ToString() ?? "",
            SystemAperture     = cmbSystemAperture.SelectedItem?.ToString() ?? "",
            ProjectType        = cmbSystemType.SelectedItem?.ToString() ?? "",
            PakaNumber         = txtPakaNumber.Text,
            ReferenceOrder     = txtReferenceOrder.Text,
            DeliveryDate       = dpDeliveryDate.SelectedDate?.ToString("yyyy-MM-dd") ?? "",
            DesignDueDate      = dpDesignDueDate.SelectedDate?.ToString("yyyy-MM-dd") ?? "",
            Participants       = new List<string>(selectedParticipants),
            SellingPrice       = txtSellingPrice.Text,
            MaterialCost       = txtMaterialCost.Text,
            ProjectHours       = txtProjectHours.Text,
            Penalties          = txtPenalties.Text,
            DORated            = chkDORated.IsChecked == true,
            ComponentNotes     = new Dictionary<string, string>(_componentNotes),
            BBType             = _bbTypeComboBox?.SelectedItem?.ToString() ?? "",
            BBSize             = _bbSizeComboBox?.SelectedItem?.ToString() ?? "",
            ISAperture         = _isExitApertureComboBox?.SelectedItem?.ToString() ?? "",
            BacklightType      = _backlightTypeComboBox?.SelectedItem?.ToString() ?? "",
            MaxWeight          = _maxWeightTextBox?.Text ?? "",
            FiniteDistance     = _finiteDistance1TextBox?.Text ?? "",
            Vrs1               = _vrsComboBox1?.SelectedItem?.ToString() ?? "",
            Vrs2               = _vrsComboBox2?.SelectedItem?.ToString() ?? "",
            Vrs3               = _vrsComboBox3?.SelectedItem?.ToString() ?? "",
            Vrs4               = _vrsComboBox4?.SelectedItem?.ToString() ?? "",
            GimbalSize         = _gimbalSizeTextBox?.Text ?? "",
            GimbalLoadCapacity = _gimbalLoadCapacityTextBox?.Text ?? "",
            GimbalAccuracy     = _gimbalAccuracyComboBox?.SelectedItem?.ToString() ?? "",
            GimbalJoystick     = _gimbalJoystickCheckBox?.IsChecked == true,
            LosHalogen         = _losHalogenCheckBox?.IsChecked == true,
            SourceStageManual  = _sourceStageManualCheckBox?.IsChecked == true,
            XyStageManual      = _xyStageManualCheckBox?.IsChecked == true,
            FrameGrabber1      = _frameGrabbersComboBox?.SelectedItem?.ToString() ?? "",
            FrameGrabber2      = _frameGrabbersComboBox2?.SelectedItem?.ToString() ?? "",
            FrameGrabber3      = _frameGrabbersComboBox3?.SelectedItem?.ToString() ?? "",
            FrameGrabber4      = _frameGrabbersComboBox4?.SelectedItem?.ToString() ?? "",
            RackmountMonitorArm = _rackmountMonitorArmCheckBox?.IsChecked == true,
            RackmountHeight     = _rackmountHeightTextBox?.Text ?? "",
            OpticalTableWidth   = _opticalTableWidthTextBox?.Text ?? "",
            OpticalTableLength  = _opticalTableLengthTextBox?.Text ?? "",
            OpticalTableHeight  = _opticalTableHeightTextBox?.Text ?? "",
            OpticalTableActive  = _opticalTableActiveCheckBox?.IsChecked == true,
            Targets            = _targets.Select(t => new TargetItem { Type = t.Type, Qty = t.Qty, Details = t.Details }).ToList(),
            PmQuestions        = _questions.Select(q => q.Text).ToList(),
            MarketingQuestions = _marketingQuestions.Select(q => q.Text).ToList(),
            MarketingNotes     = _marketingNotes.Select(n => n.Text).ToList(),
            PmNotes            = _pmNotes.Select(n => n.Text).ToList(),
        };

        foreach (var kv in _configCheckBoxes)
            data.ConfigCheckBoxes[kv.Key] = kv.Value.IsChecked == true;

        return data;
    }

    private void SaveTemplate()
    {
        try
        {
            string name = GetTemplateName();
            if (string.IsNullOrEmpty(name)) return;

            string folder = GetTemplatesFolder();
            string path   = System.IO.Path.Combine(folder, name + ".json");

            var    data    = CollectTemplateData();
            var    options = new JsonSerializerOptions { WriteIndented = true };
            string json    = JsonSerializer.Serialize(data, options);
            System.IO.File.WriteAllText(path, json);
        }
        catch
        {
            // Silent auto-save — never interrupt the user
        }
    }

    private void StartAutoSaveTimer()
    {
        _countdownSeconds = 60;

        // Per-second countdown timer — updates the label
        _countdownTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(1)
        };
        _countdownTimer.Tick += (s, e) =>
        {
            _countdownSeconds--;
            if (_countdownSeconds < 0) _countdownSeconds = 59;
            UpdateCountdownLabel();
        };
        _countdownTimer.Start();

        // Auto-save timer — fires every 60 seconds
        _autoSaveTimer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromMinutes(1)
        };
        _autoSaveTimer.Tick += (s, e) =>
        {
            SaveTemplate();
            _countdownSeconds = 60;
            UpdateCountdownLabel();
        };
        _autoSaveTimer.Start();

        UpdateCountdownLabel();
    }

    private void UpdateCountdownLabel()
    {
        if (txtAutoSaveCountdown == null) return;
        int mins = _countdownSeconds / 60;
        int secs = _countdownSeconds % 60;
        txtAutoSaveCountdown.Text = $"Automatic save in: {mins}:{secs:D2}";
    }

    private TemplateData? _pendingTemplateConfig;

    private void RefreshVariantOptions()
    {
        var saved = cmbSystemVariant.SelectedItem?.ToString() ?? "";
        cmbSystemVariant.Items.Clear();
        string[]? variants = cmbSystemType.SelectedItem?.ToString() switch
        {
            "METS" => new[] { "VS", "S", "L", "VL" },
            "ILET" => new[] { "4", "5", "6" },
            "WFOV" => new[] { "VS", "S", "L", "VL" },
            "CFT"  => new[] { "VS", "S", "L", "VL" },
            _ => null
        };
        if (variants != null)
            foreach (var v in variants) cmbSystemVariant.Items.Add(v);
        if (!string.IsNullOrEmpty(saved) && cmbSystemVariant.Items.Contains(saved))
            cmbSystemVariant.SelectedItem = saved;
    }

    private void ApplyTemplateData(TemplateData data)
    {
        txtOrderNumber.Text    = data.OrderNumber;
        txtCustomerName.Text   = data.CustomerName;
        txtFinalCustomer.Text  = data.FinalCustomer;
        // Restore System Type combos (support both new and legacy saved data)
        var sysType = !string.IsNullOrEmpty(data.SystemType) ? data.SystemType : data.ProjectType;
        RestoreComboSelection(cmbSystemType, sysType);
        RefreshVariantOptions();
        RestoreComboSelection(cmbSystemVariant, data.SystemVariant);
        RestoreComboSelection(cmbSystemAperture, data.SystemAperture);
        txtPakaNumber.Text     = data.PakaNumber;
        txtReferenceOrder.Text = data.ReferenceOrder;
        txtSellingPrice.Text   = data.SellingPrice;
        txtMaterialCost.Text   = data.MaterialCost;
        txtProjectHours.Text   = data.ProjectHours;
        txtPenalties.Text      = data.Penalties;
        chkDORated.IsChecked   = data.DORated;

        if (DateTime.TryParse(data.DeliveryDate, out var dd))
            dpDeliveryDate.SelectedDate = dd;
        if (DateTime.TryParse(data.DesignDueDate, out var ddd))
            dpDesignDueDate.SelectedDate = ddd;

        selectedParticipants.Clear();
        selectedParticipants.AddRange(data.Participants);

        foreach (var kv in data.ConfigCheckBoxes)
            if (_configCheckBoxes.TryGetValue(kv.Key, out var cb))
                cb.IsChecked = kv.Value;

        _componentNotes.Clear();
        foreach (var kv in data.ComponentNotes)
            _componentNotes[kv.Key] = kv.Value;

        _targets.Clear();
        foreach (var t in data.Targets)
            _targets.Add(new TargetItem { Type = t.Type, Qty = t.Qty, Details = t.Details });

        _questions.Clear();
        for (int i = 0; i < data.PmQuestions.Count; i++)
            _questions.Add(new QuestionItem { Text = data.PmQuestions[i], Number = i + 1 });

        _marketingQuestions.Clear();
        for (int i = 0; i < data.MarketingQuestions.Count; i++)
            _marketingQuestions.Add(new QuestionItem { Text = data.MarketingQuestions[i], Number = i + 1 });

        _marketingNotes.Clear();
        for (int i = 0; i < data.MarketingNotes.Count; i++)
            _marketingNotes.Add(new QuestionItem { Text = data.MarketingNotes[i], Number = i + 1 });

        _pmNotes.Clear();
        for (int i = 0; i < data.PmNotes.Count; i++)
            _pmNotes.Add(new QuestionItem { Text = data.PmNotes[i], Number = i + 1 });

        // Store pending inline config data for when the Config section is next loaded
        _pendingTemplateConfig = data;

        // Reload the Overview section to reflect the new data
        LoadSection(0);
    }

    private void ApplyPendingTemplateConfig()
    {
        if (_pendingTemplateConfig == null) return;
        var d = _pendingTemplateConfig;

        RestoreComboSelection(_bbTypeComboBox,          d.BBType);
        RestoreComboSelection(_bbSizeComboBox,          d.BBSize);
        RestoreComboSelection(_isExitApertureComboBox,  d.ISAperture);
        RestoreComboSelection(_backlightTypeComboBox,   d.BacklightType);
        if (_maxWeightTextBox          != null) _maxWeightTextBox.Text          = d.MaxWeight;
        if (_finiteDistance1TextBox    != null) _finiteDistance1TextBox.Text    = d.FiniteDistance;
        RestoreComboSelection(_vrsComboBox1, d.Vrs1);
        RestoreComboSelection(_vrsComboBox2, d.Vrs2);
        RestoreComboSelection(_vrsComboBox3, d.Vrs3);
        RestoreComboSelection(_vrsComboBox4, d.Vrs4);
        if (_gimbalSizeTextBox         != null) _gimbalSizeTextBox.Text         = d.GimbalSize;
        if (_gimbalLoadCapacityTextBox != null) _gimbalLoadCapacityTextBox.Text = d.GimbalLoadCapacity;
        RestoreComboSelection(_gimbalAccuracyComboBox, d.GimbalAccuracy);
        if (_gimbalJoystickCheckBox    != null) _gimbalJoystickCheckBox.IsChecked = d.GimbalJoystick;
        if (_losHalogenCheckBox        != null) _losHalogenCheckBox.IsChecked        = d.LosHalogen;
        if (_sourceStageManualCheckBox  != null) _sourceStageManualCheckBox.IsChecked = d.SourceStageManual;
        if (_xyStageManualCheckBox      != null) _xyStageManualCheckBox.IsChecked     = d.XyStageManual;
        RestoreComboSelection(_frameGrabbersComboBox,  d.FrameGrabber1);
        RestoreComboSelection(_frameGrabbersComboBox2, d.FrameGrabber2);
        RestoreComboSelection(_frameGrabbersComboBox3, d.FrameGrabber3);
        RestoreComboSelection(_frameGrabbersComboBox4, d.FrameGrabber4);
        if (_rackmountMonitorArmCheckBox != null) _rackmountMonitorArmCheckBox.IsChecked = d.RackmountMonitorArm;
        if (_rackmountHeightTextBox      != null) _rackmountHeightTextBox.Text           = d.RackmountHeight;
        if (_opticalTableWidthTextBox    != null) _opticalTableWidthTextBox.Text         = d.OpticalTableWidth;
        if (_opticalTableLengthTextBox   != null) _opticalTableLengthTextBox.Text        = d.OpticalTableLength;
        if (_opticalTableHeightTextBox   != null) _opticalTableHeightTextBox.Text        = d.OpticalTableHeight;
        if (_opticalTableActiveCheckBox  != null) _opticalTableActiveCheckBox.IsChecked  = d.OpticalTableActive;

        SyncInlineControlStates();
        _pendingTemplateConfig = null;
    }

    private void LoadTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        string folder = GetTemplatesFolder();

        var dialog = new TemplatePickerWindow(folder);
        dialog.Owner = this;
        if (dialog.ShowDialog() == true && dialog.SelectedTemplatePath != null)
        {
            try
            {
                string json = System.IO.File.ReadAllText(dialog.SelectedTemplatePath);
                var data = JsonSerializer.Deserialize<TemplateData>(json);
                if (data != null)
                {
                    ApplyTemplateData(data);
                    MessageBox.Show("Template loaded successfully!", "Load Template",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load template:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void SendSummaryButton_Click(object sender, RoutedEventArgs e)
    {
        // If not yet exported, silently generate to a temp file so we always have an attachment
        if (string.IsNullOrEmpty(_lastExportedPath) ||
            !System.IO.File.Exists(_lastExportedPath))
        {
            try
            {
                string tempPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"Kickoff {txtOrderNumber.Text.Trim()}.doc");

                GenerateFromTemplate(tempPath);
                _lastExportedPath = tempPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not auto-generate the summary document:\n\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        var win = new SendSummaryWindow(
            txtOrderNumber.Text.Trim(),
            txtCustomerName.Text.Trim(),
            string.Join(" ", new[] { cmbSystemType.SelectedItem?.ToString() ?? "", cmbSystemVariant.SelectedItem?.ToString() ?? "", cmbSystemAperture.SelectedItem?.ToString() ?? "" }.Where(s => !string.IsNullOrEmpty(s))),
            new System.Collections.Generic.List<string>(selectedParticipants),
            _lastExportedPath ?? "");
        win.Owner = this;
        win.ShowDialog();
    }
}

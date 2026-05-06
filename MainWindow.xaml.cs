using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
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

public class QuestionItem : INotifyPropertyChanged
{
    private string _text = "";
    private int _number;

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

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string propertyName) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
        "Gimbal"
    };

    private readonly Dictionary<string, WpfCheckBox> _configCheckBoxes = new();
    private readonly Dictionary<string, string> _componentNotes = new();
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
    private System.Windows.Controls.TextBox? _gimbalSizeTextBox;
    private readonly ObservableCollection<TargetItem> _targets = new();
    private readonly ObservableCollection<QuestionItem> _questions = new();
    private readonly ObservableCollection<QuestionItem> _marketingQuestions = new();
    private readonly ObservableCollection<QuestionItem> _marketingNotes = new();
    private readonly ObservableCollection<QuestionItem> _pmNotes = new();

    // Form controls
    private System.Windows.Controls.TextBox txtOrderNumber = new();
    private System.Windows.Controls.TextBox txtCustomerName = new();
    private System.Windows.Controls.TextBox txtFinalCustomer = new();
    private System.Windows.Controls.TextBox txtProjectType = new();
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
                case 4: // Notes
                    LoadNotesSection(panel);
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

        // Customer Name - spans 2 columns (2/3 width)
        AddFormField(orderFormGrid, "Customer Name:", txtCustomerName, row++, 1, 2);

        // Agent and Project Type side by side
        AddFormField(orderFormGrid, "Territory:", txtFinalCustomer, row, 0, 1);
        AddFormField(orderFormGrid, "Project Type:", txtProjectType, row++, 1, 2);

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

        // Selling Price and Material Cost
        AddFormField(marketingFormGrid, "Selling Price:", txtSellingPrice, marketingRow, 0, 1);
        AddFormField(marketingFormGrid, "Material Cost:", txtMaterialCost, marketingRow++, 1, 1);

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
            Padding = new WpfThickness(4),
            MinHeight = 32 // Reduced from 36 to 32
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
            Padding = new WpfThickness(4, 4, 4, 4), // Reduced from 4,6,4,6
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

        // Remove button (X)
        var removeButton = new System.Windows.Controls.Button
        {
            Content = "✕",
            Width = 16,
            Height = 16,
            FontSize = 10,
            Background = System.Windows.Media.Brushes.Transparent,
            BorderThickness = new WpfThickness(0),
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(100, 116, 139)),
            Cursor = System.Windows.Input.Cursors.Hand,
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            Padding = new WpfThickness(0),
            Tag = name
        };
        removeButton.Click += RemoveParticipantChip_Click;
        chipPanel.Children.Add(removeButton);

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

    private void RemoveParticipantChip_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button button) return;
        if (button.Tag is not string name) return;

        selectedParticipants.Remove(name);

        // Find and remove the chip border
        if (button.Parent is System.Windows.Controls.Panel chipPanel && 
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

                _bbTypeComboBox = MakeWhiteComboBox(138, "RR", "STD", "SR200N-33");
                _bbTypeComboBox.SelectedIndex = 0;
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

                _bbSizeComboBox = MakeWhiteComboBox(108, "1D", "2D", "4D", "8D", "12D");
                var w2 = MakeComboWrapper(_bbSizeComboBox, 108);
                w2.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w2, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w2, 4);
                radiationGrid.Children.Add(w2);
            }
            else if (item == "I.S")
            {
                var lbl1 = MakeComboLabel("Exit Aperture:");
                lbl1.Margin = new WpfThickness(0, 0, 6, 14);
                System.Windows.Controls.Grid.SetRow(lbl1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(lbl1, 1);
                radiationGrid.Children.Add(lbl1);

                _isExitApertureComboBox = MakeWhiteComboBox(138, "2\"", "3\"", "4\"", "5\"");
                var w1 = MakeComboWrapper(_isExitApertureComboBox, 138);
                w1.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w1, 2);
                radiationGrid.Children.Add(w1);
            }
            else if (item == "Backlight")
            {
                var lbl1 = MakeComboLabel("Type:");
                lbl1.Margin = new WpfThickness(0, 0, 6, 14);
                System.Windows.Controls.Grid.SetRow(lbl1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(lbl1, 1);
                radiationGrid.Children.Add(lbl1);

                _backlightTypeComboBox = MakeWhiteComboBox(138, "LED", "Fiber Optic");
                var w1 = MakeComboWrapper(_backlightTypeComboBox, 138);
                w1.Margin = new WpfThickness(0, 0, 0, 14);
                System.Windows.Controls.Grid.SetRow(w1, radRowIdx);
                System.Windows.Controls.Grid.SetColumn(w1, 2);
                radiationGrid.Children.Add(w1);
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
        // col1 = 190px  (rows 0-2: Rackmount/Device Center/Power Meter  |  rows 3/4/6: inline dim label)
        // col2 = 190px  (rows 0-2: Manual Choke/LOS alignment target/Energy Meter  |  rows 3/4/6: inline textbox)
        // col3 = Auto   (inline unit: [KG], [m], [Inches] — empty in rows 0-2)
        var componentsGrid = new System.Windows.Controls.Grid();
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(190) });  // col0
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(190) });  // col1
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new GridLength(190) });  // col2
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });       // col3 unit

        // Rows 0-2: the 9 plain checkboxes in a 3-column sub-arrangement using col0, col4, col5
        string[] top9 = { "Source Stage", "Rackmount", "Manual Choke",
                           "CTE", "Device Center", "LOS alignment target",
                           "XY Stage", "Power Meter", "Energy Meter" };

        // Column mapping: left col → grid col0, middle col → grid col1, right col → grid col2
        int[] top9GridCols = { 0, 1, 2 };

        for (int i = 0; i < top9.Length; i++)
        {
            if (i % 3 == 0)
                componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });

            var cb = _configCheckBoxes[top9[i]];
            DetachCheckBox(cb);
            StyleSysCheckBox(cb);
            cb.Margin = new WpfThickness(0, 0, 0, 0);

            var cbw = WrapCheckBoxWithNoteButton(cb, top9[i]);
            cbw.Margin = new WpfThickness(0, 0, 0, 14);

            System.Windows.Controls.Grid.SetRow(cbw, i / 3);
            System.Windows.Controls.Grid.SetColumn(cbw, top9GridCols[i % 3]);
            componentsGrid.Children.Add(cbw);
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

            var mwRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbNPWrapper = WrapCheckBoxWithNoteButton(cbNP, "NewPort Stage");
            mwRow.Children.Add(cbNPWrapper); mwRow.Children.Add(lblMW); mwRow.Children.Add(_maxWeightTextBox); mwRow.Children.Add(lblKG);
            System.Windows.Controls.Grid.SetRow(mwRow, r); System.Windows.Controls.Grid.SetColumn(mwRow, 0); System.Windows.Controls.Grid.SetColumnSpan(mwRow, 4);
            componentsGrid.Children.Add(mwRow);
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

            var fdRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbFSWrapper = WrapCheckBoxWithNoteButton(cbFS, "Focus Stage");
            fdRow.Children.Add(cbFSWrapper); fdRow.Children.Add(lblFD); fdRow.Children.Add(_finiteDistance1TextBox); fdRow.Children.Add(lblM);
            System.Windows.Controls.Grid.SetRow(fdRow, r); System.Windows.Controls.Grid.SetColumn(fdRow, 0); System.Windows.Controls.Grid.SetColumnSpan(fdRow, 4);
            componentsGrid.Children.Add(fdRow);
        }

        // ── Row 5: VRS (checkbox only) ────────────────────────────────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 5;

            var cbVRS = _configCheckBoxes["VRS"];
            DetachCheckBox(cbVRS); StyleSysCheckBox(cbVRS);
            cbVRS.Margin = new WpfThickness(0, 0, 0, 0);
            var cbVRSWrapper = WrapCheckBoxWithNoteButton(cbVRS, "VRS");
            cbVRSWrapper.Margin = new WpfThickness(0, 0, 0, 14);
            System.Windows.Controls.Grid.SetRow(cbVRSWrapper, r); System.Windows.Controls.Grid.SetColumn(cbVRSWrapper, 0);
            componentsGrid.Children.Add(cbVRSWrapper);

            _vrsComboBox1 = null;
            _vrsComboBox2 = null;
            _vrsComboBox3 = null;
        }

        // ── Row 6: Gimbal | Size: [___] [Inches] ──────────────────────────────
        {
            componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            int r = 6;

            var cbGimbal = _configCheckBoxes["Gimbal"];
            DetachCheckBox(cbGimbal); StyleSysCheckBox(cbGimbal);
            cbGimbal.Margin = new WpfThickness(0, 0, 0, 0);

            var lblSize = DimLabel("Size:"); lblSize.Margin = new WpfThickness(20, 0, 4, 0);
            _gimbalSizeTextBox = DimTextBox(55); _gimbalSizeTextBox.Margin = new WpfThickness(0);
            var lblInches = DimLabel("[Inches]"); lblInches.Margin = new WpfThickness(4, 0, 0, 0);

            var szRow = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new WpfThickness(0, 0, 0, 14) };
            var cbGimbalWrapper = WrapCheckBoxWithNoteButton(cbGimbal, "Gimbal");
            szRow.Children.Add(cbGimbalWrapper); szRow.Children.Add(lblSize); szRow.Children.Add(_gimbalSizeTextBox); szRow.Children.Add(lblInches);
            System.Windows.Controls.Grid.SetRow(szRow, r); System.Windows.Controls.Grid.SetColumn(szRow, 0); System.Windows.Controls.Grid.SetColumnSpan(szRow, 4);
            componentsGrid.Children.Add(szRow);
        }

        systemOptionsContainer.Child = componentsGrid;
        systemComponentsPanel.Children.Add(systemOptionsContainer);
        stackPanel.Children.Add(systemComponentsSection);

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
            Text = "Questions",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        stackPanel.Children.Add(title);

        // PM Questions section - compact
        var pmQuestionsLabel = new System.Windows.Controls.TextBlock
        {
            Text = "PM Questions:",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(51, 65, 85)),
            Margin = new WpfThickness(0, 0, 0, 8)
        };
        stackPanel.Children.Add(pmQuestionsLabel);

        // Add Question button - compact
        var addButton = new System.Windows.Controls.Button
        {
            Content = "+ Add Question",
            Margin = new WpfThickness(0, 0, 0, 8),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
            FontSize = 12
        };
        ApplyModernButtonStyle(addButton);
        addButton.Padding = new WpfThickness(10, 2, 10, 2);
        addButton.Click += AddQuestion_Click;
        stackPanel.Children.Add(addButton);

        // Questions ItemsControl
        questionsItemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0, 0, 0, 16)
        };

        // Numbered row: "1. [TextBox] [🗑]"
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

        var tbFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        tbFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty,
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
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

        stackPanel.Children.Add(questionsItemsControl);

        // Marketing Questions section - compact
        var marketingQuestionsLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Marketing Questions:",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(51, 65, 85)),
            Margin = new WpfThickness(0, 0, 0, 8)
        };
        stackPanel.Children.Add(marketingQuestionsLabel);

        // Add Marketing Question button - compact
        var addMarketingButton = new System.Windows.Controls.Button
        {
            Content = "+ Add Question",
            Margin = new WpfThickness(0, 0, 0, 8),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
            FontSize = 12
        };
        ApplyModernButtonStyle(addMarketingButton);
        addMarketingButton.Padding = new WpfThickness(10, 2, 10, 2);
        addMarketingButton.Click += AddMarketingQuestion_Click;
        stackPanel.Children.Add(addMarketingButton);

        // Marketing Questions ItemsControl
        marketingQuestionsItemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0, 0, 0, 16)
        };

        // Numbered row: "1. [TextBox] [🗑]"
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

        var mTbFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        mTbFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty,
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
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

        stackPanel.Children.Add(marketingQuestionsItemsControl);

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

        var tbFact = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        tbFact.SetBinding(System.Windows.Controls.TextBox.TextProperty,
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
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
        _marketingNotes.Add(new QuestionItem { Text = "", Number = 1 });
        _pmNotes.Add(new QuestionItem { Text = "", Number = 1 });
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

        void ReplaceMarkerParagraph(string markerFragment, IEnumerable<string> lines)
        {
            var para = body.Descendants(w + "p").FirstOrDefault(p =>
                string.Concat(p.Descendants(w + "t").Select(t => (string?)t ?? ""))
                      .Contains(markerFragment));
            if (para == null) return;

            var bullets = lines.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
            var replacements = (bullets.Count > 0 ? bullets : new List<string> { "" })
                .Select(MakeBulletPara).ToArray<object>();

            para.AddAfterSelf(replacements);
            para.Remove();
        }

        // ── Simple field replacements ─────────────────────────────────────────
        ReplacePlaceholder("Order Number",   txtOrderNumber.Text);
        ReplacePlaceholder("Customer Name",  txtCustomerName.Text);
        ReplacePlaceholder("Territory",      txtFinalCustomer.Text);
        ReplacePlaceholder("Paka Number",    txtPakaNumber.Text);
        ReplacePlaceholder("Project Type",   txtProjectType.Text);
        ReplacePlaceholder("Project Hours",  txtProjectHours.Text);
        ReplacePlaceholder("Selling Price",  txtSellingPrice.Text);
        ReplacePlaceholder("Material Cost",  txtMaterialCost.Text);
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
                .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Size Unit");

            if (targetItems.Count == 0)
            {
                MakeBlackRun(targetNameRun, "");
                if (sizeUnitRun != null) MakeBlackRun(sizeUnitRun, "");
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
                    MakeBlackRun(sizeUnitRun,
                        $"{targetItems[0].Qty} {targetItems[0].Details}".Trim());

                var insertAfter = templateRow;
                for (int i = 1; i < targetItems.Count; i++)
                {
                    // Clone from the pristine (unmodified) template row
                    var clonedRow = new XElement(pristineRow);
                    var cloneNameRun = clonedRow.Descendants(w + "r")
                        .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Target Name");
                    var cloneSizeRun = clonedRow.Descendants(w + "r")
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
                        MakeBlackRun(cloneSizeRun,
                            $"{targetItems[i].Qty} {targetItems[i].Details}".Trim());
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
                string component;
                if (kv.Key == "B.B")
                {
                    var parts = new List<string>();
                    var bt = _bbTypeComboBox?.SelectedItem?.ToString() ?? "";
                    var bs = _bbSizeComboBox?.SelectedItem?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(bt)) parts.Add($"Type: {bt}");
                    if (!string.IsNullOrEmpty(bs)) parts.Add($"Size: {bs}");
                    component = parts.Count > 0 ? $"B.B ({string.Join(", ", parts)})" : "B.B";
                }
                else if (kv.Key == "I.S")
                {
                    var ap = _isExitApertureComboBox?.SelectedItem?.ToString() ?? "";
                    component = !string.IsNullOrEmpty(ap) ? $"I.S (Exit Aperture: {ap})" : "I.S";
                }
                else if (kv.Key == "Backlight")
                {
                    var blt = _backlightTypeComboBox?.SelectedItem?.ToString() ?? "";
                    component = !string.IsNullOrEmpty(blt) ? $"Backlight (Type: {blt})" : "Backlight";
                }
                else
                {
                    component = kv.Key;
                }
                _componentNotes.TryGetValue(kv.Key, out var note);
                return (component, note: note ?? "");
            }).ToList();

        // Add any custom config lines (no note)
        if (!string.IsNullOrWhiteSpace(txtCustomConfig.Text))
            configItems.AddRange(txtCustomConfig.Text
                .Split('\n', StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim()).Where(l => l.Length > 0)
                .Select(l => (component: l, note: "")));

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

        // Helper: clone a table cell's tcPr (formatting) and return a fresh <w:tc>
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
            if (!string.IsNullOrEmpty(text)) para.Add(MakeDataRun(text));
            cell.Add(para);
            return cell;
        }

        var configRun = body.Descendants(w + "r")
            .FirstOrDefault(r => IsRedRun(r) && RunText(r) == "Place components from here!");
        if (configRun != null)
        {
            var templateRow = configRun.Ancestors(w + "tr").First();
            var cells = templateRow.Elements(w + "tc").ToList(); // [0]=Component, [1]=Notes

            if (configItems.Count == 0)
            {
                MakeBlackRun(configRun, "");
            }
            else
            {
                // Fill the first (template) row
                MakeBlackRun(configRun, configItems[0].component);

                // Fill notes cell of first row
                if (cells.Count >= 2)
                {
                    var notesPara = cells[1].Element(w + "p");
                    notesPara?.Elements(w + "r").Remove();
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
                        ? CloneCell(cells[0], configItems[i].component)
                        : new XElement(w + "tc", new XElement(w + "p", MakeDataRun(configItems[i].component)));

                    var notesCell = cells.Count > 1
                        ? CloneCell(cells[1], configItems[i].note)
                        : new XElement(w + "tc", new XElement(w + "p",
                            string.IsNullOrEmpty(configItems[i].note) ? null : MakeDataRun(configItems[i].note)));

                    newRow.Add(compCell);
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
        ReplaceMarkerParagraph("Start inserting PM questions",         _questions.Select(q => q.Text));
        ReplaceMarkerParagraph("Start inserting marketing questions",  _marketingQuestions.Select(q => q.Text));
        ReplaceMarkerParagraph("Start inserting marketing notes",      _marketingNotes.Select(n => n.Text));
        ReplaceMarkerParagraph("Start inserting PM notes",            _pmNotes.Select(n => n.Text));

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
}

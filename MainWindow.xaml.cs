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
    private readonly ObservableCollection<TargetItem> _targets = new();
    private readonly ObservableCollection<QuestionItem> _questions = new();

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
    private System.Windows.Controls.TextBox txtActions = new();
    private System.Windows.Controls.TextBox txtMeetingSummary = new();
    private System.Windows.Controls.TextBox txtDecisions = new();
    private System.Windows.Controls.TextBox txtRisks = new();
    private System.Windows.Controls.TextBox txtLogistics = new();
    private System.Windows.Controls.TextBox txtSoftware = new();
    private System.Windows.Controls.TextBox txtTraining = new();

    public MainWindow()
    {
        try
        {
            InitializeComponent();
            InitializeConfigItems();
            InitializeTargets();
            InitializeQuestions();

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
        AddFormField(orderFormGrid, "Agent:", txtFinalCustomer, row, 0, 1);
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
            Text = "Configuration",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 16)
        };
        stackPanel.Children.Add(title);

        // ===== RADIATION SOURCE SECTION =====
        var radiationSourceBorder = new WpfBorder
        {
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0, 0, 0)), // Changed to black
            BorderThickness = new WpfThickness(1), // Thinner border
            CornerRadius = new CornerRadius(8),
            Padding = new WpfThickness(12),
            Margin = new WpfThickness(0, 0, 0, 16)
        };

        var radiationSourcePanel = new System.Windows.Controls.StackPanel();
        radiationSourceBorder.Child = radiationSourcePanel;

        var radiationSourceTitle = new System.Windows.Controls.TextBlock
        {
            Text = "Radiation Source",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 58, 138)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        radiationSourcePanel.Children.Add(radiationSourceTitle);

        // Grid for radiation source items - now in 3 columns layout
        var radiationGrid = new System.Windows.Controls.Grid();
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });
        radiationGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });

        int radRow = 0;
        int radCol = 0;

        foreach (var item in _radiationSourceItems)
        {
            if (radCol == 0)
            {
                radiationGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            }

            var checkBox = _configCheckBoxes[item];

            // Remove from any previous parent (Border or Panel)
            if (checkBox.Parent is WpfBorder oldBorder)
            {
                oldBorder.Child = null;
            }
            else if (checkBox.Parent is System.Windows.Controls.Panel oldPanel)
            {
                oldPanel.Children.Remove(checkBox);
            }

            // Wrap checkbox in a bordered box with reduced width
            var itemBorder = new WpfBorder
            {
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0, 0, 0)), // Black border
                BorderThickness = new WpfThickness(1), // Thinner border
                CornerRadius = new CornerRadius(6),
                Background = System.Windows.Media.Brushes.White,
                Padding = new WpfThickness(8, 5, 8, 5),
                Margin = new WpfThickness(0, 0, 6, 6),
                MinWidth = 110 // Reduced width
            };

            checkBox.Margin = new WpfThickness(0);
            checkBox.FontSize = 11; // Smaller text
            checkBox.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            checkBox.VerticalContentAlignment = System.Windows.VerticalAlignment.Center;
            itemBorder.Child = checkBox;

            System.Windows.Controls.Grid.SetRow(itemBorder, radRow);
            System.Windows.Controls.Grid.SetColumn(itemBorder, radCol);
            radiationGrid.Children.Add(itemBorder);

            radCol++;
            if (radCol >= 3)
            {
                radCol = 0;
                radRow++;
            }
        }

        radiationSourcePanel.Children.Add(radiationGrid);
        stackPanel.Children.Add(radiationSourceBorder);

        // ===== SYSTEM COMPONENTS SECTION =====
        var systemComponentsBorder = new WpfBorder
        {
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0, 0, 0)), // Changed to black
            BorderThickness = new WpfThickness(1), // Thinner border
            CornerRadius = new CornerRadius(8),
            Padding = new WpfThickness(12),
            Margin = new WpfThickness(0, 0, 0, 16)
        };

        var systemComponentsPanel = new System.Windows.Controls.StackPanel();
        systemComponentsBorder.Child = systemComponentsPanel;

        var systemComponentsTitle = new System.Windows.Controls.TextBlock
        {
            Text = "System Components",
            FontSize = 14,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 58, 138)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        systemComponentsPanel.Children.Add(systemComponentsTitle);

        // Grid for system components - 3 columns layout
        var componentsGrid = new System.Windows.Controls.Grid();
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });
        componentsGrid.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = GridLength.Auto });

        int row = 0;
        int col = 0;
        foreach (var item in _systemComponentsItems)
        {
            if (col == 0)
            {
                componentsGrid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
            }

            var checkBox = _configCheckBoxes[item];
            if (checkBox.Parent is WpfBorder oldBorder)
            {
                oldBorder.Child = null;
            }
            else if (checkBox.Parent is System.Windows.Controls.Panel oldParent)
            {
                oldParent.Children.Remove(checkBox);
            }

            // Wrap checkbox in a bordered box with reduced width
            var itemBorder = new WpfBorder
            {
                BorderBrush = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0, 0, 0)), // Black border
                BorderThickness = new WpfThickness(1), // Thinner border
                CornerRadius = new CornerRadius(6),
                Background = System.Windows.Media.Brushes.White,
                Padding = new WpfThickness(8, 5, 8, 5),
                Margin = new WpfThickness(0, 0, 6, 6),
                MinWidth = 150 // Reduced width for system components
            };

            checkBox.Margin = new WpfThickness(0);
            checkBox.FontSize = 11; // Smaller text
            checkBox.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            checkBox.VerticalContentAlignment = System.Windows.VerticalAlignment.Center;
            itemBorder.Child = checkBox;

            System.Windows.Controls.Grid.SetRow(itemBorder, row);
            System.Windows.Controls.Grid.SetColumn(itemBorder, col);
            componentsGrid.Children.Add(itemBorder);

            col++;
            if (col >= 3)
            {
                col = 0;
                row++;
            }
        }

        systemComponentsPanel.Children.Add(componentsGrid);
        stackPanel.Children.Add(systemComponentsBorder);

        // Custom configuration - compact
        var customLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Custom Configuration:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(customLabel);

        ApplyModernControlStyle(txtCustomConfig);
        txtCustomConfig.AcceptsReturn = true;
        txtCustomConfig.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtCustomConfig.MinHeight = 80;
        SafeAddChild(stackPanel, txtCustomConfig);

        panel.Children.Add(card);
    }

    private void LoadTargetsSection(System.Windows.Controls.Panel panel)
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
            Text = "Targets",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        stackPanel.Children.Add(title);

        // Add Target button - compact
        var addButton = new System.Windows.Controls.Button
        {
            Content = "+ Add Target",
            Margin = new WpfThickness(0, 0, 0, 12),
            HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
            Padding = new WpfThickness(12, 6, 12, 6),
            FontSize = 12
        };
        ApplyModernButtonStyle(addButton);
        addButton.Click += AddTarget_Click;
        stackPanel.Children.Add(addButton);

        // Targets ItemsControl
        targetsItemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0, 0, 0, 0)
        };

        // Create DataTemplate for targets
        var dataTemplate = new System.Windows.DataTemplate();
        var factory = new System.Windows.FrameworkElementFactory(typeof(WpfBorder));
        factory.SetValue(WpfBorder.BackgroundProperty, new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(248, 250, 252)));
        factory.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(8));
        factory.SetValue(WpfBorder.PaddingProperty, new WpfThickness(12));
        factory.SetValue(WpfBorder.MarginProperty, new WpfThickness(0, 0, 0, 8));

        var gridFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.Grid));

        // Create column definitions using AppendChild
        var col1 = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ColumnDefinition));
        col1.SetValue(System.Windows.Controls.ColumnDefinition.WidthProperty, new GridLength(1, GridUnitType.Star));

        var col2 = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ColumnDefinition));
        col2.SetValue(System.Windows.Controls.ColumnDefinition.WidthProperty, new GridLength(80));

        var col3 = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ColumnDefinition));
        col3.SetValue(System.Windows.Controls.ColumnDefinition.WidthProperty, new GridLength(2, GridUnitType.Star));

        var col4 = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.ColumnDefinition));
        col4.SetValue(System.Windows.Controls.ColumnDefinition.WidthProperty, GridLength.Auto);

        // Type TextBox
        var typeFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        typeFactory.SetValue(System.Windows.Controls.Grid.ColumnProperty, 0);
        typeFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty, 
            new System.Windows.Data.Binding("Type") { Mode = System.Windows.Data.BindingMode.TwoWay });
        typeFactory.SetValue(System.Windows.Controls.TextBox.MarginProperty, new WpfThickness(0, 0, 8, 0));
        gridFactory.AppendChild(typeFactory);

        // Qty TextBox
        var qtyFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        qtyFactory.SetValue(System.Windows.Controls.Grid.ColumnProperty, 1);
        qtyFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty, 
            new System.Windows.Data.Binding("Qty") { Mode = System.Windows.Data.BindingMode.TwoWay });
        qtyFactory.SetValue(System.Windows.Controls.TextBox.MarginProperty, new WpfThickness(0, 0, 8, 0));
        gridFactory.AppendChild(qtyFactory);

        // Details TextBox
        var detailsFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        detailsFactory.SetValue(System.Windows.Controls.Grid.ColumnProperty, 2);
        detailsFactory.SetBinding(System.Windows.Controls.TextBox.TextProperty, 
            new System.Windows.Data.Binding("Details") { Mode = System.Windows.Data.BindingMode.TwoWay });
        detailsFactory.SetValue(System.Windows.Controls.TextBox.MarginProperty, new WpfThickness(0, 0, 8, 0));
        gridFactory.AppendChild(detailsFactory);

        // Remove button
        var removeFactory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.Button));
        removeFactory.SetValue(System.Windows.Controls.Grid.ColumnProperty, 3);
        removeFactory.SetValue(System.Windows.Controls.Button.ContentProperty, "×");
        removeFactory.SetValue(System.Windows.Controls.Button.FontSizeProperty, 18.0);
        removeFactory.SetValue(System.Windows.Controls.Button.WidthProperty, 32.0);
        removeFactory.SetValue(System.Windows.Controls.Button.HeightProperty, 32.0);
        removeFactory.SetValue(System.Windows.Controls.Button.BackgroundProperty, 
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 68, 68)));
        removeFactory.SetValue(System.Windows.Controls.Button.ForegroundProperty, System.Windows.Media.Brushes.White);
        removeFactory.SetValue(System.Windows.Controls.Button.BorderThicknessProperty, new WpfThickness(0));
        removeFactory.SetValue(System.Windows.Controls.Button.TagProperty, new System.Windows.Data.Binding("."));
        removeFactory.AddHandler(System.Windows.Controls.Button.ClickEvent, 
            new RoutedEventHandler(RemoveTarget_Click));
        gridFactory.AppendChild(removeFactory);

        factory.AppendChild(gridFactory);
        dataTemplate.VisualTree = factory;
        targetsItemsControl.ItemTemplate = dataTemplate;
        targetsItemsControl.ItemsSource = _targets;

        stackPanel.Children.Add(targetsItemsControl);

        panel.Children.Add(card);
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
            Padding = new WpfThickness(12, 6, 12, 6),
            FontSize = 12
        };
        ApplyModernButtonStyle(addButton);
        addButton.Click += AddQuestion_Click;
        stackPanel.Children.Add(addButton);

        // Questions ItemsControl
        questionsItemsControl = new System.Windows.Controls.ItemsControl
        {
            Margin = new WpfThickness(0, 0, 0, 16)
        };

        // Create DataTemplate for questions with compact styling
        var dataTemplate = new System.Windows.DataTemplate();
        var factory = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.TextBox));
        factory.SetBinding(System.Windows.Controls.TextBox.TextProperty, 
            new System.Windows.Data.Binding("Text") { Mode = System.Windows.Data.BindingMode.TwoWay });
        factory.SetValue(System.Windows.Controls.TextBox.MarginProperty, new WpfThickness(0, 0, 0, 6));
        factory.SetValue(System.Windows.Controls.TextBox.PaddingProperty, new WpfThickness(8, 6, 8, 6));
        factory.SetValue(System.Windows.Controls.TextBox.FontSizeProperty, 13.0);
        factory.SetValue(System.Windows.Controls.TextBox.BorderBrushProperty,
            new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(226, 232, 240)));
        factory.SetValue(System.Windows.Controls.TextBox.BorderThicknessProperty, new WpfThickness(1));
        dataTemplate.VisualTree = factory;
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

        // Actions section (reused for marketing questions) - compact
        ApplyModernControlStyle(txtActions);
        txtActions.AcceptsReturn = true;
        txtActions.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtActions.MinHeight = 80;
        SafeAddChild(stackPanel, txtActions);

        panel.Children.Add(card);
    }

    private void LoadNotesSection(System.Windows.Controls.Panel panel)
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
            Text = "Notes",
            FontSize = 20,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(30, 41, 59)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        stackPanel.Children.Add(title);

        // Marketing Notes subtitle
        var marketingNotesTitle = new System.Windows.Controls.TextBlock
        {
            Text = "Marketing Notes",
            FontSize = 16,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(51, 65, 85)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        stackPanel.Children.Add(marketingNotesTitle);

        // Meeting Summary - compact
        var summaryLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Meeting Summary:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(summaryLabel);

        ApplyModernControlStyle(txtMeetingSummary);
        txtMeetingSummary.AcceptsReturn = true;
        txtMeetingSummary.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtMeetingSummary.MinHeight = 60;
        txtMeetingSummary.Margin = new WpfThickness(0, 0, 0, 12);
        SafeAddChild(stackPanel, txtMeetingSummary);

        // Key Decisions - compact
        var decisionsLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Key Decisions:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(decisionsLabel);

        ApplyModernControlStyle(txtDecisions);
        txtDecisions.AcceptsReturn = true;
        txtDecisions.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtDecisions.MinHeight = 60;
        txtDecisions.Margin = new WpfThickness(0, 0, 0, 12);
        SafeAddChild(stackPanel, txtDecisions);

        // Risks - compact
        var risksLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Risks / Special Requirements:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(risksLabel);

        ApplyModernControlStyle(txtRisks);
        txtRisks.AcceptsReturn = true;
        txtRisks.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtRisks.MinHeight = 60;
        txtRisks.Margin = new WpfThickness(0, 0, 0, 16);
        SafeAddChild(stackPanel, txtRisks);

        // PM Notes subtitle
        var pmNotesTitle = new System.Windows.Controls.TextBlock
        {
            Text = "PM Notes",
            FontSize = 16,
            FontWeight = System.Windows.FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(51, 65, 85)),
            Margin = new WpfThickness(0, 0, 0, 12)
        };
        stackPanel.Children.Add(pmNotesTitle);

        // Logistics - compact
        var logisticsLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Logistics:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(logisticsLabel);

        ApplyModernControlStyle(txtLogistics);
        txtLogistics.AcceptsReturn = true;
        txtLogistics.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtLogistics.MinHeight = 50;
        txtLogistics.Margin = new WpfThickness(0, 0, 0, 12);
        SafeAddChild(stackPanel, txtLogistics);

        // Software - compact
        var softwareLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Software:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 85)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(softwareLabel);

        ApplyModernControlStyle(txtSoftware);
        txtSoftware.AcceptsReturn = true;
        txtSoftware.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtSoftware.MinHeight = 50;
        txtSoftware.Margin = new WpfThickness(0, 0, 0, 12);
        SafeAddChild(stackPanel, txtSoftware);

        // Training - compact
        var trainingLabel = new System.Windows.Controls.TextBlock
        {
            Text = "Training:",
            FontSize = 12,
            FontWeight = System.Windows.FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(71, 85, 105)),
            Margin = new WpfThickness(0, 0, 0, 4)
        };
        stackPanel.Children.Add(trainingLabel);

        ApplyModernControlStyle(txtTraining);
        txtTraining.AcceptsReturn = true;
        txtTraining.TextWrapping = System.Windows.TextWrapping.Wrap;
        txtTraining.MinHeight = 50;
        SafeAddChild(stackPanel, txtTraining);

        panel.Children.Add(card);
    }

    private void InitializeConfigItems()
    {
        // Initialize Radiation Source items
        foreach (var item in _radiationSourceItems)
        {
            var checkBox = new WpfCheckBox
            {
                Content = item,
                FontSize = 12
            };
            _configCheckBoxes[item] = checkBox;
        }

        // Initialize System Components items
        foreach (var item in _systemComponentsItems)
        {
            var checkBox = new WpfCheckBox
            {
                Content = item,
                FontSize = 12
            };
            _configCheckBoxes[item] = checkBox;
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

        // ItemsSource will be set when LoadTargetsSection is called
    }

    private void InitializeQuestions()
    {
        _questions.Add(new QuestionItem { Text = "Define all target sizes" });
        _questions.Add(new QuestionItem { Text = "Confirm frame grabber compatibility" });
        _questions.Add(new QuestionItem { Text = "Confirm rack / table / logistics details" });

        // ItemsSource will be set when LoadActionsSection is called
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
            LoadSection(0); // Navigate to Overview section
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

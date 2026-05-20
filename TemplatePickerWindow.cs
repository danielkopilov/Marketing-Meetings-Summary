using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WpfBorder = System.Windows.Controls.Border;
using WpfThickness = System.Windows.Thickness;

namespace Marketing_Meetings_Summary;

/// <summary>
/// A restricted file-picker dialog that only shows the Templates folder.
/// No navigation outside that folder is possible.
/// </summary>
public class TemplatePickerWindow : Window
{
    public string? SelectedTemplatePath { get; private set; }

    private readonly string _templatesFolder;
    private ListBox _listBox = new();
    private List<string> _templateFiles = new();

    public TemplatePickerWindow(string templatesFolder)
    {
        _templatesFolder = templatesFolder;

        Title = "Load Template";
        Width = 500;
        Height = 380;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = System.Windows.Media.Brushes.White;

        BuildUI();
        RefreshList();
    }

    private void BuildUI()
    {
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // ── Header ─────────────────────────────────────────────────────────
        var header = new WpfBorder
        {
            Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(239, 246, 255)),
            Padding = new WpfThickness(16, 12, 16, 12),
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(0, 0, 0, 1)
        };

        var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
        headerStack.Children.Add(new TextBlock
        {
            Text = "📁  ",
            FontSize = 18,
            VerticalAlignment = VerticalAlignment.Center
        });
        headerStack.Children.Add(new TextBlock
        {
            Text = "Templates",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(37, 99, 235)),
            VerticalAlignment = VerticalAlignment.Center
        });
        header.Child = headerStack;
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // ── File list ───────────────────────────────────────────────────────
        _listBox = new ListBox
        {
            Margin = new WpfThickness(12, 10, 12, 6),
            FontSize = 14,
            BorderBrush = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(226, 232, 240)),
            BorderThickness = new WpfThickness(1)
        };
        System.Windows.Controls.ScrollViewer.SetHorizontalScrollBarVisibility(_listBox, System.Windows.Controls.ScrollBarVisibility.Disabled);
        _listBox.MouseDoubleClick += (s, e) => SelectAndClose();
        Grid.SetRow(_listBox, 1);
        root.Children.Add(_listBox);

        // ── Buttons ─────────────────────────────────────────────────────────
        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new WpfThickness(12, 6, 12, 12)
        };

        var loadBtn = MakeButton("Load", "#2563EB", "White");
        loadBtn.Click += (s, e) => SelectAndClose();

        var cancelBtn = MakeButton("Cancel", "White", "#374151");
        cancelBtn.BorderBrush = new System.Windows.Media.SolidColorBrush(
            System.Windows.Media.Color.FromRgb(209, 213, 219));
        cancelBtn.BorderThickness = new WpfThickness(1);
        cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };

        btnRow.Children.Add(cancelBtn);
        btnRow.Children.Add(loadBtn);
        Grid.SetRow(btnRow, 2);
        root.Children.Add(btnRow);

        Content = root;
    }

    private static Button MakeButton(string label, string bgHex, string fgHex)
    {
        var btn = new Button
        {
            Content = label,
            Width = 90,
            Height = 32,
            FontSize = 13,
            FontWeight = FontWeights.SemiBold,
            Cursor = Cursors.Hand,
            Margin = new WpfThickness(6, 0, 0, 0),
            BorderThickness = new WpfThickness(0)
        };

        bool isBgNamedColor = !bgHex.StartsWith("#");
        btn.Background = isBgNamedColor
            ? (System.Windows.Media.Brush)System.Windows.Media.Brushes.White
            : new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(bgHex));

        btn.Foreground = new System.Windows.Media.SolidColorBrush(
            (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(fgHex));

        // Rounded template
        var tmpl = new ControlTemplate(typeof(Button));
        var bf = new FrameworkElementFactory(typeof(WpfBorder));
        bf.SetValue(WpfBorder.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
        bf.SetValue(WpfBorder.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
        bf.SetValue(WpfBorder.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
        bf.SetValue(WpfBorder.CornerRadiusProperty, new CornerRadius(8));
        var cf = new FrameworkElementFactory(typeof(ContentPresenter));
        cf.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        cf.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
        bf.AppendChild(cf);
        tmpl.VisualTree = bf;
        btn.Template = tmpl;

        return btn;
    }

    private void RefreshList()
    {
        _listBox.Items.Clear();
        _templateFiles.Clear();

        if (!Directory.Exists(_templatesFolder))
        {
            _listBox.Items.Add(new ListBoxItem
            {
                Content = "No templates found.",
                IsEnabled = false,
                Foreground = System.Windows.Media.Brushes.Gray
            });
            return;
        }

        var files = Directory.GetFiles(_templatesFolder, "*.json")
            .OrderByDescending(f => File.GetLastWriteTime(f))
            .ToList();

        if (files.Count == 0)
        {
            _listBox.Items.Add(new ListBoxItem
            {
                Content = "No templates found.",
                IsEnabled = false,
                Foreground = System.Windows.Media.Brushes.Gray
            });
            return;
        }

        foreach (var file in files)
        {
            string name     = Path.GetFileNameWithoutExtension(file);
            var    lastSave = File.GetLastWriteTime(file);

            var item = new ListBoxItem
            {
                Tag     = file,
                Padding = new WpfThickness(10, 6, 10, 6)
            };

            var stack = new StackPanel { Orientation = Orientation.Horizontal };
            stack.Children.Add(new TextBlock
            {
                Text = "📄  ",
                FontSize = 14,
                VerticalAlignment = VerticalAlignment.Center
            });
            var namePanel = new StackPanel();
            namePanel.Children.Add(new TextBlock
            {
                Text = name,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(30, 41, 59))
            });
            namePanel.Children.Add(new TextBlock
            {
                Text = $"Last saved: {lastSave:dd/MM/yyyy  HH:mm}",
                FontSize = 11,
                Foreground = System.Windows.Media.Brushes.Gray
            });
            stack.Children.Add(namePanel);
            item.Content = stack;

            _listBox.Items.Add(item);
            _templateFiles.Add(file);
        }
    }

    private void SelectAndClose()
    {
        if (_listBox.SelectedItem is ListBoxItem selected && selected.Tag is string path)
        {
            SelectedTemplatePath = path;
            DialogResult = true;
            Close();
        }
        else
        {
            MessageBox.Show("Please select a template first.", "No Selection",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}

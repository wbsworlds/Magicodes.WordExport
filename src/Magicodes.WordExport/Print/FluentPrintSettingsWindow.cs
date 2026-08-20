using System;
using System.Collections.Generic;
using System.Linq;
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using WpfButton = System.Windows.Controls.Button;
using WpfComboBox = System.Windows.Controls.ComboBox;
using WpfRadioButton = System.Windows.Controls.RadioButton;
using WpfTextBox = System.Windows.Controls.TextBox;
using WpfTextBlock = System.Windows.Controls.TextBlock;
using WpfMessageBox = System.Windows.MessageBox;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfFontFamily = System.Windows.Media.FontFamily;
using WpfCursors = System.Windows.Input.Cursors;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using Orientation = System.Windows.Controls.Orientation;
using VerticalAlignment = System.Windows.VerticalAlignment;

namespace Magicodes.WordExport.Print;

/// <summary>
/// 用户选择：打印机 / 方向 / 份数 / 页码范围。
/// </summary>
public sealed record PrintSettings(
    PrintQueue PrintQueue,
    PageOrientation Orientation,
    int Copies,
    int FromPage,
    int ToPage,
    bool AllPages);

/// <summary>
/// 自绘打印设置窗口。
/// </summary>
public sealed class FluentPrintSettingsWindow : Window
{
    private readonly int _totalPages;
    private readonly PageOrientation _defaultOrientation;
    private readonly WpfComboBox _printerBox;
    private readonly WpfComboBox _orientationBox;
    private readonly IntegerUpDown _copiesBox;
    private readonly WpfRadioButton _allPagesRadio;
    private readonly WpfRadioButton _rangeRadio;
    private readonly IntegerUpDown _fromBox;
    private readonly IntegerUpDown _toBox;
    private readonly Border _fromBorder;
    private readonly Border _toBorder;

    public PrintSettings? Settings { get; private set; }

    private const string PageRangeGroup = "PageRange";
    private static readonly WpfColor PrimaryColor = WpfColor.FromRgb(0, 122, 204);
    private static readonly WpfColor TextColor = WpfColor.FromRgb(51, 51, 51);
    private static readonly WpfColor SecondaryTextColor = WpfColor.FromRgb(102, 102, 102);
    private static readonly WpfColor BorderColor = WpfColor.FromRgb(210, 210, 210);

    public FluentPrintSettingsWindow(int totalPages, PageOrientation defaultOrientation)
    {
        _totalPages = totalPages;
        _defaultOrientation = defaultOrientation;

        Title = "打印";
        Width = 420;
        Height = 380;
        MinWidth = 380;
        MinHeight = 360;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.CanResize;
        WindowStyle = WindowStyle.SingleBorderWindow;
        Background = WpfBrushes.White;
        FontFamily = new WpfFontFamily("Microsoft YaHei UI");
        FontSize = 12;
        Foreground = new SolidColorBrush(TextColor);

        var root = new Grid
        {
            Margin = new Thickness(20, 16, 20, 12),
            RowDefinitions =
            {
                new RowDefinition { Height = new GridLength(1, GridUnitType.Star) },
                new RowDefinition { Height = GridLength.Auto },
            }
        };

        var panel = new StackPanel();
        Grid.SetRow(panel, 0);

        var server = new LocalPrintServer();
        var queues = server.GetPrintQueues(
                new[] { EnumeratedPrintQueueTypes.Local, EnumeratedPrintQueueTypes.Connections })
            .OrderBy(q => q.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        _printerBox = CreateComboBox();
        _printerBox.ItemsSource = queues;
        _printerBox.DisplayMemberPath = nameof(PrintQueue.Name);
        var defaultQueue = queues.FirstOrDefault(q => q.Name == server.DefaultPrintQueue?.Name)
                           ?? queues.FirstOrDefault();
        if (defaultQueue != null) _printerBox.SelectedItem = defaultQueue;

        _orientationBox = CreateComboBox();
        _orientationBox.ItemsSource = new[]
        {
            new { Label = "纵向", Value = PageOrientation.Portrait },
            new { Label = "横向", Value = PageOrientation.Landscape },
        };
        _orientationBox.DisplayMemberPath = "Label";
        _orientationBox.SelectedValuePath = "Value";
        _orientationBox.SelectedValue = defaultOrientation;

        _copiesBox = new IntegerUpDown(1, 1, 999, 70);
        _copiesBox.Height = 26;

        BuildSection(panel, "打印机", _printerBox);
        BuildSection(panel, "方向", _orientationBox);
        BuildCompactSection(panel, "份数", _copiesBox);

        panel.Children.Add(new Separator
        {
            Margin = new Thickness(0, 2, 0, 2),
            Background = new SolidColorBrush(WpfColor.FromRgb(235, 235, 235)),
        });

        var rangeLabel = new WpfTextBlock
        {
            Text = "页码范围",
            FontWeight = FontWeights.Medium,
            Foreground = new SolidColorBrush(TextColor),
            Margin = new Thickness(0, 8, 0, 6),
        };
        panel.Children.Add(rangeLabel);

        _allPagesRadio = CreateRadio($"所有页（共 {totalPages} 页）", PageRangeGroup);
        _allPagesRadio.IsChecked = true;
        _allPagesRadio.Checked += (_, _) => SetRangeEnabled(false);
        _allPagesRadio.Unchecked += (_, _) => SetRangeEnabled(true);
        panel.Children.Add(_allPagesRadio);

        var customRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Height = 30,
            Margin = new Thickness(0, 4, 0, 0),
        };

        _rangeRadio = CreateRadio("自定义：", PageRangeGroup);
        _rangeRadio.VerticalAlignment = VerticalAlignment.Center;
        customRow.Children.Add(_rangeRadio);

        _fromBox = new IntegerUpDown(1, 1, totalPages, 55)
        {
            Height = 24,
            VerticalAlignment = VerticalAlignment.Center,
        };

        _fromBorder = new Border
        {
            Child = _fromBox,
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(BorderColor),
            CornerRadius = new CornerRadius(3),
            Margin = new Thickness(12, 0, 0, 0),
            Height = 24,
        };
        customRow.Children.Add(_fromBorder);

        var dashText = new WpfTextBlock
        {
            Text = "—",
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 4, 0),
            Foreground = new SolidColorBrush(SecondaryTextColor),
            FontSize = 12,
        };
        customRow.Children.Add(dashText);

        _toBox = new IntegerUpDown(totalPages, 1, totalPages, 55)
        {
            Height = 24,
            VerticalAlignment = VerticalAlignment.Center,
        };

        _toBorder = new Border
        {
            Child = _toBox,
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(BorderColor),
            CornerRadius = new CornerRadius(3),
            Height = 24,
        };
        customRow.Children.Add(_toBorder);

        panel.Children.Add(customRow);

        SetRangeEnabled(false);

        root.Children.Add(panel);

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 12, 0, 0),
        };
        Grid.SetRow(btnPanel, 1);

        var cancelBtn = CreateButton("取消", 80, 30, false);
        cancelBtn.Click += (_, _) => { DialogResult = false; Close(); };

        var okBtn = CreateButton("打印", 80, 30, true);
        okBtn.Click += (_, _) =>
        {
            if (_printerBox.SelectedItem is not PrintQueue queue)
            {
                WpfMessageBox.Show(this, "请选择打印机", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            int copies = Math.Clamp(_copiesBox.Value, 1, 999);
            int from = 1, to = totalPages;
            bool allPages = _allPagesRadio.IsChecked == true;
            if (!allPages)
            {
                from = Math.Clamp(_fromBox.Value, 1, totalPages);
                to = Math.Clamp(_toBox.Value, 1, totalPages);
                if (from > to) (from, to) = (to, from);
            }
            Settings = new PrintSettings(
                queue,
                (PageOrientation)(_orientationBox.SelectedValue ?? PageOrientation.Portrait),
                copies, from, to, allPages);
            DialogResult = true;
            Close();
        };
        btnPanel.Children.Add(cancelBtn);
        btnPanel.Children.Add(okBtn);
        root.Children.Add(btnPanel);

        Content = root;
    }

    private void SetRangeEnabled(bool enabled)
    {
        _fromBox.IsEnabled = enabled;
        _toBox.IsEnabled = enabled;
        _fromBorder.Opacity = enabled ? 1.0 : 0.5;
        _toBorder.Opacity = enabled ? 1.0 : 0.5;
    }

    private static void BuildSection(StackPanel panel, string label, FrameworkElement control)
    {
        var labelBlock = new WpfTextBlock
        {
            Text = label,
            FontWeight = FontWeights.Medium,
            Foreground = new SolidColorBrush(TextColor),
            Margin = new Thickness(0, 0, 0, 4),
        };
        panel.Children.Add(labelBlock);

        var border = new Border
        {
            Child = control,
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(BorderColor),
            CornerRadius = new CornerRadius(3),
            Background = WpfBrushes.White,
            Margin = new Thickness(0, 0, 0, 12),
        };

        if (control is WpfComboBox combo)
        {
            combo.BorderThickness = new Thickness(0);
            combo.Background = WpfBrushes.Transparent;
            combo.Margin = new Thickness(6, 3, 6, 3);
            combo.Height = 26;
        }

        panel.Children.Add(border);
    }

    private static void BuildCompactSection(StackPanel panel, string label, FrameworkElement control)
    {
        var row = new Grid
        {
            Margin = new Thickness(0, 0, 0, 10),
        };
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var labelBlock = new WpfTextBlock
        {
            Text = label,
            FontWeight = FontWeights.Medium,
            Foreground = new SolidColorBrush(TextColor),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 12, 0),
        };
        Grid.SetColumn(labelBlock, 0);
        row.Children.Add(labelBlock);

        var border = new Border
        {
            Child = control,
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(BorderColor),
            CornerRadius = new CornerRadius(3),
            Background = WpfBrushes.White,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Left,
        };
        Grid.SetColumn(border, 1);

        if (control is IntegerUpDown upDown)
        {
            upDown.Margin = new Thickness(6, 1, 6, 1);
        }

        row.Children.Add(border);
        panel.Children.Add(row);
    }

    private static WpfComboBox CreateComboBox()
    {
        return new WpfComboBox
        {
            Height = 28,
            Padding = new Thickness(6, 3, 6, 3),
            FontSize = 12,
        };
    }

    private static WpfRadioButton CreateRadio(string content, string groupName)
    {
        return new WpfRadioButton
        {
            Content = content,
            GroupName = groupName,
            Height = 26,
            FontSize = 12,
            Foreground = new SolidColorBrush(TextColor),
        };
    }

    private static WpfButton CreateButton(string content, double width, double height, bool isPrimary)
    {
        return new WpfButton
        {
            Content = content,
            Width = width,
            Height = height,
            Margin = new Thickness(8, 0, 0, 0),
            IsDefault = isPrimary,
            FontSize = 12,
            FontWeight = isPrimary ? FontWeights.SemiBold : FontWeights.Normal,
            Cursor = WpfCursors.Hand,
            Background = isPrimary
                ? new SolidColorBrush(PrimaryColor)
                : WpfBrushes.White,
            Foreground = isPrimary
                ? WpfBrushes.White
                : new SolidColorBrush(TextColor),
            BorderBrush = isPrimary
                ? new SolidColorBrush(PrimaryColor)
                : new SolidColorBrush(BorderColor),
            BorderThickness = new Thickness(1),
        };
    }
}

/// <summary>
/// 数字增减控件（TextBox + 上下箭头按钮）。
/// </summary>
internal sealed class IntegerUpDown : Grid
{
    private int _value;
    private readonly WpfTextBox _box;
    private readonly WpfButton _upBtn;
    private readonly WpfButton _dnBtn;

    public int Value
    {
        get => _value;
        set
        {
            _value = Math.Clamp(value, Min, Max);
            if (_box.Text != _value.ToString())
                _box.Text = _value.ToString();
        }
    }

    public int Min { get; set; }
    public int Max { get; set; }

    public IntegerUpDown(int initialValue, int min, int max, double width = 120)
    {
        Min = min;
        Max = max;
        Width = width;

        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        _box = new WpfTextBox
        {
            VerticalContentAlignment = VerticalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Stretch,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            TextAlignment = TextAlignment.Center,
            BorderThickness = new Thickness(0),
            Background = WpfBrushes.Transparent,
            Padding = new Thickness(2, 0, 2, 0),
            Text = initialValue.ToString(),
            CaretIndex = initialValue.ToString().Length,
            FontSize = 12,
        };
        SetColumn(_box, 0);
        Children.Add(_box);

        _value = Math.Clamp(initialValue, min, max);
        _box.Text = _value.ToString();
        _box.LostFocus += (_, _) =>
        {
            if (int.TryParse(_box.Text, out var v)) Value = v;
            else _box.Text = _value.ToString();
        };
        _box.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter)
            {
                if (int.TryParse(_box.Text, out var v)) Value = v;
                else _box.Text = _value.ToString();
            }
        };

        var btnStack = new StackPanel
        {
            Orientation = Orientation.Vertical,
            VerticalAlignment = VerticalAlignment.Stretch,
        };
        SetColumn(btnStack, 1);

        _upBtn = new WpfButton
        {
            Content = "▲",
            FontSize = 6,
            Padding = new Thickness(0),
            BorderThickness = new Thickness(0),
            Background = WpfBrushes.Transparent,
            Width = 14,
            Cursor = WpfCursors.Hand,
            VerticalAlignment = VerticalAlignment.Top,
            Height = 12,
        };
        _upBtn.Click += (_, _) => Value++;
        btnStack.Children.Add(_upBtn);

        _dnBtn = new WpfButton
        {
            Content = "▼",
            FontSize = 6,
            Padding = new Thickness(0),
            BorderThickness = new Thickness(0),
            Background = WpfBrushes.Transparent,
            Width = 14,
            Cursor = WpfCursors.Hand,
            VerticalAlignment = VerticalAlignment.Bottom,
            Height = 12,
        };
        _dnBtn.Click += (_, _) => Value--;
        btnStack.Children.Add(_dnBtn);

        Children.Add(btnStack);

        IsEnabledChanged += (_, e) =>
        {
            var enabled = (bool)e.NewValue;
            _box.IsEnabled = enabled;
            _upBtn.IsEnabled = enabled;
            _dnBtn.IsEnabled = enabled;
        };
    }
}

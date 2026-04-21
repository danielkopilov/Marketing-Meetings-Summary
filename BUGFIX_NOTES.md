# Bug Fix - Tab Switching Crashes

## Issues Fixed

### 1. **Questions Data Binding Issue**
**Problem:** TextBoxes were bound directly to strings in an `ObservableCollection<string>`. WPF cannot properly handle two-way binding with primitive types in collections, causing crashes when editing question text.

**Solution:** Created a new `QuestionItem` class that implements `INotifyPropertyChanged` with a `Text` property. This allows proper two-way data binding.

```csharp
public class QuestionItem : INotifyPropertyChanged
{
    private string _text = "";
    public string Text
    {
        get => _text;
        set { _text = value; OnPropertyChanged(nameof(Text)); }
    }
}
```

### 2. **Unnecessary Event Handlers**
**Problem:** Empty event handlers (`checkBox.Checked += (s, e) => { };`) were added but served no purpose and could potentially cause issues.

**Solution:** Removed all empty event handlers that were left over from the preview panel removal.

### 3. **Missing Error Handling**
**Problem:** Button click handlers had no try-catch blocks, so any exception would crash the application.

**Solution:** Added comprehensive try-catch blocks around:
- `AddTarget_Click()`
- `RemoveTarget_Click()`
- `AddQuestion_Click()`

### 4. **XAML Binding Configuration**
**Problem:** The questions TextBox binding didn't specify binding mode or update trigger.

**Solution:** Updated XAML binding to:
```xml
<TextBox Text="{Binding Text, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
```

## Testing Recommendations

1. Switch between all tabs multiple times
2. Add and remove targets
3. Add and remove questions
4. Edit text in question fields
5. Fill in various form fields while switching tabs
6. Generate a Word document to ensure all data is captured correctly

## Files Modified

- `MainWindow.xaml.cs` - Added QuestionItem class, updated collection types, added error handling
- `MainWindow.xaml` - Updated questions binding to use the Text property

## Build Status

✅ Debug build: Successful
✅ Release build: Successful

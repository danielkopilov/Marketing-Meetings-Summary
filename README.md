# Marketing Kickoff Meeting Summary

A modern WPF desktop application built with .NET 10 for creating structured meeting summaries for marketing and new-project kickoff meetings. The application provides a clean, professional interface with tabbed navigation and automated Word document generation.

![.NET Version](https://img.shields.io/badge/.NET-10.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

## 🎯 Features

- **Modern Tab-Based Interface**: Clean, professional UI with 5 organized tabs
- **Automated Word Document Generation**: Export meeting summaries to properly formatted .docx files
- **Dynamic Form Fields**: Add/remove targets and questions as needed
- **Configuration Checklist**: Pre-populated standard configuration items
- **Data Validation**: Ensures required fields are filled before export
- **Professional Styling**: Rounded corners, modern colors, and intuitive layout

## 📋 Tabs Overview

### 1. 📋 Overview
- Order Number (required)
- Customer & Final Customer information
- System/Project Type
- Meeting Date, Delivery Date, Design Due Date
- PAKA number
- Participants
- Reference Order

### 2. ⚙️ Configuration
- Standard configuration items (checkboxes):
  - Motorized Target Wheel
  - Motorized Source Stage
  - Motorized Focus Stage
  - Integrating Sphere
  - Blackbody / SR800N
  - PC + CTE
  - Frame Grabber
  - Installation + Training
  - Rackmount / PDU
  - Optical Table
- Additional configuration details (free text)

### 3. 🎯 Targets
- Dynamic target list with Type, Quantity, and Details
- Pre-populated with common target types: 4Bar, Pin Hole, Square, Cross, Step, USAF, Boresight, LOS Alignment
- Add/Remove targets as needed

### 4. ❗ Actions
- Open Questions (dynamic list)
- Action Items (free text)

### 5. 📝 Notes
- Meeting Summary
- Key Decisions
- Risks / Special Requirements
- Logistics
- Software / FG / CTE
- Installation / Training

## 🚀 Getting Started

### Prerequisites

- Windows 10/11
- .NET 10 Runtime (or SDK for development)

### Installation

1. **Download the latest release** from the [Releases](../../releases) page
2. **Extract the ZIP file** to your preferred location
3. **Run** `Marketing Meetings Summary.exe`

### Building from Source

```powershell
# Clone the repository
git clone https://github.com/YOUR_USERNAME/marketing-meeting-summary.git
cd marketing-meeting-summary

# Build the project
dotnet build -c Release

# Run the application
dotnet run
```

## 💾 Usage

1. **Fill in the meeting details** across the various tabs
2. **Required field**: Order Number must be filled
3. **Add dynamic items**: Use the "+ Add Target" and "+ Add Question" buttons as needed
4. **Export**: Click "📄 Export Summary" to generate a Word document
5. **Save**: Choose your save location and filename
6. **Open**: Optionally open the generated document immediately

## 📁 Output Format

The generated Word document includes:
- **Header**: Order number, customer information
- **System details**: Project type, dates
- **Configuration**: Bulleted list of selected items
- **Targets**: Formatted list with quantities and specifications
- **PAKA number**: Bold and underlined
- **Questions**: Red text section for open items
- **Summary sections**: All notes and decisions formatted clearly

## 🛠️ Technical Details

- **Framework**: .NET 10, WPF (Windows Presentation Foundation)
- **Language**: C# with implicit usings
- **Document Generation**: DocumentFormat.OpenXml 3.2.0
- **Architecture**: MVVM-lite pattern with ObservableCollections
- **Data Binding**: Two-way binding with INotifyPropertyChanged

## 🔧 Dependencies

- `DocumentFormat.OpenXml` (v3.2.0) - For .docx file creation

## 📝 Recent Updates

### Version 1.1 (Latest)
- ✅ Fixed tab switching crashes
- ✅ Improved data binding stability with QuestionItem class
- ✅ Added comprehensive error handling
- ✅ Removed live preview panel for cleaner interface
- ✅ Narrower tabs and textboxes for better layout
- ✅ Updated title to "Marketing Kickoff meeting Summary"

### Version 1.0
- Initial release with tab-based interface
- Word document generation
- Dynamic targets and questions

## 🐛 Known Issues

None currently. Please report issues on the [Issues](../../issues) page.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👤 Author

Created for streamlining marketing kickoff meeting documentation and project initiation workflows.

## 🙏 Acknowledgments

- Built with modern WPF best practices
- Designed based on real-world meeting documentation needs
- Inspired by contemporary web application aesthetics

---

**Note**: The MeetingNotes folder (containing sample meeting documents) is excluded from the repository for privacy reasons.

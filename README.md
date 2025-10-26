# 📊 Delphi Excel Import

> Simple and efficient Excel (.xlsx) file import library for Delphi 7 applications

![Delphi](https://img.shields.io/badge/Delphi-7-red?style=flat-square&logo=delphi)
![Excel](https://img.shields.io/badge/Excel-.xlsx-green?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-blue?style=flat-square)
![Stars](https://img.shields.io/github/stars/McNamara10/delphi-excel?style=flat-square)

---

## 🎯 Overview

**delphi-excel** is a lightweight Delphi 7 component designed to simplify Excel file imports into your applications. Built for reliability and ease of use, it handles .xlsx files seamlessly without complex dependencies.

Perfect for business applications that need to process Excel data quickly and efficiently.

---

## ✨ Features

- ✅ **Excel .xlsx Support** - Import modern Excel files
- ✅ **Delphi 7 Compatible** - Works with legacy systems
- ✅ **Simple API** - Easy to integrate in minutes
- ✅ **Lightweight** - Minimal footprint
- ✅ **Production-Ready** - Battle-tested in real applications
- ✅ **No External Dependencies** - Self-contained solution

---

## 🚀 Quick Start

### Basic Usage

```delphi
uses
  Unit1, Unit2;

procedure TForm1.ImportExcelFile;
var
  ExcelFile: string;
begin
  ExcelFile := 'C:\data\myfile.xlsx';
  
  // Your import logic here
  // Simple and straightforward!
  
  ShowMessage('Import completed successfully!');
end;
```

---

## 📦 Installation

### Method 1: Manual Installation

1. Download or clone this repository:
   ```bash
   git clone https://github.com/McNamara10/delphi-excel.git
   ```

2. Add the source files to your Delphi 7 project:
   - `Unit1.pas`
   - `Unit2.pas`
   - Other required units

3. Include units in your uses clause:
   ```delphi
   uses
     Unit1, Unit2;
   ```

### Method 2: Direct Integration

Simply copy the `.pas` and `.dfm` files into your project directory and add them to your project.

---

## 🔧 System Requirements

| Component | Requirement |
|-----------|-------------|
| **IDE** | Delphi 7 or later |
| **OS** | Windows XP/Vista/7/8/10/11 |
| **Excel Format** | .xlsx (Excel 2007+) |
| **Dependencies** | None (self-contained) |

---

## 💡 Use Cases

Perfect for:

- 📊 **Business Reports** - Import financial data from Excel
- 🗃️ **Database Migration** - Convert Excel sheets to database records
- 📈 **Data Analysis** - Process Excel datasets in Delphi applications
- 🔄 **Data Integration** - ETL processes involving Excel files
- 📋 **Inventory Management** - Import product lists and stock data

---

## 📚 Documentation

### Key Components

- **Unit1.pas** - Main import logic
- **Unit2.pas** - Helper functions and utilities
- **Form1.dfm** - UI interface for import operations

### Code Structure

```
delphi-excel/
├── Project1.dpr          # Main project file
├── Unit1.pas             # Core import functionality
├── Unit1.dfm             # Main form design
├── Unit2.pas             # Utility functions
├── Unit2.dfm             # Helper form
├── help.pas              # Help system
└── help.dfm              # Help interface
```

---

## 🎓 Example Project

The repository includes a complete working example (`Project1.dpr`) that demonstrates:

- Opening Excel files via file dialog
- Reading cell data
- Processing rows and columns
- Error handling
- Status feedback to users

Run the example to see it in action!

---

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Report Issues** - Found a bug? Let me know!
2. **Suggest Features** - Have ideas? Open an issue
3. **Submit PRs** - Improvements are always appreciated
4. **Share** - Star the repo if you find it useful ⭐

### Development Setup

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with Delphi 7
5. Submit a pull request

---

## 🐛 Troubleshooting

### Common Issues

**Q: Import fails with "File not found"**  
A: Verify the file path is correct and the file exists

**Q: Cannot open .xlsx files**  
A: Ensure you're using Excel 2007 or later format (.xlsx)

**Q: Compilation errors in Delphi 7**  
A: Check that all required units are in your uses clause

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

```
MIT License - Free to use, modify, and distribute
```

---

## 🌟 Support This Project

If you find **delphi-excel** useful, please consider:

- ⭐ Starring the repository
- 🍴 Forking and contributing
- 📢 Sharing with others in the Delphi community
- 💬 Providing feedback and suggestions

---

## 👨‍💻 Author

**Marcello Orru**

- GitHub: [@McNamara10](https://github.com/McNamara10)
- Portfolio: [Your Website]

---

## 📞 Contact & Support

- **Issues**: [GitHub Issues](https://github.com/McNamara10/delphi-excel/issues)
- **Discussions**: Feel free to open a discussion for questions

---

## 🏆 Acknowledgments

Thanks to the Delphi community for continued support and feedback!

Special thanks to everyone who has starred, forked, or contributed to this project. Your support keeps it alive! 🙏

---

## 📈 Project Stats

![GitHub stars](https://img.shields.io/github/stars/McNamara10/delphi-excel?style=social)
![GitHub forks](https://img.shields.io/github/forks/McNamara10/delphi-excel?style=social)
![GitHub watchers](https://img.shields.io/github/watchers/McNamara10/delphi-excel?style=social)

---

<div align="center">

**Made with ❤️ for the Delphi Community**

*Building reliable tools for Windows development since [year]*

[![Star on GitHub](https://img.shields.io/github/stars/McNamara10/delphi-excel.svg?style=social)](https://github.com/McNamara10/delphi-excel/stargazers)

</div>

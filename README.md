# 🏪 VBA POS - Community Edition

**Free & Open Source Point of Sale System for Excel**

[![License](https://img.shields.io/badge/License-GPL%20v2-blue.svg)](LICENSE)
[![Excel](https://img.shields.io/badge/Excel-2016%2B-green.svg)](https://www.microsoft.com/excel)
[![Database](https://img.shields.io/badge/Database-SQLite-orange.svg)](https://www.sqlite.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

---

## 📋 Overview

**VBA POS Community Edition** is a complete, production-ready Point of Sale (POS) system built entirely in Excel VBA with SQLite database backend. Perfect for small businesses, kiosks, and retail stores.

### ✨ Key Features

- ✅ **Complete POS System** - Sales, inventory, customers, employees
- ✅ **SQLite Database** - Fast, reliable, no server required
- ✅ **100% Free** - GPL v2 open source license
- ✅ **No Internet Required** - Works completely offline
- ✅ **Easy to Use** - Familiar Excel interface
- ✅ **Customizable** - Full source code access
- ✅ **Multi-language** - Spanish/English support
- ✅ **Professional** - Invoice printing, reports, analytics

### 🎯 Perfect For

- 🏪 Small retail stores
- ☕ Coffee shops and cafes
- 🛒 Mini markets and kiosks
- 📦 Warehouse management
- 🎨 Artisan shops
- 🍕 Small restaurants (basic POS)

---

## 🚀 Quick Start

### System Requirements

- **Excel:** 2016 or later (Windows only)
- **OS:** Windows 10/11 (64-bit recommended)
- **RAM:** 2GB minimum, 4GB recommended
- **Storage:** 50MB for application + database
- **Dependencies:** SQLite ODBC Driver

### Installation (5 Minutes)

1. **Download Project**
   ```
   Code > Download ZIP
   ```

2. **Extract Files**
   - Extract ZIP to desired folder (e.g., `C:\POS\`)
   - Files needed: `pos-excel-sqlite.xlsm` + `DBVentas.db`

3. **Install SQLite ODBC Driver**
   - Download from: http://www.ch-werner.de/sqliteodbc/
   - Install 32-bit or 64-bit (match your Excel version)
   - Run installer and complete setup

4. **Configure Database**
   - If using multi-workstation: Place `DBVentas.db` in shared network folder
   - If single workstation: Keep `DBVentas.db` with Excel file

5. **Open Application**
   - Double-click `pos-excel-sqlite.xlsm`
   - Click "Enable Content" when prompted (macros required)
   - First time: Browse and select `DBVentas.db` location

6. **Login**
   - Default email: `admin@gmail.com`
   - Default password: `admin`
   - **⚠️ Change password immediately after first login!**

### First Sale in 2 Minutes

1. Open POS → Sales
2. Scan/enter product barcode
3. Enter customer DNI (or use default)
4. Click "Complete Sale" (F3)
5. Print receipt ✅

---

## 📚 Core Features

### 1. 💰 Sales Management

**Features:**
- Fast barcode scanning
- Product search by name/code
- Customer lookup by DNI
- Discount application
- Multiple payment methods (cash, card, transfer)
- Receipt printing
- Sale returns/refunds
- Daily sales reports

**Keyboard Shortcuts:**
- `F2` - Search product
- `F3` - Complete sale
- `F5` - New sale
- `ESC` - Cancel

### 2. 📦 Inventory Management

**Features:**
- Product CRUD (Create, Read, Update, Delete)
- Categories and units
- Stock tracking
- Low stock alerts
- Product search and filtering
- Barcode support
- Cost and price management
- Stock adjustments

**Reports:**
- Current inventory
- Stock movements
- Low stock products
- Product valuation
- Sales by product

### 3. 👥 Customer Management

**Features:**
- Customer database (name, DNI, phone, email, address)
- Purchase history
- Customer search
- Customer loyalty tracking
- Customer reports

**Data Collected:**
- Personal info (DNI, name, phone)
- Purchase history
- Total spent
- Last purchase date

### 4. 👔 Employee Management

**Features:**
- Employee registration
- Role-based permissions (Admin, Cashier, Warehouse)
- Sales tracking by employee
- Login/logout tracking
- Performance reports

**Roles:**
- **Admin:** Full access
- **Cashier:** Sales, customers
- **Warehouse:** Inventory only

### 5. 📊 Reports & Analytics

**Built-in Reports:**
- Daily sales summary
- Sales by date range
- Sales by product
- Sales by customer
- Sales by employee
- Inventory valuation
- Low stock report
- Profit & loss (basic)

**Export Options:**
- Excel
- PDF (if installed)
- Print

---

## 🗄️ Database Schema

**Technology:** SQLite 3

### Core Tables

```sql
-- Products
products (idProduct, barcode, product, cost, price, stock, idCategory, idUnit, idState)
categories (idCategory, category, idState)
units (idUnit, unit, idState)

-- Sales
sales (idSale, date, total, idCustomer, idEmployee, idState)
saleDetails (idSaleDetail, idSale, idProduct, quantity, price, subtotal)

-- Customers
customers (idCustomer, dni, name, surname, phone, email, address, idState)

-- Employees
employees (idEmployee, dni, name, surname, phone, email, address, role, idState)

-- System
states (idState, state) -- Active, Inactive, Deleted
```

**Total Tables:** 8 core tables

**idState Convention:**
- `1` = Active
- `2` = Inactive
- `3` = Deleted (soft delete)

---

## 🔌 Premium Add-On Available

### TurboPOS Premium Edition

Extend the Community Edition with enterprise features:

| Feature | Community | Premium |
|---------|-----------|---------|
| **Basic POS** | ✅ | ✅ |
| **Multi-Store Management** | ❌ | ✅ |
| **AI Inventory Predictions** | ❌ | ✅ |
| **Cloud Backup & Sync** | ❌ | ✅ |
| **Advanced Analytics** | ❌ | ✅ |
| **Loyalty Program** | ❌ | ✅ |
| **Email/SMS Marketing** | ❌ | ✅ |
| **Advanced Accounting** | ❌ | ✅ |
| **Professional Quotations** | ❌ | ✅ |
| **Priority Support** | ❌ | ✅ |

**Pricing:** Starting at $49/month
**Trial:** 30 days FREE (all features)
**Repository:** https://github.com/yourusername/vba-pos-premium

### How Premium Works

The premium add-in (.xla) loads automatically and works alongside the Community Edition:

```
Community Edition (FREE)
    ↓
Loads Premium Add-In (OPTIONAL)
    ↓
Premium Features Activated
```

**Benefits:**
- ✅ Community stays 100% free forever
- ✅ Premium is optional (no forced upgrade)
- ✅ Same database (seamless integration)
- ✅ Cancel anytime, keep community features

---

## 🛠️ Project Structure

```
vba-pos-excel-sqlite/
├── src/
│   ├── pos-excel-sqlite.xlsm    # Main workbook
│   ├── DBVentas.db               # SQLite database
│   ├── vba/                      # VBA modules
│   ├── cls/                      # Class modules
│   ├── frm/                      # UserForms
│   ├── icons/                    # UI icons
│   └── others/                   # Other resources
├── database/
│   ├── schema.sql                # Database schema
│   └── seed-data.sql             # Sample data
├── docs/                         # Documentation
│   ├── API.md                    # Public API docs
│   ├── USER_GUIDE.md             # User manual
│   └── DATABASE_SCHEMA.md        # DB documentation
├── examples/                     # Plugin examples
└── README.md                     # This file
```

---

## 🐛 Troubleshooting

### Common Issues

**1. Database Connection Error**
```
Error: "Could not connect to database"

Solution:
1. Verify SQLite ODBC driver is installed
2. Check database path in Config sheet
3. Ensure DBVentas.db file exists
4. Check file permissions
```

**2. Macros Disabled**
```
Error: "Macros have been disabled"

Solution:
1. File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Enable all macros
3. Or add folder to trusted locations
4. Restart Excel
```

**3. Missing References**
```
Error: "Compile error: Can't find project or library"

Solution:
1. Open VBA Editor (Alt+F11)
2. Tools → References
3. Uncheck missing references (marked as MISSING)
4. Re-add required references:
   - Microsoft ActiveX Data Objects 6.1 Library
   - Microsoft Scripting Runtime
```

**4. Print Issues**
```
Error: Receipt won't print

Solution:
1. Check printer is connected
2. Set default printer in Windows
3. Try Print Preview first
4. Check receipt template in Config sheet
```

---

## 🤝 Contributing

We welcome contributions!

### How to Contribute

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test thoroughly
5. Commit (`git commit -m 'Add amazing feature'`)
6. Push (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Contribution Guidelines

- Follow existing code style
- Add comments for complex logic
- Test all changes
- Update documentation
- Don't break existing functionality

---

## 📄 License

**GNU General Public License v2.0 (GPL v2)**

This software is free and open source. You can:
- ✅ Use it commercially
- ✅ Modify the source code
- ✅ Distribute it
- ✅ Use it privately

**Requirements:**
- ❗ Disclose source code
- ❗ Keep same license (GPL v2)
- ❗ Include copyright notice
- ❗ State changes made

See [LICENSE](LICENSE) for full terms.

---

## 🙏 Acknowledgments

**Original Project:**
- Created by WilfredoHQ
- GPL v2 Licensed
- Community maintained

**Built with:**
- Excel VBA
- SQLite database
- Open source love ❤️

---

## 📞 Support

### Community Support (Free)

- **GitHub Issues:** Report bugs and request features
- **Discussions:** Ask questions, share ideas
- **Wiki:** Documentation and guides

### Premium Support

Upgrade to Premium for:
- ✅ Email support (24-hour response)
- ✅ Phone support (Professional/Enterprise)
- ✅ Priority bug fixes
- ✅ Custom development

---

## 🌍 Localization

**Supported Languages:**
- 🇪🇸 Spanish (Bolivia, Peru, Mexico)
- 🇬🇧 English

**Currency Support:**
- Bolivianos (Bs)
- Soles (S/)
- Pesos (MXN)
- US Dollars ($)

---

## 📈 Statistics

- **Lines of Code:** ~8,000
- **VBA Modules:** 12
- **UserForms:** 15
- **Database Tables:** 8
- **Performance:** Handles 10,000+ transactions

---

## 🔗 Links

- **Premium Edition:** https://github.com/yourusername/vba-pos-premium
- **Documentation:** Coming soon
- **Download:** [Releases](https://github.com/yourusername/vba-pos-excel-sqlite/releases)

---

**Made with ❤️ for small businesses**

**Free Forever | GPL v2 Licensed | Community Driven**

**⭐ Star us on GitHub if this helps your business!**

# 📦 Delivery Management System

A Django-based delivery management system that imports data from Google Sheets and generates professional invoices. The system allows you to manage deliveries, track shipments, and generate both HTML and Word invoices.

## ✨ Features

- **Google Sheets Integration**: Import delivery data directly from Google Sheets
- **Modern Web Interface**: Beautiful, responsive UI with gradient designs
- **Invoice Generation**: Create professional invoices in HTML and Word formats
- **Data Filtering**: Filter deliveries by client and shipment type
- **Bulk Operations**: Select multiple deliveries for batch invoice generation
- **Duplicate Prevention**: Automatic detection and prevention of duplicate entries
- **Multi-format Support**: Handle different timestamp formats from various sheets

## 🛠️ Technologies Used

- **Backend**: Django (Python)
- **Frontend**: HTML5, CSS3, JavaScript
- **Database**: Django ORM (SQLite/PostgreSQL/MySQL)
- **Google Sheets API**: gspread, oauth2client
- **Document Generation**: python-docx
- **Styling**: Custom CSS with modern gradients and animations

## 📋 Prerequisites

- Python 3.8+
- Django 3.2+
- Google Sheets API credentials
- Required Python packages (see requirements.txt)

## 🚀 Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/abdelhak-kadir/Facturation-App.git
   cd delivery-management-system
   ```

2. **Create virtual environment**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**

   ```bash
   pip install django gspread oauth2client python-docx pytz python-num2words
   ```

4. **Google Sheets API Setup**

   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select existing one
   - Enable Google Sheets API and Google Drive API
   - Create service account credentials
   - Download the JSON credentials file
   - Place it in `importer/credentials/`

5. **Configure Django**

   ```bash
   python manage.py makemigrations
   python manage.py migrate
   python manage.py createsuperuser
   ```

6. **Run the server**
   ```bash
   python manage.py runserver
   ```

## 📊 Data Model

### Livraison (Delivery) Model

| Field            | Type                 | Description                       |
| ---------------- | -------------------- | --------------------------------- |
| timestamp        | DateTimeField        | Unique timestamp for the delivery |
| nom_chauffeur    | CharField            | Driver name                       |
| client           | CharField            | Client name                       |
| chargements      | CharField            | Loading location/type             |
| date_chargement  | DateField            | Loading date                      |
| dechargement     | CharField            | Unloading location                |
| bon_livraison    | CharField            | Delivery receipt number           |
| tarif            | CharField            | Rate/pricing                      |
| deplacement      | DecimalField         | Travel/displacement cost          |
| avance           | DecimalField         | Advance payment                   |
| charge_variable  | DecimalField         | Variable charges                  |
| prix_cv          | DecimalField         | CV price                          |
| operateur        | CharField            | Operator                          |
| commercial       | CharField            | Commercial representative         |
| ice              | CharField            | ICE number (tax identifier)       |
| qte              | PositiveIntegerField | Quantity                          |
| nom_destinataire | CharField            | Recipient name                    |
| facturation      | CharField            | Billing status                    |

## 🔧 Configuration

### Google Sheets Setup

1. **Sheet Structure**: The system expects sheets with headers matching the model fields
2. **Supported Headers**: The import function handles multiple header variations:
   - French: "NOM DE CHAUFFEUR", "CLIENTS", "DATE DE CHARGE"
   - English: "nom_chauffeur", "client", "date_chargement"
3. **Timestamp Format**: Supports multiple formats including:
   - `DD/MM/YYYY HH:MM:SS`
   - `MM/DD/YYYY HH:MM:SS`
   - Various date-only formats

### Sheet IDs Configuration

Update the Google Sheet IDs in `sheet_importer.py`:

```python
# Sheet 1
sheet = client.open_by_key("sheet ID").sheet1

# Sheet 2
sheet = client.open_by_key("sheet ID").sheet1
```

## 🎯 Usage

### Importing Data

1. Navigate to the delivery list page
2. Click the "📥 Importer" button
3. The system will import data from both configured Google Sheets
4. Duplicates are automatically detected and skipped

### Generating Invoices

1. **Select Deliveries**: Use checkboxes to select deliveries
2. **Choose Format**:
   - **HTML Invoice**: Click "📄 Générer Facture HTML"
   - **Word Invoice**: Click "📝 Générer Facture Word"
3. **Download**: The invoice will be generated and downloaded automatically

### Filtering Data

- **By Client**: Enter client name in the client filter field
- **By Shipment**: Enter shipment type in the chargements filter
- Click "🔍 Filtrer" to apply filters
- Click "🔄 Réinitialiser" to reset filters

## 📑 Invoice Features

### Word Invoice Includes:

- Professional company header with logo support
- Client information and ICE number
- Detailed delivery table with dates, quantities, and rates
- Automatic calculation of subtotals, VAT (20%), and total
- Amount in French words
- Company footer with contact information

### Customization:

- Company logo: Place `logo.png` in `importer/media/`
- Company details: Update in the `_add_document_footer()` function
- VAT rate: Modify `tva_rate` in the invoice generation code

## 🎨 UI Features

- **Modern Design**: Gradient backgrounds and smooth animations
- **Responsive Layout**: Works on desktop, tablet, and mobile
- **Interactive Elements**: Hover effects and smooth transitions
- **Accessibility**: Proper contrast and semantic markup
- **Real-time Updates**: Dynamic selection counters and button states

## 🔍 Troubleshooting

### Common Issues:

1. **Import Failures**:

   - Check Google Sheets API credentials
   - Verify sheet IDs are correct
   - Ensure proper column headers

2. **Timestamp Parsing Errors**:

   - The system supports multiple formats
   - Check the `formats_to_try` array in the import function

3. **Invoice Generation Issues**:
   - Verify `python-docx` is installed
   - Check logo file path for Word invoices
   - Ensure `num2words` package is installed for French number conversion

## 📝 Development

### Adding New Features:

1. **New Model Fields**:

   - Add to `models.py`
   - Update `forms.py`
   - Run migrations

2. **New Import Sources**:

   - Create new import function in `sheet_importer.py`
   - Add to the view's import process

3. **Custom Invoice Templates**:
   - Modify Word generation functions
   - Update HTML templates

### Code Structure:

```
├── models.py          # Data models
├── views.py           # View controllers
├── forms.py           # Django forms
├── utils/
│   └── sheet_importer.py  # Google Sheets import logic
├── templates/
│   ├── livraison_list.html    # Main list view
│   ├── edit_livraison.html    # Edit form
│   └── invoice.html           # HTML invoice template
└── static/           # CSS, JS, images
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 📞 Support

For support and questions:

- Create an issue in the repository
- Check the troubleshooting section
- Review the Google Sheets API documentation

## 🔄 Updates

The system automatically handles:

- Duplicate detection during imports
- Multiple timestamp formats
- Empty cell handling
- Column header variations
- Error logging and reporting

# Cloud R tech — HR Portal

A Django-based internal HR document portal for generating professional PDF letters and pay slips.

---

## 📋 Modules

| Module | URL | Description |
|---|---|---|
| Dashboard | `/` | Home screen with all module cards |
| Appraisal Letter | `/appraisal/` | Generate appraisal increment letters |
| Experience Letter | `/experience/` | Generate experience certificates |
| Offer Letter | `/offer/` | Generate job offer letters |
| Contract Extension | `/contract/` | Generate contract renewal letters |
| **Pay Slip** | `/payslip/` | Generate monthly pay slips with auto-calculated salary |

---

## ⚙️ Requirements

- Python **3.10+**
- pip

All Python dependencies are listed in `requirements.txt`:

```
Django>=4.2,<6.0
reportlab>=4.0
Pillow>=10.0
python-dotenv>=1.0
```

---

## 🚀 Setup & Run

### Step 1 — Unzip the project

```bash
unzip cloudrtech_hr_v8.zip
cd cloudrtech_hr_v7/hr_project
```

### Step 2 — Create a virtual environment

```bash
python -m venv venv
```

Activate it:

- **Windows:**
  ```bash
  venv\Scripts\activate
  ```
- **macOS / Linux:**
  ```bash
  source venv/bin/activate
  ```

### Step 3 — Install dependencies

```bash
pip install -r requirements.txt
```

### Step 4 — Configure environment variables

Copy the example env file and edit it:

```bash
cp .env.example .env
```

Open `.env` and set your values:

```env
SECRET_KEY=your-very-secret-key-here
DEBUG=True
ALLOWED_HOSTS=*

# HR portal login credentials
ADMIN_USERNAME=admin
ADMIN_PASSWORD=cloudrtech@45
```

> **Note:** For local development the defaults work fine as-is. Change `SECRET_KEY` and the credentials before deploying to production.

### Step 5 — Set up the database

```bash
python manage.py migrate
```

### Step 6 — Run the development server

```bash
python manage.py runserver
```

Open your browser and go to:

```
http://127.0.0.1:8000/
```

Log in with the credentials from your `.env` file (default: `admin` / `cloudrtech@45`).

---

## Pay Slip Module

Navigate to **`/payslip/`** and fill in the form:

| Field | Example |
|---|---|
| Employee No | `CRT887` (auto-prefixed with `CRT`, type the 3-digit number) |
| Name | `Marri Vishnu Vardhan Reddy` |
| Designation | Choose from dropdown (add/edit/delete options) |
| Category | Choose from dropdown — `Contract / Permanent / Probation / Intern / Trainee` (add/edit/delete options) |
| Sex | `Male / Female` |
| Date of Joining | `20/05/2024` |
| Pay Slip Month | `April 2026` |
| Working Days | `30` |
| Paid Holiday | `1` |
| Gross Salary (INR) | `37500` |

The salary breakdown on the generated PDF is **auto-calculated** from Gross Salary:

| Component | Formula |
|---|---|
| Basic | 40% of Gross |
| HRA | 15% of Gross |
| Medical Allowance | 5% of Gross |
| Performance Bonus | 20% of Gross |
| Conveyance | 10% of Gross |
| Management Allowance | 10% of Gross |
| **Total Earnings** | Sum of above |
| Professional Tax | ₹ 250.00 (fixed) |
| **Net Pay** | Total Earnings − ₹ 250 |

The Pay Slip page also has quick-navigation buttons to the Appraisal,
Experience, Offer, and Contract letter forms.

---

## Dropdown fields (add / edit / delete)

The following fields use a custom dropdown with built-in **add / edit /
delete** management (click **⊕ Manage** or the field itself to open it):

| Field | Used in |
|---|---|
| Designation | Appraisal, Experience, Offer, Contract, Pay Slip |
| Reason for Relieving | Experience Letter |
| Conduct | Experience Letter |
| Category | Pay Slip |

Options you add/edit/delete are stored in your browser session (no database
setup required) and shared across all forms that use the same field.

### Reference No. / Employee No.

The Reference No. (Experience, Contract) and Employee No. (Pay Slip) fields
default to `CRT` — just type the 3-digit number that follows
(e.g. `CRT001`, `CRT042`).

---

## 📁 Project Structure

```
hr_project/
├── config/
│   ├── settings.py          # Django settings (reads from .env)
│   ├── urls.py              # Root URL config
│   ├── wsgi.py
│   └── asgi.py
├── letters/
│   ├── views.py             # All form + PDF generate views
│   ├── urls.py              # All URL routes
│   ├── pdf_generator.py     # ReportLab PDF builders (all modules incl. payslip)
│   ├── models.py
│   └── admin.py
├── templates/
│   └── letters/
│       ├── base.html
│       ├── dashboard.html
│       ├── payslip_form.html       ← Pay Slip form
│       ├── appraisal_form.html
│       ├── experience_form.html
│       ├── offer_form.html
│       └── contract_form.html
├── static/
│   └── letters/
│       └── img/
│           ├── letterhead.png
│           ├── payslip_header.jpg  ← Cloud R tech banner header
│           ├── signature.png
│           └── seal_stamp.png
├── .env.example             # Copy to .env and fill in values
├── manage.py
└── requirements.txt
```

---

## 🔐 Login & Session

- Login URL: `/login/`
- Credentials are set via `ADMIN_USERNAME` and `ADMIN_PASSWORD` in `.env`
- Sessions expire when the browser is closed, or after **12 hours** maximum

---

## 🌐 Production Deployment

When going live, update your `.env`:

```env
DEBUG=False
SECRET_KEY=a-long-random-secret-key
ALLOWED_HOSTS=yourdomain.com,www.yourdomain.com
ADMIN_PASSWORD=a-strong-password
```

Then collect static files:

```bash
python manage.py collectstatic
```

Serve with **Gunicorn** behind **Nginx**:

```bash
pip install gunicorn
gunicorn config.wsgi:application --bind 0.0.0.0:8000
```

Configure Nginx to proxy `http://localhost:8000` and serve the `staticfiles/` directory directly.

---

## 🛠️ Common Issues

**`ModuleNotFoundError: No module named 'dotenv'`**
```bash
pip install python-dotenv
```

**Static files not loading (CSS / images missing)**
Make sure `DEBUG=True` in development. In production run `collectstatic` and configure Nginx to serve `/static/`.

**Port already in use**
```bash
python manage.py runserver 8080
# then open http://127.0.0.1:8080/
```

**Database errors after updating**
```bash
python manage.py migrate
```

---

## 📞 Support

**Cloud R tech**
Address: 187/188, Thiruvalluvar Salai, Near DAV Girls School, Paneer Nagar, Mogappair East, Chennai, Tamil Nadu 600037
Email: hr@cloudrtech.com
Phone: +91 9080385995

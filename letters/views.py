"""
letters/views.py — Cloud R tech HR Portal
"""

import json
import logging
import re
from functools import wraps
from io import BytesIO
import os

from django.conf import settings
import hashlib
from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_POST

from . import pdf_generator as pg
from . import docx_generator as dg

logger = logging.getLogger(__name__)

DEFAULT_HR_NAME = 'Raj Padmanaban'

# Default dropdown lists (session-persisted)
DEFAULT_DESIGNATIONS = [
    'Software Engineer', 'Senior Software Engineer', 'Python Developer',
    'Full Stack Developer', 'HR Manager', 'Project Manager', 'Team Lead',
]
DEFAULT_REASONS = [
    'Career Growth', 'Personal Reasons', 'Higher Studies',
    'Relocation', 'Better Opportunity', 'Health Issues',
]
DEFAULT_CONDUCTS = ['Good', 'Excellent', 'Satisfactory', 'Outstanding']
DEFAULT_CATEGORIES = ['Contract', 'Permanent', 'Probation', 'Intern', 'Trainee']


# ── Auth helpers ──────────────────────────────────────────────────────────────

def _login_required(view_fn):
    @wraps(view_fn)
    def wrapper(request, *args, **kwargs):
        if not request.session.get('is_admin'):
            return redirect(f'/login/?next={request.path}')
        return view_fn(request, *args, **kwargs)
    return wrapper


def _require_post(view_fn):
    @wraps(view_fn)
    def wrapper(request, *args, **kwargs):
        if request.method != 'POST':
            return HttpResponse('Method Not Allowed', status=405)
        return view_fn(request, *args, **kwargs)
    return wrapper


def _sanitize_name(name: str) -> str:
    return re.sub(r'[^\w\s\-.]', '', name).strip().replace(' ', '_') or 'document'


def _pdf_response(pdf_bytes: bytes, filename: str) -> HttpResponse:
    resp = HttpResponse(pdf_bytes, content_type='application/pdf')
    safe = _sanitize_name(filename)
    resp['Content-Disposition'] = f'attachment; filename="{safe}.pdf"'
    return resp


def _docx_response(docx_bytes: bytes, filename: str) -> HttpResponse:
    # Compute SHA256 so client can verify the downloaded file matches server output
    sha = hashlib.sha256(docx_bytes).hexdigest()
    resp = HttpResponse(
        docx_bytes,
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    safe = _sanitize_name(filename)
    resp['Content-Disposition'] = f'attachment; filename="{safe}.docx"'
    # Expose checksum and explicit content-length for easy verification after download
    resp['X-Docx-SHA256'] = sha
    resp['Content-Length'] = str(len(docx_bytes))
    return resp


def _int_field(post, key, default=0):
    raw = post.get(key, default)
    if raw == '' or raw is None:
        return default
    try:
        return int(float(str(raw)))
    except (TypeError, ValueError):
        raise ValueError(f"Field '{key}' must be a number (got: {raw!r})")


def _hr_context(request):
    """Return hr_name + sig_src (BytesIO or None) from session."""
    hr_name = request.session.get('hr_name') or DEFAULT_HR_NAME
    sig_bytes = request.session.get('sig_bytes')   # stored as list of ints
    sig_src = BytesIO(bytes(sig_bytes)) if sig_bytes else None
    return hr_name, sig_src


# ── Dropdown helpers ──────────────────────────────────────────────────────────

def _get_list(request, key, default):
    return request.session.get(key, list(default))


def _save_list(request, key, lst):
    request.session[key] = lst
    request.session.modified = True


# ── Auth views ────────────────────────────────────────────────────────────────

def login_view(request):
    if request.session.get('is_admin'):
        return redirect('/')
    error = None
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        if (username == settings.ADMIN_USERNAME and
                password == settings.ADMIN_PASSWORD):
            request.session['is_admin'] = True
            request.session.set_expiry(86400 * 7)
            next_url = request.GET.get('next', '/')
            return redirect(next_url)
        else:
            error = 'Invalid username or password.'
    return render(request, 'letters/login.html', {'error': error})


def logout_view(request):
    request.session.flush()
    return redirect('/login/')


# ── Settings view ─────────────────────────────────────────────────────────────

@_login_required
def hr_settings(request):
    message = None
    error = None
    current_name = request.session.get('hr_name') or DEFAULT_HR_NAME
    has_custom_sig = bool(request.session.get('sig_bytes'))

    if request.method == 'POST':
        action = request.POST.get('action', 'save')

        if action == 'reset':
            request.session.pop('hr_name', None)
            request.session.pop('sig_bytes', None)
            message = f'Reset to default: {DEFAULT_HR_NAME} with original signature.'
            current_name = DEFAULT_HR_NAME
            has_custom_sig = False

        else:  # save
            name = request.POST.get('hr_name', '').strip()
            sig_file = request.FILES.get('signature_png')

            if not name:
                error = 'HR Manager name cannot be empty.'
            else:
                if sig_file:
                    if not sig_file.name.lower().endswith('.png'):
                        error = 'Signature file must be a .png file.'
                    elif sig_file.size > 2 * 1024 * 1024:
                        error = 'PNG file is too large (max 2 MB).'
                    else:
                        raw = sig_file.read()
                        if not raw.startswith(b'\x89PNG'):
                            error = 'Uploaded file does not appear to be a valid PNG.'
                        else:
                            request.session['sig_bytes'] = list(raw)
                            has_custom_sig = True

                if not error:
                    request.session['hr_name'] = name
                    current_name = name
                    sig_msg = 'with new signature' if sig_file and not error else ('with custom signature' if has_custom_sig else 'with default signature')
                    message = f'Saved! Active HR Manager: {name} {sig_msg}.'

    return render(request, 'letters/hr_settings.html', {
        'current_name':   current_name,
        'has_custom_sig': has_custom_sig,
        'default_name':   DEFAULT_HR_NAME,
        'message':        message,
        'error':          error,
    })


# ── Dropdown AJAX endpoints ───────────────────────────────────────────────────

@_login_required
def dropdown_get(request):
    """GET /dropdown/?key=designations  → JSON list"""
    key = request.GET.get('key')
    defaults = {
        'designations': DEFAULT_DESIGNATIONS,
        'reasons':      DEFAULT_REASONS,
        'conducts':     DEFAULT_CONDUCTS,
        'categories':   DEFAULT_CATEGORIES,
    }
    if key not in defaults:
        return JsonResponse({'error': 'unknown key'}, status=400)
    return JsonResponse({'items': _get_list(request, key, defaults[key])})


@_login_required
def dropdown_save(request):
    """POST /dropdown/save/  body: {key, items:[]}  → JSON ok"""
    if request.method != 'POST':
        return JsonResponse({'error': 'POST only'}, status=405)
    try:
        body = json.loads(request.body)
        key   = body['key']
        items = body['items']
    except (KeyError, json.JSONDecodeError):
        return JsonResponse({'error': 'bad body'}, status=400)
    valid_keys = {'designations', 'reasons', 'conducts', 'categories'}
    if key not in valid_keys:
        return JsonResponse({'error': 'unknown key'}, status=400)
    items = [str(i).strip() for i in items if str(i).strip()]
    _save_list(request, key, items)
    return JsonResponse({'ok': True, 'items': items})


# ── Page views ────────────────────────────────────────────────────────────────

@_login_required
def dashboard(request):
    return render(request, 'letters/dashboard.html')


@_login_required
def appraisal_form(request):
    return render(request, 'letters/appraisal_form.html')


@_login_required
def experience_form(request):
    return render(request, 'letters/experience_form.html')


@_login_required
def offer_form(request):
    return render(request, 'letters/offer_form.html')


@_login_required
def contract_form(request):
    return render(request, 'letters/contract_form.html')


@_login_required
def payslip_form(request):
    return render(request, 'letters/payslip_form.html')


# ── PDF generation views ──────────────────────────────────────────────────────

@_login_required
@_require_post
def generate_appraisal(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'name':            request.POST.get('name', '').strip(),
            'gender':          request.POST.get('gender', 'male'),
            'address':         request.POST.get('address', '').strip(),
            'pincode':         request.POST.get('pincode', '').strip(),
            'date':            request.POST.get('date', '').strip(),
            'current_monthly': _int_field(request.POST, 'current_monthly'),
            'new_monthly':     _int_field(request.POST, 'new_monthly'),
            'designation':     request.POST.get('designation', '').strip(),
            'ctc_lpa':         request.POST.get('ctc_lpa', '').strip(),
            'hr_name':         hr_name,
            'sig_src':         sig_src,
        }
        pdf = pg.generate_appraisal(data)
        logger.info("Appraisal PDF generated for %s", data['name'])
        return _pdf_response(pdf, f"Appraisal_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating appraisal PDF")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_experience(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'ref':               request.POST.get('ref', '').strip(),
            'name':              request.POST.get('name', '').strip(),
            'gender':            request.POST.get('gender', 'male'),
            'closing_date':      request.POST.get('closing_date', '').strip(),
            'designation':       request.POST.get('designation', '').strip(),
            'date_of_joining':   request.POST.get('date_of_joining', '').strip(),
            'date_of_relieving': request.POST.get('date_of_relieving', '').strip(),
            'reason':            request.POST.get('reason', '').strip(),
            'conduct':           request.POST.get('conduct', '').strip(),
            'hr_name':           hr_name,
            'sig_src':           sig_src,
        }
        pdf = pg.generate_experience(data)
        logger.info("Experience PDF generated for %s", data['name'])
        return _pdf_response(pdf, f"Experience_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating experience PDF")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_offer(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'date':         request.POST.get('date', '').strip(),
            'name':         request.POST.get('name', '').strip(),
            'gender':       request.POST.get('gender', 'male'),
            'address':      request.POST.get('address', '').strip(),
            'pincode':      request.POST.get('pincode', '').strip(),
            'designation':  request.POST.get('designation', '').strip(),
            'joining_date': request.POST.get('joining_date', '').strip(),
            'ctc_monthly':  _int_field(request.POST, 'ctc_monthly'),
            'ctc_lpa':      request.POST.get('ctc_lpa', '').strip(),
            'hr_name':      hr_name,
            'sig_src':      sig_src,
        }
        pdf = pg.generate_offer(data)
        logger.info("Offer PDF generated for %s", data['name'])
        return _pdf_response(pdf, f"Offer_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating offer PDF")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_contract(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'ref':               request.POST.get('ref', '').strip(),
            'name':              request.POST.get('name', '').strip(),
            'gender':            request.POST.get('gender', 'male'),
            'address':           request.POST.get('address', '').strip(),
            'pincode':           request.POST.get('pincode', '').strip(),
            'extended_date':     request.POST.get('extended_date', '').strip(),
            'current_ctc_lpa':   request.POST.get('current_ctc_lpa', '').strip(),
            'increment_ctc_lpa': request.POST.get('increment_ctc_lpa', '').strip(),
            'effective_month':   request.POST.get('effective_month', '').strip(),
            'designation':       request.POST.get('designation', '').strip(),
            'ctc_monthly':       _int_field(request.POST, 'ctc_monthly'),
            'hr_name':           hr_name,
            'sig_src':           sig_src,
        }
        pdf = pg.generate_contract(data)
        logger.info("Contract PDF generated for %s", data['name'])
        return _pdf_response(pdf, f"Contract_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating contract PDF")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_payslip(request):
    try:
        data = {
            'emp_no':          request.POST.get('emp_no', '').strip(),
            'name':            request.POST.get('name', '').strip(),
            'designation':     request.POST.get('designation', '').strip(),
            'category':        request.POST.get('category', '').strip(),
            'sex':             request.POST.get('sex', '').strip(),
            'date_of_joining': request.POST.get('date_of_joining', '').strip(),
            'month_year':      request.POST.get('month_year', '').strip(),
            'working_days':    request.POST.get('working_days', '').strip(),
            'paid_holiday':    request.POST.get('paid_holiday', '0').strip(),
            'gross_salary':    _int_field(request.POST, 'gross_salary'),
        }
        pdf = pg.generate_payslip(data)
        logger.info("Payslip PDF generated for %s", data['name'])
        return _pdf_response(pdf, f"Payslip_{data['name']}_{data['month_year']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating payslip PDF")
        return HttpResponse('Internal Server Error', status=500)


# ── DOCX generation views ─────────────────────────────────────────────────────

@_login_required
@_require_post
def generate_appraisal_docx(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'name':            request.POST.get('name', '').strip(),
            'gender':          request.POST.get('gender', 'male'),
            'address':         request.POST.get('address', '').strip(),
            'pincode':         request.POST.get('pincode', '').strip(),
            'date':            request.POST.get('date', '').strip(),
            'current_monthly': _int_field(request.POST, 'current_monthly'),
            'new_monthly':     _int_field(request.POST, 'new_monthly'),
            'designation':     request.POST.get('designation', '').strip(),
            'ctc_lpa':         request.POST.get('ctc_lpa', '').strip(),
            'hr_name':         hr_name,
            'sig_src':         sig_src,
        }
        docx = dg.generate_appraisal(data)
        logger.info("Appraisal DOCX generated for %s", data['name'])
        return _docx_response(docx, f"Appraisal_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating appraisal DOCX")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_experience_docx(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'ref':               request.POST.get('ref', '').strip(),
            'name':              request.POST.get('name', '').strip(),
            'gender':            request.POST.get('gender', 'male'),
            'closing_date':      request.POST.get('closing_date', '').strip(),
            'designation':       request.POST.get('designation', '').strip(),
            'date_of_joining':   request.POST.get('date_of_joining', '').strip(),
            'date_of_relieving': request.POST.get('date_of_relieving', '').strip(),
            'reason':            request.POST.get('reason', '').strip(),
            'conduct':           request.POST.get('conduct', '').strip(),
            'hr_name':           hr_name,
            'sig_src':           sig_src,
        }
        docx = dg.generate_experience(data)
        logger.info("Experience DOCX generated for %s", data['name'])
        return _docx_response(docx, f"Experience_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating experience DOCX")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_offer_docx(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'date':         request.POST.get('date', '').strip(),
            'name':         request.POST.get('name', '').strip(),
            'gender':       request.POST.get('gender', 'male'),
            'address':      request.POST.get('address', '').strip(),
            'pincode':      request.POST.get('pincode', '').strip(),
            'designation':  request.POST.get('designation', '').strip(),
            'joining_date': request.POST.get('joining_date', '').strip(),
            'ctc_monthly':  _int_field(request.POST, 'ctc_monthly'),
            'ctc_lpa':      request.POST.get('ctc_lpa', '').strip(),
            'hr_name':      hr_name,
            'sig_src':      sig_src,
        }
        docx = dg.generate_offer(data)
        logger.info("Offer DOCX generated for %s", data['name'])
        return _docx_response(docx, f"Offer_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating offer DOCX")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_contract_docx(request):
    try:
        hr_name, sig_src = _hr_context(request)
        data = {
            'ref':               request.POST.get('ref', '').strip(),
            'name':              request.POST.get('name', '').strip(),
            'gender':            request.POST.get('gender', 'male'),
            'address':           request.POST.get('address', '').strip(),
            'pincode':           request.POST.get('pincode', '').strip(),
            'extended_date':     request.POST.get('extended_date', '').strip(),
            'current_ctc_lpa':   request.POST.get('current_ctc_lpa', '').strip(),
            'increment_ctc_lpa': request.POST.get('increment_ctc_lpa', '').strip(),
            'effective_month':   request.POST.get('effective_month', '').strip(),
            'designation':       request.POST.get('designation', '').strip(),
            'ctc_monthly':       _int_field(request.POST, 'ctc_monthly'),
            'hr_name':           hr_name,
            'sig_src':           sig_src,
        }
        docx = dg.generate_contract(data)
        logger.info("Contract DOCX generated for %s", data['name'])
        return _docx_response(docx, f"Contract_{data['name']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating contract DOCX")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
@_require_post
def generate_payslip_docx(request):
    try:
        data = {
            'emp_no':          request.POST.get('emp_no', '').strip(),
            'name':            request.POST.get('name', '').strip(),
            'designation':     request.POST.get('designation', '').strip(),
            'category':        request.POST.get('category', '').strip(),
            'sex':             request.POST.get('sex', '').strip(),
            'date_of_joining': request.POST.get('date_of_joining', '').strip(),
            'month_year':      request.POST.get('month_year', '').strip(),
            'working_days':    request.POST.get('working_days', '').strip(),
            'paid_holiday':    request.POST.get('paid_holiday', '0').strip(),
            'gross_salary':    _int_field(request.POST, 'gross_salary'),
        }
        docx = dg.generate_payslip(data)
        logger.info("Payslip DOCX generated for %s", data['name'])
        return _docx_response(docx, f"Payslip_{data['name']}_{data['month_year']}")
    except ValueError as exc:
        return HttpResponse(f'Bad Request: {exc}', status=400)
    except Exception:
        logger.exception("Error generating payslip DOCX")
        return HttpResponse('Internal Server Error', status=500)


@_login_required
def docx_debug(request):
    """
    Debug endpoint: GET /docx/debug/?file=sample_appraisal.docx
    Returns JSON with `file`, `size` and `sha256` for quick verification
    """
    fname = request.GET.get('file', 'sample_appraisal.docx')
    path = os.path.join(settings.BASE_DIR, fname)
    if not os.path.exists(path):
        return JsonResponse({'error': 'file not found', 'file': fname}, status=404)
    try:
        with open(path, 'rb') as fh:
            data = fh.read()
        sha = hashlib.sha256(data).hexdigest()
        return JsonResponse({'file': fname, 'size': len(data), 'sha256': sha})
    except Exception as exc:
        logger.exception('Error reading debug file')
        return JsonResponse({'error': str(exc)}, status=500)

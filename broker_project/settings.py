import os
from pathlib import Path
# import dj_database_url
from dotenv import load_dotenv   # ← ADD
load_dotenv()    
# -------------------------
# Base
# -------------------------
BASE_DIR = Path(__file__).resolve().parent.parent

# -------------------------
# Secrets & debug
# -------------------------
# Use environment SECRET_KEY in Render / production
SECRET_KEY = os.environ.get("SECRET_KEY", "fallback-secret-key-for-local-dev")

# Control debug from environment. Default True (local). In Render set DEBUG=False.
#DEBUG = os.environ.get("DEBUG", "True").lower() in ("1", "true", "yes")
DEBUG = True
# -------------------------
# Hosts / CSRF
# -------------------------
# During dev we allow all hosts (you can restrict later)
ALLOWED_HOSTS = [
    "pooja11121980.pythonanywhere.com",
    "127.0.0.1",
    "localhost",
]


# Trusted origins for CSRF when served from onrender.com / render.com
# Add your actual Render service URL to this env var if needed.
# CSRF_TRUSTED_ORIGINS = os.environ.get(
#     "CSRF_TRUSTED_ORIGINS", "https://*.onrender.com,https://*.render.com"
# ).split(",")
CSRF_TRUSTED_ORIGINS = [
    "https://pooja11121980.pythonanywhere.com"
]

# -------------------------
# Applications
# -------------------------
INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "widget_tweaks",
    "brokerapp",
]

# -------------------------
# Middleware
# -------------------------
MIDDLEWARE = [
    "django.middleware.security.SecurityMiddleware",
    # "whitenoise.middleware.WhiteNoiseMiddleware",  # serve static files on Render
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
    "brokerapp.middleware.SingleOrgMiddleware",
]

ROOT_URLCONF = "broker_project.urls"

# -------------------------
# Templates
# -------------------------
TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "brokerapp" / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
                "brokerapp.context_processors.current_org",
            ],
        },
    },
]

WSGI_APPLICATION = "broker_project.wsgi.application"

# -------------------------
# Database (Render-compatible)
# -------------------------
# Priority: use DATABASE_URL env var on Render. For local dev, fallback to local Postgres (if present)
# DATABASE_URL = os.environ.get("DATABASE_URL")

# if DATABASE_URL:
#     # Production / Render
#     DATABASES = {
#         "default": dj_database_url.parse(
#             DATABASE_URL,
#             conn_max_age=600,
#             ssl_require=True
#         )
#     }
# else:
#     # Local development: if you run Postgres locally keep these values,
#     # otherwise install sqlite fallback by uncommenting the sqlite block below.
#     DATABASES = {
#         "default": {
#             "ENGINE": "django.db.backends.postgresql",
#             "NAME": "broker_project_db",
#             "USER": "postgres",
#             "PASSWORD": "keshav1604",
#             "HOST": "localhost",
#             "PORT": "5432",
#             "OPTIONS": {
#                 "sslmode": "disable"
#             },
#         }
#     }
#     # If you want sqlite local fallback instead, comment the above and use:
#     # BASE_DIR = Path(__file__).resolve().parent.parent
#     # DATABASES = {
#     #     "default": {
#     #         "ENGINE": "django.db.backends.sqlite3",
#     #         "NAME": BASE_DIR / "db.sqlite3",
#     #     }
#     # }
# -------------------------
# Database (PythonAnywhere FREE – SQLite)
# -------------------------
DATABASES = {
    "default": {
        "ENGINE": "django.db.backends.sqlite3",
        "NAME": BASE_DIR / "db.sqlite3",
    }
}

# -------------------------
# Password validators
# -------------------------
AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]

# -------------------------
# Internationalization
# -------------------------
LANGUAGE_CODE = "en-us"
TIME_ZONE = "UTC"
USE_I18N = True
USE_TZ = True

# -------------------------
# Static files
# -------------------------
STATIC_URL = "/static/"
STATICFILES_DIRS = []
STATIC_ROOT = BASE_DIR / "staticfiles"
STATICFILES_STORAGE = "whitenoise.storage.CompressedManifestStaticFilesStorage"

# -------------------------
# Default primary key field type
# -------------------------
DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# -------------------------
# Email – Gmail SMTP (LIVE)
# -------------------------

EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True

EMAIL_HOST_USER = os.environ.get("EMAIL_USER", "")
EMAIL_HOST_PASSWORD = os.environ.get("EMAIL_PASS", "")

DEFAULT_FROM_EMAIL = EMAIL_HOST_USER


# -------------------------
# Auth / Login redirects
# -------------------------
LOGIN_URL = "/login/"
LOGIN_REDIRECT_URL = "/dashboard/"
LOGOUT_REDIRECT_URL = "/login/"

# -------------------------
# App-specific defaults
# -------------------------
DEFAULT_ORG_NAME = os.environ.get("DEFAULT_ORG_NAME", "Rathi Trading Co.")

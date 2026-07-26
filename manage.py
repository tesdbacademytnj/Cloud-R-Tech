#!/usr/bin/env python
"""Django's command-line utility for administrative tasks."""
import os
import sys

# Use the modern SQLite bundled by pysqlite3-binary instead of the
# older system SQLite (AlmaLinux 8 ships 3.26, Django needs 3.31+).
import pysqlite3
sys.modules['sqlite3'] = sys.modules.pop('pysqlite3')


def main():
    """Run administrative tasks."""
    os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
    try:
        from django.core.management import execute_from_command_line
    except ImportError as exc:
        raise ImportError(
            "Couldn't import Django. Are you sure it's installed and "
            "available on your PYTHONPATH environment variable? Did you "
            "forget to activate a virtual environment?"
        ) from exc
    execute_from_command_line(sys.argv)


if __name__ == '__main__':
    main()

import os
import django

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "crm.settings")
django.setup()

from django.contrib.auth import get_user_model

username = os.getenv("DJANGO_SUPERUSER_USERNAME")
email = os.getenv("DJANGO_SUPERUSER_EMAIL", "")
password = os.getenv("DJANGO_SUPERUSER_PASSWORD")

if not username or not password:
    print("Superuser env vars not set, skipping.")
    raise SystemExit(0)

User = get_user_model()

if User.objects.filter(username=username).exists():
    print("Superuser already exists, skipping.")
else:
    user = User.objects.create_superuser(username=username, email=email, password=password)
    print(f"Superuser created: {user.username}")

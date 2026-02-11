import os
import django
from django.urls import reverse
from django.conf import settings

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'broker_project.settings')
django.setup()

party_name = "SANTOSH SAHU (BANK A/C)"
broker_name = "BROKER (SPL A/C)"

print("--- Verifying Party URL ---")
try:
    url = reverse('party_edit', args=[party_name])
    print(f"Successfully reversed Party URL for '{party_name}': {url}")
except Exception as e:
    print(f"Failed to reverse Party URL for '{party_name}': {e}")

print("\n--- Verifying Broker URL ---")
try:
    url = reverse('broker_edit', args=[broker_name])
    print(f"Successfully reversed Broker URL for '{broker_name}': {url}")
except Exception as e:
    print(f"Failed to reverse Broker URL for '{broker_name}': {e}")

import traceback

try:
    from src.backend.app.services import emails_utils as eu
    from src.backend.app.api.v1.endpoints import codes as codes_mod
    from src.backend.app.api.v1 import api as api_mod
    print('Import check OK: emails_utils, codes, api loaded')
except Exception:
    print('Import check FAILED')
    traceback.print_exc()

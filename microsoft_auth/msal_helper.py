# microsoft_auth/msal_helper.py
import msal
from django.conf import settings

def build_msal_app():
    return msal.ConfidentialClientApplication(
        client_id=settings.MICROSOFT_CLIENT_ID,
        client_credential=settings.MICROSOFT_CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{settings.MICROSOFT_TENANT_ID}"
    )

def get_auth_url(msal_app, state):
    return msal_app.get_authorization_request_url(
        scopes=["User.Read"],
        state=state,
        redirect_uri=settings.MICROSOFT_REDIRECT_URI
    )

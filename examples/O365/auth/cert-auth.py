#!/usr/bin/env python3
import os
from O365 import Account

from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_CERT_PASSWORD = os.getenv("CLIENT_CERT_PASSWORD")

CERT_PFX_PATH = os.getenv("CERT_PFX_PATH")
CERT_PFX_PASSWORD = os.getenv("CERT_PFX_PASSWORD")

CERT_PEM_PATH = os.getenv("CERT_PEM_PATH")
CERT_PEM_PASSWORD = os.getenv("CERT_PEM_PASSWORD")

# Reference for Python MSAL
# https://learn.microsoft.com/en-us/python/api/msal/msal.application.confidentialclientapplication?view=msal-py-latest


def using_pfx_path(
        client_id: str = CLIENT_ID,
        cert_pfx_path: str = CERT_PFX_PATH,
        cert_pfx_password: str = CERT_PFX_PASSWORD,
        tenant_id: str = TENANT_ID,
    ):
    """
      Supporting reading client certificates from PFX files (>=v1.29.0)
    """
    client_secret = {
       "private_key_pfx_path":cert_pfx_path,
       "passphrase": cert_pfx_password,
    }
    credentials = client_id, client_secret
    account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant_id)
    if account.authenticate():
        print('[pfx_cert] Authenticated!')

    scopes = ['basic']
    if account.authenticate(scopes=scopes):
       print(f'[pfx_cert] Authenticated with scope {scopes}!')


def using_pem_path(
        client_id: str = CLIENT_ID,
        cert_pem_path: str = CERT_PEM_PATH,
        cert_pem_password: str = CERT_PEM_PASSWORD,
        cert_thumbprint: str = None,
        tenant_id: str = TENANT_ID,
    ):
    """
      Support using a certificate in X.509 (.pem) formatFeed in a dict in this form
    """
    with open(cert_pem_path) as fp:
        pem_certificate = fp.read()

    if cert_thumbprint is None:
        cert_thumbprint = get_thumbprint(pem_certificate)

    client_secret = {
       "private_key": pem_certificate,
       "thumbprint": cert_thumbprint,
       # "passphrase": cert_pem_password,
    }
    credentials = client_id, client_secret
    account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant_id)
    if account.authenticate():
        print('[pem_cert] Authenticated!')

    scopes = ['basic']
    if account.authenticate(scopes=scopes):
       print(f'[pem_cert] Authenticated with scope {scopes}!')


def get_thumbprint(pem_certificate: str) -> str:
    """
    Extract the thumbprint from a PEM certificate.
    """
    from cryptography import x509
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.backends import default_backend
    import hashlib

    # Load the certificate
    cert = x509.load_pem_x509_certificate(pem_certificate.encode(), default_backend())
    # Get the DER-encoded bytes
    der_bytes = cert.public_bytes(encoding=serialization.Encoding.DER)
    # Compute SHA-1 hash (thumbprint)
    thumbprint = hashlib.sha1(der_bytes).hexdigest().upper()
    return thumbprint


if __name__ == "__main__":
    print(f"Using Azure Tenant ID: {TENANT_ID}")
    # Example usage
    using_pfx_path()
    # or
    using_pem_path()
    # Note: Ensure that the environment variables are set correctly before running this script.

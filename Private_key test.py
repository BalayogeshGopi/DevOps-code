import base64
import msal
from cryptography.hazmat.primitives import serialization, hashes
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext

# Variables
client_id = '0f985ad1-d73b-4e80-9baf-b61d70479ca3'  # Your client ID
tenant_id = 'ce849bab-cc1c-465b-b62e-18f07c9ac198'  # Your tenant ID
site_url = 'https://bmwgroup.sharepoint.com/teams/TMP-Dev'  # Your SharePoint site URL
thumbprint = 'aec6d07ce9147f561afffe705c6763a9eab05e35'  # Your certificate thumbprint
private_key_pem_path = 'C:/Users/qxz3anc/Downloads/3477a86f-575d-4f4c-962e-bee6f0d80671-YdprF.pem'  # Path to your private key in PEM format

# Load the private key from the PEM file
with open(private_key_pem_path, "rb") as key_file:
    private_key = key_file.read()

# Initialize MSAL Confidential Client with the private key and thumbprint
authority = f"https://login.microsoftonline.com/{tenant_id}"
app = msal.ConfidentialClientApplication(
    client_id,
    client_credential={
        "private_key": private_key.decode("utf-8"),
        "thumbprint": thumbprint
    },
    authority=authority
)

# Get token for SharePoint Online
scopes = ["https://bmwgroup.sharepoint.com/.default"]
result = app.acquire_token_for_client(scopes=scopes)

if "access_token" in result:
    access_token = result['access_token']
    print("Access token acquired!")

    # Connect to SharePoint using the access token
    ctx_auth = AuthenticationContext(site_url)
    ctx_auth.acquire_token_for_user(result['access_token'], ClientCredential(client_id, None))
    ctx = ClientContext(site_url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"Connected to SharePoint site: {web.properties['Title']}")
else:
    print("Failed to acquire token")

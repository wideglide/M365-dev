#!/usr/bin/env python3
import os
from O365 import Account


"""
SharePoint site information retrieval script

Site Properties:
    o create_list()
        Creates a SharePoint list.
    o get_default_document_library()
        Returns the default document library of this site (Drive instance)
    o get_document_library()
        Returns a Document Library (a Drive instance)
    o get_list_by_name()
        Returns a sharepoint list based on the display name of the list
    o get_lists()
        Returns a collection of lists within this site
    o get_subsites()
        Returns a list of subsites defined for this site
    o list_constructor()
    o list_document_libraries()
        Returns a collection of document libraries for this site
    o new_query()
        Create a new query to filter results
    o created:
    o description:
    o display_name:
    o main_resource:
    o modified:
    o name:
    o object_id:
    o protocol:
    o root:
    o site_storage:
    o web_url:


List Properties:
    o create_list_item()
        Create new list item
    o delete_list_item()
        Delete an existing list item
    o get_item_by_id()
        Returns a sharepoint list item based on id
    o get_items()
        Returns a collection of Sharepoint Items
    o get_list_columns()
        Returns the sharepoint list columns
    o list_column_constructor()
        A Sharepoint List column within a SharepointList
    o list_item_constructor()
    o new_query()
        Create a new query to filter results
    o column_name_cw:
    o content_types_enabled:
    o created:
    o created_by:
    o description:
    o display_name:
    o hidden:
    o main_resource:
    o modified:
    o modified_by:
    o name:
    o object_id:
    o protocol:
    o template:
    o web_url:


Document Library Properties
    o get_child_folders()
        Returns all the folders inside this folder
    o get_item()
        Returns a DriveItem by it's Id
    o get_item_by_path()
        Returns a DriveItem by it's absolute path: /path/to/file
    o get_items()
        Returns a collection of drive items from the root folder
    o get_recent()
        Returns a collection of recently used DriveItems
    o get_root_folder()
        Returns the Root Folder of this drive
    o get_shared_with_me()
        Returns a collection of DriveItems shared with me
    o get_special_folder()
        Returns the specified Special Folder
    o new_query()
        Create a new query to filter results
    o refresh()
        Updates this drive with data from the server
    o search()
        Search for DriveItems under this drive.
    o created:
    o description:
    o drive_type:
    o main_resource:
    o modified:
    o name:
    o object_id:
    o owner:
    o parent:
    o protocol:
    o quota:
    o web_url:
"""


def authenticate():
    from dotenv import load_dotenv

    load_dotenv()

    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    pem_certificate = os.getenv("CLIENT_CERTIFICATE")
    cert_pfx_path = os.getenv("CERT_PFX_PATH")

    if cert_pfx_path is not None:
        client_secret = {
        "private_key_pfx_path": cert_pfx_path,
        "passphrase": os.getenv("CERT_PFX_PASSWORD"),
        }
    elif pem_certificate is not None:
        client_secret = {
            "private_key": pem_certificate,
            "thumbprint": os.getenv("CERT_THUMBPRINT"),
        }
    else:
        print("No certificate or PFX path provided. Exiting.")
        exit(1)

    credentials = client_id, client_secret
    account = Account(credentials, auth_flow_type='credentials', tenant_id=tenant_id)
    if account.authenticate():
        print('[pfx_cert] Authenticated!')
    else:
        print('[-] ERROR: Authentication failed.')
        exit(1)
    return account


def dump_attrs(ob):
    for name in dir(ob):
        if not name.startswith('_'):
            attr = getattr(ob, name)
            if callable(attr):
                short_doc = attr.__doc__.strip().split('\n')[0] if attr.__doc__ else ''
                print(f"    o {name}() {short_doc}")
            else:
                print(f"    o {name}:") # {getattr(ob, name)}")


def get_site_info(account, site_path):
    site = account.sharepoint().get_site('root', site_path)
    print(f"SharePoint site: {site.name}")
    print(f"  site_id: {site.object_id}")
    print(f"  site_url: {site.web_url}")
    print(f"  site_description: {site.description}")
    print(f"  site_created: {site.created}")
    print(f"  site_last_modified: {site.modified}")

    # Get site lists
    print("=== Site Lists ===")
    lists = site.get_lists()
    for lst in lists:
        num_items = len(lst.get_items())
        print(f"  list: {lst.name} (id: {lst.object_id}) - {num_items} items")

    # Get site document libraries
    print("=== Document Libraries ===")
    doc_libs = site.list_document_libraries()
    for lib in doc_libs:
        print(f"  document library: {lib.name} (id: {lib.object_id})")

    return site


if __name__ == "__main__":
    account = authenticate()
    site = get_site_info(account, os.getenv('SITE_PATH'))

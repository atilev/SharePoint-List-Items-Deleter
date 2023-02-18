#============================Importing Library============================
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

print("SHAREPOINT LIST ITEMS DELETER")

#============================Autentication============================
#enter your sharepoint main site url 
url = 'https://xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
#enter your user name as email
username = 'xxxxxxxxxxxxxxx@xxxxxx.com'
#enter your password
password = 'xxxxxxxxxxx'
 
ctx_auth = AuthenticationContext(url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Authentication successful") 


#============================Main============================
#number of items to get at once (max 5000)
batch = 5000

#while&try&except to check whether list file exist
while True:
    try:
        #Sharepoint List file name
        sharepoint_list = input("Please enter the Sharepoint list file name: ")
        
        #connecting to sharepoint list to read and write
        target_list = ctx.web.lists.get_by_title(sharepoint_list)
        list_tasks = ctx.web.lists.get_by_title(sharepoint_list)
        #5000 items limit to read from sharepoint
        items = list_tasks.items.get().top(batch).execute_query()
        break
    
    except:
        next_step = input("File is not found! To try again press enter or (e) to exit").lower() 
        if next_step == "e":
            exit()
        else:
            continue

#checking whether list is empty
if len(items) == 0:
    print(f"No items found in \"{sharepoint_list}\" list file.")

# # Option 2: remove a list item (with an option to restore from a recycle bin)
# item_id = items[i].id
# target_list.get_item_by_id(item_id).recycle().execute_query()

else:
    print("\nDeleting items...")

    #counter for deleted items
    total = 0

    #deleting all items with while & for loop in batches
    while len(items) > 0:

        # Option 1: Permanently remove a list item
        for i in range(len(items)):
            total += 1
            print(f"{total}", end="\r")
            item_id = items[i].id
            target_item = target_list.get_item_by_id(item_id).delete_object().execute_query()
            
        #Autentication again due to timeout
        ctx_auth = AuthenticationContext(url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(url, ctx_auth)
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            
        target_list = ctx.web.lists.get_by_title(sharepoint_list)
        list_tasks = ctx.web.lists.get_by_title(sharepoint_list)
        items = list_tasks.items.get().top(batch).execute_query()

    print(f"{total} items have been deleted in \"{sharepoint_list}\" list file.")
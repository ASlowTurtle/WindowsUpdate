# This script requires pywin32
import win32com.client
# https://learn.microsoft.com/en-us/windows/win32/api/wuapi/nf-wuapi-iupdatesearcher-search
# Inspiration: https://github.com/OSDeploy/OSD/blob/master/Public/OSDCloudTS/Start-WindowsUpdate.ps1
"""
IsInstalled=0 and Type='Software'
IsAssigned=0
AutoSelectOnWebSites=0
BrowseOnly=1
"""

"""
Exitcodes:
1       Com error during instance creation.
2       Error searching for updates. Probably dua to a bad search string.
3       Error downloading updates.
4       Error starting update installation.
"""

search_string = "IsInstalled=0 and Type='Software'"
error_encountered = 0
print("[+] Step 1 creating com instances.")
try:
    wu_object = win32com.client.Dispatch('Microsoft.Update.Session')
    wu_searcher = wu_object.CreateUpdateSearcher()
    wu_downloader = wu_object.CreateUpdateDownloader()
    wu_installer = wu_object.CreateUpdateInstaller()
    wu_update_collection = win32com.client.Dispatch('Microsoft.Update.UpdateColl')
except:
    print("[-] Error creating com instances. Exiting.")
    exit(1)

print("[+] Step 2 searching for updates.")
try:
    updates = wu_searcher.Search(search_string)
except:
    print("[-] Error searching for updates using string [{0}]".format(search_string))
    exit(2)

try:
    print("\t*******************")
    print("\tFound [{0}] Updates using search string [{1}]".format(updates.Updates.Count,search_string))
    for single_update in updates.Updates:
        print("\t-------------------")
        print("\tTitle:\t\t\t{0}".format(single_update.Title))
        print("\tRebootRequired:\t\t{0}".format(single_update.RebootRequired))
    print("\t-------------------")
except:
    print("[-] Aborting due to error while printing update information.")

print("[+] Step 3 creating update collection.")
for single_update in updates.Updates:
    print("\tTrying to add update [{0}] to install list.".format(single_update.Title))
    try:
        wu_update_collection.Add(single_update)
    except:
        print("[-] Error adding update [{0}] to install list. Trying to continue with other updates.".format(single_update.Title))

if input("Do you want to continue? [y/n]") != "y":
    print("User exit.")
    exit(0)


print("[+] Step 4 downloading updates.")
if wu_update_collection.Count > 0:
    try:
        wu_downloader.Updates = wu_update_collection
        wu_downloader.Download()
    except:
        print("[-] Error downloading updates.")
        exit(3)
    
print("[+] Step 5 installing updates.")
try:
    wu_installer.Updates = wu_update_collection
    wu_installer.Install()
except:
    print("[-] Error installting updates.")
    exit(4)


"""
object_methods = [method_name for method_name in dir(Updates)
                  if callable(getattr(Updates, method_name))]
                 
dir(Updates)
dir(Updates.Updates[0])
Updates.Updates[0].Title
"""
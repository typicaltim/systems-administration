# My Systems Administration Repository

## Deploy-O365
This script utilizes the workflow powershell feature to run O365 installations in parallel for rapid deployment. I used this to deploy O365 in an organization since O365 does not have an MSI installer that plays nicely with Group Policy. It had about a fairly high success rate for machines that were online and had WinRM running. A few pieces of software would cause conflicts with the O365 installer and I had to go and touch those machines manually. But this script saved somewhere around 100 hours of time for me. I hope that someone out there finds this and can use it to save even more time for themselves. For that special someone, here are some things you might want to know:

* After pulling a list of computers from Active Directory
* The script first checks if it can ping the device.
* Then it checks to see if it can use WinRM.
* Then it checks that the hostname reported by the machine matches the computer name from AD it is attempting to reach (this ensures your DNS doesn't screw you over)
* Then it checks if O365 is installed or not
* Then it checks if there is a user logged in or not (so that when the uninstaller runs and closes Office apps, users don't lose data they have not saved). If a user is logged in it does not touch the machine.
* The source files at the top of the script are copied to the target machine to be used locally when the uninstall/install commands are invoked.
* The uninstallers/installers run.
* The script verifies the install worked.
* The local copy of the source files are removed from the target machine.
* After each host, the console outputs a status of it's progress
* After the script finishes the list of machines from AD, it outputs a summary of the results

This script is not perfect. There are plenty of ways to improve it. But I don't have reason/time to improve it any further right now. However, there are so many wonderful little snippets of code in here that I wanted to save for myself and others that I had to put it on Github.

This script is a decent framework/starting point for all sorts of tasks and can be gutted to be used for a plethora of systems administration chores. Good luck out there! <3

UserLocal.vbs
- Run on every system using admin privileges (System account) -- only needs to be run once on each system, can be run multiple times without impact
- Creates a WMI Namespace called UserLocal (change this on line 3 if you want to name it something else)
- Sets permissions on the namespace to allow normal users to write to this namespace
- Creates a WMI Class called Outlook_Configuration (change this on line 4 if you want to name it something else)

Outlook-Config.vbs
- Can be run on a system after 
- Run on every system as the user (important because the configuration is specific to the user) for every user you want to collect data for as often you want to collect inventory
- Creates an object instance of the WMI Class Outlook_Configuration in the namespace called UserLocal (both created by running UserLocal.vbs and renamable on lines 3 and 4)
- WMI Object instance contains the username as well as wether cached mode exists or not for every profile the user has. This is extracted from the registry and then decoded.

Test
- Manually run UserLocal.vbs on a test system as a local admin.
- Manually run Outlook-Config.vbs on the same test system as a user with an Outlook profile.
- Verify that the namespace is properly created, that the class is properly created, and that at least one instance of the class is created with appropriate data

Hardware Inventory
- Add the Outlook_Configuration class to hardware inventory.
- Create queries, collections, and reports once the data is collected.

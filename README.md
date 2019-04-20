# remote-assistant
A short Powershell script to run a GUI for easier remote support.
You can enter all or part of the computer name in the text box. The script searches the AD for the name.
When a PC is found, the connection is initiated. If more than one PC is found, a list will be provided.
If the computer name was not found in the AD, a connection to the "raw" input is established. E.g. for an IP address

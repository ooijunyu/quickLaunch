# quickLaunch
# Purpose
I am working in a company that has a lot of documents, Lotus Notes application, desktop and web applcations scattered everywhere which I need to access on a daily basis for my work. Hence, I wrote the script to help me to launch the app easily instead of searching them up and down in various places.

# How to use
1. Grab all 3 files to a folder.
2. Populate mySource.txt in the format: name, "link"
3. The link part can take in any link that can be open by Windows Run
4. Double click quickLaunch.vbs to start.
5. Type in a matching pattern to launch the link.

# How it works
1. The script first read in mySource.txt as key:item pair
2. The script search the keys of mySource.txt for that matches the characters of user input in order (e.g. gg matches Google)
3. The script launches the link by explorer.exe

# Pin to taskbar
1. Right click quickLaunch.vbs to create shortcut
2. Right click the shortcut to edit the Target field. At WScript.exe and a whitespace in front of the text (script location) in the Target field.
3. Right click the shortcut and select pin to taskbar.

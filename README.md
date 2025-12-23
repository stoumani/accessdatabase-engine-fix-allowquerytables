# accessdatabase-engine-fix-allowquerytables
To fix the issue that is occurring for Excel Pivot Tables since update KB5072033 for Windows, adds a registry value that allows query remote tables. Includes a .ps1 script that can be run, as well as a .bat script

This fix is for companies that use files with old pivot tables using 2016 Access Database Engine (in this case, 32-bit) to connect to an Access Database. The update broke refreshing the data when attempting to connect to the database, but after the fix, it was able to connect to the DSNSQL. Have not tested 64-bit, as it is not used in my case, but you are free to try it

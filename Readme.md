Main Focused Program File_Usage_Merge.py

For Output:
File Used - All_Inventory - Keystone Medium.xlsx
Sheet Used - Online PGM
start_col = column number for first file end_col = last file column number

Outputs Folder: Contains text files having console output


Programs:

Commonality_Score.py

The grouping logic is commonality score based
The program having maximum commonality score is taken and check with other files where it is used along with other files
Based on that it will formed a group 

If unique grouping is possible then different groups will be created but if file with maximum commonality is used with all other files then it will create only one group of all files

------------------------------------------------------------------------------------------------------

File_Usage.py

Grouping Logic
There will be a program which is using maximum files for e.g. program using 20 files 
other program using files less than 20

We will create group of that 20 files and display the program which are using those
and continue the same logic

the second group will contain files which are not in first group 
but the programs will be listed which are using files from group 2 also some files from group 1
and continue the same logic 

------------------------------------------------------------------------------------------------------

File_Usage_Merge.py

The grouping logic is similar to temp.py

There will be a program which is using maximum files for e.g. program using 20 files 
other program using files less than 20

We will create group of that 20 files and display the program which are using those
and continue the same logic

the second group will contain files which are not in first group 
but the programs will be listed which are using files from group 2 also some files from group 1
and continue the same logic 

----Changes---
as it is creating grouping as temp.py 
but for example if there is a group with 1 or 2 files and the program which is using files from this group
the remaining other files of this program are available in other group so this group will be get merged into that
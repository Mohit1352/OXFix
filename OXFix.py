#!/usr/bin/env python3
#Place this file in the directory of all broken PPTs or DOCs, or provide directory name as input, or pipe it from another command.
import os
import shutil
import sys

directories=["."]
filetypes=[".ppt",".doc"]
mode='0' #default mode, will go through all formats, else index+1 for each option in the list.

for i in sys.argv[1:]:
    if "filemode" in i:
        mode=i.split("=")[-1]
    else:
        directories.append(i)

if mode!='0':
    filetypes=filetypes[int(mode)-1:int(mode)]

print("Pipe may be empty, press Enter key to continue if no output is seen.")

for line in sys.stdin:
    if line=="\n":
        break
    line=line.replace(" ","")
    line=line.replace("\n","")
    directories.append(line)

print("Mode: ",end="")
[print(i[1:],end=" ") for i in filetypes]
print()
print("Current Working Directory:",os.getcwd())

fc=0

#Will work only on PPT files in the directory.
for d in directories:
    files=os.listdir(f"./{d}")
    files2=[]
    for i in filetypes:
        files2+=[j for j in files if i in j]
    files2=list(set(files2))
    if len(files2)>0:
        print("\nFolder",d,":")
        print(len(files2),"files to fix.")
        os.chdir(f"./{d}")
        for i in files2:
            try:
                print(f"[{files2.index(i)+1}]",i,end=" ...")
                shutil._unpack_zipfile(i,f"./{i}_Temp")
                os.chdir(f"./{i}_Temp")
                shutil._make_zipfile(f"../{i}",".")
                os.chdir("..")
                os.remove(i)
                shutil.rmtree(f"{i}_Temp",ignore_errors=True)
                os.rename(f"{i}.zip",i)
                print("Done.")
                fc+=1
            except:
                print("No issue detected.")
        if d!=".":
            os.chdir("..")
if fc>0:
    print("All files fixed.")
else:
    print("No files found matching criteria in current directory.")

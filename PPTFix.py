#Place this file in the directory of all broken PPTs, or provide directory name as input, or pipe it from another command.
import os
import shutil
import sys

directories=["."]

for i in sys.argv[1:]:
    directories.append(i)

print("Pipe may be empty, press Enter key to continue if no output is seen.")

for line in sys.stdin:
    if line=="\n":
        break
    line=line.replace(" ","")
    line=line.replace("\n","")
    directories.append(line)

#Will work only on PPT files in the directory.
for d in directories:
    ppts=[i for i in os.listdir(f"./{d}") if ".ppt" in i]
    if len(ppts)>0:
        print("Folder",d,":")
        print(len(ppts),"files to fix.")
        os.chdir(f"./{d}")
        for i in ppts:
            print(f"[{ppts.index(i)+1}]",i,end=" ...")
            shutil._unpack_zipfile(i,f"./{i}_Temp")
            os.chdir(f"./{i}_Temp")
            shutil._make_zipfile(f"../{i}",".")
            os.chdir("..")
            os.remove(i)
            shutil.rmtree(f"{i}_Temp",ignore_errors=True)
            os.rename(f"{i}.zip",i)
            print("Done.")
        if d!=".":
            os.chdir("..")
print("All files fixed.")

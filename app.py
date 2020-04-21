import os
import binhex
import re
import zipfile
import shutil
import glob
import time


def main():
    print("Starting ...")

    # Empty tmp
    if(os.path.isdir('.tmp')):
        print("Empting .tmp...")
        empty_folder('.tmp')

    if os.path.exists("out.xlsm"):
        print("Removing old out.xlsm...")
        os.remove("out.xlsm")

    # password = "macro"
    password = b'EAE846DB88F888F8770889F8D3CA304AB2466415D496EBE5B7A36F62928053F0F248C55F73'
    pattern = b'DPB\=\"(.*?)\"'


    xlsmCounter = len(glob.glob1('.',"*.xlsm"))
    if(xlsmCounter > 1):
        print("Only one .xlsm file must be on this folder!")
        time.sleep(3)
        exit()

    # Extract all the contents of zip file in current directory
    os.chdir(".")
    with zipfile.ZipFile(glob.glob("*.xlsm")[0], 'r') as zipObj:
        print("Extracting xlsm...")
        zipObj.extractall('.tmp')

    with open('.tmp/xl/vbaProject.bin', 'rb') as f:
        print("Reading vbaProject.bin...")
        hex = f.read()

        # Adjust password lenght
        match = re.search(pattern, hex).group(1)
        if len(match) > len(password):
            print("Adjusting password lenght...")
            nzeros = len(match)-len(password)
            print("Added {} zeros after the password".format(nzeros))

            for i in range(nzeros):
                password = password + b'0'

            print("New password: " + str(password))

        # Define new bin
        password = b'DPB="' + password + b'"'

        print("Changing DPB value...")
        out = re.sub(pattern, password, hex)

        # Write bin
        with open('.tmp/xl/vbaProject.bin', 'wb') as f:
            print("Writing vbaProject.bin...")
            f.write(out)

            # Create out.xlsm
            print("Creating out.xlsm...")
            shutil.make_archive('out', 'zip', '.tmp')
            os.rename("out.zip", "out.xlsm")

            # Success
            print("SUCCESS: The file out.xml has the new VBA password of 'macro'")
            input("Press Enter to exit...")


def empty_folder(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


if __name__ == "__main__":
    main()

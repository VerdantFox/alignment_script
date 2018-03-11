How run alignment script:


1.) If not already done so, install python (version 3.X) from
    "https://www.python.org/downloads/" making sure to add python to PATH
    when that question comes up. (Without adding to PATH windows command
    prompt will not recognize python).

2.) From the start menu type:

    cmd

    The command prompt will open.

3.) Open the folder "alignment_script" in file explorer.

4.) Copy the address to this folder from the top bar. The address
    should look something like:

    C:\Users\Yourname\Desktop\Twin_Alignment\alignment_script

5.) In the command prompt type "cd " then right-click to paste folder path,
    hit enter. This should change your path from you command prompt to the
    path you pasted (where the files are located). e.g.

    cd C:\Users\Yourname\Desktop\Twin_Alignment\alignment_script

6.) Activate the virtual environment (contains outside libraries needed
    for alignment script) by copy and pasting in the following line:

    venv\Script\activate.bat

    (change back-slashes to forward slashes if working on Mac)
    You should notice a "(venv)"" before your address in the command
    prompt now. e.g.

    (venv) C:\Users\Yourname\Desktop\Twin_Alignment\alignment_script>

7.) Move all excel files you would like to align into the folder you
    have open which contains the file alignment.py and remove any other
    excel file you don't want aligned.

8.) Into the command prompt type:

    python alignment.py

    This should lead to a series of messages explaining where the script
    is at in the alignment process, ending in "****** FILE SAVED **********"

9.) Your alignment file should now be found in your current folder. The file
    will be named "xxxx-xx-xx_combined_matrix" where the x's are today's date.

Note: I don't think this will happen, but if while in (venv), after running
      "python alignment.py" you get an error message saying something like
      "Could not find blahblah module", type into the command prompt:

      pip install -r requirements.txt

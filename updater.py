import time, os, sys, shutil

old = sys.argv[1]
new = sys.argv[2]

time.sleep(2)  # wait for main app to exit

shutil.move(new, old)

os.startfile(old)
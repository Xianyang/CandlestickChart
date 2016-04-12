import subprocess,os
import shutil
import time
subprocess.call("C:\Program Files (x86)\WinSCP\WinSCP.exe /console /script=M:\Dashboard_v2\conf\dailyUpdateToServer.txt /xmllog=M:\Dashboard_v2\conf\conf\updateFilesToServerRecord.log")

time.sleep(20)

paths = ['../Daily/data/','../Daily/chart/']
folderNames = ['COMMODITIES','Asia_INDEX','Other_INDEX','EQUITY','CURNCY','METAL']
for path in paths:
    for folderName in folderNames:
        folder = path+folderName
        for the_file in os.listdir(folder):
            file_path = os.path.join(folder, the_file)
            print "delete "+file_path
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                #elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception, e:
                print e

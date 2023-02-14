import sys
import os
import shutil
from collections import defaultdict
from PyPDF2 import PdfMerger

input_dir = sys.argv[1]
output_dir = sys.argv[2]
try:
    os.mkdir(output_dir)
except:
    print(output_dir + ' already exists')

#create dictionary of files
transcript_dict = defaultdict(list)
for filename in os.listdir(input_dir):
    parts = filename.split('_')
    appid = parts[1]
    transcript_dict[appid].append(input_dir + os.sep + filename)

merged_dict = defaultdict(list)
for appid in transcript_dict:
    files = transcript_dict[appid]
    merger = PdfMerger()
    try:
        if len(files) > 1:
            for f in files:
                merger.append(f)
            file_name = output_dir + os.sep + appid + '_merged_transcript.pdf'
            merger.write(file_name)
            merger.close()
            merged_dict[appid].append(file_name)
        else:
            file_name = output_dir + os.sep + appid + '_single_transcript.pdf'
            shutil.copy(files[0],file_name)
            merged_dict[appid].append(file_name)
    except:
        print('Error: ' + appid)
        num = 1
        for f in files:
            file_name = output_dir + os.sep + appid + '_' + str(num) + '.pdf'
            shutil.copy(f,file_name)
            num += 1
            merged_dict[appid].append(file_name)
            
print(merged_dict)
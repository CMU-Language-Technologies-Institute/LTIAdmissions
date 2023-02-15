import sys
from collections import defaultdict

#Create a csv that combines all transcripts for an applicant into one row

transcript_csv = sys.argv[1]
output_csv = sys.argv[2]

transcript_dict = defaultdict(list)
max_len = 4
with open(transcript_csv, 'r') as f:
    for line in f.readlines():
        line_parts = line.split(',')
        appid = line_parts[0].strip()
        url = line_parts[1].strip()
        transcript_dict[appid].append(url)

output_file = open(output_csv, 'w')
output_file.write('appid,transcript URL 1,transcript URL 2,transcript URL 3,transcript URL 4\n')
num_ids = 0
for appid in transcript_dict:
    #print(appid)
    num_ids += 1
    url_list = transcript_dict[appid]
    final_list = []
    for i in range(0,4):
        url = url_list[i] if len(url_list) > i else ''
        final_list.append(url)
    output_file.write(appid + ',' + ','.join(final_list) + '\n')
    
output_file.flush()
output_file.close()
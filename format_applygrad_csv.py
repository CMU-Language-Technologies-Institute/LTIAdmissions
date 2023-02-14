import sys

###################################################################################
# Takes a txt file from the apply grad dump for a program (such as MIIS, MCDS,
# ans MSAII) and creates a csv, which can be loaded by pandas.  Renames
# duplicate column names with "_N" so that each column has a unique name.
###################################################################################
def format_csv(input_file):
    output_file_name = input_file.split('.')[0] + '.csv'
    output_file = open(output_file_name, 'w', encoding="utf-8", errors='ignore')
    is_header = True
    header_list = []
    header_duplicate_dict = {}
    linenum = 1
    with open(input_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.replace('"', '')
            linenum += 1
            line_parts = line.split('^')
            if is_header:
                for header_name in line_parts:
                    if header_name in header_list:
                        header_duplicate_dict[header_name] = header_duplicate_dict.get(header_name, 1) + 1
                        header_name = header_name + '_' + str(header_duplicate_dict[header_name])
                    header_list.append(header_name)
                new_string = ','.join(map(lambda x: "\"" + x.strip() + "\"", header_list))
                is_header = False
            else:
                new_string = ','.join(map(lambda x: "\"" + x.strip() + "\"", line_parts))
            #print(new_string)
            output_file.write(new_string)
            output_file.write('\n')
            output_file.flush()
    output_file.close()

if __name__ == '__main__':
    input_file = sys.argv[1]
    format_csv(input_file)
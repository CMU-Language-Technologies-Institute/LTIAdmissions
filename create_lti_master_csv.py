import sys
import pandas as pd
import re
import math
import xlsxwriter
from annotate_applicants_api import annotate_applicants_api
import argparse

# Initialize parser
parser = argparse.ArgumentParser()
 
# Adding optional argument
parser.add_argument("-mcds", "--mcds", help = "MCDS csv formatted input from admissions system")
parser.add_argument("-msaii", "--msaii", help = "MSAII csv formatted input from admissions system")
parser.add_argument("-miis", "--miis", help = "MIIS csv formatted input from admissions system")
parser.add_argument("-r", "--resumes", help = "csv with columns appid and Resume URL")
parser.add_argument("-t", "--transcripts", help = "csv with columns appid, transcript URL 1, transcript URL 2, transcript URL 3, transcript URL 4")
parser.add_argument("-o", "--output", help = "The name of the csv output file to write to")
 
# Read arguments from command line
args = parser.parse_args()

if args.mcds == None or args.miis == None or args.msaii == None or args.output == None:
    raise ValueError("MCDS (-mcds), MIIS(-miis), MSAII(-msaii), and output (-o) csv files are required")

mcds_input_file = args.mcds
msaii_input_file = args.msaii
miis_input_file = args.miis

mcds_df = pd.read_csv(mcds_input_file)
msaii_df = pd.read_csv(msaii_input_file)
miis_df = pd.read_csv(miis_input_file)
df = pd.concat([mcds_df, msaii_df, miis_df]).drop_duplicates(subset=['appid'], keep='last').reset_index(drop=True)

annotate_applicants = annotate_applicants_api(df)
annotate_applicants.find_top_program()
annotate_applicants.map_applygrad_columns()
annotate_applicants.add_gpa_and_gre_columns()
annotate_applicants.add_language_scores()
if args.resumes != None:
    annotate_applicants.add_resume_column(args.resumes)
if args.transcripts != None:
    annotate_applicants.add_transcript_columns(args.transcripts)
annotate_applicants.sort_master()
annotate_applicants.write_excel(args.output)
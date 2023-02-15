import sys
import pandas as pd
import re
import math
import xlsxwriter
from annotate_applicants_api import annotate_applicants_api
import argparse

# Initialize parser
parser = argparse.ArgumentParser()
 
# Adding arguments
parser.add_argument("-i", "--input", help = "The initial program csv formatted input from admissions system")
parser.add_argument("-n", "--name", help = "The name of the program: MCDS, MIIS, MSAII")
parser.add_argument("-r", "--resumes", help = "csv with columns appid and Resume URL")
parser.add_argument("-t", "--transcripts", help = "csv with columns appid, transcript URL 1, transcript URL 2, transcript URL 3, transcript URL 4")
parser.add_argument("-o", "--output", help = "The name of the csv output file to write to")
parser.add_argument("-ra", "--resume_annotations", help = "A csv file with the resume annotations")
parser.add_argument("-ta", "--transcript_annotations", help = "A csv file with the transcript annotations")
 
# Read arguments from command line
args = parser.parse_args()

if args.input == None or args.name == None or args.output == None:
    raise ValueError("Input (-i), name (-n), and output (-o) csv files are required")

df = pd.read_csv(args.input)

annotate_applicants = annotate_applicants_api(df)
annotate_applicants.map_applygrad_columns()
annotate_applicants.program_rank(args.name)
annotate_applicants.add_gpa_and_gre_columns()
annotate_applicants.add_language_scores()
if args.resumes != None:
    annotate_applicants.add_resume_column(args.resumes)
if args.transcripts != None:
    annotate_applicants.add_transcript_columns(args.transcripts)
if args.resume_annotations != None:
    annotate_applicants.add_resume_annotations(args.resume_annotations)
if args.transcript_annotations != None:
    annotate_applicants.add_transcript_annotations(args.transcript_annotations)
annotate_applicants.sort_program()
annotate_applicants.write_excel(args.output)
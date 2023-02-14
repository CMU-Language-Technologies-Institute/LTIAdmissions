import sys
import pandas as pd
import re
import math
import xlsxwriter

#Program constants
MCDS = 'Master of Computational Data Science'
MSAII = 'Master of Science in Artificial Intelligence and Innovation'
MIIS = 'M.S. in Intelligent Information Systems'
LTI_PROGRAM_MAP = {MCDS: 'MCDS', MSAII: 'MSAII', MIIS: 'MIIS'}
PROGRAM_ABV_MAP = {'MCDS': MCDS, 'MSAII': MSAII, 'MIIS': MIIS}

#Language constants
#Maps IELTS and DUO scores to TOEFL scores
ENGLISH_ISO_CODES = ['US', 'UK', 'AU']
IELTS_TOEFL_SPK_DICT = {0.0: 0, 1.0:3, 2.0:6, 3.0:8, 4.0:11, 4.5:13, 5.0:15, 5.5:17, 
                        6.0:19, 6.5:21, 7.0:23, 7.5:25, 8.0:28, 8.5:29, 9.0:30}
IELTS_TOEFL_TOT_DICT = {0.0: 0, 1.0:8, 2.0:16, 3.0:24, 4.0:31, 4.5:33, 5.0:40, 5.5:53, 
                        6.0:69, 6.5:86, 7.0:97, 7.5:106, 8.0:112, 8.5:116, 9.0:119}
DUO_TOEFL_SPK_DICT = {110:18, 105:21, 110:22, 115:23, 120:25, 125:26, 130:27, 135:28, 140:30, 145:30, 150:30, 155:30, 160:30}
DUO_TOEFL_TOT_DICT = {100:75, 105:82, 110:88, 115:94, 120:100, 125:105, 130:110, 135:114, 140:119, 145:120, 150:120, 155:120, 160:120}

class annotate_applicants_api:
    
    def __init__(self, df):
        self.df = df
        grades_df = pd.read_csv('grade_ranges.csv')
        self.grades_df = grades_df
        
        
    #Determines which program is the applicant's first choice and whether 
    #the first choice is an LTI program
    def find_top_program(self):
        df = self.df
        top_program_list = df['program_name'].to_list()
        top_programs = []
        is_lti_program_list = []
        #program_ranks.append('Program Rank')
        for program_choice in top_program_list:
            is_lti_program = 'N'
            if program_choice in LTI_PROGRAM_MAP:
                is_lti_program = 'Y'
                program_choice = LTI_PROGRAM_MAP[program_choice]
            is_lti_program_list.append(is_lti_program)
            top_programs.append(program_choice)
        df.insert(loc=9, column='Is LTI Program', value=is_lti_program_list)
        df.insert(loc=10, column='TOP program', value=top_programs)
     
    #Determines whether the top program choice is an LTI program and
    def program_rank(self, program_abv):
        df = self.df
        #Top Program        
        top_program_list = df['program_name'].to_list()
        top_programs = []
        is_lti_program_list = []
        #program_ranks.append('Program Rank')
        for program_choice in top_program_list:
            is_lti_program = 'N'
            if program_choice in LTI_PROGRAM_MAP:
                is_lti_program = 'Y'
                program_choice = LTI_PROGRAM_MAP[program_choice]
            is_lti_program_list.append(is_lti_program)
            top_programs.append(program_choice)
        df.insert(loc=9, column='Is LTI Program', value=is_lti_program_list)
        
        #Program Rank        
        program_list_1 = df['program_name'].to_list()
        program_list_2 = df['program_name_2'].to_list()
        program_list_3 = df['program_name_3'].to_list()
        program_list_4 = df['program_name_4'].to_list()
        program_lists = zip(program_list_1,program_list_2,program_list_3,program_list_4)
        program_ranks = []
        #program_ranks.append('Program Rank')
        for program_choices in program_lists:
            program = PROGRAM_ABV_MAP[program_abv]
            choice_num = 1
            program_rank = str(5)
            for program_choice in program_choices:
                #print(program_choice)
                program_choice_string = str(program_choice)
                if program_choice_string.lower() == program.lower():
                    program_rank = str(choice_num)
                choice_num += 1
            program_ranks.append(program_rank)
        df.insert(loc=10, column='Prog Rank', value=program_ranks)
    
    #Rename ambiguous columns to more logical names
    #Removed unneeded columns
    def map_applygrad_columns(self):
        header_dict = self.df.head().to_dict()
        #Applygrad specific - drop columns with SPACER as name 
        for header in header_dict:
            if 'SPACER' in header:
                self.df = self.df.drop(header, axis=1)
        #Rename Applygrad columns to logical names
        self.df = self.df.rename(columns={'section1': 'TOEFL_speak', 
                                'section2': 'TOEFL_listen', 
                                'section3': 'TOEFL_write', 
                                'essay': 'TOEFL_essay', 
                                'total': 'TOEFL_total', 
                                'essay_2': 'TOEFL_essay_2', 
                                'totalmb': 'TOEFL_total_mb',
                                'section1mb': 'TOEFL_speak_mb', 
                                'section2mb': 'TOEFL_listen_mb', 
                                'section3mb': 'TOEFL_write_mb', 
                                'listeningscore': 'IELTS_listeningscore', 
                                'readingscore': 'IELTS_readingscore', 
                                'writingscore': 'IELTS_writingscore',
                                'speakingscore': 'IELTS_speakingscore', 
                                'overallscore': 'IELTS_overallscore',
                                'conversationscore': 'DUO_conversationscore', 
                                'productionscore': 'DUO_productionscore',
                                'literacyscore': 'DUO_literacyscore', 
                                'comprehensionscore': 'DUO_comprehensionscore',
                                'overallscore_2': 'DUO_overallscore', 
                                'scale': 'DUO_scale'})
        self.df.pop('email')
        self.df.pop('title')
        self.df.pop('middlename')
        self.df.pop('organization')
        self.df.pop('gender_identity')

    #Calculates the normalized GPA and adds a Norm GPA column to start of spreadsheet
    #Adds GRE columns to the front of the spreadsheet where they are more visible
    #Adds a column that combines GRE quantitative score and GPA, which can be used as an initial ranking
    def add_gpa_and_gre_columns(self):
        grades_df = self.grades_df
        df = self.df
        univ_list = df['UnivName'].to_list()
        gpa_list = df['gpa'].to_list()
        gpa_major_list = df['gpa_major'].to_list()
        gpa_scale_list = df['GPAScale'].to_list()
        gpa_tuple_list = zip(gpa_list, gpa_scale_list, univ_list, gpa_major_list)
        gpa_norm_list = []
        gpa_med_list = []
        gpa_major_norm_list = []
        for gpa_tuple in gpa_tuple_list:
            gpa = str(gpa_tuple[0])
            gpa = gpa.replace(',','.')
            gpa = gpa.replace('%','')
            gpa = float(re.sub("[^0-9.]", "0", gpa))
            gpa_scale = 4.0
            #Corner cases
            if gpa_tuple[1] == 'Other':
                gpa_scale = 30.0
            elif gpa_tuple[1] == '7':
                gpa_scale = 10.0
            #Happy Path
            else:
                gpa_scale = float(re.sub("[^0-9,.]", "", gpa_tuple[1]))
            gpa_norm = gpa/gpa_scale
            gpa_norm_list.append(gpa_norm)

            univ_df = grades_df[grades_df['University Name'] == gpa_tuple[2]]
            univ_df_2 = univ_df[univ_df['Scale'] == gpa_scale]
            if not univ_df_2.empty:
                gpa_med_list.append(float(univ_df_2['Med'].values[0])/gpa_scale)
            else:
                gpa_med_list.append('N/A')
            gpa_major_value = str(gpa_tuple[3])
            gpa_major_value = gpa_major_value.replace(',','.')
            gpa_major_value = gpa_major_value.replace('%','')
            gpa_major_value = float(re.sub("[^0-9.]", "0", gpa_major_value))  
            gpa_major_norm = gpa_major_value/gpa_scale
            gpa_major_norm_list.append(gpa_major_norm)

        df.insert(loc=11, column='Norm GPA', value=gpa_norm_list)
        df.insert(loc=12, column='Median GPA/univerity', value=gpa_med_list)
        df.insert(loc=13, column='Norm GPA major', value=gpa_major_norm_list)

        #GRE columns
        grev_scores = df['verbalscore'].to_list()
        grev_percentiles = df['verbalpercentile'].to_list()
        greq_scores = df['quantitativescore'].to_list()
        greq_percentiles = df['quantitativepercentile'].to_list()

        df.insert(loc=13, column='GRE V Score', value=grev_scores)
        df.insert(loc=14, column='GRE V %', value=grev_percentiles)
        df.insert(loc=15, column='GRE Q Score', value=greq_scores)
        df.insert(loc=16, column='GRE Q %', value=greq_percentiles)
        
        #GRE quatitative + GPA column
        norm_gpas = df['Norm GPA'].to_list()
        greq_scores = df['quantitativepercentile'].to_list()
        gpa_gre_tuple = zip(norm_gpas, greq_scores)
        gpa_gre_combined = []
        for gpa_gre in gpa_gre_tuple:
            gpa_score = 100 * float(gpa_gre[0])
            gre_score = 0
            if not math.isnan(gpa_gre[1]):
                gre_score = int(gpa_gre[1])
            combinded_score = gpa_score + gre_score
            gpa_gre_combined.append(combinded_score)
        df.insert(loc=17, column='GPA+GREQ', value=gpa_gre_combined)
        
    def add_language_scores(self):
        #Language Additions
        #English requires
        df = self.df
        grades_df = self.grades_df
        iso_codes = df['iso_code'].to_list()
        native_tongues = df['native_tongue'].to_list()
        univ_list = df['UnivName'].to_list()
        language_tuple = zip(iso_codes, native_tongues, univ_list)
        eng_univ_list = []
        engreq_list = []
        for language_values in language_tuple:
            iso_code = language_values[0]
            native_tongue = language_values[1].lower()
            univ_df = grades_df[grades_df['University Name'] == language_values[2]]
            english_univ = 'N'
            if not univ_df.empty:
                english_univ = univ_df['English Speaking'].values[0]
            eng_univ_list.append(english_univ)
            engreq = 'Y'
            if (iso_code in ENGLISH_ISO_CODES and native_tongue == 'english'):
                engreq = 'N'
            engreq_list.append(engreq)

        df.insert(loc=18, column='ENG University', value=eng_univ_list)
        df.insert(loc=19, column='ENGREQ', value=engreq_list)

        #Speaking scores
        toefl_speak_scores = df['TOEFL_speak'].to_list()
        toefl_speak_mb_scores = df['TOEFL_speak_mb'].to_list()
        ielts_speak_scores = df['IELTS_speakingscore'].to_list()
        duo_speak_scores = df['DUO_conversationscore'].to_list()
        speak_scores_tuple = zip(toefl_speak_scores, toefl_speak_mb_scores,ielts_speak_scores, duo_speak_scores) 
        speak_scores = []
        for speak_score in speak_scores_tuple:
            toefl_speak_score = 0
            #print(speak_score[1])
            if not math.isnan(speak_score[0]):
                toefl_speak_score = int(speak_score[0])
            toefl_speak_mb_score = 0
            if not math.isnan(speak_score[1]):
                toefl_speak_mb_score = int(speak_score[1])
            toefl_max_speak = max(toefl_speak_score, toefl_speak_mb_score)

            norm_ielts_speak_score = 0
            if not math.isnan(speak_score[2]):
                ielts_speak_score = float(speak_score[2])
                norm_ielts_speak_score = IELTS_TOEFL_SPK_DICT.get(ielts_speak_score, 0)

            norm_duo_speak_score = 0
            if not math.isnan(speak_score[3]):
                duo_speak_score = float(speak_score[3])
                norm_duo_speak_score = DUO_TOEFL_SPK_DICT.get(duo_speak_score, 0)

            norm_speak_score = max(toefl_max_speak, norm_ielts_speak_score, norm_duo_speak_score)
            speak_scores.append(norm_speak_score)

        df.insert(loc=20, column='SPEAK', value=speak_scores)

        #Total scores
        toefl_total_scores = df['TOEFL_total'].to_list()
        toefl_total_mb_scores = df['TOEFL_total_mb'].to_list()
        ielts_total_scores = df['IELTS_overallscore'].to_list()
        duo_total_scores = df['DUO_overallscore'].to_list()
        eng_req_scores = df['ENGREQ'].to_list()
        total_scores_tuple = zip(toefl_total_scores, toefl_total_mb_scores,ielts_total_scores,duo_total_scores,eng_req_scores) 
        total_scores = []
        tot_threshold_list = []
        for total_score in total_scores_tuple:
            toefl_total_score = 0
            #print(speak_score[1])
            if not math.isnan(total_score[0]):
                toefl_total_score = int(total_score[0])
            toefl_total_mb_score = 0
            if not math.isnan(total_score[1]):
                toefl_total_mb_score = int(total_score[1])
            toefl_max_total = max(toefl_total_score, toefl_total_mb_score)

            norm_ielts_total_score = 0
            if not math.isnan(total_score[2]):
                ielts_total_score = float(total_score[2])
                norm_ielts_total_score = IELTS_TOEFL_TOT_DICT.get(ielts_total_score, 0)

            norm_duo_total_score = 0
            if not math.isnan(total_score[3]):
                duo_total_score = float(total_score[3])
                norm_duo_total_score = DUO_TOEFL_TOT_DICT.get(duo_total_score, 0)

            norm_total_score = max(toefl_max_total, norm_ielts_total_score, norm_duo_total_score)
            total_scores.append(norm_total_score)

            above_tot_threshold = 'N'
            eng_req_score = total_score[4]
            if norm_total_score == 0 or norm_total_score > 99:
                above_tot_threshold = 'Y'
            tot_threshold_list.append(above_tot_threshold)

        df.insert(loc=21, column='TOT', value=total_scores)
        df.insert(loc=22, column='Abv TOT Threshold', value=tot_threshold_list)
        
    def add_resume_column(self, resume_csv):
        resume_df = pd.read_csv(resume_csv)
        self.df = pd.merge(self.df, resume_df, on='appid', how='left')  
        col = self.df.pop('resume URL').to_list()
        resume_links = []
        for url in col:
            formula = ''
            if type(url) == str: 
                formula = '=HYPERLINK(\"' + url + '\")'
            resume_links.append(formula)
        #Add links to CVs and transcripts
        self.df.insert(loc=23, column='Resume URL', value=resume_links)
        
    def add_transcript_columns(self, transcript_csv):
        transcript_df = pd.read_csv(transcript_csv)
        self.df = pd.merge(self.df, transcript_df, on='appid', how='left')

        col1 = self.df.pop('transcript URL 1').to_list()
        col2 = self.df.pop('transcript URL 2').to_list()
        col3 = self.df.pop('transcript URL 3').to_list()
        col4 = self.df.pop('transcript URL 4').to_list()

        transcripts = zip(col1, col2, col3, col4)
        transcript1_links = []
        transcript2_links = []
        transcript3_links = []
        transcript4_links = []
        for urls in transcripts:
            formula1 = ''
            if type(urls[0]) == str: 
                formula1 = '=HYPERLINK(\"' + urls[0] + '\")'
            transcript1_links.append(formula1)
            formula2 = ''
            if type(urls[1]) == str: 
                formula2 = '=HYPERLINK(\"' + urls[1] + '\")'
            transcript2_links.append(formula2)
            formula3 = ''
            if type(urls[2]) == str:
                formula3 = '=HYPERLINK(\"' + urls[2] + '\")'
            transcript3_links.append(formula3)
            formula4 = ''
            if type(urls[3]) == str:
                formula4 = '=HYPERLINK(\"' + urls[3] + '\")'
            transcript4_links.append(formula4)

        #Add links to CVs and transcripts
        self.df.insert(loc=24, column='Transcript 1 URL', value=transcript1_links)
        self.df.insert(loc=25, column='Transcript 2 URL', value=transcript2_links)
        self.df.insert(loc=26, column='Transcript 3 URL', value=transcript3_links)
        self.df.insert(loc=27, column='Transcript 4 URL', value=transcript4_links)
        
    def add_resume_annotations(self, resume_annotations):
        resume_ann_df = pd.read_csv(resume_annotations)
        resume_ann_df = resume_ann_df.rename(columns={'Notes': 'CV Notes'})
        merged_df = pd.merge(self.df, resume_ann_df, on='appid', how='left')

        resume_col1 = merged_df.pop('Number of first authored publications at international conferences (ACL, ACM, IEEE)').to_list()
        merged_df.insert(loc=24, column='Number of first authored publications at international conferences (ACL, ACM, IEEE)', value=resume_col1)
        resume_col2 = merged_df.pop('Number of internships during undergrad or former masters').to_list()
        merged_df.insert(loc=25, column='Number of internships during undergrad or former masters', value=resume_col2)
        resume_col3 = merged_df.pop('Number of years working fulltime in a CS related job').to_list()
        merged_df.insert(loc=26, column='Number of years working fulltime in a CS related job', value=resume_col3)
        resume_col4 = merged_df.pop('Evidence of project work with real world applications').to_list()
        merged_df.insert(loc=27, column='Evidence of project work with real world applications', value=resume_col4)
        resume_col5 = merged_df.pop('Number of patents').to_list()
        merged_df.insert(loc=28, column='Number of patents', value=resume_col5)
        resume_col6 = merged_df.pop('Awards for technical work').to_list()
        merged_df.insert(loc=29, column='Awards for technical work', value=resume_col6)
        resume_col7 = merged_df.pop('Evidence of having worked in a team ').to_list()
        merged_df.insert(loc=30, column='Evidence of having worked in a team ', value=resume_col7)
        resume_col8 = merged_df.pop('Github URL').to_list()
        merged_df.insert(loc=31, column='Github URL', value=resume_col8)
        resume_col9 = merged_df.pop('CV Notes').to_list()
        merged_df.insert(loc=32, column='CV Notes', value=resume_col9)
        
        self.df = merged_df
        
    def add_transcript_annotations(self, transcript_annotations):
        transcript_ann_df = pd.read_csv(transcript_annotations)
        transcript_ann_df = transcript_ann_df.rename(columns={'Notes': 'Transcript Notes'})
        merged_df = pd.merge(self.df, transcript_ann_df, on='appid', how='left')

        trans_col1 = merged_df.pop('Advanced Math Courses').to_list()
        merged_df.insert(loc=33, column='Advanced Math Courses', value=trans_col1)
        trans_col2 = merged_df.pop('Number of advanced math courses (total/As/Bs/Cs)').to_list()
        merged_df.insert(loc=34, column='Number of advanced math courses (total/As/Bs/Cs)', value=trans_col2)
        trans_col3 = merged_df.pop('Core Computer Science Courses').to_list()
        merged_df.insert(loc=35, column='Core Computer Science Courses', value=trans_col3)
        trans_col4 = merged_df.pop('Number core computer science courses  (total/As/Bs/Cs)').to_list()
        merged_df.insert(loc=36, column='Number core computer science courses  (total/As/Bs/Cs)', value=trans_col4)
        trans_col5 = merged_df.pop('AI and Machine Learning courses').to_list()
        merged_df.insert(loc=37, column='AI and Machine Learning courses', value=trans_col5)
        trans_col6 = merged_df.pop('Number AI and Machine Learning courses  (total/As/Bs/Cs)').to_list()
        merged_df.insert(loc=38, column='Number AI and Machine Learning courses  (total/As/Bs/Cs)', value=trans_col6)
        trans_col7 = merged_df.pop('Transcript Notes').to_list()
        merged_df.insert(loc=39, column='Transcript Notes', value=trans_col7)
        
        self.df = merged_df
        
    def write_excel(self, output_file):
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        self.df.to_excel(writer, index=False)
        writer.save() 

# LTIAdmissions

## Requirements
- Python 3.10.6
- pandas
- xlsxwriter
- PyPDF2

## Creating csv with links to Resumes
1. Download the zip file with all resumes for MCDS, MIIS, and MSAII
2. Unzip files and combine all directories (there will be duplicates, which can be skipped)
3. Upload all resumes to a Google drive folder
4. Find the Google drive directory id by navigating to the resume directory: https://drive.google.com/drive/folders/[DRIVE_ID]
    1. For example: The drive ID for the url: https://drive.google.com/drive/folders/14hqpGTIpPyW4C1nxsYTezeTrSaWP2d_m is: 14hqpGTIpPyW4C1nxsYTezeTrSaWP2d_m
5. Create links for all resumes in the folder
    1. Create a new spreadsheet in Google Drive (not in the same folder as resumes)
    2. Click on Extensions -> Apps Script
    3. Copy and paste the code snippet below, filling in the DRIVE_ID
```
function myFunction() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var s=ss.getActiveSheet();
  var c=s.getActiveCell();

  // get Header row range
  var headerRow = s.getRange(1, 1, 1, 2);

  // add Header values
  headerRow.setValues([['appid', 'resume URL']]);
  
  var fldr=DriveApp.getFolderById("1PBmZx71vlXFBlCgUs9iT7Z7k-xiA5cjV");
  var files=fldr.getFiles();
  var ids=[],names=[],f,str;
  while (files.hasNext()) {
    f=files.next();
    app_id = f.getName().split('_')[1];
    ids.push([app_id]);
    url_string = f.getUrl();
    names.push([url_string]);
  }
  s.getRange(2,1,names.length).setValues(ids);
  s.getRange(2,2,names.length).setValues(names);
}
```
6. Navigate back to the spreadsheet.  The applicant IDs and resume URLs will now be populated.  Name the spreadsheet and download it as a csv.

## Creating csv with links to Transcripts
1. Download the zip file with all transcripts for MCDS, MIIS, and MSAII
2. Unzip files and combine all directories (there will be duplicates, which can be skipped)
3. Combine all transcript PDFs (that can be combined) by running combine transcripts script
```
python combine_transcripts.py [input_transcript_dir] [output_transcript_dir]
```
5. Upload all transcripts to a Google drive folder
6. Find the Google drive directory id by navigating to the resume directory: https://drive.google.com/drive/folders/[DRIVE_ID]
    1. For example: The drive ID for the url: https://drive.google.com/drive/folders/14hqpGTIpPyW4C1nxsYTezeTrSaWP2d_m is: 14hqpGTIpPyW4C1nxsYTezeTrSaWP2d_m
7. Create links for all transcripts in the folder
    1. Create a new spreadsheet in Google Drive (not in the same folder as resumes)
    2. Click on Extensions -> Apps Script
    3. Copy and paste the code snippet below, filling in the DRIVE_ID
```
function myFunction() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var s=ss.getActiveSheet();
  var c=s.getActiveCell();

  // get Header row range
  var headerRow = s.getRange(1, 1, 1, 2);

  // add Header values
  headerRow.setValues([['appid', 'transcript URL']]);

  var fldr=DriveApp.getFolderById("[DRIVE_ID");
  var files=fldr.getFiles();
  var ids=[],names=[],f,str;
  while (files.hasNext()) {
    f=files.next();
    app_id = f.getName().split('_')[0];
    ids.push([app_id]);
    url_string = f.getUrl();
    names.push([url_string]);
  }
  s.getRange(1,1,names.length).setValues(ids);
  s.getRange(1,2,names.length).setValues(names);
}
```
6. Navigate back to the spreadsheet.  The applicant IDs and transcript URLs will now be populated.  Name the spreadsheet and download it as a csv.
7. Collate transcripts for students with more than one
```
python collate_transcripts.py [input_transcript_csv] [output_transcript_csv] 
```

## Creating Master LTI spreadsheet with all applicants 
This part of the process creates a master spreadsheet with all of the LTI applicants for MCDS, MIIS, MSAII.  The spreadsheet provides an initial ranking based on whether applicants passed the language requirements, whether they applied to an LTI program as their first choice, and the sum of the GRE quantitative score and the GPA.
1. Download applicant csvs for MCDS, MIIS, and MSAII from applygrad.
2. Format the applygrad data to a csv  
```
python format_applygrad_csv.py [input_applygrad_file]
```
3. Combine all applicants and add links to resumes and transcripts.  This outputs an excel (xlsx) file.
```
python create_lti_master_csv.py -mcds MCDS.csv -msaii MSAII.csv -miis MIIS.csv -r ResumeLinks.csv -t Collated_transcripts.csv -o [output_file].xlsx
```

## Creating a Program Specific spreadsheet adding annotations and links to resumes and transcripts
This part of the process creates a program specific (MCDS, MIIS, or MSAII) spreadsheet.  The spreadsheet provides an initial ranking based on whether applicants passed the language requirements, whether they applied to the program as their first choice, and the sum of the GRE quantitative score and the GPA.
1. Download program specific CSV from applygrad (you may already have this from creating a master spreadsheet)
2. Format the applygrad data to a csv (Skip this step if you have formatted data from creating the master spreadsheet)
```
python format_applygrad_csv.py [input_applygrad_file]
```
3. Gather files with resume URLs, transcript URLs, resume annotations, and transcript annotations
4. Create program spreadsheet 
```
python create_program_csv.py -i [PROGRAM].csv -n [PROGRAM_NAME - MCDS|MIIS|MSAII] -r [RESUME_URLS].csv -t [TRANSCRIPT_URLS].csv -o [OUTPUT_FILE].xlsx -ra [RESUME_ANNOTATIONS].csv -ta [TRANSCRIPT_ANNOTATIONS].csv
```




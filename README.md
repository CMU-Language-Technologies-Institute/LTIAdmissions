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

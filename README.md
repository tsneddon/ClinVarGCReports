# ClinVarGCReports
Python scripts to generate ClinGen GenomeConnect reports from ClinVar FTP files.

## About this project
The scripts in this project use ClinVar FTP files to generate the following files in the subdirectory ClinVarGCReports/:

**ClinVarGCReports.py** - this script outputs an Excel file for ClinVar variants that have a GenomeConnect submission. The Excel contains a README with summary stats and 6 structured tabs as detailed below:
  
  * \#1. AllSubs: All ClinVar variants where there is a GenomeConnect submission.
  * \#2. AllNovel: All ClinVar variants where the only submission is from GenomeConnect.
  * \#3. LabConflicts: ClinVar variants where the GenomeConnect testing lab clinical significance [P] vs [LP] vs [VUS] vs [LB] vs [B] differs from the clinical lab with same name.
  * \#4. LabConsensus: ClinVar variants where the GenomeConnect testing lab clinical significance [P] vs [LP] vs [VUS] vs [LB] vs [B] is the same as that from the clinical lab with same name.
  * \#5. EPConflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from an Expert Panel or Practice Guideline.
  * \#6. Outlier: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from at least one 1-star or above (or clinical testing) submitter.
  * \#7. LabNotSubmitted: GenomeConnect SCVs where the testing lab has an OrgID but has NOT independently submitted an SCV on the variant.
  * \#8. LabNoOrgID: GenomeConnect SCVs where the testing lab was submitted without an OrgID.


## How to run these scripts
All scripts are run as 'python3 *filename.py*
All scripts use FTP to take the most recent ClinVar FTP files as input and to output the files with the date of the FTP submission_summary.txt.gz file appended:

  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/submission_summary.txt.gz
  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/variation_allele.txt.gz
  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/variant_summary.txt.gz
  * ftp.ncbi.nih.gov/pub/clinvar/tab_delimited/organization_summary.txt
  * ftp.ncbi.nih.gov/pub//pub/GTR/data/gtr_ftp.xml.gz
  * ftp.ncbi.nih.gov/pub/clinvar/xml/clinvar_variation/ClinVarVariationRelease_00-latest.xml.gz

These ClinVar files are then removed when finished.

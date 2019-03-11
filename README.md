# ClinVarGCReports
Python scripts to generate ClinGen GenomeConnect reports from ClinVar FTP files.

## About this project
The scripts in this project use ClinVar FTP files to generate the following files in the subdirectory ClinVarGCReports/:

**ClinVarGCReports.py** - this script outputs a tab-delimted file for all GenomeConnect SCVs with the following content:
  * \#1. varID
  * \#2. SCV
  * \#3. Clinical significance of testing lab
  * \#4. Date Last Evaluated of testing lab
  * \#5. Name of testing lab
  * \#6. OrgID of testing lab
  * \#7. Review Status (should be 'no assertion provided')
  * \#8. Collection Method (should be 'phenotyping only')
  * \#9. Submitter (should be GenomeConnect)
  
This script then uses the file above to output an Excel file for ClinVar variants that have a GenomeConnect submission. The Excel contains a README with summary stats and 5 structured tabs as detailed below:
  
  * \#1. All_subs: All ClinVar variants where there is a GenomeConnect submission.
  * \#2. All_novel: All ClinVar variants where the only submission is from GenomeConnect.
  * \#3. Lab_Conflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from the clinical lab with same name.
  * \#4. EP_Conflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from an Expert Panel or Practice Guideline.
  * \#5. Outlier: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from at least one 1-star or above (or clinical testing) submitter.



%let fpath=C:\Users\bmant\Downloads;
%let dssfile=cdisc_crf_specializations_draft;

/*Import dss*/
proc import datafile="&fpath.\&dssfile..xlsx" out=dss dbms=xlsx replace;sheet='CRF Specializations';
run;

/*Create subcodelists for vlm in P21 */
data cls;
format name $200. id $200.;
set dss;
if value_list ne '' or prepopulated_term ne '';

if substr(crf_item,length(trim(crf_item))-3,4)='RESU' then do;
 name='Unit, subset for '||trim(left(short_name))||' - Original';
 id=trim(left(codelist_submission_value))||'_'||trim(left(vlm_group_id))||'_OR';
end;


if codelist_submission_value ne '' and substr(crf_item,length(trim(crf_item))-3,4) ne 'RESU' then do;
 name='Subset for '||trim(left(short_name))||' - '||trim(left(crf_item))||' '||trim(left(codelist_submission_value));
id=trim(left(codelist_submission_value))||'_'||trim(left(crf_group_id))||'_'||trim(left(crf_item));
end;


if codelist_submission_value eq '' and substr(crf_item,length(trim(crf_item))-3,4) ne 'RESU' then do;
 name='Subset for '||trim(left(short_name))||' - '||trim(left(crf_item));
id=trim(left(crf_item))||'_'||trim(left(crf_group_id));
end;



if value_list='' and prepopulated_term ne '' then value_list=prepopulated_term;
listcount=countc(value_list,';')+1;
do order=1 to listcount;
term=scan(value_list,order,';');
decoded_value=scan(value_display_list,order,';');
output;

end;

keep id name  order term  decoded_value crf_group_id crf_item prepopulated_term value_list codelist_submission_value variable_name;

run;

data dss_ chks;
format typchk $200.;
set dss;
if significant_digits ne '' and length ne '200' and data_type ne 'float' then typchk="Data type should be float";
if length='200' and significant_digits ne '' then typchk="Data type likely text and digits not required"; 
if typchk ne '' then output chks;
if substr(crf_item,length(trim(crf_item))-3,4)='RESU' or  substr(crf_item,length(trim(crf_item))-2,3)='RES' then vlm_target='Y';
if significant_digits ne '' and length ne '200' then data_type='float';/*Fixes for discrepant details i.e. non float with digits*/
if length='200' then significant_digits =''; 
output dss_;

run;

proc sort data=dss_ out=dss__;
by crf_group_id crf_item ;
run;

data wc;

set dss__;

*if comparator ='' and sdtm_variable eq 'LBLOINC' and assigned_value ne '' then comparator='EQ';/*fix for duplicates-better qualifiers should be used i.e. LBRESTYP*/


if (value_list ne '' or prepopulated_term ne '') and implementation_option='Normalized' and vlm_target='';/*VLM ONLY APPROPRIATE FOR NORMALIZED DATA SOURCES*/
assigned_value=prepopulated_term;
if assigned_value ne '' and index(trim(left(assigned_value)),' ') gt 0 then assigned_value='"'||trim(left(assigned_value))||'"';

if value_list ne '' then vlx='("'||trim(left(tranwrd(value_list,';','","')))||'")';

if value_list ne '' then comparator='IN';
if prepopulated_term  ne '' then comparator='EQ';

if comparator='EQ' then wc=trim(left(crf_item))||' '||trim(left(comparator))||' '||trim(left(assigned_value));
if comparator='IN' then wc=trim(left(crf_item))||' '||trim(left(comparator))||' '||trim(left(vlx));



run;





/*Generate where clause*/

data wc_;
format combinedwc $500.;
set wc;
retain combinedwc;
by   crf_group_id crf_item  ;
if first.crf_group_id then combinedwc=wc;
  else  combinedwc=catx(' and ', combinedwc, wc);

if last.crf_group_id then output;
run;

/*Merges*/
proc sql;
create table vlmerge(keep=domain short_name crf_group_id combinedwc crf_item data_type length format codelist_submission_value significant_digits mandatory_value assigned_value) 
as select distinct a.*,b.combinedwc from dss_ a left join wc_ b on a.vlm_group_id=b.vlm_group_id where a.vlm_target='Y';/*Merge full where caluse to targets*/
create table vlmerge2 as select  distinct a.*,b.id from vlmerge a left join  cls b on a.crf_group_id=b.crf_group_id and a.crf_item=b.crf_item;/*merge unit codelist*/
quit;

/*Format to P21 VL template*/
data vlm(rename=(domain=Dataset sdtm_variable=variable combinedwc=where_clause codelist_submission_value=codelist short_name=label mandatory_value=mandatory));
format codelist_submission_value Decoded_Variable Codelist_Expected	Expected_Codelist_ID Expected_Codelist_Name	Origin	Source	Method	Predecessor	Comment	Developer_Notes variant $100.;
set vlmerge2;
if id ne '' then codelist_submission_value=id;
origin='Collected';
order+1;
drop id vlm_group_id;
run;

/*Reorder*/
data vlmfinal;
retain Order	Dataset	Variable	Variant	Where_Clause	Label	Data_Type	Length	Significant_Digits	Format	Mandatory	Assigned_Value	Codelist	Decoded_Variable
Codelist_Expected	Expected_Codelist_ID	Expected_Codelist_Name	Origin	Source	Method	Predecessor	Comment	Developer_Notes;
set vlm;
run;

/*Export of vlm in P21 format*/
proc export data=vlmfinal outfile="&fpath.\crf_vlm.xlsx" dbms=xlsx replace;
run;





data _form;
set dss (rename=(domain=form
crf_item=id
significant_digits=digits
mandatory_variable=mandatory crf_group_id=section codelist=nci_codelist_id sdtm_target_variable=sdtm_target));
if substr(id,length(trim(id))-3,4)='RESU' then measurement_units=translate(value_list,',',';');
sdtm_target=translate(sdtm_target,',',';');

keep Form	Section	short_name Order	ID	Source_Variable_Name	Question_Text	Prompt	Description	Data_Type	Core	Length	Digits	Mandatory	Codelist	Measurement_Units	Completion_Instructions	Implementation_Notes	Mapping_Instructions	SDTM_Target	Reason_Not_Mapped sdtm_annotation	Developer_Notes;
run;

proc sql;
create table form_cl
as select distinct a.*, b.id as codelist,b.prepopulated_term from _form a left join cls b on a.id=b.crf_item and a.section=b.crf_group_id;
quit;


data form_cl_ord;

set form_cl;
order +1;
if  first.section then order=1;

by form section;
run;

data units;
attrib
id label='ID'
unit label='Unit'
;
set cls;
if indexw(name,'Unit,') gt 0;
id=term;
if decoded_value ne '' then unit=decoded_value;
else unit=term;
run;

proc sort data=units(keep= id unit) nodupkey;
by id;
run;



data questions_final;
attrib 

Form label='Form'
Section label='Section'	
short_name label='Short name'
Order	label='Order'
ID	label='ID'
Source_Variable_Name label="Source Variable Name"
Question_Text label="Question Text"	
prompt label="Prompt"	
Description	label="Description"
Data_Type	label="Data Type"
Core	label="Core"
Length label="Length"
Digits label="Digits"
Mandatory label="Mandatory" format=$3.
Codelist label="Codelist"
Measurement_Units label="Measurement Units "
Completion_Instructions label="Completion Instructions"
Implementation_Notes label= "Implementation Notes"
Mapping_Instructions label="Mapping Instructions"
SDTM_Target label="SDTM Target"
Reason_Not_Mapped label="Reason Not Mapped"
Developer_Notes label="Developer Notes"
prepopulated_term label="Prepopulated Term"
sdtm_annotation label="SDTM Annotation"
;
set form_cl_ord;

if sdtm_target ne '' then do;
if countc(sdtm_target,',') eq 0 then sdtm_target=trim(left(form))||'.'||sdtm_target;
if countc(sdtm_target,',') gt 0 then do;
sdtm_target=tranwrd(sdtm_target,',',','||trim(left(form))||'.');
sdtm_target=trim(left(form))||'.'||sdtm_target;
end;

end;
if data_type='decimal' then data_type='float';

if mandatory='N' then Mandatory='No';
if mandatory='Y' then Mandatory='Yes';

run;


proc sort data=questions_final out=forms(keep=form) nodupkey;
by form;
run;

data form_final;
attrib Form label='ID'
sdtm_target_domain label='SDTM Target Domains'
;
set forms;
rename form=id;
name=form;
order+1;
sdtm_target_domain=form;
run;

proc sort data=questions_final out=sections(keep=form section short_name) nodupkey;
by section;
run;

proc sort data=sections;
by form section;
run;
data sections_final;
attrib 
short_name lable='Name'
order label='Order'
section label='ID'
mandatory label='Mandatory'
repeating label='Repeating'
;
set sections;
rename section=id ;
rename short_name=name;

mandatory='No';
repeating='No';
order+1;

if first.form then order=1;
by form;


run;

data terms(keep=codelist display_term recommended_term order);
attrib 
recommended_term label="Recommended Term"
id label='Codelist'
display_term label='Display Term'
order label='Order'
;

set cls;
rename id=codelist;
recommended_term=term;
display_term=decoded_value;
if decoded_value='' then display_term=term;

run;

proc sort data=terms out=terms_final nodupkey;
by codelist recommended_term;
run;


proc sort data=cls out=codelist nodupkey;
by id;
run;

data codelist_final;
attrib 
recommended_codelist format=$50. label="Recommended Codelist"
type label='Type'
name label='Name'
id label='ID'
;
set codelist;

type='text';
keep id name type;

run;


data formspec;
 infile datalines delimiter=','; 
input Attribute: $20. Value: $100.;
datalines;
StudyName,	Ben Testing
StudyDescription,	CDISC COSmos Data crf Specialization
ProtocolName,	
Language,	en
Source Spec(s),
Type, EDC	
;
run;

proc export data=questions_final outfile="&fpath\crf_form.xlsx" dbms=xlsx replace label;sheet='Questions';
run;

proc export data=sections_final outfile="&fpath\crf_form.xlsx" dbms=xlsx replace label;sheet='Sections';
run;
proc export data=form_final outfile="&fpath\crf_form.xlsx" dbms=xlsx replace label;sheet='Forms';
run;
proc export data=units outfile="&fpath\crf_form.xlsx" dbms=xlsx replace label;sheet='Units';
run;

proc export data=codelist_final outfile="&fpath\crf_form.xlsx" dbms=xlsx replace label;sheet='Codelists';
run;
proc export data=terms_final outfile="&fpath\dss_form.xlsx" dbms=xlsx replace label;sheet='Terms';
run;
proc export data=formspec outfile="&fpath\crf_form.xlsx" dbms=xlsx replace label;sheet='FormSpec';
run;


proc sort data=cls out=cl_nodups dupout=cl_chk nodupkey ;
by term value_list;

run;

data cl_chk_;
set cl_chk;
*if substr(id,1,4) ne 'UNIT' and substr(id,3,4) ne 'RESU';

if codelist_submission_value ne '' then do;
if countc(value_list,';') le 1 then newid=trim(left(codelist_submission_value))||'_'||translate(trim(left(value_list)),'_',';','_',' ');
end;

if codelist_submission_value eq '' and variable_name ne '' then do;
if countc(value_list,';') le 1 then newid=trim(left(variable_name))||'_'||translate(trim(left(value_list)),'',';','_',' ');
end;

if countc(value_list,';') gt 1 then do;

newid=trim(left(codelist_submission_value))||'_'||trim(left(variable_name));
end;

newname='Subset for '||trim(left(variable_name));


run;






* Load your data
import excel "C:\Users\Administrator\Desktop\Damalie work\DAMS_DATA.xlsx", sheet("Sheet1") firstrow clear

des

codebook age totalafc amhngml
codebook weight height
replace weight = 73 if height > 1 & weight == .
replace height =1.6 if height == . | height > 3
replace bmi = (weight/(height*height))
codebook bmi


***********

***********
gen AFC = totalafc if amhngml != .
gen AMH = amhngml if amhngml <= 20


gen bmi_category = 0
replace bmi_category = 1 if bmi < 18.5
replace bmi_category = 2 if bmi >= 18.5 & bmi < 25
replace bmi_category = 3 if bmi >= 25 & bmi <= 29.9
replace bmi_category = 4 if bmi >= 30
label define bmi_category 1 "underweight" 2 "normal" 3 "overweight" 4 "obese"
label values bmi_category bmi_category
label variable bmi_category "bmi category"
codebook bmi_category
order bmi_category, after (bmi)

codebook selfrecipientod
replace selfrecipientod = "EMBRYO ADOPTION" if selfrecipientod == "EA"
replace selfrecipientod = "RECIPIENT" if selfrecipientod == "EMBRYO ADOPTION" | selfrecipientod == "RECIPIENT CYCLE"
replace selfrecipientod = "SELF" if selfrecipientod == "SELF CYCLE"
replace selfrecipientod = "NA" if selfrecipientod == ""

gen self_recipient = 1 if selfrecipientod == "SELF"
replace self_recipient = 2 if selfrecipientod == "RECIPIENT" | selfrecipientod == "OOCYTE DONOR"
label define self_recipient 1 "SELF" 2 "RECIPIENT"
label values self_recipient self_recipient
label variable self_recipient "SELF or RECIPIENT"

gen self_donor = 1 if selfrecipientod == "SELF"
replace self_donor = 2 if selfrecipientod == "OOCYTE DONOR"
label define self_donor 1 "SELF" 2 "OOCYTE DONOR"
label values self_donor self_donor
label variable self_donor "SELF or RECIPIENT"

gen recipient_self = 1 if selfrecipientod == "SELF"
replace recipient_self = 2 if selfrecipientod == "RECIPIENT"
label define recipient_self 1 "SELF" 2 "RECIPIENT"
label values recipient_self recipient_self
label variable recipient_self "SELF or RECIPIENT"

gen orpi = (amhngml * totalafc)/ age 
label variable orpi "Ovarian Response Prediction Index"
codebook orpi
summarize orpi

gen foi = numberofoocytesretrieved / totalafc
label variable foi "follicle-to-oocyte index"
codebook foi
summarize foi

gen fort = (totalnumberoffolliclesdevelo / totalafc) * 100
label variable fort "follicular output rate (%)"
codebook fort
summarize fort 


gen fertilizationrate = (numberofeggsfertilized / numberofoocytesretrieved) * 100
label variable fertilizationrate "fertilization rate (%)"
summarize fertilizationrate

gen blasto_rate = (numberofday5blastocysts / numberofoocytesretrieved) * 100
label variable blasto_rate "Blastocyst rate (%)"

gen eggCleavagerate = (numberofeggscleaved / numberofoocytesretrieved) * 100
label variable eggCleavagerate "Egg cleavage rate (%)"

gen implantationrate = (numberoffetusesat6weekstra / numberofembryostransferred) * 100
label variable implantationrate "Implantation rate (%)"

// ongoing pregnancy rate calc
// gen ongoingpreg = 0 if numberoffetusesat6weekstra == . & (self_recipient == "SELF" | self_recipient == "RECIPIENT")
// replace ongoingpreg = 1 if numberoffetusesat6weekstra == 1 | numberoffetusesat6weekstra == 2 | numberoffetusesat6weekstra == 3
// codebook ongoingpreg


//
// gen clinicalpregrate = (numberoffetusesat6weekstra / numberofembryostransferred) * 100
// label variable eggCleavagerate "Ongoing pregnancy rate (%)"

gen urinePreg = 0 if urinepregnancyresults == "NEAGATIVE" | urinepregnancyresults == "NEGATIVE"
replace urinePreg = 1 if urinepregnancyresults == "POSITIVE" | urinepregnancyresults == "POSITIVE  2" | urinepregnancyresults == "POSITIVE  1" | urinepregnancyresults == "POSITIVE  3" | urinepregnancyresults == "POSITIVE 1" | urinepregnancyresults == "POSITIVE 2" | urinepregnancyresults == "POSITIVE 3" 
label define urinePreg 0 "Negative" 1 "Positive"
label variable urinePreg "Urine pregnancy results"
label values urinePreg urinePreg

gen urinepregnancy = 0 if urinePreg == 0 
replace urinepregnancy = 1 if urinepregnancyresults == "POSITIVE" | urinepregnancyresults == "POSITIVE  1" | urinepregnancyresults == "POSITIVE 1"
replace urinepregnancy = 2 if urinepregnancyresults == "POSITIVE  2" | urinepregnancyresults == "POSITIVE 2"
replace urinepregnancy = 3 if urinepregnancyresults == "POSITIVE  3" | urinepregnancyresults == "POSITIVE 3"
codebook urinepregnancy

// br selfrecipientod urinepregnancyresults urinePreg cyclecancelled if urinePreg == 0

*****
gen urinepregrate = (urinepregnancy / (numberofetrecipient)) * 100
label variable urinepregrate "Urine pregnancy rate (%)"

gen clinicalpregrate = (numberoffetusesat6weekstra / (numberofembryostransferred / numberofetrecipient)) * 100
label variable clinicalpregrate "Clinical pregnancy rate (%)"



encode selfrecipientod, gen(selfRecipientOd)

gen cycleCancel = 0 if cyclecancelled == "N0" | cyclecancelled == "NO" | cyclecancelled == "NA" | cyclecancelled == ""
replace cycleCancel = 1 if cyclecancelled == "YES" | cyclecancelled == "YES  " | cyclecancelled == "YES  "

label define cycleCancel 0 "No" 1 "Yes"
label variable cycleCancel "cycle cancellation"
label values cycleCancel cycleCancel

tab cycleCancel

gen diminishedOvR_amh = 0 if amhngml < 1.2 & amhngml != .
replace diminishedOvR_amh = 1 if amhngml >= 1.2 & amhngml != .
label define diminishedOvR_amh 0 "diminished" 1 "undiminish"
label values diminishedOvR_amh diminishedOvR_amh
label variable diminishedOvR_amh "diminished ovarian reserve (amh)"

gen diminishedOvR_afc = 0 if totalafc < 5 & totalafc != .
replace diminishedOvR_afc = 1 if totalafc >= 5 & totalafc != .
label define diminishedOvR_afc 0 "diminished" 1 "undiminish"
label values diminishedOvR_afc diminishedOvR_afc
label variable diminishedOvR_afc "diminished ovarian reserve (afc)"

** 
gen diminishedOvR_afc_new = 0 if AFC < 5 & AFC != .
replace diminishedOvR_afc_new = 1 if AFC >= 5 & AFC != .
label define diminishedOvR_afc_new 0 "diminished" 1 "undiminish"
label values diminishedOvR_afc_new diminishedOvR_afc
label variable diminishedOvR_afc_new "diminished ovarian reserve (afc)"

**

gen ovarianResponse = 0 if numberofoocytesretrieved < 4
replace ovarianResponse = 1 if numberofoocytesretrieved >= 4
label define ovarianResponse 0 "Poor" 1 "Good"
label values ovarianResponse ovarianResponse
label variable ovarianResponse "Poor ovarian response"


gen osi = numberofoocytesretrieved / totaldoseofgonadotropinsiu * 1000
label variable osi "Ovarian Sensitivity Index"
br osi

// br age amhngml totalafc orpi fertilityrate
tab selfrecipientod
tab self_recipient

******************
gen maritalstatus = 1 if marritalstatus == "SINGLE"
replace maritalstatus =2 if marritalstatus == "MARIED" | marritalstatus == "MARRIED" 
replace maritalstatus = 3 if marritalstatus == "WITH A PARTNER" | marritalstatus == "COHABITING" | marritalstatus == "CO-HABITING"
label define maritalstatus 1 "Single" 2 "Married" 3 "Co-habitting"
label values maritalstatus maritalstatus
label variable maritalstatus "Marital Status"
order maritalstatus, after (marritalstatus)
tab maritalstatus
drop marritalstatus

replace previousinfertilitytreammentr = "OI PLUS TI" if previousinfertilitytreammentr == "OI AND TI"

replace vdrl = "NON REACTIVE" if vdrl == "NON REACTIVE " | vdrl == "NON REACTVIVE"
encode vdrl, gen (Vdrl)
order Vdrl, after (vdrl)
drop vdrl

replace hiv = "NON REACTIVE" if hiv == "NON REACTIVE " | hiv == "NON REACTVIVE" | hiv == "NON RACTIVE"
encode hiv, gen (HIV)
order HIV, after (hiv)
drop hiv

replace hbsag = "NON REACTIVE" if hbsag == "NON REACTIVE " | hbsag == "NON REACTVIVE" | hbsag == "NON RACTIVE"
encode hbsag, gen (HBsag)
order HBsag, after (hbsag)
drop hbsag

replace hapatitisc = "NON REACTIVE" if hapatitisc == "NON REACTIVE " | hapatitisc == "NON REACTVIVE" | hapatitisc == "NON RACTIVE"
encode hapatitisc, gen (hepC)
order hepC, after (hapatitisc)
drop hapatitisc

rename bloodgroup Bloodgroup
gen bloodgroup = "A POSITIVE" if Bloodgroup == "A   POSITIVE" |  Bloodgroup == "A  POSITIVE" | Bloodgroup == "A POSITIVE"
replace bloodgroup = "A NEGATIVE" if Bloodgroup == "A  NEGATIVE" | Bloodgroup == "A   NEGATIVE" | Bloodgroup == "A NEGATIVE"
replace bloodgroup = "AB POSITIVE" if Bloodgroup == "AB   POSITIVE" | Bloodgroup == "AB  POSITIVE" | Bloodgroup == "AB POSITIVE"
replace bloodgroup = "AB NEGATIVE" if Bloodgroup == "AB  NEGATIVE" | Bloodgroup == "AB   NEGATIVE" | Bloodgroup == "AB NEGATIVE"
replace bloodgroup = "B POSITIVE" if Bloodgroup == "B   POSITIVE" | Bloodgroup == "B  POSITIVE" | Bloodgroup == "B POSITIVE"
replace bloodgroup = "B NEGATIVE" if Bloodgroup == "B  NEGATIVE" | Bloodgroup == "B   NEGATIVE" | Bloodgroup == "B NEGATIVE"
replace bloodgroup = "O POSITIVE" if Bloodgroup == "O   POSITIVE" | Bloodgroup == "O  POSITIVE" | Bloodgroup == "O POSITIVE"
replace bloodgroup = "O NEGATIVE" if Bloodgroup == "O  NEGATIVE" | Bloodgroup == "O   NEGATIVE" | Bloodgroup == "O NEGATIVE"
drop Bloodgroup
tab bloodgroup

replace	bloodsugar = "13.0" if bloodsugar == "13.0 (RBS)"
destring bloodsugar, replace 
summarize bloodsugar


gen adenomyosispresence = 1
replace adenomyosispresence = 0 if adenomyosis == "NIL" |adenomyosis == "NO"
label define adenomyosispresence 0 "Absent" 1 "Present"
label values adenomyosispresence adenomyosispresence
tab adenomyosispresence


replace fibroids_cat = "Present Insig" if fibroids_cat == "PRESENT AND INSIGNIFICANT" | fibroids_cat == "PRESENT INSIGNIFICANT" | fibroids_cat == "PRESENT NOT SIGNIFICANT"
replace fibroids_cat = "Present Sig" if fibroids_cat == "PRESENT AND SIGNIFICANT" | fibroids_cat == "PRESENT SIGNIFICANT" | fibroids_cat == "PRESENT  SIGNIFICANT" | fibroids_cat == "PRESENT" | fibroids_cat == "PRESENT GIGNIFICANT"
br fibroids_cat fibroidpresence myomectomy





rename adenomyosis Adenomyosis
tab Adenomyosis
gen adenomyosis = 0 if Adenomyosis == "NIL" | Adenomyosis == "NO"
replace adenomyosis = 1 if Adenomyosis == "YES"
replace adenomyosis = 2 if Adenomyosis == "PRESENT AND SIGNIFICANT"
replace adenomyosis = 3 if Adenomyosis == "PRESENT AND INSIGNIFICANT"
label define adenomyosis 0 "Nil" 1 "Yes" 2 "Present Sig" 3 "Present Insig"
label values adenomyosis adenomyosis

rename ovariancystpresent Ovariancystpresent
gen ovariancystpresent = 0 if Ovariancystpresent == "NIL" | Ovariancystpresent == "NO"
replace ovariancystpresent = 1 if Ovariancystpresent == "YES"
label define ovariancystpresent 0 "No" 1 "Yes"
label values ovariancystpresent ovariancystpresent

replace previoushistoryofovariancyst = "NO" if previoushistoryofovariancyst == "NIL"

rename previoushistoryofovariancyst Previoushistoryofovariancyst
encode Previoushistoryofovariancyst, gen(previoushistoryofovariancyst)

tab emdometriosispresent 
replace emdometriosispresent = "NO" if emdometriosispresent == "NIL"
rename emdometriosispresent Endometriosispresent
encode Endometriosispresent, gen (endometriosispresent)
tab endometriosispresent

replace endometriosisgrade = "NO" if endometriosisgrade == "NIL" 
replace endometriosisgrade = "IV" if endometriosisgrade == "4"
rename endometriosisgrade Endometriosisgrade
encode Endometriosisgrade, gen(endometriosisgrade)
tab endometriosisgrade


replace previoushistoryofsalpingectom = "NO" if previoushistoryofsalpingectom == "NIL"
rename previoushistoryofsalpingectom Previoushistoryofsalpingectom
encode Previoushistoryofsalpingectom, gen(previoushistoryofsalpingectom)

replace previoushistoryofappendicecto = "NO" if previoushistoryofappendicecto == "NIL"
rename previoushistoryofappendicecto Previoushistoryofappendicecto
encode Previoushistoryofappendicecto, gen(previoushistoryofappendicecto)

replace previoushistoryoflaparoscopy = "NO" if previoushistoryoflaparoscopy == "NIL" | previoushistoryoflaparoscopy == "NA"
rename previoushistoryoflaparoscopy Previoushistoryoflaparoscopy
encode Previoushistoryoflaparoscopy, gen(previoushistoryoflaparoscopy)

rename endometrialmorphology Endometrialmorphology
encode Endometrialmorphology, gen(endometrialmorphology)

replace ethnicity = "NORTHERN" if ethnicity == "BANA" | ethnicity == "BARSARE (NORTH)" | ethnicity == "BENHYIRA" | ethnicity == "DAAGARI" | ethnicity == "FRAFRA" | ethnicity == "BONO" | ethnicity == "SISALA"
replace ethnicity = "EXPATRIATE" if ethnicity == "IGBO (NIGERIAN)" | ethnicity == "NIGERIAN" 

* Create a new variable called 'employment_status'
gen employment_status = ""

* Replace values for the 'employment_status' variable based on the 'occupation' variable

rename ocuppation occupation
* Group under "self-employed"
replace employment_status = "self-employed" if strpos(occupation, "Farmer") > 0 | strpos(occupation, "Trader") > 0 | strpos(occupation, "Shop Owner") > 0 | strpos(occupation, "Freelancer") > 0 | strpos(occupation, "Consultant") > 0 | strpos(occupation, "Business Owner") > 0 | strpos(occupation, "Contractor") > 0 | strpos(occupation, "Entrepreneur") > 0 | strpos(occupation, "Business Woman") > 0 | strpos(occupation, "Bussiness Woman") > 0 | strpos(occupation, "Businesswoman") > 0 | strpos(occupation, "School Proprietress") > 0 | strpos(occupation, "Hairdresser") > 0 | strpos(occupation, "Hair Dresser") > 0 | strpos(occupation, "Hair Dressing") > 0 | strpos(occupation, "Self Employed") > 0 | strpos(occupation, "Seamstress") > 0 | strpos(occupation, "Fashion Designer") > 0 | strpos(occupation, "Business  Woman") > 0 | strpos(occupation, "Events Planner") > 0 | strpos(occupation, "Events Planer") > 0  | strpos(occupation, "Lotto") > 0  | strpos(occupation, "Beautician") > 0  | strpos(occupation, "Care Giver") > 0 | strpos(occupation, "Décor") > 0 | strpos(occupation, "Caring Assistance") > 0  | strpos(occupation, "Auditor") > 0 | strpos(occupation, "Care Assistant") > 0 | strpos(occupation, "Receptionist") > 0 | strpos(occupation, "Self-Employed") > 0

* Group under "government employed"
replace employment_status = "government employed" if strpos(occupation, "Teacher") > 0 | strpos(occupation, "Police Officer") > 0 | strpos(occupation, "Soldier") > 0 | strpos(occupation, "Nurse") > 0 | strpos(occupation, "Doctor") > 0 | strpos(occupation, "Civil Servant") > 0 | strpos(occupation, "Public Health Worker") > 0 | strpos(occupation, "Government Official") > 0 | strpos(occupation, "Ges") > 0 | strpos(occupation, "Midwife") > 0 | strpos(occupation, "Teaching") > 0 | strpos(occupation, "Secretary") > 0 | strpos(occupation, "Service Personnel") > 0 | strpos(occupation, "Service  Personnel") > 0 | strpos(occupation, "Custom Relation Officer") > 0  | strpos(occupation, "Administrator") > 0  | strpos(occupation, "Forestry Commision") > 0  | strpos(occupation, "Public Servant") > 0 | strpos(occupation, "Health Assistant") > 0 | strpos(occupation, "Administration Assistant") > 0 | strpos(occupation, "Nursing") > 0 | strpos(occupation, "Customs Officer") > 0 | strpos(occupation, "Health Research Officer") > 0 | strpos(occupation, "Nealth Service") > 0 | strpos(occupation, "Nealth Service") > 0 | strpos(occupation, "Health And Safety Officer") > 0 | strpos(occupation, "Clinical Research") > 0 | strpos(occupation, "Extension Officer") > 0 | strpos(occupation, "Facility Officer") > 0 | strpos(occupation, "Dispensing Assistant") > 0  | strpos(occupation, "Administrative Assistant") > 0 | strpos(occupation, "Health  Assistant") > 0  |strpos(occupation, "Police") > 0   | strpos(occupation, "Pharmacist") > 0 

* Group under "private employed"
replace employment_status = "private employed" if strpos(occupation, "Engineer") > 0 | strpos(occupation, "Accountant") > 0 | strpos(occupation, "Banker") > 0 | strpos(occupation, "Manager") > 0 | strpos(occupation, "Salesperson") > 0 | strpos(occupation, "Clerk") > 0 | strpos(occupation, "Technician") > 0 | strpos(occupation, "HR") > 0 | strpos(occupation, "Private Sector") > 0 | strpos(occupation, "Corporate") > 0 | strpos(occupation, "Caterer") > 0 | strpos(occupation, "Cosmetologist") > 0 | strpos(occupation, "Sales Assistant") > 0  | strpos(occupation, "Sales Girl") > 0 | strpos(occupation, "Cashier") > 0 | strpos(occupation, "Research Assistant") > 0 | strpos(occupation, "Broadcasting") > 0 | strpos(occupation, "Home Care") > 0 | strpos(occupation, "Support Work") > 0 | strpos(occupation, "Sales") > 0 | strpos(occupation, "Miner") > 0 | strpos(occupation, "Apprentice") > 0 | strpos(occupation, "Apprentiship") > 0 | strpos(occupation, "Waitress") > 0 | strpos(occupation, "Cleaner") > 0 | strpos(occupation, "Support  Worker") > 0  | strpos(occupation, "Supporting Worker") > 0  | strpos(occupation, "Recycler") > 0   | strpos(occupation, "Educator") > 0 

* Group under "unemployed"
replace employment_status = "unemployed" if strpos(occupation, "Unemployed") > 0 | strpos(occupation, "None") > 0 | strpos(occupation, "Looking For Work") > 0 | strpos(occupation, "Job Seeker") > 0 | strpos(occupation, "Student") > 0 | strpos(occupation, "Housewife") > 0 

* Check the new variable
list occupation employment_status in 1/20
order employment_status, after (occupation)

pwcorr AFC AMH age, sig


graph matrix AFC AMH age

******************
save "C:\Users\Administrator\Desktop\Damalie work\damalie_data.dta", replace
export excel using "damalie_data", sheetreplace firstrow(variables)
************************************************




capture log close
log using "C:\Users\Administrator\Desktop\Damalie work\damali_output.smcl", replace

use damalie_data

// 1.	BMI
// What proportion of infertile women were obese? 
tab bmi_category selfrecipientod, col
by selfrecipientod, sort: summarize bmi

tab bmi_category self_recipient, col
by self_recipient, sort: summarize bmi

// 2.	Ovarian Response Prediction Index (ORPI) calculated by multiplying the AMH (ng/ml) level by the number of antral follicles (2–9 mm), and the result was divided by the age (years) of the patient.
summarize orpi, detail
codebook orpi
pwcorr age amhngml totalafc orpi foi fort, sig
graph matrix age amhngml totalafc orpi foi fort
by selfrecipientod, sort: summarize orpi

**** Check graphs again *****

// 3.	FOI: the ratio of the number of retrieved oocytes and AFC before COS
summarize foi
codebook foi
// 4.	FORT: preovulatory follicle (with a mean diameter between 16 and 22 mm) count divided by AFC × 100
summarize fort
codebook fort
// 5.	Mean age of self cycle patients 

by selfrecipientod, sort: summarize age
codebook age

// 6.	Mean age of recipient cycle patienst & 7.	Mean age of donors 
by selfrecipientod, sort: summarize age
oneway age selfrecipientod // anova
// 8.	Total dose of FSH 
// startdose
by selfrecipientod, sort: summarize startdose, detail
oneway startdose selfrecipientod // anova

by selfrecipientod, sort: summarize totaldoseofgonadotropinsiu, detail
oneway totaldoseofgonadotropinsiu selfrecipientod // anova

by selfrecipientod, sort: summarize doseofovulationtriggerbhcg, detail
oneway doseofovulationtriggerbhcg selfrecipientod // anova

// 9.	No eggs retrieved 
by selfrecipientod, sort: summarize numberofoocytesretrieved, detail
oneway numberofoocytesretrieved selfRecipientOd // anova
anova numberofoocytesretrieved selfRecipientOd

** (stats difference between eggs retrieved between self and egg donors. Same for dosage of gonad, fertilization rate, blast, implantation rate, pregnancy rate, ongoing rate, cancellation rate)

// 10.	Fertilization rate 
summarize fertilizationrate, detail
by self_recipient, sort: summarize fertilizationrate, detail
oneway fertilizationrate selfrecipientod // anova
anova fertilizationrate selfRecipientOd

// Cleavage
summarize eggCleavagerate, detail
by self_recipient, sort: summarize eggCleavagerate, detail
oneway eggCleavagerate selfrecipientod // anova
anova eggCleavagerate selfRecipientOd

// 11.	Blastocyst rate 
summarize blasto_rate, detail
by self_recipient, sort: summarize blasto_rate, detail
oneway blasto_rate selfrecipientod // anova
anova blasto_rate selfRecipientOd

// 12.	Implantation rate 
summarize implantationrate, detail
by self_recipient, sort: summarize implantationrate, detail
oneway implantationrate selfrecipientod // anova
anova implantationrate selfRecipientOd

// 13.	Pregnancy rate 
codebook urinePreg
tab urinePreg
tab self_recipient urinePreg, row chi

summarize urinepregrate
by self_recipient, sort: summarize urinepregrate
oneway urinepregrate self_recipient // anova
anova urinepregrate self_recipient

// 14.	Ongoing pregnancy rate ****
tab numberoffetusesat6weekstra

summarize clinicalpregrate
by self_recipient, sort: summarize clinicalpregrate
oneway clinicalpregrate self_recipient // anova
anova clinicalpregrate self_recipient

// 15.	Cycle cancellation rate ** The NA should be revisited and entered

// tab cycleCancel selfrecipientod if selfRecipientOd != 1 & selfRecipientOd != 3  & cycleCancel != 0, col

tab cycleCancel self_recipient if cycleCancel != 0, col

tab cycleCancel if cycleCancel != 0

// 16.	What proportion of self cycles have diminished ovarian reserve
// i.e. AMH < 1.2ng/ml or Total AFC < 7 

tab selfrecipientod diminishedOvR_amh, row chi
tab selfrecipientod diminishedOvR_afc, row chi // afc < 5

// How do these compare between the AMH arm and AFC arm? 
by diminishedOvR_amh, sort: summarize amhngml
by diminishedOvR_afc, sort: summarize totalafc

pwcorr totalafc amhngml, sig

// What is the corresponding age at which DOR occurs? 
by diminishedOvR_amh, sort: summarize age
by diminishedOvR_afc, sort: summarize age

lowess diminishedOvR_amh age, mean recast(scatter) mcolor(%0)
lowess diminishedOvR_afc age, recast(scatter) mcolor(%0)


by diminishedOvR_afc, sort: summarize numberofoocytesretrieved
by diminishedOvR_amh, sort: summarize numberofoocytesretrieved

// 17.	What proportion of self cycles have poor ovarian response i.e., no oocyte retrieved less than 4
tab selfrecipientod ovarianResponse, row chi

summarize osi,detail
by urinePreg, sort: summarize osi
by self_recipient, sort: summarize osi
oneway osi self_recipient // anova
anova osi self_recipient

*****************
****************************

tab ethnicity
tab employment_status
tab bloodgroup
*** Some Anatomical distibutions

tab fibroidpresence
tab fibroid selfrecipientod, col chi

tab adenomyosis
tab adenomyosispresence
tab adenomyosis selfrecipientod, col chi

tab ovariancystpresent
tab ovariancystpresent

tab previoushistoryofovariancyst
tab previoushistoryofovariancyst selfrecipientod, col chi

tab endometriosispresent 
tab endometriosispresent selfrecipientod, col chi
 
tab endometriosisgrade 
tab endometriosisgrade selfrecipientod, col chi
 
tab previoushistoryofsalpingectom 
tab previoushistoryofsalpingectom selfrecipientod, col chi
 
tab previoushistoryofappendicecto
tab previoushistoryofappendicecto selfrecipientod, col chi
 
tab previoushistoryoflaparoscopy 
tab previoushistoryoflaparoscopy selfrecipientod, col chi
 
tab endometrialmorphology
tab endometrialmorphology selfrecipientod, col chi
 
***************
******************
***********************

// Objective 2:
// Determine the factors associated with diminished ovarian reserve and poor ovarian response
//
// A.	PHYSIOLOGICAL FACTORS 
// i.	Age
logistic diminishedOvR_afc age
logistic diminishedOvR_amh age
logistic ovarianResponse age
regress osi age

// ii.	BMI 	_________________
logistic diminishedOvR_afc bmi
logistic diminishedOvR_amh bmi
logistic ovarianResponse bmi
regress osi bmi

logistic diminishedOvR_afc i.bmi_category
logistic diminishedOvR_amh i.bmi_category
logistic ovarianResponse i.bmi_category
regress osi i.bmi_category

logistic diminishedOvR_afc age bmi
logistic diminishedOvR_amh age bmi
logistic ovarianResponse age bmi
regress osi age bmi

// B.	ANATOMICAL FACTORS
// i.	Fibroids 
// a.	Present and significant (submucous or intramural more than 5cm)
// b.	Present and insignificant (subserous or intramural less than 5cm)
// c.	Previous history/previous myomectomy 
logistic diminishedOvR_afc i.fibroidpresence
logistic diminishedOvR_amh i.fibroidpresence
logistic ovarianResponse i.fibroidpresence
regress osi i.fibroidpresence

logistic diminishedOvR_afc i.fibroid
logistic diminishedOvR_amh i.fibroid
logistic ovarianResponse i.fibroid
regress osi i.fibroid

// ii.	Adenomyosis
// a.	Present and significant (distorting the endometrium)
// b.	Present and insignificant (not distorting the endometrium)
// c.	Previous history (excision of adenomyosis)
logistic ovarianResponse i.adenomyosispresence
regress osi i.adenomyosispresence

// iii.	 Ovarian Cyst 
// a.	Present 
// b.	Previous history of ovarian cystectomy 	Yes __________ No ___________
logistic diminishedOvR_afc i.ovariancystpresent
logistic diminishedOvR_amh i.ovariancystpresent
logistic ovarianResponse i.ovariancystpresent
regress osi i.ovariancystpresent

logistic diminishedOvR_afc i.previoushistoryofovariancyst
logistic diminishedOvR_amh i.previoushistoryofovariancyst
logistic ovarianResponse i.previoushistoryofovariancyst
regress osi age i.previoushistoryofovariancyst

// iv.	 Endometriosis 					
// a.	Present					Yes __________ No ___________
// b.	Grade: 	I ______ II ______ III ______ IV ______
logistic diminishedOvR_afc i.endometriosispresent
logistic diminishedOvR_amh i.endometriosispresent
logistic ovarianResponse i.endometriosispresent
regress osi i.endometriosispresent

logistic diminishedOvR_afc i.endometriosisgrade
logistic diminishedOvR_amh i.endometriosisgrade
logistic ovarianResponse i.endometriosisgrade
regress osi i.endometriosisgrade

// v.	Previous history of salpingectomy 		Yes __________ No ___________
logistic diminishedOvR_afc i.previoushistoryofsalpingectom
logistic diminishedOvR_amh i.previoushistoryofsalpingectom
logistic ovarianResponse i.previoushistoryofsalpingectom
regress osi i.previoushistoryofsalpingectom

// vi.	 Previous history of appendicectomy 	Yes __________ No ___________
logistic diminishedOvR_afc i.previoushistoryofappendicecto
logistic diminishedOvR_amh i.previoushistoryofappendicecto
logistic ovarianResponse i.previoushistoryofappendicecto
regress osi i.previoushistoryofappendicecto

tab previoushistoryofappendicecto

// vii. Previous history of laparoscopy
logistic diminishedOvR_afc i.previoushistoryoflaparoscopy
logistic diminishedOvR_amh i.previoushistoryoflaparoscopy
logistic ovarianResponse i.previoushistoryoflaparoscopy
regress osi i.previoushistoryoflaparoscopy

// Objective 3
// Establish the association between 
// i.	Age
// ii.	BMI 
// iii.	AMH 
// iv.	AFC
// v.	ORPI 
// vi.  OSI
// vii.	Presence of fibroids 
// viii.Presence of adenomyosis 
// x.	presence of endometriosis 
// xi.	endometrial thickness 
// xii.	endometrial morphology 
//



// And 
//
// i.	Cycle cancellation
logistic cycleCancel age
logistic cycleCancel bmi 
logistic cycleCancel amhngml
logistic cycleCancel totalafc
logistic cycleCancel orpi
logistic cycleCancel i.ovarianResponse
logistic cycleCancel osi
logistic cycleCancel i.fibroidpresence
logistic cycleCancel i.adenomyosispresence
logistic cycleCancel i.endometriosispresent
logistic cycleCancel i.endometrialmorphology
logistic cycleCancel endometrialthickness

// ii.	FOI 
regress foi age
regress foi bmi 
regress foi amhngml
regress foi totalafc
regress foi orpi
regress foi i.ovarianResponse
regress foi osi
regress foi i.fibroidpresence
regress foi i.adenomyosispresence
regress foi i.endometriosispresent
regress foi i.endometrialmorphology
regress foi endometrialthickness

// iii.	FORT 
regress fort age
regress fort bmi 
regress fort amhngml
regress fort totalafc
regress fort orpi
regress fort i.ovarianResponse
regress fort osi
regress fort i.fibroidpresence
regress fort i.adenomyosispresence
regress fort i.endometriosispresent
regress fort i.endometrialmorphology
regress fort endometrialthickness

use damalie_data
// iv.	Fertilization rate
regress fertilizationrate age
regress fertilizationrate bmi 
regress fertilizationrate amhngml
regress fertilizationrate totalafc
regress fertilizationrate orpi
regress fertilizationrate i.ovarianResponse
regress fertilizationrate osi
regress fertilizationrate i.fibroidpresence
regress fertilizationrate i.adenomyosispresence
regress fertilizationrate i.endometriosispresent
regress fertilizationrate i.endometrialmorphology
regress fertilizationrate endometrialthickness
 
// v.	Blastocyst rate 
regress blasto_rate age
regress blasto_rate bmi 
regress blasto_rate amhngml
regress blasto_rate totalafc
regress blasto_rate orpi
regress blasto_rate i.ovarianResponse
regress blasto_rate osi
regress blasto_rate i.fibroidpresence
regress blasto_rate i.adenomyosispresence
regress blasto_rate i.endometriosispresent
regress blasto_rate i.endometrialmorphology
regress blasto_rate endometrialthickness

// vi.	Implantation rate 
regress implantationrate age
regress implantationrate bmi 
regress implantationrate amhngml
regress implantationrate totalafc
regress implantationrate orpi
regress implantationrate i.ovarianResponse
regress implantationrate osi
regress implantationrate i.fibroidpresence
regress implantationrate i.adenomyosispresence
regress implantationrate i.endometriosispresent
regress implantationrate i.endometrialmorphology
regress implantationrate endometrialthickness

// vii.	Pregnancy rates 
regress urinepregrate age
regress urinepregrate bmi 
regress urinepregrate amhngml
regress urinepregrate totalafc
regress urinepregrate orpi
regress urinepregrate i.ovarianResponse
regress urinepregrate osi
regress urinepregrate i.fibroidpresence
regress urinepregrate i.adenomyosispresence
regress urinepregrate i.endometriosispresent
regress urinepregrate i.endometrialmorphology
regress urinepregrate endometrialthickness

// viii.	Ongoing pregnancy (number of fetuses at 6 weeks)
regress clinicalpregrate age
regress clinicalpregrate bmi 
regress clinicalpregrate amhngml
regress clinicalpregrate totalafc
regress clinicalpregrate orpi
regress clinicalpregrate i.ovarianResponse
regress clinicalpregrate osi
regress clinicalpregrate i.fibroidpresence
regress clinicalpregrate i.adenomyosispresence
regress clinicalpregrate i.endometriosispresent
regress clinicalpregrate i.endometrialmorphology
regress clinicalpregrate endometrialthickness



log close


stop
****************
lowess totalafc age
lowess totalafc age, mean mcolor(%0)
lowess amhngml age
lowess amhngml age, mean mcolor(%0)
lowess orpi age
lowess orpi age, mean mcolor(%0)
lowess foi age
lowess foi age, mean mcolor(%0)
lowess fort age
lowess fort age, mean mcolor(%0)


twoway scatter fort age, mcolor(*.6) || lfit fort age || lowess fort age ||, by(diminishedOvR_afc)

twoway scatter fort age, mcolor(%0) || lfit fort age || lowess fort age ||, by(diminishedOvR_afc)

lowess fort age, mcolor(%0) by(diminishedOvR_afc)



# Install and load necessary packages
# install.packages("rms")
# install.packages("nomogramEx")
# install.packages("DescTools")
# Install and load necessary packages
library(Hmisc)
library(dplyr)
library(tidyverse)
library(ggplot2)
library(splines)
library(labelled)
library(sjlabelled)
library(gtsummary)
library(broom)
library(psych)
library(forcats)

# library(haven)
# dams_data <- read_dta("C:/Users/Administrator/Desktop/Damalie work/damalie_data.dta")
library(readxl)
dams_data <- read_excel("C:/Users/Administrator/Desktop/Damalie work/damalie_data.xls")
View(dams_data)

str(dams_data)

summary(dams_data$amhngml)

# Assuming you h# Assuming you h# Assuming you have a dataframe with the variable names as column names
# Create a named vector for labelling the variables
variable_labels <- list(
  age = "Age",
  selfrecipientod = "Self Recipient OD",
  address = "Address",
  numberofpregnancies = "Number of Pregnancies",
  numberofdeliveries = "Number of Deliveries",
  numberofmiscarriages = "Number of Miscarriages",
  numberinducedabortions = "Number of Induced Abortions",
  numberectopicpregnancies = "Number of Ectopic Pregnancies",
  religion = "Religion",
  occupation = "Occupation",
  employment_status = "Employment Status",
  ethnicity = "Ethnicity",
  educationalevel = "Educational Level",
  maritalstatus = "Marital Status",
  durationofinfertilityyrs = "Duration of Infertility (years)",
  previousinfertilitytreammentr = "Previous Infertility Treatment",
  durationoftreatmentofoiplus = "Duration of Treatment for OI+",
  weight = "Weight (kg)",
  height = "Height (cm)",
  bmi = "Body Mass Index (BMI)",
  bmi_category = "BMI Category",
  systolicbp = "Systolic Blood Pressure (mmHg)",
  diastolicbp = "Diastolic Blood Pressure (mmHg)",
  bloodsugar = "Blood Sugar (mg/dL)",
  thyroidfunction = "Thyroid Function",
  historyofmedicalillness = "History of Medical Illness",
  hb = "Hemoglobin (Hb)",
  hbelectrophoresis = "Hb Electrophoresis",
  HBsag = "HBsAg",
  hepC = "Hepatitis C",
  HIV = "HIV",
  Vdrl = "VDRL",
  longtermmedicaltion = "Long-term Medication",
  alcoholorsocialdruguse = "Alcohol or Social Drug Use",
  fibroids = "Fibroids",
  fibroidpresence = "Fibroid Presence",
  fibroids_cat = "Fibroids Category",
  myomectomy = "Myomectomy",
  Adenomyosis = "Adenomyosis",
  Ovariancystpresent = "Ovarian Cyst Present",
  Previoushistoryofovariancyst = "Previous History of Ovarian Cyst",
  Endometriosispresent = "Endometriosis Present",
  Endometriosisgrade = "Endometriosis Grade",
  Previoushistoryofsalpingectom = "Previous History of Salpingectomy",
  Previoushistoryofappendicecto = "Previous History of Appendicectomy",
  Previoushistoryoflaparoscopy = "Previous History of Laparoscopy",
  laproscopyuterus = "Laparoscopy of Uterus",
  laparoscopyovaries = "Laparoscopy of Ovaries",
  laparoscopyfallopiantubes = "Laparoscopy of Fallopian Tubes",
  modeofpreviousdeliveries = "Mode of Previous Deliveries",
  amhngml = "Anti-Müllerian Hormone (ng/mL)",
  afcrightovary = "AFC Right Ovary",
  afcleftovary = "AFC Left Ovary",
  totalafc = "Total Antral Follicle Count",
  causeofinfertilitydiagnosed = "Cause of Infertility Diagnosed",
  recommendedtreatmentivfet = "Recommended Treatment: IVF-ET",
  recommendedtreatmenticsiet = "Recommended Treatment: ICSI-ET",
  recommededtreatmentsurrogacy = "Recommended Treatment: Surrogacy",
  startdose = "Start Dose",
  totaldoseofgonadotropinsiu = "Total Dose of Gonadotropins (IU)",
  doseofovulationtriggerbhcg = "Dose of Ovulation Trigger (β-hCG)",
  numberoffolliclesdevelopedri = "Number of Follicles Developed (Right)",
  numberoffolliclesdevelopedle = "Number of Follicles Developed (Left)",
  totalnumberoffolliclesdevelo = "Total Number of Follicles Developed",
  cyclecancelled = "Cycle Cancelled",
  numberofoocytesretrieved = "Number of Oocytes Retrieved",
  numberofeggsfertilized = "Number of Eggs Fertilized",
  numberofeggscleaved = "Number of Eggs Cleaved",
  numberofday5blastocysts = "Number of Day 5 Blastocysts",
  numberofetrecipient = "Number of ET Recipient",
  numberofembryostransferred = "Number of Embryos Transferred",
  urinepregnancyresults = "Urine Pregnancy Results",
  serumbhcglevel = "Serum β-hCG Level",
  chemicalpregnancy = "Chemical Pregnancy",
  numberoffetusesat6weekstra = "Number of Fetuses at 6 Weeks",
  endometrialthickness = "Endometrial Thickness",
  Endometrialmorphology = "Endometrial Morphology",
  AFC = "Antral Follicle Count",
  AMH = "Anti-Müllerian Hormone",
  self_recipient = "Self Recipient",
  self_donor = "Self Donor",
  recipient_self = "Recipient Self",
  orpi = "Ovarian Response Prediction Index",
  foi = "Follicle Output Rate",
  fort = "Follicular Output Rate",
  fertilizationrate = "Fertilization Rate",
  blasto_rate = "Blastocyst Rate",
  eggCleavagerate = "Egg Cleavage Rate",
  implantationrate = "Implantation Rate",
  urinePreg = "Urine Pregnancy",
  urinepregnancy = "Urine Pregnancy",
  urinepregrate = "Urine Pregnancy Rate",
  clinicalpregrate = "Clinical Pregnancy Rate",
  selfRecipientOd = "Self Recipient OD",
  cycleCancel = "Cycle Cancel",
  diminishedOvR_amh = "Diminished Ovarian Reserve (AMH)",
  diminishedOvR_afc = "Diminished Ovarian Reserve (AFC)",
  diminishedOvR_afc_new = "Diminished Ovarian Reserve (AFC New)",
  ovarianResponse = "Ovarian Response",
  osi = "Ovarian Sensitivity Index",
  bloodgroup = "Blood Group",
  adenomyosispresence = "Adenomyosis Presence",
  adenomyosis = "Adenomyosis",
  ovariancystpresent = "Ovarian Cyst Present",
  previoushistoryofovariancyst = "Previous History of Ovarian Cyst",
  endometriosispresent = "Endometriosis Present",
  endometriosisgrade = "Endometriosis Grade",
  previoushistoryofsalpingectom = "Previous History of Salpingectomy",
  previoushistoryofappendicecto = "Previous History of Appendicectomy",
  previoushistoryoflaparoscopy = "Previous History of Laparoscopy",
  endometrialmorphology = "Endometrial Morphology"
)

# Apply the labels to your dataset
for (var in names(variable_labels)) {
  var_label(dams_data[[var]]) <- variable_labels[[var]]
}


# Load your data
dams_data <- transform(dams_data, 
                       urinepregrate = as.numeric(urinepregrate)
) 
dams_data <- dams_data %>%
  mutate(across(where(is.character), as.factor))

# Perform Welch's ANOVA
welch_test <- oneway.test(numberofoocytesretrieved ~ selfrecipientod, data = dams_data, var.equal = FALSE)
welch_test <- oneway.test(fertilizationrate ~ selfrecipientod, data = dams_data, var.equal = TRUE)
welch_test <- oneway.test(blasto_rate ~ selfrecipientod, data = dams_data, var.equal = FALSE)
welch_test <- oneway.test(implantationrate ~ selfrecipientod, data = dams_data, var.equal = FALSE)
# Print the results
print(welch_test)

###############################
table(dams_data$hepC)
summary(dams_data$ovarianResponse)

## Demographic Distribution

# Load your data
# Assuming you have already loaded your data into dams_data dataframe

# Table 1: Demographic Distribution
demographic <- dams_data %>% 
  select(age, selfrecipientod, employment_status, religion, ethnicity, educationalevel, maritalstatus) %>%
  tbl_summary(by=selfrecipientod,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>% add_overall() %>% add_p(
    test = list(
      maritalstatus ~ fisher.test
    ),
    test.args = list(
      maritalstatus ~ list(simulate.p.value = TRUE, B = 1e5)
    )
  )

demographic1 <- dams_data %>% 
  select(age, bmi, bmi_category, systolicbp, diastolicbp, bloodsugar, selfrecipientod) %>%
  tbl_summary(by=selfrecipientod,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>% add_overall() %>% add_p(
    test = list(
      bmi_category ~ fisher.test
    ),
    test.args = list(
      bmi_category ~ list(simulate.p.value = TRUE, B = 1e5)
    )
  )

preghistory <- dams_data %>% 
  select(numberofpregnancies, numberofdeliveries, numberofmiscarriages, numberinducedabortions,
         numberectopicpregnancies, durationofinfertilityyrs, previousinfertilitytreammentr, 
         durationoftreatmentofoiplus, self_recipient) %>%
  tbl_summary(by=self_recipient,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>% add_overall() %>% add_p()

# Table 2: Fertility Treatment Summary

ftsummary1 <- dams_data %>% 
  select(cycleCancel, ovarianResponse, fertilizationrate, eggCleavagerate, blasto_rate, self_donor) %>%
  tbl_summary(by=self_donor,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>%
  add_overall() %>%
  add_p()


ftsummary2 <- dams_data %>% 
  select(implantationrate, urinepregrate, clinicalpregrate, self_recipient) %>%
  tbl_summary(by=self_recipient,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>%
  add_overall() %>%
  add_p()

ftsummary3 <- dams_data %>% 
  select(fibroidpresence, fibroids_cat, myomectomy, self_recipient) %>%
  tbl_summary(by=self_recipient,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>%
  add_overall() %>%
  add_p()


# ## Graphical presentations
# # Remove NA values from the dataset
# dams_data_clean <- dams_data %>%
#   drop_na(cycleCancel, ovarianResponse, fertilizationrate, eggCleavagerate, blasto_rate, 
#           self_donor, implantationrate, urinepregrate, clinicalpregrate, self_recipient)
# 
# # Fertility Treatment Summary (ftsummary1)
# # Bar chart for cycleCancel by self_donor
# ggplot(dams_data_clean, aes(x = cycleCancel, fill = as.factor(self_donor))) +
#   geom_bar(position = "dodge") +
#   labs(title = "Cycle Cancel by Self Donor",
#        x = "Cycle Cancel",
#        fill = "Self Donor") +
#   theme_minimal()
# 
# # Bar chart for ovarianResponse by self_donor
# ggplot(dams_data_clean, aes(x = ovarianResponse, fill = as.factor(self_donor))) +
#   geom_bar(position = "dodge") +
#   labs(title = "Ovarian Response by Self Donor",
#        x = "Ovarian Response",
#        fill = "Self Donor") +
#   theme_minimal()
# 
# # Histogram for fertilizationrate by self_donor
# ggplot(dams_data_clean, aes(x = fertilizationrate, fill = as.factor(self_donor))) +
#   geom_histogram(position = "dodge", binwidth = 0.05) +
#   labs(title = "Fertilization Rate by Self Donor",
#        x = "Fertilization Rate",
#        fill = "Self Donor") +
#   theme_minimal()
# 
# # Histogram for eggCleavagerate by self_donor
# ggplot(dams_data_clean, aes(x = eggCleavagerate, fill = as.factor(self_donor))) +
#   geom_histogram(position = "dodge", binwidth = 0.05) +
#   labs(title = "Egg Cleavage Rate by Self Donor",
#        x = "Egg Cleavage Rate",
#        fill = "Self Donor") +
#   theme_minimal()
# 
# # Histogram for blasto_rate by self_donor
# ggplot(dams_data_clean, aes(x = blasto_rate, fill = as.factor(self_donor))) +
#   geom_histogram(position = "dodge", binwidth = 0.05) +
#   labs(title = "Blastocyst Rate by Self Donor",
#        x = "Blastocyst Rate",
#        fill = "Self Donor") +
#   theme_minimal()
# 
# # Fertility Treatment Summary (ftsummary2)
# # Bar chart for implantationrate by self_recipient
# ggplot(dams_data_clean, aes(x = as.factor(implantationrate), fill = as.factor(self_recipient))) +
#   geom_bar(position = "dodge") +
#   labs(title = "Implantation Rate by Self Recipient",
#        x = "Implantation Rate",
#        fill = "Self Recipient") +
#   theme_minimal()
# 
# # Bar chart for urinepregrate by self_recipient
# ggplot(dams_data_clean, aes(x = as.factor(urinepregrate), fill = as.factor(self_recipient))) +
#   geom_bar(position = "dodge") +
#   labs(title = "Urine Pregnancy Rate by Self Recipient",
#        x = "Urine Pregnancy Rate",
#        fill = "Self Recipient") +
#   theme_minimal()
# 
# # Bar chart for clinicalpregrate by self_recipient
# ggplot(dams_data_clean, aes(x = as.factor(clinicalpregrate), fill = as.factor(self_recipient))) +
#   geom_bar(position = "dodge") +
#   labs(title = "Clinical Pregnancy Rate by Self Recipient",
#        x = "Clinical Pregnancy Rate",
#        fill = "Self Recipient") +
#   theme_minimal()



# Table 3: Medical History Summary
medhistsummary <- dams_data %>% 
  select(historyofmedicalillness, hb, hbelectrophoresis, bloodgroup, 
         HBsag, hepC, HIV, Vdrl, longtermmedicaltion, alcoholorsocialdruguse) %>%
  tbl_summary(by=NULL,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels()

orpiFoiFortOsi_OR_summary <- dams_data %>% 
  select(orpi, foi, fort, osi, diminishedOvR_amh, diminishedOvR_afc_new, 
         ovarianResponse, selfrecipientod) %>%
  tbl_summary(by=selfrecipientod,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>%
  add_overall() %>%
  add_p()


dose_Oocyte_summary <- dams_data %>% 
  select(startdose, totaldoseofgonadotropinsiu, doseofovulationtriggerbhcg, cycleCancel, 
         totalnumberoffolliclesdevelo, numberofoocytesretrieved, ovarianResponse, orpi,
         foi, fort, osi, self_donor) %>%
  tbl_summary(by=self_donor,
              statistic = list(),
              type = all_dichotomous() ~ "categorical",
              digits = all_continuous() ~ 2,
              missing = "no"
  ) %>% bold_labels() %>%
  add_overall() %>%
  add_p()

# Print the tables
demographic
demographic1
preghistory
ftsummary1
ftsummary2
medhistsummary
orpiFoiFortOsi_OR_summary
dose_Oocyte_summary

#####
summary(dams_data$urinepregrate)
# Using tapply
# Calculate mean urinepregrate for each group in self_recipient
means <- tapply(dams_data$urinepregrate, dams_data$self_recipient, mean, na.rm = TRUE)
"Means:"
means

# Calculate standard deviation of urinepregrate for each group in self_recipient
sds <- tapply(dams_data$urinepregrate, dams_data$self_recipient, sd, na.rm = TRUE)
"Standard Deviations:"
sds

summary_stats_mean_sd <- dams_data %>%
  group_by(self_recipient) %>%
  summarize(
    mean_urinepregrate = mean(urinepregrate, na.rm = TRUE),
    sd_urinepregrate = sd(urinepregrate, na.rm = TRUE)
  )


"Summary Statistics of fertilization rate (mean and sd):"
summary_stats_mean_sd


# Calculate summary statistics median and iqr
summary_stats_median_iqr1 <- dams_data %>%
  group_by(self_donor) %>%
  summarize(
    median_startdose = median(startdose, na.rm = TRUE),
    iqr_startdose = IQR(startdose, na.rm = TRUE),
    lower_quartile = quantile(startdose, 0.25, na.rm = TRUE),
    upper_quartile = quantile(startdose, 0.75, na.rm = TRUE)
  )

"Start dose Median and iqr"
summary(dams_data$startdose)
"Summary Statistics startdose (median and iqr):"
summary_stats_median_iqr1

# Calculate standard deviation of urinepregrate for each group in self_recipient
sds1 <- tapply(dams_data$doseofovulationtriggerbhcg, dams_data$self_donor, sd, na.rm = TRUE)
"Standard Deviations:"
sds1

summary_stats_mean_sd_totaldose <- dams_data %>%
  group_by(self_donor) %>%
  summarize(
    mean_doseofovulationtriggerbhcg = mean(doseofovulationtriggerbhcg, na.rm = TRUE),
    sd_doseofovulationtriggerbhcg = sd(doseofovulationtriggerbhcg, na.rm = TRUE)
  )

"Start dose Median and iqr"
summary(dams_data$doseofovulationtriggerbhcg)

"Summary Statistics of fertilization rate (mean and sd):"
summary_stats_mean_sd_totaldose


##############################
##### Graphs using spline ####
##############################

## AFC (all)
ggplot(dams_data, aes(x = age, y = totalafc)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of AFC over age",
    x = "Age in years",
    y = "Antral Follicle Count") +
  theme_minimal()

## AFC with intercept (all)
ggplot(dams_data, aes(x = age, y = totalafc)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 5, linetype = "dashed") +
  labs(#title = "Nomogram Plot of AFC over age",
    x = "Age in years",
    y = "Antral Follicle Count") +
  theme_minimal()

##############################

## AFC (new)
ggplot(dams_data, aes(x = age, y = AFC)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of AFC over age",
    x = "Age in years",
    y = "Antral Follicle Count") +
  theme_minimal()

## AFC with intercept (new)
ggplot(dams_data, aes(x = age, y = AFC)) +
  geom_point(alpha = 0.5) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 5, linetype = "dashed") +
  labs(#title = "Nomogram Plot of AFC over age",
    x = "Age in years",
    y = "Antral Follicle Count") +
  theme_minimal()

##############################
## AMH
ggplot(dams_data, aes(x = age, y = AMH)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of AMH over age",
    x = "Age in years",
    y = "Anti-Müllerian Hormone (ng/ml)") +
  theme_minimal()

## AMH with intercept
ggplot(dams_data, aes(x = age, y = AMH)) +
  geom_point(alpha = 0.4) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 1.2, linetype = "dashed") +
  labs(#title = "Nomogram Plot of AMH over age",
    x = "Age in years",
    y = "Anti-Müllerian Hormone (ng/ml)") +
  theme_minimal()

##########################################

# Add a new variable to distinguish between the two plots
dams_data_long <- dams_data %>%
  pivot_longer(cols = c(AFC, AMH), 
               names_to = "variable", 
               values_to = "value")

ggplot(dams_data_long, aes(x = age, y = value)) +
  geom_point(alpha = 0.4) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(data = subset(dams_data_long, variable == "AMH"), aes(yintercept = 1.2), linetype = "dashed", color = "orange") +
  geom_hline(data = subset(dams_data_long, variable == "AFC"), aes(yintercept = 5), linetype = "dashed", color = "blue") +
  labs(x = "Age in years", y = "") +
  theme_minimal() +
  facet_wrap(~ variable, scales = "free_y", 
             labeller = labeller(variable = c(AFC = "Antral Follicle Count", AMH = "Anti-Müllerian Hormone (ng/mL)")))

# Create the combined plot
ggplot(dams_data_long, aes(x = age, y = value)) +
  geom_point(alpha = 0.5) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 1.2, linetype = "dashed") +
  geom_hline(yintercept = 5, linetype = "dashed") +
  labs(x = "Age in years", y = "") +
  theme_minimal() +
  facet_wrap(~ variable, scales = "free_y", 
             labeller = labeller(variable = c(AFC = "Antral Follicle Count", AMH = "Anti-Müllerian Hormone (ng/mL)")))



# Create the combined plot with fixed y-axis scale
ggplot(dams_data_long, aes(x = age, y = value)) +
  geom_point(alpha = 0.5) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(x = "Age in years", y = "") +
  theme_minimal() +
  facet_wrap(~ variable, scales = "free_y", 
             labeller = labeller(variable = c(AFC = "Antral Follicle Count", AMH = "Anti-Müllerian Hormone (ng/mL)")))


# corelation between AFC, AMH and age
pairs.panels(dams_data[c("AFC", "AMH", "age")])

pairs.panels(dams_data[c("AFC", "AMH")])

pairs.panels(dams_data[c("AFC", "age")])

pairs.panels(dams_data[c("AMH", "age")])

## AFC agaisnt AMH
ggplot(dams_data, aes(x = AFC, y = amhngml)) +
  geom_point(alpha = 0.5) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 1.2, linetype = "dashed") +
  labs(#title = "Nomogram Plot of AMH against AFC",
    x = "Antral Follicle Count",
    y = "Anti-Müllerian Hormone") +
  theme_minimal()

ggplot(dams_data, aes(x = age)) +
  geom_point(aes(y = totalafc, color = "AFC"), alpha = 0.5) +
  geom_smooth(aes(y = totalafc, color = "AFC"), method = "lm", formula = y ~ ns(x, 3), se = FALSE) +
  geom_point(aes(y = amhngml, color = "AMH"), alpha = 0.5) +
  geom_smooth(aes(y = amhngml, color = "AMH"), method = "lm", formula = y ~ ns(x, 3), se = FALSE) +
  labs(#title = "Nomogram Plot of AFC and AMH over age",
    x = "Age in years",
    y = "count / (ng/ml)",
    color = "Variable") +
  theme_minimal() +
  scale_color_manual(values = c("AFC" = "blue", "AMH" = "orange"))



## to use this graph = excellent
ggplot(dams_data, aes(x = age)) +
  geom_point(aes(y = AFC, color = "AFC"), alpha = 0.5) +
  geom_smooth(aes(y = AFC, color = "AFC"), method = "lm", formula = y ~ ns(x, 3), se = FALSE) +
  geom_point(aes(y = amhngml, color = "AMH"), alpha = 0.5) +
  geom_smooth(aes(y = amhngml, color = "AMH"), method = "lm", formula = y ~ ns(x, 3), se = FALSE) +
  geom_hline(yintercept = 1.2, linetype = "dashed", color = "orange") +  # Horizontal line for AMH
  geom_hline(yintercept = 5, linetype = "dashed", color = "blue") +  # Horizontal line for AFC
  labs(#title = "Nomogram Plot of AFC and AMH over age",
    x = "Age in years",
    y = "count / (ng/ml)",
    color = "Variable") +
  theme_minimal() +
  scale_color_manual(values = c("AFC" = "blue", "AMH" = "orange"))




###########################
############################
### afc and amh against bmi
# Add a new variable to distinguish between the two plots
dams_data_long <- dams_data %>%
  pivot_longer(cols = c(AFC, AMH), 
               names_to = "variable", 
               values_to = "value")

ggplot(dams_data_long, aes(x = bmi, y = value)) +
  geom_point(alpha = 0.5) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(data = subset(dams_data_long, variable == "AMH"), aes(yintercept = 1.2), linetype = "dashed", color = "orange") +
  geom_hline(data = subset(dams_data_long, variable == "AFC"), aes(yintercept = 5), linetype = "dashed", color = "blue") +
  labs(x = "Body Mass Index (kg/m²)", y = "") +
  theme_minimal() +
  facet_wrap(~ variable, scales = "free_y", 
             labeller = labeller(variable = c(AFC = "Antral Follicle Count", AMH = "Anti-Müllerian Hormone (ng/mL)")))

# Create the combined plot
ggplot(dams_data_long, aes(x = bmi, y = value)) +
  geom_point(alpha = 0.5) +  # Add points with some transparency
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 1.2, linetype = "dashed") +
  geom_hline(yintercept = 5, linetype = "dashed") +
  labs(x = "Body Mass Index (kg/m²)", y = "") +
  theme_minimal() +
  facet_wrap(~ variable, scales = "free_y", 
             labeller = labeller(variable = c(AFC = "Antral Follicle Count", AMH = "Anti-Müllerian Hormone (ng/mL)")))




##########################################
#############################
## ORPI
ggplot(dams_data, aes(x = age, y = orpi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of ORPI over age",
    x = "Age in years",
    y = "Ovarian Response Prediction Index") +
  theme_minimal()

## ORPI with intercept
ggplot(dams_data, aes(x = age, y = orpi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 1.5, linetype = "dashed") +
  labs(#title = "Nomogram Plot of ORPI over age",
    x = "Age in years",
    y = "Ovarian Response Prediction Index") +
  theme_minimal()

#############################
## FORT
ggplot(dams_data, aes(x = age, y = fort)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of FORT over age",
    x = "Age in years",
    y = "FORT") +
  theme_minimal()

## FORT with intercept
ggplot(dams_data, aes(x = age, y = fort)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 144, linetype = "dashed") +
  labs(#title = "Nomogram Plot of FORT over age",
    x = "Age in years",
    y = "FORT") +
  theme_minimal()

#############################
## FOI
ggplot(dams_data, aes(x = age, y = foi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of FOI over age",
    x = "Age in years",
    y = "FOI") +
  theme_minimal()

## FOI with intercept
ggplot(dams_data, aes(x = age, y = foi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 1.2, linetype = "dashed") +
  labs(#title = "Nomogram Plot of FOI over age",
    x = "Age in years",
    y = "FOI") +
  theme_minimal()

###############################

## OSI
ggplot(dams_data, aes(x = age, y = osi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of Ovarian Sensitivity Index over age",
    x = "Age in years",
    y = "Ovarian Sensitivity Index") +
  theme_minimal()

## OSI with intercept
ggplot(dams_data, aes(x = age, y = osi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_hline(yintercept = 4, linetype = "dashed") +
  labs(#title = "Nomogram Plot of Ovarian Sensitivity Index over age",
    x = "Age in years",
    y = "Ovarian Sensitivity Index") +
  theme_minimal()
###############################

## OSI and oocytes retrieved
ggplot(dams_data, aes(x = numberofoocytesretrieved, y = osi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  labs(#title = "Nomogram Plot of Ovarian Sensitivity Index over oocytes retrieved",
    x = "Oocytes retrieved",
    y = "Ovarian Sensitivity Index") +
  theme_minimal()

## OSI and oocytes retrieved with intercept
ggplot(dams_data, aes(x = numberofoocytesretrieved, y = osi)) +
  geom_smooth(method = "lm", formula = y ~ ns(x, 3), col = "orange") +
  geom_vline(xintercept = 8, linetype = "dashed") +
  labs(#title = "Nomogram Plot of Ovarian Sensitivity Index over oocytes retrieved",
    x = "Oocyte retrieved",
    y = "Ovarian Sensitivity Index") +
  theme_minimal()


################################################################################
################################################################################
###########################################
###########################################

## Regressions

# Define a function to perform logistic regression and create a gtsummary table
logistic_regression_gtsummary <- function(formula, data) {
  model <- glm(formula, data = data, family = binomial)
  tbl_regression(model, exponentiate = TRUE)
}

# Define a function to perform linear regression and create a gtsummary table
linear_regression_gtsummary <- function(formula, data) {
  model <- lm(formula, data = data)
  tbl_regression(model)
}

# Objective 2 - Physiological Factors
# # Adjusted regressions
# # linear_afc_physio_anato <- linear_regression_gtsummary(AFC ~ age + bmi + factor(fibroidpresence) + 
# #                                                          factor(fibroid) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
# #                                                          factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
# #                                                          factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
# #                                                          previoushistoryofappendicecto, dams_data)
# # 
# # linear_amh_physio_anato <- linear_regression_gtsummary(amhngml ~ age + bmi + factor(fibroidpresence) + 
# #                                                          factor(fibroid) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
# #                                                          factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
# #                                                          factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
# #                                                          previoushistoryofappendicecto, dams_data)
# 
# logistic_diminishedOvR_amh_physio_anato <- logistic_regression_gtsummary(diminishedOvR_amh ~ age + bmi + factor(fibroidpresence) + 
#                                                                            factor(fibroid) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
#                                                                            factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
#                                                                            factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
#                                                                            previoushistoryofappendicecto, dams_data)
# 
# 
# logistic_diminishedOvR_afc_physio_anato <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ age + bmi + factor(fibroidpresence) + 
#                                                                            factor(fibroid) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
#                                                                            factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
#                                                                            factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
#                                                                            previoushistoryofappendicecto, dams_data)
# 
# 
# linear_orpi_physio_anato <- linear_regression_gtsummary(orpi ~ age + bmi + factor(fibroidpresence) + 
#                                                                            factor(fibroid) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
#                                                                            factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
#                                                                            factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
#                                                                            previoushistoryofappendicecto, dams_data)
# 
# # undajusted regressions
# # Age
# logistic_diminishedOvR_afc_new_age <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ age, dams_data)
# logistic_diminishedOvR_amh_age <- logistic_regression_gtsummary(diminishedOvR_amh ~ age, dams_data)
# linear_orpi_age <- linear_regression_gtsummary(orpi ~ age, dams_data)
# 
# 
# # BMI
# logistic_diminishedOvR_afc_bmi <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ bmi, dams_data)
# logistic_diminishedOvR_amh_bmi <- logistic_regression_gtsummary(diminishedOvR_amh ~ bmi, dams_data)
# linear_orpi_bmi <- linear_regression_gtsummary(orpi ~ bmi, dams_data)
# 
# logistic_diminishedOvR_afc_bmi_category <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(bmi_category), dams_data)
# logistic_diminishedOvR_amh_bmi_category <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(bmi_category), dams_data)
# linear_orpi_bmi_category <- linear_regression_gtsummary(orpi ~ factor(bmi_category), dams_data)
# 
# logistic_diminishedOvR_afc_age_bmi <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ age + bmi, dams_data)
# logistic_diminishedOvR_amh_age_bmi <- logistic_regression_gtsummary(diminishedOvR_amh ~ age + bmi, dams_data)
# linear_orpi_age_bmi <- linear_regression_gtsummary(orpi ~ age + bmi, dams_data)
# 
# # Anatomical Factors
# # Fibroids
# logistic_diminishedOvR_afc_fibroidpresence <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(fibroidpresence), dams_data)
# logistic_diminishedOvR_amh_fibroidpresence <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(fibroidpresence), dams_data)
# linear_orpi_fibroidpresence <- linear_regression_gtsummary(orpi ~ factor(fibroidpresence), dams_data)
# 
# logistic_diminishedOvR_afc_fibroid <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(fibroid), dams_data)
# logistic_diminishedOvR_amh_fibroid <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(fibroid), dams_data)
# linear_orpi_fibroid <- linear_regression_gtsummary(orpi ~ factor(fibroid), dams_data)
# 
# # Adenomyosis
# logistic_ovarianResponse_adenomyosispresence <- logistic_regression_gtsummary(ovarianResponse ~ factor(adenomyosispresence), dams_data)
# linear_orpi_adenomyosispresence <- linear_regression_gtsummary(orpi ~ adenomyosispresence, dams_data)
# 
# # Ovarian Cyst
# logistic_diminishedOvR_afc_ovariancystpresent <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(ovariancystpresent), dams_data)
# logistic_diminishedOvR_amh_ovariancystpresent <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(ovariancystpresent), dams_data)
# linear_orpi_ovariancystpresent <- linear_regression_gtsummary(orpi ~ factor(ovariancystpresent), dams_data)
# 
# logistic_diminishedOvR_afc_previoushistoryofovariancyst <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(previoushistoryofovariancyst), dams_data)
# logistic_diminishedOvR_amh_previoushistoryofovariancyst <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryofovariancyst), dams_data)
# linear_orpi_previoushistoryofovariancyst <- linear_regression_gtsummary(orpi ~ factor(previoushistoryofovariancyst), dams_data)
# 
# # Endometriosis
# logistic_diminishedOvR_afc_endometriosispresent <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(endometriosispresent), dams_data)
# logistic_diminishedOvR_amh_endometriosispresent <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(endometriosispresent), dams_data)
# linear_orpi_endometriosispresent <- linear_regression_gtsummary(orpi ~ factor(endometriosispresent), dams_data)
# 
# logistic_diminishedOvR_afc_endometriosisgrade <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(endometriosisgrade), dams_data)
# logistic_diminishedOvR_amh_endometriosisgrade <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(endometriosisgrade), dams_data)
# linear_orpi_endometriosisgrade <- linear_regression_gtsummary(orpi ~ factor(endometriosisgrade), dams_data)
# 
# # Previous history of salpingectomy
# logistic_diminishedOvR_afc_previoushistoryofsalpingectomy <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(previoushistoryofsalpingectom), dams_data)
# logistic_diminishedOvR_amh_previoushistoryofsalpingectomy <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryofsalpingectom), dams_data)
# linear_orpi_previoushistoryofsalpingectomy <- linear_regression_gtsummary(orpi ~ factor(previoushistoryofsalpingectom), dams_data)
# 
# # Previous history of appendicectomy
# logistic_diminishedOvR_afc_previoushistoryofappendicectomy <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(previoushistoryofappendicecto), dams_data)
# logistic_diminishedOvR_amh_previoushistoryofappendicectomy <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryofappendicecto), dams_data)
# linear_orpi_previoushistoryofappendicectomy <- linear_regression_gtsummary(orpi ~ factor(previoushistoryofappendicecto), dams_data)
# 
# # Previous history of laparoscopy
# logistic_diminishedOvR_afc_previoushistoryoflaparoscopy <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ factor(previoushistoryoflaparoscopy), dams_data)
# logistic_diminishedOvR_amh_previoushistoryoflaparoscopy <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryoflaparoscopy), dams_data)
# linear_orpi_previoushistoryoflaparoscopy <- linear_regression_gtsummary(orpi ~ factor(previoushistoryoflaparoscopy), dams_data)
# 
# 
# # linear_afc_physio_anato
# # linear_amh_physio_anato
# logistic_diminishedOvR_amh_physio_anato
# logistic_diminishedOvR_afc_physio_anato
# linear_orpi_physio_anato
# 
# logistic_diminishedOvR_afc_bmi
# logistic_diminishedOvR_amh_bmi
# linear_orpi_bmi
# logistic_diminishedOvR_afc_bmi_category
# logistic_diminishedOvR_amh_bmi_category
# linear_orpi_bmi_category
# logistic_diminishedOvR_afc_age_bmi
# logistic_diminishedOvR_amh_age_bmi
# linear_orpi_age_bmi
# logistic_diminishedOvR_afc_fibroidpresence
# logistic_diminishedOvR_amh_fibroidpresence
# linear_orpi_fibroidpresence
# logistic_diminishedOvR_afc_fibroid
# logistic_diminishedOvR_amh_fibroid
# linear_orpi_fibroid
# logistic_ovarianResponse_adenomyosispresence
# linear_orpi_adenomyosispresence
# logistic_diminishedOvR_afc_ovariancystpresent
# logistic_diminishedOvR_amh_ovariancystpresent
# linear_orpi_ovariancystpresent
# logistic_diminishedOvR_afc_previoushistoryofovariancyst
# logistic_diminishedOvR_amh_previoushistoryofovariancyst
# linear_orpi_previoushistoryofovariancyst
# logistic_diminishedOvR_afc_endometriosispresent
# logistic_diminishedOvR_amh_endometriosispresent
# linear_orpi_endometriosispresent
# logistic_diminishedOvR_afc_endometriosisgrade
# logistic_diminishedOvR_amh_endometriosisgrade
# linear_orpi_endometriosisgrade
# logistic_diminishedOvR_afc_previoushistoryofsalpingectomy
# logistic_diminishedOvR_amh_previoushistoryofsalpingectomy
# linear_orpi_previoushistoryofsalpingectomy
# logistic_diminishedOvR_afc_previoushistoryofappendicectomy
# logistic_diminishedOvR_amh_previoushistoryofappendicectomy
# linear_orpi_previoushistoryofappendicectomy
# logistic_diminishedOvR_afc_previoushistoryoflaparoscopy
# logistic_diminishedOvR_amh_previoushistoryoflaparoscopy
# linear_orpi_previoushistoryoflaparoscopy

#####
#####
# Define the variables of interest for logistic regression models
variables_logistic <- c("age", "bmi", "factor(fibroidpresence)", "factor(myomectomy)", "factor(adenomyosispresence)", 
                        "previoushistoryofsalpingectom", "factor(ovariancystpresent)", "factor(previoushistoryofovariancyst)", 
                        "factor(endometriosispresent)", "factor(previoushistoryoflaparoscopy)", "previoushistoryofappendicecto")

## diminished AMH model
# Create logistic regression models for each variable and the combined model
logistic_diminishedOvR_amh_models <- lapply(variables_logistic, function(var) {
  glm(as.formula(paste("diminishedOvR_amh ~", var)), data = dams_data, family = binomial(link = "logit"))
})

logistic_diminishedOvR_amh_physio_anato <- glm(diminishedOvR_amh ~ age + bmi + factor(fibroidpresence) + 
                                                 factor(myomectomy) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
                                                 factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
                                                 factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
                                                 previoushistoryofappendicecto, data = dams_data, family = binomial(link = "logit"))

# Summarize the models using gtsummary
logistic_summaries_d_amh <- lapply(logistic_diminishedOvR_amh_models, tbl_regression, exponentiate = TRUE)
logistic_combined_summary_d_amh <- tbl_regression(logistic_diminishedOvR_amh_physio_anato, exponentiate = TRUE) %>% bold_labels()

# Combine each unadjusted summary with the adjusted summary and stack them
combined_summaries_d_amh <- lapply(logistic_summaries_d_amh, function(summary) {
  tbl_merge(tbls = list(summary, logistic_combined_summary_d_amh), tab_spanner = c("**Unadjusted**", "**Adjusted**"))
})

final_combined_summaries_d_amh <- tbl_stack(combined_summaries_d_amh)

# Display the final summary
final_combined_summaries_d_amh


## diminishes AFC model
# Create logistic regression models for each variable and the combined model
logistic_diminishedOvR_afc_new_models <- lapply(variables_logistic, function(var) {
  glm(as.formula(paste("diminishedOvR_afc_new ~", var)), data = dams_data, family = binomial(link = "logit"))
})

logistic_diminishedOvR_afc_new_physio_anato <- glm(diminishedOvR_afc_new ~ age + bmi + factor(fibroidpresence) + 
                                                     factor(myomectomy) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
                                                     factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
                                                     factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
                                                     previoushistoryofappendicecto, data = dams_data, family = binomial(link = "logit"))

# Summarize the models using gtsummary
logistic_summaries_d_afc <- lapply(logistic_diminishedOvR_afc_new_models, tbl_regression, exponentiate = TRUE)
logistic_combined_summary_d_afc <- tbl_regression(logistic_diminishedOvR_afc_new_physio_anato, exponentiate = TRUE) %>% bold_labels()

# Combine each unadjusted summary with the adjusted summary and stack them
combined_summaries_d_afc <- lapply(logistic_summaries_d_afc, function(summary) {
  tbl_merge(tbls = list(summary, logistic_combined_summary_d_afc), tab_spanner = c("**Unadjusted**", "**Adjusted**"))
})

final_combined_summaries_d_afc <- tbl_stack(combined_summaries_d_afc)

# Display the final summary
final_combined_summaries_d_afc


## ORPI model
# Define the variables of interest for linear regression models
variables_linear <- c("age", "bmi", "factor(fibroidpresence)", "factor(myomectomy)", "factor(adenomyosispresence)", 
                      "previoushistoryofsalpingectom", "factor(ovariancystpresent)", "factor(previoushistoryofovariancyst)", 
                      "factor(endometriosispresent)", "factor(previoushistoryoflaparoscopy)", "previoushistoryofappendicecto")

# Create a custom function to perform linear regression and summarize it using gtsummary
linear_regression_gtsummary <- function(formula, data) {
  model <- lm(formula, data = data)
  tbl_regression(model)
}

# Create linear regression models for each variable and the combined model
linear_orpi_model <- lapply(variables_linear, function(var) {
  linear_regression_gtsummary(as.formula(paste("orpi ~", var)), dams_data)
})

linear_orpi_physio_anato <- linear_regression_gtsummary(orpi ~ age + bmi + factor(fibroidpresence) + 
                                                          factor(myomectomy) + factor(adenomyosispresence) + previoushistoryofsalpingectom +
                                                          factor(ovariancystpresent) + factor(previoushistoryofovariancyst) + 
                                                          factor(endometriosispresent) + factor(previoushistoryoflaparoscopy) +
                                                          previoushistoryofappendicecto, dams_data)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_orpi <- lapply(linear_orpi_model, function(summary) {
  tbl_merge(
    tbls = list(summary, linear_orpi_physio_anato),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_orpi <- tbl_stack(combined_summaries_orpi)

# Display the final summary
final_combined_summaries_orpi
#####
#####


# Objective 3 - Cycle cancellation and other outcomes

# Filter the data to exclude rows with NA in self_donor
filtered_data_SD <- dams_data %>%
  filter(!is.na(self_donor))

# Filter the data to exclude rows with NA in recipient_self
filtered_data_SR <- dams_data %>%
  filter(!is.na(recipient_self))

# Filter the data to exclude rows with NA in self_recipient
filtered_data_S_R <- dams_data %>%
  filter(!is.na(self_recipient))

# Filter the data to exclude rows with NA in self_donor
filtered_data_S <- dams_data %>%
  filter(!is.na(self_donor))

# Filter the data to exclude rows with NA in recipient_self
filtered_data_R <- dams_data %>%
  filter(!is.na(recipient_self))

# Filter the data to exclude rows with NA in self_recipient
filtered_data_D <- dams_data %>%
  filter(!is.na(self_recipient))


##########################################
##########################################
# Implantation rate
linear_implantationrate_age <- linear_regression_gtsummary(implantationrate ~ age, filtered_data_D)

linear_implantationrate_bmi <- linear_regression_gtsummary(implantationrate ~ bmi, filtered_data_D)

linear_implantationrate_amh <- linear_regression_gtsummary(implantationrate ~ amhngml, filtered_data_R)

linear_implantationrate_afc <- linear_regression_gtsummary(implantationrate ~ AFC, filtered_data_R)

linear_implantationrate_orpi <- linear_regression_gtsummary(implantationrate ~ orpi, filtered_data_R)

linear_implantationrate_fibpres <- linear_regression_gtsummary(implantationrate ~ factor(fibroidpresence), filtered_data_R)

linear_implantationrate_fibroid <- linear_regression_gtsummary(implantationrate ~ myomectomy, filtered_data_R)

linear_implantationrate_endomorph <- linear_regression_gtsummary(implantationrate ~ endometrialmorphology, filtered_data_R)

linear_implantationrate_endothick <- linear_regression_gtsummary(implantationrate ~ endometrialthickness, filtered_data_R)

linear_implantationrate_ovarianResp <- linear_regression_gtsummary(implantationrate ~ ovarianResponse, filtered_data_R)

linear_implantationrate_osi <- linear_regression_gtsummary(implantationrate ~ osi, filtered_data_R)

dams_data$implantationrate
# adjusted
linear_implantationrate_adj <- linear_regression_gtsummary(implantationrate ~ age + bmi + amhngml+ 
                                                             AFC + orpi, filtered_data_S)

linear_implantationrate_adj1 <- linear_regression_gtsummary(implantationrate ~ factor(fibroidpresence) + ovarianResponse +
                                                              osi + endometrialmorphology + endometrialthickness
                                                            , filtered_data_SR)

# fertilization rate
linear_fertilizationrate_age <- linear_regression_gtsummary(fertilizationrate ~ age, filtered_data_S)

linear_fertilizationrate_bmi <- linear_regression_gtsummary(fertilizationrate ~ bmi, filtered_data_S)

linear_fertilizationrate_amh <- linear_regression_gtsummary(fertilizationrate ~ amhngml, filtered_data_S_R)

linear_fertilizationrate_afc <- linear_regression_gtsummary(fertilizationrate ~ AFC, filtered_data_S_R)

linear_fertilizationrate_orpi <- linear_regression_gtsummary(fertilizationrate ~ orpi, filtered_data_S_R)

linear_fertilizationrate_fibpres <- linear_regression_gtsummary(fertilizationrate ~ factor(fibroidpresence), filtered_data_S_R)

linear_fertilizationrate_fibroid <- linear_regression_gtsummary(fertilizationrate ~  myomectomy, filtered_data_S_R)

linear_fertilizationrate_adenom <- linear_regression_gtsummary(fertilizationrate ~ adenomyosispresence, filtered_data_S_R)

linear_fertilizationrate_endomorph <- linear_regression_gtsummary(fertilizationrate ~ endometrialmorphology, filtered_data_S_R)

linear_fertilizationrate_endom <- linear_regression_gtsummary(fertilizationrate ~ Endometriosispresent, filtered_data_S_R)

linear_fertilizationrate_endothick <- linear_regression_gtsummary(fertilizationrate ~ endometrialthickness, filtered_data_S_R)

linear_fertilizationrate_ovarianResp <- linear_regression_gtsummary(fertilizationrate ~ ovarianResponse, filtered_data_S_R)

linear_fertilizationrate_osi <- linear_regression_gtsummary(fertilizationrate ~ osi, filtered_data_S_R)

#adjusted

linear_fertilizationrate_adj <- linear_regression_gtsummary(fertilizationrate ~ age + bmi + amhngml+ 
                                                              AFC + orpi, filtered_data_S_R)

linear_fertilizationrate_adj1 <- linear_regression_gtsummary(fertilizationrate ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                               osi + endometrialmorphology + Endometriosispresent + endometrialthickness
                                                             , filtered_data_S_R)


# Clinical pregnancy rate
linear_clinicalpregrate_age <- linear_regression_gtsummary(clinicalpregrate ~ age, filtered_data_SR)

linear_clinicalpregrate_bmi <- linear_regression_gtsummary(clinicalpregrate ~ bmi, filtered_data_SR)

linear_clinicalpregrate_amh <- linear_regression_gtsummary(clinicalpregrate ~ amhngml, filtered_data_SR)

linear_clinicalpregrate_afc <- linear_regression_gtsummary(clinicalpregrate ~ AFC, filtered_data_SR)

linear_clinicalpregrate_orpi <- linear_regression_gtsummary(clinicalpregrate ~ orpi, filtered_data_SR)

linear_clinicalpregrate_fibpres <- linear_regression_gtsummary(clinicalpregrate ~ factor(fibroidpresence), filtered_data_SR)

linear_clinicalpregrate_fibroid <- linear_regression_gtsummary(clinicalpregrate ~ myomectomy, filtered_data_SR)

linear_clinicalpregrate_adenom <- linear_regression_gtsummary(clinicalpregrate ~ adenomyosispresence, filtered_data_SR)

linear_clinicalpregrate_endomorph <- linear_regression_gtsummary(clinicalpregrate ~ endometrialmorphology, filtered_data_SR)

linear_clinicalpregrate_endom <- linear_regression_gtsummary(clinicalpregrate ~ Endometriosispresent, filtered_data_SR)

linear_clinicalpregrate_endothick <- linear_regression_gtsummary(clinicalpregrate ~ endometrialthickness, filtered_data_SR)

linear_clinicalpregrate_ovarianResp <- linear_regression_gtsummary(clinicalpregrate ~ ovarianResponse, filtered_data_SR)

linear_clinicalpregrate_osi <- linear_regression_gtsummary(clinicalpregrate ~ osi, filtered_data_SR)

#adjusted

linear_clinicalpregrate_adj <- linear_regression_gtsummary(clinicalpregrate ~ age + bmi + amhngml+ 
                                                             AFC + orpi, filtered_data_SR)

linear_clinicalpregrate_adj1 <- linear_regression_gtsummary(clinicalpregrate ~ fibroidpresence + myomectomy + ovarianResponse +
                                                              osi + endometrialmorphology + endometrialthickness
                                                            , filtered_data_SR)
##########################################
##########################################


# Fertilization rate
linear_fertilizationrate_age <- linear_regression_gtsummary(fertilizationrate ~ age, filtered_data_SD)

linear_fertilizationrate_bmi <- linear_regression_gtsummary(fertilizationrate ~ bmi, filtered_data_SD)

linear_fertilizationrate_amh <- linear_regression_gtsummary(fertilizationrate ~ amhngml, filtered_data_SD)

linear_fertilizationrate_afc <- linear_regression_gtsummary(fertilizationrate ~ AFC, filtered_data_SD)

linear_fertilizationrate_orpi <- linear_regression_gtsummary(fertilizationrate ~ orpi, filtered_data_SD)

linear_fertilizationrate_fibpres <- linear_regression_gtsummary(fertilizationrate ~ factor(fibroidpresence), filtered_data_SR)

linear_fertilizationrate_fibroid <- linear_regression_gtsummary(fertilizationrate ~ adenomyosispresence, filtered_data_SR)

linear_fertilizationrate_adenom <- linear_regression_gtsummary(fertilizationrate ~ adenomyosispresence, filtered_data_SR)

linear_fertilizationrate_endomorph <- linear_regression_gtsummary(fertilizationrate ~ endometrialmorphology, filtered_data_SR)

linear_fertilizationrate_endom <- linear_regression_gtsummary(fertilizationrate ~ Endometriosispresent, filtered_data_SR)

linear_fertilizationrate_endothick <- linear_regression_gtsummary(fertilizationrate ~ endometrialthickness, filtered_data_SR)

linear_fertilizationrate_ovarianResp <- linear_regression_gtsummary(fertilizationrate ~ ovarianResponse, filtered_data_SR)

linear_fertilizationrate_osi <- linear_regression_gtsummary(fertilizationrate ~ osi, filtered_data_SR)

#adjusted
linear_fertilizationrate_adj <- linear_regression_gtsummary(fertilizationrate ~ age + bmi + amhngml+ 
                                                              AFC + orpi, filtered_data_SD)

linear_fertilizationrate_adj1 <- linear_regression_gtsummary(fertilizationrate ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                               osi + endometrialmorphology + 
                                                               Endometriosispresent + endometrialthickness, filtered_data_SR)


# Blastosis rate
linear_blasto_rate_age <- linear_regression_gtsummary(blasto_rate ~ age, filtered_data_SD)

linear_blasto_rate_bmi <- linear_regression_gtsummary(blasto_rate ~ bmi, filtered_data_SD)

linear_blasto_rate_amh <- linear_regression_gtsummary(blasto_rate ~ amhngml, filtered_data_SD)

linear_blasto_rate_afc <- linear_regression_gtsummary(blasto_rate ~ AFC, filtered_data_SD)

linear_blasto_rate_orpi <- linear_regression_gtsummary(blasto_rate ~ orpi, filtered_data_SD)

linear_blasto_rate_fibpres <- linear_regression_gtsummary(blasto_rate ~ factor(fibroidpresence), filtered_data_SR)

linear_blasto_rate_fibroid <- linear_regression_gtsummary(blasto_rate ~ myomectomy, filtered_data_SR)

linear_blasto_rate_adenom <- linear_regression_gtsummary(blasto_rate ~ adenomyosispresence, filtered_data_SR)

linear_blasto_rate_endomorph <- linear_regression_gtsummary(blasto_rate ~ endometrialmorphology, filtered_data_SR)

linear_blasto_rate_endom <- linear_regression_gtsummary(blasto_rate ~ Endometriosispresent, filtered_data_SR)

linear_blasto_rate_endothick <- linear_regression_gtsummary(blasto_rate ~ endometrialthickness, filtered_data_SR)

linear_blasto_rate_ovarianResp <- linear_regression_gtsummary(blasto_rate ~ ovarianResponse, filtered_data_SR)

linear_blasto_rate_osi <- linear_regression_gtsummary(blasto_rate ~ osi, filtered_data_SR)

# adjusted
linear_blasto_rate_adj <- linear_regression_gtsummary(blasto_rate ~ age + bmi + amhngml+ 
                                                        AFC + orpi, filtered_data_SD)

linear_blasto_rate_adj1 <- linear_regression_gtsummary(blasto_rate ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                         osi + endometrialmorphology + Endometriosispresent + endometrialthickness
                                                       , filtered_data_SR)

# Implantation rate
linear_implantationrate_age <- linear_regression_gtsummary(implantationrate ~ age, filtered_data_SR)

linear_implantationrate_bmi <- linear_regression_gtsummary(implantationrate ~ bmi, filtered_data_SR)

linear_implantationrate_amh <- linear_regression_gtsummary(implantationrate ~ amhngml, filtered_data_SR)

linear_implantationrate_afc <- linear_regression_gtsummary(implantationrate ~ AFC, filtered_data_SR)

linear_implantationrate_orpi <- linear_regression_gtsummary(implantationrate ~ orpi, filtered_data_SR)

linear_implantationrate_fibpres <- linear_regression_gtsummary(implantationrate ~ factor(fibroidpresence), filtered_data_SR)

linear_implantationrate_fibroid <- linear_regression_gtsummary(implantationrate ~ myomectomy, filtered_data_SR)

linear_implantationrate_adenom <- linear_regression_gtsummary(implantationrate ~ adenomyosispresence, filtered_data_SR)

linear_implantationrate_endomorph <- linear_regression_gtsummary(implantationrate ~ endometrialmorphology, filtered_data_SR)

linear_implantationrate_endom <- linear_regression_gtsummary(implantationrate ~ Endometriosispresent, filtered_data_SR)

linear_implantationrate_endothick <- linear_regression_gtsummary(implantationrate ~ endometrialthickness, filtered_data_SR)

linear_implantationrate_ovarianResp <- linear_regression_gtsummary(implantationrate ~ ovarianResponse, filtered_data_SR)

linear_implantationrate_osi <- linear_regression_gtsummary(implantationrate ~ osi, filtered_data_SR)

# adjusted
linear_implantationrate_adj <- linear_regression_gtsummary(implantationrate ~ age + bmi + amhngml+ 
                                                             AFC + orpi, filtered_data_SR)

linear_implantationrate_adj1 <- linear_regression_gtsummary(implantationrate ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                              osi + endometrialmorphology + endometrialthickness
                                                            , filtered_data_SR)

# Urine pregnacy rate
linear_urinepregrate_age <- linear_regression_gtsummary(urinepregrate ~ age, filtered_data_SR)

linear_urinepregrate_bmi <- linear_regression_gtsummary(urinepregrate ~ bmi, filtered_data_SR)

linear_urinepregrate_amh <- linear_regression_gtsummary(urinepregrate ~ amhngml, filtered_data_SR)

linear_urinepregrate_afc <- linear_regression_gtsummary(urinepregrate ~ AFC, filtered_data_SR)

linear_urinepregrate_orpi <- linear_regression_gtsummary(urinepregrate ~ orpi, filtered_data_SR)

linear_urinepregrate_fibpres <- linear_regression_gtsummary(urinepregrate ~ factor(fibroidpresence), filtered_data_SR)

linear_urinepregrate_fibroid <- linear_regression_gtsummary(urinepregrate ~ myomectomy, filtered_data_SR)

linear_urinepregrate_adenom <- linear_regression_gtsummary(urinepregrate ~ adenomyosispresence, filtered_data_SR)

linear_urinepregrate_endomorph <- linear_regression_gtsummary(urinepregrate ~ endometrialmorphology, filtered_data_SR)

linear_urinepregrate_endom <- linear_regression_gtsummary(urinepregrate ~ Endometriosispresent, filtered_data_SR)

linear_urinepregrate_endothick <- linear_regression_gtsummary(urinepregrate ~ endometrialthickness, filtered_data_SR)

linear_urinepregrate_ovarianResp <- linear_regression_gtsummary(urinepregrate ~ ovarianResponse, filtered_data_SR)

linear_urinepregrate_osi <- linear_regression_gtsummary(urinepregrate ~ osi, filtered_data_SR)

#adjusted

linear_urinepregrate_adj <- linear_regression_gtsummary(urinepregrate ~ age + bmi + amhngml+ 
                                                          AFC + orpi, filtered_data_SR)

linear_urinepregrate_adj1 <- linear_regression_gtsummary(urinepregrate ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                           osi + endometrialmorphology + Endometriosispresent + endometrialthickness
                                                         , filtered_data_SR)


# Clinical pregnancy rate
linear_clinicalpregrate_age <- linear_regression_gtsummary(clinicalpregrate ~ age, filtered_data_SR)

linear_clinicalpregrate_bmi <- linear_regression_gtsummary(clinicalpregrate ~ bmi, filtered_data_SR)

linear_clinicalpregrate_amh <- linear_regression_gtsummary(clinicalpregrate ~ amhngml, filtered_data_SR)

linear_clinicalpregrate_afc <- linear_regression_gtsummary(clinicalpregrate ~ AFC, filtered_data_SR)

linear_clinicalpregrate_orpi <- linear_regression_gtsummary(clinicalpregrate ~ orpi, filtered_data_SR)

linear_clinicalpregrate_fibpres <- linear_regression_gtsummary(clinicalpregrate ~ factor(fibroidpresence), filtered_data_SR)

linear_clinicalpregrate_fibroid <- linear_regression_gtsummary(clinicalpregrate ~ myomectomy, filtered_data_SR)

linear_clinicalpregrate_adenom <- linear_regression_gtsummary(clinicalpregrate ~ adenomyosispresence, filtered_data_SR)

linear_clinicalpregrate_endomorph <- linear_regression_gtsummary(clinicalpregrate ~ endometrialmorphology, filtered_data_SR)

linear_clinicalpregrate_endom <- linear_regression_gtsummary(clinicalpregrate ~ Endometriosispresent, filtered_data_SR)

linear_clinicalpregrate_endothick <- linear_regression_gtsummary(clinicalpregrate ~ endometrialthickness, filtered_data_SR)

linear_clinicalpregrate_ovarianResp <- linear_regression_gtsummary(clinicalpregrate ~ ovarianResponse, filtered_data_SR)

linear_clinicalpregrate_osi <- linear_regression_gtsummary(clinicalpregrate ~ osi, filtered_data_SR)

#adjusted

linear_clinicalpregrate_adj <- linear_regression_gtsummary(clinicalpregrate ~ age + bmi + amhngml+ 
                                                             AFC + orpi, filtered_data_SR)

linear_clinicalpregrate_adj1 <- linear_regression_gtsummary(clinicalpregrate ~ fibroidpresence + myomectomy + ovarianResponse +
                                                              osi + endometrialmorphology + endometrialthickness
                                                            , filtered_data_SR)
dams_data$endometrialthickness

# cycle cancellation
# Age
logistic_cycleCancel_age <- logistic_regression_gtsummary(cycleCancel ~ age, dams_data)
# BMI
logistic_cycleCancel_bmi <- logistic_regression_gtsummary(cycleCancel ~ bmi, dams_data)
# AMH
logistic_cycleCancel_amhngml <- logistic_regression_gtsummary(cycleCancel ~ amhngml, dams_data)
# AFC
logistic_cycleCancel_totalafc <- logistic_regression_gtsummary(cycleCancel ~ AFC, dams_data)
# ORPI
logistic_cycleCancel_orpi <- logistic_regression_gtsummary(cycleCancel ~ orpi, dams_data)
# Ovarian Response
logistic_cycleCancel_ovarianResponse <- logistic_regression_gtsummary(cycleCancel ~ factor(ovarianResponse), dams_data)
# OSI
logistic_cycleCancel_osi <- logistic_regression_gtsummary(cycleCancel ~ osi, dams_data)
# Fibroids
logistic_cycleCancel_fibroid <- logistic_regression_gtsummary(cycleCancel ~ factor(myomectomy), dams_data)
# Presence of Fibroids
logistic_cycleCancel_fibroidpresence <- logistic_regression_gtsummary(cycleCancel ~ factor(fibroidpresence), dams_data)
# Presence of Adenomyosis
logistic_cycleCancel_adenomyosispresence <- logistic_regression_gtsummary(cycleCancel ~ factor(adenomyosispresence), dams_data)
# Presence of Endometriosis
logistic_cycleCancel_endometriosispresent <- logistic_regression_gtsummary(cycleCancel ~ factor(endometriosispresent), dams_data)
# Endometrial Thickness
logistic_cycleCancel_endometrialthickness <- logistic_regression_gtsummary(cycleCancel ~ endometrialthickness, dams_data)
# Endometrial Morphology
logistic_cycleCancel_endometrialmorphology <- logistic_regression_gtsummary(cycleCancel ~ endometrialmorphology, dams_data)

#adjusted
logistic_cycleCancel_adj <- logistic_regression_gtsummary(cycleCancel ~ age + bmi + amhngml+ 
                                                            AFC + orpi, filtered_data_SR)

logistic_cycleCancel_adj1 <- logistic_regression_gtsummary(cycleCancel ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                             osi + endometrialmorphology + Endometriosispresent + 
                                                             endometrialthickness, filtered_data_SR)

# FOI
# Age
linear_foi_age <- linear_regression_gtsummary(foi ~ age, dams_data)
# BMI
linear_foi_bmi <- linear_regression_gtsummary(foi ~ bmi, dams_data)
# AMH
linear_foi_amhngml <- linear_regression_gtsummary(foi ~ amhngml, dams_data)
# AFC
linear_foi_totalafc <- linear_regression_gtsummary(foi ~ AFC, dams_data)
# ORPI
linear_foi_orpi <- linear_regression_gtsummary(foi ~ orpi, dams_data)
# Ovarian Response
linear_foi_ovarianResponse <- linear_regression_gtsummary(foi ~ factor(ovarianResponse), dams_data)
# OSI
linear_foi_osi <- linear_regression_gtsummary(foi ~ osi, dams_data)
# Fibroids
linear_foi_fibroid <- linear_regression_gtsummary(foi ~ factor(myomectomy), dams_data)
# Presence of Fibroids
linear_foi_fibroidpresence <- linear_regression_gtsummary(foi ~ factor(fibroidpresence), dams_data)
# Presence of Endometriosis
linear_foi_endometriosispresent <- linear_regression_gtsummary(foi ~ factor(endometriosispresent), dams_data)
# Endometrial Thickness
linear_foi_endometrialthickness <- linear_regression_gtsummary(foi ~ endometrialthickness, dams_data)
# Endometrial Morphology
linear_foi_endometrialmorphology <- linear_regression_gtsummary(foi ~ endometrialmorphology, dams_data)

#adjusted

linear_foi_adj <- linear_regression_gtsummary(foi ~ age + bmi + amhngml+ 
                                                AFC + orpi, filtered_data_SR)

linear_foi_adj1 <- linear_regression_gtsummary(foi ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                 osi + endometrialmorphology + 
                                                 Endometriosispresent + endometrialthickness, filtered_data_SR)

# FORT
# Age
linear_fort_age <- linear_regression_gtsummary(fort ~ age, dams_data)
# BMI
linear_fort_bmi <- linear_regression_gtsummary(fort ~ bmi, dams_data)
# AMH
linear_fort_amhngml <- linear_regression_gtsummary(fort ~ amhngml, dams_data)
# AFC
linear_fort_totalafc <- linear_regression_gtsummary(fort ~ AFC, dams_data)
# ORPI
linear_fort_orpi <- linear_regression_gtsummary(fort ~ orpi, dams_data)
# Ovarian Response
linear_fort_ovarianResponse <- linear_regression_gtsummary(fort ~ factor(ovarianResponse), dams_data)
# OSI
linear_fort_osi <- linear_regression_gtsummary(fort ~ osi, dams_data)
# Fibroids
linear_fort_fibroid <- linear_regression_gtsummary(fort ~ factor(myomectomy), dams_data)
# Presence of Fibroids
linear_fort_fibroidpresence <- linear_regression_gtsummary(fort ~ factor(fibroidpresence), dams_data)
# Presence of Endometriosis
linear_fort_endometriosispresent <- linear_regression_gtsummary(fort ~ factor(endometriosispresent), dams_data)
# Endometrial Thickness
linear_fort_endometrialthickness <- linear_regression_gtsummary(fort ~ endometrialthickness, dams_data)
# Endometrial Morphology
linear_fort_endometrialmorphology <- linear_regression_gtsummary(fort ~ endometrialmorphology, dams_data)

#adjusted

linear_fort_adj <- linear_regression_gtsummary(fort ~ age + bmi + amhngml+ 
                                                 AFC + orpi, filtered_data_SR)

linear_fort_adj1 <- linear_regression_gtsummary(fort ~ factor(fibroidpresence) + myomectomy + ovarianResponse +
                                                  osi + endometrialmorphology + 
                                                  Endometriosispresent + endometrialthickness, filtered_data_SR)



# # Age
# logistic_cycleCancel_age <- logistic_regression_gtsummary(cycleCancel ~ age, dams_data)
# linear_foi_age <- linear_regression_gtsummary(foi ~ age, dams_data)
# linear_fort_age <- linear_regression_gtsummary(fort ~ age, dams_data)
# linear_fertilizationrate_age <- linear_regression_gtsummary(fertilizationrate ~ age, dams_data)
# linear_blasto_rate_age <- linear_regression_gtsummary(blasto_rate ~ age, dams_data)
# linear_implantationrate_age <- linear_regression_gtsummary(implantationrate ~ age, dams_data)
# linear_urinepregrate_age <- linear_regression_gtsummary(urinepregrate ~ age, dams_data)
# linear_clinicalpregrate_age <- linear_regression_gtsummary(clinicalpregrate ~ age, dams_data)
# 
# # BMI
# logistic_cycleCancel_bmi <- logistic_regression_gtsummary(cycleCancel ~ bmi, dams_data)
# linear_foi_bmi <- linear_regression_gtsummary(foi ~ bmi, dams_data)
# linear_fort_bmi <- linear_regression_gtsummary(fort ~ bmi, dams_data)
# linear_fertilizationrate_bmi <- linear_regression_gtsummary(fertilizationrate ~ bmi, dams_data)
# linear_blasto_rate_bmi <- linear_regression_gtsummary(blasto_rate ~ bmi, dams_data)
# linear_implantationrate_bmi <- linear_regression_gtsummary(implantationrate ~ bmi, dams_data)
# linear_urinepregrate_bmi <- linear_regression_gtsummary(urinepregrate ~ bmi, dams_data)
# linear_clinicalpregrate_bmi <- linear_regression_gtsummary(clinicalpregrate ~ bmi, dams_data)
# 
# # AMH
# logistic_cycleCancel_amhngml <- logistic_regression_gtsummary(cycleCancel ~ amhngml, dams_data)
# linear_foi_amhngml <- linear_regression_gtsummary(foi ~ amhngml, dams_data)
# linear_fort_amhngml <- linear_regression_gtsummary(fort ~ amhngml, dams_data)
# linear_fertilizationrate_amhngml <- linear_regression_gtsummary(fertilizationrate ~ amhngml, dams_data)
# linear_blasto_rate_amhngml <- linear_regression_gtsummary(blasto_rate ~ amhngml, dams_data)
# linear_implantationrate_amhngml <- linear_regression_gtsummary(implantationrate ~ amhngml, dams_data)
# linear_urinepregrate_amhngml <- linear_regression_gtsummary(urinepregrate ~ amhngml, dams_data)
# linear_clinicalpregrate_amhngml <- linear_regression_gtsummary(clinicalpregrate ~ amhngml, dams_data)
# 
# # AFC
# logistic_cycleCancel_totalafc <- logistic_regression_gtsummary(cycleCancel ~ totalafc, dams_data)
# linear_foi_totalafc <- linear_regression_gtsummary(foi ~ totalafc, dams_data)
# linear_fort_totalafc <- linear_regression_gtsummary(fort ~ totalafc, dams_data)
# linear_fertilizationrate_totalafc <- linear_regression_gtsummary(fertilizationrate ~ totalafc, dams_data)
# linear_blasto_rate_totalafc <- linear_regression_gtsummary(blasto_rate ~ totalafc, dams_data)
# linear_implantationrate_totalafc <- linear_regression_gtsummary(implantationrate ~ totalafc, dams_data)
# linear_urinepregrate_totalafc <- linear_regression_gtsummary(urinepregrate ~ totalafc, dams_data)
# linear_clinicalpregrate_totalafc <- linear_regression_gtsummary(clinicalpregrate ~ totalafc, dams_data)
# 
# # ORPI
# logistic_cycleCancel_orpi <- logistic_regression_gtsummary(cycleCancel ~ orpi, dams_data)
# linear_foi_orpi <- linear_regression_gtsummary(foi ~ orpi, dams_data)
# linear_fort_orpi <- linear_regression_gtsummary(fort ~ orpi, dams_data)
# linear_fertilizationrate_orpi <- linear_regression_gtsummary(fertilizationrate ~ orpi, dams_data)
# linear_blasto_rate_orpi <- linear_regression_gtsummary(blasto_rate ~ orpi, dams_data)
# linear_implantationrate_orpi <- linear_regression_gtsummary(implantationrate ~ orpi, dams_data)
# linear_urinepregrate_orpi <- linear_regression_gtsummary(urinepregrate ~ orpi, dams_data)
# linear_clinicalpregrate_orpi <- linear_regression_gtsummary(clinicalpregrate ~ orpi, dams_data)
# 
# # Ovarian Response
# logistic_cycleCancel_ovarianResponse <- logistic_regression_gtsummary(cycleCancel ~ factor(ovarianResponse), dams_data)
# linear_foi_ovarianResponse <- linear_regression_gtsummary(foi ~ factor(ovarianResponse), dams_data)
# linear_fort_ovarianResponse <- linear_regression_gtsummary(fort ~ factor(ovarianResponse), dams_data)
# linear_fertilizationrate_ovarianResponse <- linear_regression_gtsummary(fertilizationrate ~ factor(ovarianResponse), dams_data)
# linear_blasto_rate_ovarianResponse <- linear_regression_gtsummary(blasto_rate ~ factor(ovarianResponse), dams_data)
# linear_implantationrate_ovarianResponse <- linear_regression_gtsummary(implantationrate ~ factor(ovarianResponse), dams_data)
# linear_urinepregrate_ovarianResponse <- linear_regression_gtsummary(urinepregrate ~ factor(ovarianResponse), dams_data)
# linear_clinicalpregrate_ovarianResponse <- linear_regression_gtsummary(clinicalpregrate ~ factor(ovarianResponse), dams_data)
# 
# # OSI
# logistic_cycleCancel_osi <- logistic_regression_gtsummary(cycleCancel ~ osi, dams_data)
# linear_foi_osi <- linear_regression_gtsummary(foi ~ osi, dams_data)
# linear_fort_osi <- linear_regression_gtsummary(fort ~ osi, dams_data)
# linear_fertilizationrate_osi <- linear_regression_gtsummary(fertilizationrate ~ osi, dams_data)
# linear_blasto_rate_osi <- linear_regression_gtsummary(blasto_rate ~ osi, dams_data)
# linear_implantationrate_osi <- linear_regression_gtsummary(implantationrate ~ osi, dams_data)
# linear_urinepregrate_osi <- linear_regression_gtsummary(urinepregrate ~ osi, dams_data)
# linear_clinicalpregrate_osi <- linear_regression_gtsummary(clinicalpregrate ~ osi, dams_data)
# 
# # Presence of fibroids
# logistic_cycleCancel_fibroidpresence <- logistic_regression_gtsummary(cycleCancel ~ factor(fibroidpresence), dams_data)
# linear_foi_fibroidpresence <- linear_regression_gtsummary(foi ~ factor(fibroidpresence), dams_data)
# linear_fort_fibroidpresence <- linear_regression_gtsummary(fort ~ factor(fibroidpresence), dams_data)
# linear_fertilizationrate_fibroidpresence <- linear_regression_gtsummary(fertilizationrate ~ factor(fibroidpresence), dams_data)
# linear_blasto_rate_fibroidpresence <- linear_regression_gtsummary(blasto_rate ~ factor(fibroidpresence), dams_data)
# linear_implantationrate_fibroidpresence <- linear_regression_gtsummary(implantationrate ~ factor(fibroidpresence), dams_data)
# linear_urinepregrate_fibroidpresence <- linear_regression_gtsummary(urinepregrate ~ factor(fibroidpresence), dams_data)
# linear_clinicalpregrate_fibroidpresence <- linear_regression_gtsummary(clinicalpregrate ~ factor(fibroidpresence), dams_data)
# 
# # Presence of adenomyosis
# logistic_cycleCancel_adenomyosispresence <- logistic_regression_gtsummary(cycleCancel ~ factor(adenomyosispresence), dams_data)
# linear_foi_adenomyosispresence <- linear_regression_gtsummary(foi ~ factor(adenomyosispresence), dams_data)
# linear_fort_adenomyosispresence <- linear_regression_gtsummary(fort ~ factor(adenomyosispresence), dams_data)
# linear_fertilizationrate_adenomyosispresence <- linear_regression_gtsummary(fertilizationrate ~ factor(adenomyosispresence), dams_data)
# linear_blasto_rate_adenomyosispresence <- linear_regression_gtsummary(blasto_rate ~ factor(adenomyosispresence), dams_data)
# linear_implantationrate_adenomyosispresence <- linear_regression_gtsummary(implantationrate ~ factor(adenomyosispresence), dams_data)
# linear_urinepregrate_adenomyosispresence <- linear_regression_gtsummary(urinepregrate ~ factor(adenomyosispresence), dams_data)
# linear_clinicalpregrate_adenomyosispresence <- linear_regression_gtsummary(clinicalpregrate ~ factor(adenomyosispresence), dams_data)
# 
# # Presence of endometriosis
# logistic_cycleCancel_endometriosispresent <- logistic_regression_gtsummary(cycleCancel ~ factor(endometriosispresent), dams_data)
# linear_foi_endometriosispresent <- linear_regression_gtsummary(foi ~ factor(endometriosispresent), dams_data)
# linear_fort_endometriosispresent <- linear_regression_gtsummary(fort ~ factor(endometriosispresent), dams_data)
# linear_fertilizationrate_endometriosispresent <- linear_regression_gtsummary(fertilizationrate ~ factor(endometriosispresent), dams_data)
# linear_blasto_rate_endometriosispresent <- linear_regression_gtsummary(blasto_rate ~ factor(endometriosispresent), dams_data)
# linear_implantationrate_endometriosispresent <- linear_regression_gtsummary(implantationrate ~ factor(endometriosispresent), dams_data)
# linear_urinepregrate_endometriosispresent <- linear_regression_gtsummary(urinepregrate ~ factor(endometriosispresent), dams_data)
# linear_clinicalpregrate_endometriosispresent <- linear_regression_gtsummary(clinicalpregrate ~ factor(endometriosispresent), dams_data)
# 
# # Endometrial thickness
# logistic_cycleCancel_endometrialthickness <- logistic_regression_gtsummary(cycleCancel ~ endometrialthickness, dams_data)
# linear_foi_endometrialthickness <- linear_regression_gtsummary(foi ~ endometrialthickness, dams_data)
# linear_fort_endometrialthickness <- linear_regression_gtsummary(fort ~ endometrialthickness, dams_data)
# linear_fertilizationrate_endometrialthickness <- linear_regression_gtsummary(fertilizationrate ~ endometrialthickness, dams_data)
# linear_blasto_rate_endometrialthickness <- linear_regression_gtsummary(blasto_rate ~ endometrialthickness, dams_data)
# linear_implantationrate_endometrialthickness <- linear_regression_gtsummary(implantationrate ~ endometrialthickness, dams_data)
# linear_urinepregrate_endometrialthickness <- linear_regression_gtsummary(urinepregrate ~ endometrialthickness, dams_data)
# linear_clinicalpregrate_endometrialthickness <- linear_regression_gtsummary(clinicalpregrate ~ endometrialthickness, dams_data)
# 
# 
# # Endometrial morphology
# logistic_cycleCancel_endometrialthickness <- logistic_regression_gtsummary(cycleCancel ~ endometrialmorphology, dams_data)
# linear_foi_endometrialthickness <- linear_regression_gtsummary(foi ~ endometrialmorphology, dams_data)
# linear_fort_endometrialthickness <- linear_regression_gtsummary(fort ~ endometrialmorphology, dams_data)
# linear_fertilizationrate_endometrialthickness <- linear_regression_gtsummary(fertilizationrate ~ endometrialmorphology, dams_data)
# linear_blasto_rate_endometrialthickness <- linear_regression_gtsummary(blasto_rate ~ endometrialmorphology, dams_data)
# linear_implantationrate_endometrialthickness <- linear_regression_gtsummary(implantationrate ~ endometrialmorphology, dams_data)
# linear_urinepregrate_endometrialthickness <- linear_regression_gtsummary(urinepregrate ~ endometrialmorphology, dams_data)
# linear_clinicalpregrate_endometrialthickness <- linear_regression_gtsummary(clinicalpregrate ~ endometrialmorphology, dams_data)
# 

#######################
#cycle cancel
logistic_cycleCancel_age
logistic_cycleCancel_bmi
logistic_cycleCancel_amhngml
logistic_cycleCancel_totalafc
logistic_cycleCancel_orpi
logistic_cycleCancel_ovarianResponse
logistic_cycleCancel_osi
logistic_cycleCancel_fibroid
logistic_cycleCancel_fibroidpresence
logistic_cycleCancel_adenomyosispresence
logistic_cycleCancel_endometriosispresent
logistic_cycleCancel_endometrialthickness
logistic_cycleCancel_endometrialmorphology

logistic_cycleCancel_adj
logistic_cycleCancel_adj1

# foi
linear_foi_age
linear_foi_bmi
linear_foi_amhngml
linear_foi_totalafc
linear_foi_orpi
linear_foi_ovarianResponse
linear_foi_osi
logistic_foi_fibroid
linear_foi_fibroidpresence
linear_foi_endometriosispresent
linear_foi_endometrialthickness
linear_foi_endometrialmorphology

logistic_foi_adj
logistic_foi_adj1

# fort
linear_fort_age
linear_fort_bmi
linear_fort_amhngml
linear_fort_totalafc
linear_fort_orpi
linear_fort_ovarianResponse
linear_fort_osi
logistic_fort_fibroid
linear_fort_fibroidpresence
linear_fort_endometriosispresent
linear_fort_endometrialthickness
linear_fort_endometrialmorphology

logistic_fort_adj
logistic_fort_adj1

# fertilization
linear_fertilizationrate_age
linear_fertilizationrate_bmi
linear_fertilizationrate_amh
linear_fertilizationrate_afc
linear_fertilizationrate_orpi
linear_fertilizationrate_fibpres
linear_fertilizationrate_fibroid
linear_fertilizationrate_adenom
linear_fertilizationrate_endomorph
linear_fertilizationrate_endom
linear_fertilizationrate_endothick
linear_fertilizationrate_ovarianResp
linear_fertilizationrate_osi

linear_fertilizationrate_adj
linear_fertilizationrate_adj1

#blastosis
linear_blasto_rate_age
linear_blasto_rate_bmi
linear_blasto_rate_amh
linear_blasto_rate_afc
linear_blasto_rate_orpi
linear_blasto_rate_fibpres
linear_blasto_rate_fibroid
linear_blasto_rate_adenom
linear_blasto_rate_endomorph
linear_blasto_rate_endom
linear_blasto_rate_endothick
linear_blasto_rate_ovarianResp
linear_blasto_rate_osi

linear_blasto_rate_adj
linear_blasto_rate_adj1

# implantation
linear_implantationrate_age
linear_implantationrate_bmi
linear_implantationrate_amh
linear_implantationrate_afc
linear_implantationrate_orpi
linear_implantationrate_fibpres
linear_implantationrate_fibroid
linear_implantationrate_adenom
linear_implantationrate_endomorph
linear_implantationrate_endom
linear_implantationrate_endothick
linear_implantationrate_ovarianResp
linear_implantationrate_osi

linear_implantationrate_adj
linear_implantationrate_adj1

# urine pregnancy
linear_urinepregrate_age
linear_urinepregrate_bmi
linear_urinepregrate_amh
linear_urinepregrate_afc
linear_urinepregrate_orpi
linear_urinepregrate_fibpres
linear_urinepregrate_fibroid
linear_urinepregrate_adenom
linear_urinepregrate_endomorph
linear_urinepregrate_endom
linear_urinepregrate_endothick
linear_urinepregrate_ovarianResp
linear_urinepregrate_osi

linear_urinepregrate_adj
linear_urinepregrate_adj1

# clinical pregnancy
linear_clinicalpregrate_age
linear_clinicalpregrate_bmi
linear_clinicalpregrate_amh
linear_clinicalpregrate_afc
linear_clinicalpregrate_orpi
linear_clinicalpregrate_fibpres
linear_clinicalpregrate_fibroid
linear_clinicalpregrate_adenom
linear_clinicalpregrate_endomorph
linear_clinicalpregrate_endom
linear_clinicalpregrate_endothick
linear_clinicalpregrate_ovarianResp
linear_clinicalpregrate_osi

linear_clinicalpregrate_adj
linear_clinicalpregrate_adj1



########################




















# Define functions for logistic and linear regressions with gtsummary
logistic_regression_gtsummary <- function(formula, data) {
  model <- glm(formula, data = data, family = binomial)
  tbl <- tbl_regression(model, exponentiate = TRUE)
  return(tbl)
}

linear_regression_gtsummary <- function(formula, data) {
  model <- lm(formula, data = data)
  tbl <- tbl_regression(model)
  return(tbl)
}

# Create and combine regression models and summary tables for each group

# Age
logistic_diminishedOvR_afc_age <- logistic_regression_gtsummary(diminishedOvR_afc_new ~ age, dams_data)
logistic_diminishedOvR_amh_age <- logistic_regression_gtsummary(diminishedOvR_amh ~ age, dams_data)
linear_orpi_age <- linear_regression_gtsummary(orpi ~ age, dams_data)

tbl_combined_age <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_age,
    logistic_diminishedOvR_amh_age,
    linear_orpi_age
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**",
                  "**ORPI**")
)

# BMI
logistic_diminishedOvR_afc_bmi <- logistic_regression_gtsummary(diminishedOvR_afc ~ bmi, dams_data)
logistic_diminishedOvR_amh_bmi <- logistic_regression_gtsummary(diminishedOvR_amh ~ bmi, dams_data)
linear_orpi_bmi <- linear_regression_gtsummary(osi ~ bmi, dams_data)

tbl_combined_bmi <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_bmi,
    logistic_diminishedOvR_amh_bmi,
    linear_orpi_bmi
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**",
                  "**OSI**")
)

# BMI Categories
logistic_diminishedOvR_afc_bmi_category <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(bmi_category), dams_data)
logistic_diminishedOvR_amh_bmi_category <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(bmi_category), dams_data)
linear_orpi_bmi_category <- linear_regression_gtsummary(orpi ~ factor(bmi_category), dams_data)

tbl_combined_bmi_category <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_bmi_category,
    logistic_diminishedOvR_amh_bmi_category,
    linear_orpi_bmi_category
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**",
                  "**ORPI**")
)

# Age + BMI
logistic_diminishedOvR_afc_age_bmi <- logistic_regression_gtsummary(diminishedOvR_afc ~ age + bmi, dams_data)
logistic_diminishedOvR_amh_age_bmi <- logistic_regression_gtsummary(diminishedOvR_amh ~ age + bmi, dams_data)
logistic_ovarianResponse_age_bmi <- logistic_regression_gtsummary(ovarianResponse ~ age + bmi, dams_data)
linear_osi_age_bmi <- linear_regression_gtsummary(osi ~ age + bmi, dams_data)

tbl_combined_age_bmi <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_age_bmi,
    logistic_diminishedOvR_amh_age_bmi,
    linear_orpi_age_bmi
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**",
                  "**ORPI**")
)

# Fibroid Presence
logistic_diminishedOvR_afc_fibroidpresence <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(fibroidpresence), dams_data)
logistic_diminishedOvR_amh_fibroidpresence <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(fibroidpresence), dams_data)
linear_orpi_fibroidpresence <- linear_regression_gtsummary(orpi ~ factor(fibroidpresence), dams_data)

tbl_combined_fibroidpresence <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_fibroidpresence,
    logistic_diminishedOvR_amh_fibroidpresence,
    linear_orpi_fibroidpresence
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**",
                  "**orpi**")
)

# Fibroid Type
logistic_diminishedOvR_afc_fibroid <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(fibroid), dams_data)
logistic_diminishedOvR_amh_fibroid <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(fibroid), dams_data)
linear_orpi_fibroid <- linear_regression_gtsummary(orpi ~ factor(fibroid), dams_data)

tbl_combined_fibroid <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_fibroid,
    logistic_diminishedOvR_amh_fibroid,
    linear_orpi_fibroid
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)


# Ovarian Cyst Presence
logistic_diminishedOvR_afc_ovariancystpresent <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(ovariancystpresent), dams_data)
logistic_diminishedOvR_amh_ovariancystpresent <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(ovariancystpresent), dams_data)
linear_orpi_ovariancystpresent <- linear_regression_gtsummary(orpi ~ factor(ovariancystpresent), dams_data)

tbl_combined_ovariancystpresent <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_ovariancystpresent,
    logistic_diminishedOvR_amh_ovariancystpresent,
    linear_orpi_ovariancystpresent
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)

# Previous History of Ovarian Cyst
logistic_diminishedOvR_afc_previoushistoryofovariancyst <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(previoushistoryofovariancyst), dams_data)
logistic_diminishedOvR_amh_previoushistoryofovariancyst <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryofovariancyst), dams_data)
linear_orpi_previoushistoryofovariancyst <- linear_regression_gtsummary(orpi ~ factor(previoushistoryofovariancyst), dams_data)

tbl_combined_previoushistoryofovariancyst <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_previoushistoryofovariancyst,
    logistic_diminishedOvR_amh_previoushistoryofovariancyst,
    linear_orpi_previoushistoryofovariancyst
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)

# Endometriosis Presence
logistic_diminishedOvR_afc_endometriosispresent <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(endometriosispresent), dams_data)
logistic_diminishedOvR_amh_endometriosispresent <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(endometriosispresent), dams_data)
linear_orpi_endometriosispresent <- linear_regression_gtsummary(orpi ~ factor(endometriosispresent), dams_data)

tbl_combined_endometriosispresent <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_endometriosispresent,
    logistic_diminishedOvR_amh_endometriosispresent,
    linear_orpi_endometriosispresent
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)

# Endometriosis Grade
logistic_diminishedOvR_afc_endometriosisgrade <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(endometriosisgrade), dams_data)
logistic_diminishedOvR_amh_endometriosisgrade <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(endometriosisgrade), dams_data)
linear_orpi_endometriosisgrade <- linear_regression_gtsummary(orpi ~ factor(endometriosisgrade), dams_data)

tbl_combined_endometriosisgrade <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_endometriosisgrade,
    logistic_diminishedOvR_amh_endometriosisgrade,
    linear_orpi_endometriosisgrade
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)

# Previous History of Salpingectomy
logistic_diminishedOvR_afc_previoushistoryofsalpingectom <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(previoushistoryofsalpingectom), dams_data)
logistic_diminishedOvR_amh_previoushistoryofsalpingectom <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryofsalpingectom), dams_data)
linear_orpi_previoushistoryofsalpingectom <- linear_regression_gtsummary(orpi ~ factor(previoushistoryofsalpingectom), dams_data)

tbl_combined_previoushistoryofsalpingectom <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_previoushistoryofsalpingectom,
    logistic_diminishedOvR_amh_previoushistoryofsalpingectom,
    linear_orpi_previoushistoryofsalpingectom
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**Ovarian Response**", 
                  "**orpi**")
)

# Previous History of Appendicectomy
logistic_diminishedOvR_afc_previoushistoryofappendicecto <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(previoushistoryofappendicecto), dams_data)
logistic_diminishedOvR_amh_previoushistoryofappendicecto <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryofappendicecto), dams_data)
logistic_ovarianResponse_previoushistoryofappendicecto <- logistic_regression_gtsummary(ovarianResponse ~ factor(previoushistoryofappendicecto), dams_data)
linear_orpi_previoushistoryofappendicecto <- linear_regression_gtsummary(orpi ~ factor(previoushistoryofappendicecto), dams_data)

tbl_combined_previoushistoryofappendicecto <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_previoushistoryofappendicecto,
    logistic_diminishedOvR_amh_previoushistoryofappendicecto,
    linear_orpi_previoushistoryofappendicecto
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**",
                  "**orpi**"
  )
)


# Previous History of Laparoscopy
logistic_diminishedOvR_afc_previoushistoryoflaparoscopy <- logistic_regression_gtsummary(diminishedOvR_afc ~ factor(previoushistoryoflaparoscopy), dams_data)
logistic_diminishedOvR_amh_previoushistoryoflaparoscopy <- logistic_regression_gtsummary(diminishedOvR_amh ~ factor(previoushistoryoflaparoscopy), dams_data)
linear_orpi_previoushistoryoflaparoscopy <- linear_regression_gtsummary(orpi ~ factor(previoushistoryoflaparoscopy), dams_data)

tbl_combined_previoushistoryoflaparoscopy <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_previoushistoryoflaparoscopy,
    logistic_diminishedOvR_amh_previoushistoryoflaparoscopy,
    linear_orpi_previoushistoryoflaparoscopy
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)

# Endometrial thickness
logistic_diminishedOvR_afc_endometrialthickness <- logistic_regression_gtsummary(diminishedOvR_afc ~ endometrialthickness, dams_data)
logistic_diminishedOvR_amh_endometrialthickness <- logistic_regression_gtsummary(diminishedOvR_amh ~ endometrialthickness, dams_data)
linear_orpi_endometrialthickness <- linear_regression_gtsummary(orpi ~ endometrialthickness, dams_data)

tbl_combined_endometrialthickness <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_endometrialthickness,
    logistic_diminishedOvR_amh_endometrialthickness,
    linear_orpi_endometrialthickness
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**orpi**")
)

# Endometrial morphology
logistic_diminishedOvR_afc_endometrialmorphology <- logistic_regression_gtsummary(diminishedOvR_afc ~ endometrialmorphology, dams_data)
logistic_diminishedOvR_amh_endometrialmorphology <- logistic_regression_gtsummary(diminishedOvR_amh ~ endometrialmorphology, dams_data)
linear_orpi_endometrialmorphology <- linear_regression_gtsummary(orpi ~ endometrialmorphology, dams_data)

tbl_combined_endometrialmorphology <- tbl_merge(
  tbls = list(
    logistic_diminishedOvR_afc_endometrialmorphology,
    logistic_diminishedOvR_amh_endometrialmorphology,
    linear_orpi_endometrialmorphology
  ),
  tab_spanner = c("**Diminished Ovarian Reserve (AFC)**", 
                  "**Diminished Ovarian Reserve (AMH)**", 
                  "**Ovarian Response**", 
                  "**orpi**")
)



## OBJ 3

# Define the list of independent variables
variables <- c("age", "bmi", "amhngml", "AFC", "orpi",
               "i.ovarianResponse", "osi", "i.fibroidpresence",
               "i.adenomyosispresence", "i.endometriosispresent",
               "i.endometrialmorphology", "endometrialthickness")

# Function to run logistic regression and create summary table for cycleCancel
run_logistic_and_create_table <- function(variable, outcome) {
  formula <- as.formula(paste(outcome, "~", variable))
  model <- try(glm(formula, data = dams_data, family = binomial), silent = TRUE)
  if (inherits(model, "try-error")) {
    return(NULL)
  }
  
  tbl <- tbl_regression(model, exponentiate = TRUE) %>%
    as_gt() %>%
    gt::tab_header(title = paste("Logistic Regression:", outcome, "~", variable))
  
  return(tbl)
}

# Function to run linear regression and create summary table for other outcomes
run_linear_and_create_table <- function(variable, outcome) {
  formula <- as.formula(paste(outcome, "~", variable))
  model <- try(lm(formula, data = dams_data), silent = TRUE)
  if (inherits(model, "try-error")) {
    return(NULL)
  }
  
  tbl <- tbl_regression(model, exponentiate = FALSE) %>%
    as_gt() %>%
    gt::tab_header(title = paste("Linear Regression:", outcome, "~", variable))
  
  return(tbl)
}

# Outcomes to be analyzed
outcomes <- list(
  cycleCancel = "logistic",
  foi = "linear",
  fort = "linear",
  fertilizationrate = "linear",
  blasto_rate = "linear",
  implantationrate = "linear",
  urinepregrate = "linear",
  clinicalpregrate = "linear"
)

# Loop through each outcome and variable, creating and displaying tables
for (outcome in names(outcomes)) {
  regression_type <- outcomes[[outcome]]
  
  for (variable in variables) {
    if (regression_type == "logistic") {
      tbl <- run_logistic_and_create_table(variable, outcome)
    } else {
      tbl <- run_linear_and_create_table(variable, outcome)
    }
    
    if (!is.null(tbl)) {
      print(tbl)
    }
  }
}







# Print the combined tables
print(tbl_combined_age)
print(tbl_combined_bmi)
print(tbl_combined_bmi_category)
print(tbl_combined_age_bmi)
print(tbl_combined_fibroidpresence)
print(tbl_combined_fibroid)
print(logistic_ovarianResponse_adenomyosispresence)
print(tbl_combined_ovariancystpresent)
print(tbl_combined_previoushistoryofovariancyst)
print(tbl_combined_endometriosispresent)
print(tbl_combined_endometriosisgrade)
print(tbl_combined_previoushistoryofsalpingectom)
print(logistic_ovarianResponse_previoushistoryofappendicecto)
print(tbl_combined_previoushistoryoflaparoscopy)
print(tbl_combined_endometrialthickness)
print(tbl_combined_endometrialmorphology)








# Objective 3 - Cycle cancellation and other outcomes


# Filter the data to exclude rows with NA in self_donor
filtered_data_SD <- dams_data %>%
  filter(!is.na(self_donor))

# Filter the data to exclude rows with NA in recipient_self
filtered_data_SR <- dams_data %>%
  filter(!is.na(recipient_self))

# Filter the data to exclude rows with NA in self_recipient
filtered_data_S_R <- dams_data %>%
  filter(!is.na(self_recipient))

# Filter the data to exclude rows with NA in self_donor
filtered_data_S <- dams_data %>%
  filter(!is.na(selfcycle))

# Filter the data to exclude rows with NA in recipient_self
filtered_data_R <- dams_data %>%
  filter(!is.na(recipientcycle))

# Filter the data to exclude rows with NA in self_recipient
filtered_data_D <- dams_data %>%
  filter(!is.na(oocytedonorcycle))


## Fertilization rate

# Define the variables of interest for linear regression models
variables_fertilization <- c("age", "bmi", "amhngml", "AFC", "orpi", "fibroidpresence", "ovarianResponse", "osi", 
                             "endometrialmorphology", "endometrialthickness")

# Create a custom function to perform linear regression and summarize it using gtsummary
linear_regression_gtsummary <- function(formula, data) {
  model <- lm(formula, data = data)
  tbl_regression(model)
}

# Create linear regression models for each variable and the combined model
linear_fertilization_model <- lapply(variables_fertilization, function(var) {
  linear_regression_gtsummary(as.formula(paste("fertilizationrate ~", var)), filtered_data_SD)
})

linear_fertilization_combined <- linear_regression_gtsummary(fertilizationrate ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness, filtered_data_SD)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_fertilization <- lapply(linear_fertilization_model, function(summary) {
  tbl_merge(
    tbls = list(summary, linear_fertilization_combined),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_fertilization <- tbl_stack(combined_summaries_fertilization)

# Display the final summary
print(final_combined_summaries_fertilization)


## Blastocyst Formation Rate
# Define the variables of interest for linear regression models
variables_blastocyst <- c("age", "bmi", "amhngml", "AFC", "orpi", "fibroidpresence", "ovarianResponse", "osi", 
                          "endometrialmorphology", "endometrialthickness")

# Create linear regression models for each variable and the combined model
linear_blastocyst_model <- lapply(variables_blastocyst, function(var) {
  linear_regression_gtsummary(as.formula(paste("blasto_rate ~", var)), filtered_data_SD)
})

linear_blastocyst_combined <- linear_regression_gtsummary(blasto_rate ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness, filtered_data_SD)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_blastocyst <- lapply(linear_blastocyst_model, function(summary) {
  tbl_merge(
    tbls = list(summary, linear_blastocyst_combined),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_blastocyst <- tbl_stack(combined_summaries_blastocyst)

# Display the final summary
print(final_combined_summaries_blastocyst)



## Implantation Rate
# Define the variables of interest for linear regression models
variables_implantation <- c("age", "bmi", "amhngml", "AFC", "orpi", "fibroidpresence", "ovarianResponse", "osi", 
                            "endometrialmorphology", "endometrialthickness")

# Create linear regression models for each variable and the combined model
linear_implantation_model <- lapply(variables_implantation, function(var) {
  linear_regression_gtsummary(as.formula(paste("implantationrate ~", var)), filtered_data_SR)
})

linear_implantation_combined <- linear_regression_gtsummary(implantationrate ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness, filtered_data_SD)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_implantation <- lapply(linear_implantation_model, function(summary) {
  tbl_merge(
    tbls = list(summary, linear_implantation_combined),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_implantation <- tbl_stack(combined_summaries_implantation)

# Display the final summary
print(final_combined_summaries_implantation)


dams_data$urinepregrate
## Urine Pregnancy Rate
# Define the variables of interest for linear regression models
variables_urine_pregnancy <- c("age", "bmi", "amhngml", "AFC", "orpi", "fibroidpresence", "ovarianResponse", "osi", 
                               "endometrialmorphology", "endometrialthickness")

# Create linear regression models for each variable and the combined model
linear_urine_pregnancy_model <- lapply(variables_urine_pregnancy, function(var) {
  linear_regression_gtsummary(as.formula(paste("urinepregrate ~", var)), filtered_data_SR)
})

linear_urine_pregnancy_combined <- linear_regression_gtsummary(urinepregrate ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness, filtered_data_SD)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_urine_pregnancy <- lapply(linear_urine_pregnancy_model, function(summary) {
  tbl_merge(
    tbls = list(summary, linear_urine_pregnancy_combined),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_urine_pregnancy <- tbl_stack(combined_summaries_urine_pregnancy)

# Display the final summary
print(final_combined_summaries_urine_pregnancy)

dams_data$clinicalpregrate
## Clinical Pregnancy Rate
# Define the variables of interest for linear regression models
variables_clinical_pregnancy <- c("age", "bmi", "amhngml", "AFC", "orpi", "fibroidpresence", "ovarianResponse", "osi", 
                                  "endometrialmorphology", "endometrialthickness")

# Create linear regression models for each variable and the combined model
linear_clinical_pregnancy_model <- lapply(variables_clinical_pregnancy, function(var) {
  linear_regression_gtsummary(as.formula(paste("clinicalpregrate ~", var)), filtered_data_SR)
})

linear_clinical_pregnancy_combined <- linear_regression_gtsummary(clinicalpregrate ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness, filtered_data_SD)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_clinical_pregnancy <- lapply(linear_clinical_pregnancy_model, function(summary) {
  tbl_merge(
    tbls = list(summary, linear_clinical_pregnancy_combined),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_clinical_pregnancy <- tbl_stack(combined_summaries_clinical_pregnancy)

# Display the final summary
print(final_combined_summaries_clinical_pregnancy)



## Cycle Cancellation
# Define the variables of interest for logistic regression models
variables_cycle_cancellation <- c("age", "bmi", "amhngml", "AFC", "orpi", "fibroidpresence", "ovarianResponse", "osi", 
                                  "endometrialmorphology", "endometrialthickness")

# Create a custom function to perform logistic regression and summarize it using gtsummary
logistic_regression_gtsummary <- function(formula, data) {
  model <- glm(formula, data = data, family = binomial())
  tbl_regression(model)
}

# Create logistic regression models for each variable and the combined model
logistic_cycle_cancellation_model <- lapply(variables_cycle_cancellation, function(var) {
  logistic_regression_gtsummary(as.formula(paste("cycleCancel ~", var)), dams_data)
})

logistic_cycle_cancellation_combined <- logistic_regression_gtsummary(cycleCancel ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness, filtered_data_SD)

# Combine each unadjusted summary with the adjusted summary
combined_summaries_cycle_cancellation <- lapply(logistic_cycle_cancellation_model, function(summary) {
  tbl_merge(
    tbls = list(summary, logistic_cycle_cancellation_combined),
    tab_spanner = c("**Unadjusted**", "**Adjusted**")
  )
})

# Stack all combined summaries into a single final summary table
final_combined_summaries_cycle_cancellation <- tbl_stack(combined_summaries_cycle_cancellation)

# Display the final summary
print(final_combined_summaries_cycle_cancellation)















# 
# library(gtsummary)
# library(dplyr)
# 
# # Define a function to remove rows with blank unadjusted values
# remove_blank_unadjusted <- function(tbl) {
#   if (is.null(tbl)) return(NULL)
#   tbl %>%
#     filter(!is.na(log_OR) & log_OR != "")
# }
# 
# # Create logistic regression models for each variable and the combined model
# logistic_cycle_cancellation_model <- lapply(variables_cycle_cancellation, function(var) {
#   logistic_regression_gtsummary(as.formula(paste("cycleCancel ~", var)), filtered_data_SD)
# })
# 
# logistic_cycle_cancellation_combined <- logistic_regression_gtsummary(
#   cycleCancel ~ age + bmi + amhngml + AFC + orpi + fibroidpresence + ovarianResponse + osi + endometrialmorphology + endometrialthickness,
#   filtered_data_SD
# )
# 
# # Extract tables from the summaries
# tables_unadjusted <- lapply(logistic_cycle_cancellation_model, function(summary) {
#   if (!is.null(summary)) remove_blank_unadjusted(summary$tbl) else NULL
# })
# 
# table_combined <- if (!is.null(logistic_cycle_cancellation_combined)) logistic_cycle_cancellation_combined$tbl else NULL
# 
# # Check if tables are correctly extracted and not NULL
# if (any(sapply(tables_unadjusted, is.null))) {
#   message("One or more unadjusted tables are NULL or have blank rows. Check the logistic regression models.")
# }
# 
# if (is.null(table_combined)) {
#   stop("The combined adjusted table is NULL. Check the logistic regression combined model.")
# }
# 
# # Combine each unadjusted summary with the adjusted summary
# combined_summaries_cycle_cancellation <- mapply(function(tbl_unadj) {
#   if (!is.null(tbl_unadj) && !is.null(table_combined)) {
#     tbl_merge(
#       tbls = list(tbl_unadj, table_combined),
#       tab_spanner = c("**Unadjusted**", "**Adjusted**")
#     )
#   } else {
#     NULL
#   }
# }, tables_unadjusted, SIMPLIFY = FALSE)
# 
# # Stack all combined summaries into a single final summary table
# final_combined_summaries_cycle_cancellation <- tbl_stack(combined_summaries_cycle_cancellation)
# 
# # Display the final summary
# print(final_combined_summaries_cycle_cancellation)

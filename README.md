# Paracetamol-Administration-Analysis
An R Script to look at the times of administration, calculate the patient's total dose within a 24hr period and identify any doses that were subsequently administered within a 6hr window of the previous dose.


# Codex Prompt

You are an expert R (Posit/RStudio) data analyst. Write a complete, production-quality R script to analyse Paracetamol administration records exported from an Excel spreadsheet.

GOALS
1) Import an Excel file containing Paracetamol administrations from electronic health records.
2) Parse administration date/time robustly.
3) For each patient, calculate the TOTAL Paracetamol dose administered within a rolling 24-hour window (at each administration time), and identify administrations that would make the rolling 24h total exceed a safe threshold (default 4000 mg, but make this user-configurable).
4) Identify any Paracetamol doses administered within a 6-hour window of the previous Paracetamol administration for the same patient (flag as “<6h since previous”).
5) Produce clean summary outputs and export them to Excel.

ASSUMPTIONS / FLEXIBILITY
- The Excel may have varying column names. Implement a “column mapping” section near the top where the user can specify which columns correspond to:
  * patient_id (MRN or equivalent)
  * administration_datetime
  * dose_mg (numeric dose in mg; if dose given as grams or includes text, parse safely)
  * medication_name (optional; if present, filter to Paracetamol/Acetaminophen variants)
  * route (optional)
  * ward/location (optional)
  * prescriber/order_id (optional)
- If medication_name exists, filter to paracetamol synonyms: “paracetamol”, “acetaminophen”, “APAP”, “calpol”, etc. Case-insensitive.
- If medication_name does not exist, assume the file already contains only Paracetamol administrations.
- Handle timezones safely (use Europe/Dublin as default) and keep datetimes in POSIXct.
- Handle missing/invalid datetimes and doses: report counts of dropped rows and why.

REQUIRED PACKAGES
Use tidyverse, readxl, lubridate, janitor, stringr, slider, writexl (or openxlsx). Include package install checks (but do not auto-install in hospital environments; instead, message the user clearly).

OUTPUTS (WRITE ALL OF THESE)
A) “clean_admins” dataset:
- One row per administration after cleaning, sorted by patient_id and administration_datetime.
- Include:
  * patient_id
  * administration_datetime
  * dose_mg
  * time_since_prev_hours (difference from previous admin for same patient)
  * within_6h_flag (TRUE/FALSE where time_since_prev_hours < 6)
  * rolling_24h_total_mg (sum of dose_mg in the previous 24 hours inclusive of current admin, by patient)
  * exceeds_24h_threshold_flag (TRUE/FALSE where rolling_24h_total_mg > threshold_mg)
  * rolling_24h_window_start (administration_datetime - hours(24))

B) “patient_day_summary” dataset (calendar-day based, optional but implement):
- For each patient and calendar date, compute:
  * total_dose_mg_calendar_day
  * max_rolling_24h_total_mg_that_day
  * n_admins
  * n_within_6h
  * any_exceeds_threshold

C) “exceptions” dataset:
- Only rows where within_6h_flag == TRUE OR exceeds_24h_threshold_flag == TRUE
- Add a human-readable “exception_reason” column.

D) Export an Excel workbook with separate sheets:
- clean_admins
- patient_day_summary
- exceptions
- data_quality (row counts, missingness, dropped rows summary, date range)

VISUALS (BASIC, OPTIONAL BUT NICE)
- A histogram or bar chart of time_since_prev_hours (excluding NA)
- A plot of rolling_24h_total_mg over time for the top 5 patients by max rolling total (faceted)

IMPLEMENTATION NOTES (IMPORTANT)
- For time since previous: group_by(patient_id) arrange(administration_datetime) then compute difftime.
- For rolling 24h total: use slider::slide_index_dbl with administration_datetime as index, summing dose_mg over the .before window of 24 hours (use lubridate::hours(24)).
- Be careful: administrations may occur seconds apart; include current administration in rolling sum.
- Ensure dose_mg is numeric; if the source is like “1 g” or “1000mg”, parse to mg.
- Use functions and clear sections: config, import, clean, derive flags, summaries, export, plots.
- Add clear console messages (glue or message()) describing progress.

USER INPUTS AT TOP OF SCRIPT
- file_path (string): path to the Excel file
- sheet (string or index): which sheet to read
- threshold_mg (default 4000)
- min_interval_hours (default 6)
- tz (default "Europe/Dublin")
- column mapping list

DELIVERABLE
Return ONLY the final R script in one code block, no extra commentary.
